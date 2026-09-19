import io
import os
import time
import queue
import shutil
import zipfile
import threading
from pathlib import Path
from datetime import datetime

import streamlit as st

from scraper_core import (
    DB, OUTPUT_DIR, All4CarScraper, dedupe_keep_order
)

st.set_page_config(
    page_title="All4Car Batch VIN Scraper",
    page_icon="🚗",
    layout="wide",
)

# Prevent several browser sessions from launching several Chrome scrapers
# in the same small server at exactly the same time.
@st.cache_resource
def global_scraper_lock():
    return threading.Lock()

def init_state():
    defaults = {
        "events": queue.Queue(),
        "scraper": None,
        "worker": None,
        "batch_id": None,
        "status": "IDLE",
        "current_vin": "-",
        "current_category": "-",
        "vin_progress": "-",
        "sub_progress": "-",
        "logs": [],
        "output_dir": str(OUTPUT_DIR.resolve()),
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

def drain_events():
    q = st.session_state.events
    try:
        while True:
            e = q.get_nowait()
            msg = e.get("message")
            if msg:
                st.session_state.logs.append(
                    f"[{datetime.now():%H:%M:%S}] {msg}"
                )
                st.session_state.logs = st.session_state.logs[-1200:]

            kind = e.get("kind")
            if kind == "vin":
                st.session_state.current_vin = e.get("vin", "-")
                st.session_state.vin_progress = f"{e.get('position', 0)} / {e.get('total', 0)}"
            elif kind == "sub":
                st.session_state.current_category = e.get("category", "-")
                st.session_state.sub_progress = f"{e.get('position', 0)} / {e.get('total', 0)}"
            elif kind == "site":
                st.session_state.status = e.get("status", "-")
            elif kind == "done":
                st.session_state.status = "DONE"
            elif kind == "fatal":
                st.session_state.status = "ERROR"
    except queue.Empty:
        pass

def worker_target(scraper, batch_id):
    lock = global_scraper_lock()
    acquired = lock.acquire(blocking=False)
    if not acquired:
        scraper.emit(
            kind="fatal",
            message="الخادم يشغّل Batch أخرى حاليًا. أعد المحاولة بعد انتهائها."
        )
        return
    try:
        scraper.run_batch(batch_id)
    finally:
        try:
            scraper.close_driver()
        except Exception:
            pass
        lock.release()

def launch_batch(batch_id, output_dir):
    if st.session_state.worker and st.session_state.worker.is_alive():
        st.error("هناك Batch تعمل حاليًا في هذه الجلسة.")
        return

    st.session_state.events = queue.Queue()
    st.session_state.logs = []
    scraper = All4CarScraper(DB, output_dir, st.session_state.events)
    worker = threading.Thread(
        target=worker_target,
        args=(scraper, batch_id),
        daemon=True,
    )
    st.session_state.scraper = scraper
    st.session_state.worker = worker
    st.session_state.batch_id = batch_id
    st.session_state.status = f"WORKING — Batch #{batch_id}"
    worker.start()

def zip_output_folder(folder: Path):
    if not folder.exists():
        return None
    files = [p for p in folder.rglob("*") if p.is_file()]
    if not files:
        return None
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in files:
            z.write(p, p.relative_to(folder))
    buf.seek(0)
    return buf.getvalue()

st.title("🚗 All4Car Batch VIN Scraper")
st.caption("Web Edition — نفس محرك V9 الناجح، بواجهة Streamlit تعمل من المتصفح.")

left, right = st.columns(2)
with left:
    vins_text = st.text_area(
        "VIN List — كل VIN في سطر",
        height=240,
        placeholder="WAUZZZ4H7CN006593\n...",
    )
with right:
    categories_text = st.text_area(
        "Main Categories — كل مجموعة في سطر",
        height=240,
        placeholder="IC engine cooling\nFuel system",
    )

force = st.checkbox("Force Re-scrape completed VIN/categories", value=False)

st.subheader("مكان حفظ النتائج")
save_dir_text = st.text_input(
    "Output Folder",
    value=st.session_state.output_dir,
    help="في التشغيل المحلي على Windows يمكنك كتابة أو لصق مسار المجلد الذي تريد الحفظ داخله، مثل C:\\Users\\ASUS\\Desktop\\All4Car Results",
)
st.caption(
    "مهم: المتصفح لا يسمح لصفحة ويب باختيار مجلد من جهاز المستخدم ثم الكتابة إليه مباشرة. "
    "عند التشغيل المحلي، المسار أعلاه هو مسار على نفس الكمبيوتر ويُحفظ إليه مباشرة. "
    "عند الاستضافة على الإنترنت، المسار يكون على السيرفر وتقوم بتنزيل النتائج من زر Download."
)

c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    if st.button("▶ Start New Batch", type="primary", use_container_width=True):
        vins = dedupe_keep_order(vins_text.splitlines())
        cats = dedupe_keep_order(categories_text.splitlines())
        if not vins:
            st.error("أدخل VIN واحدًا على الأقل.")
        elif not cats:
            st.error("أدخل Main Category واحدة على الأقل.")
        else:
            # Desktop-style: save directly to the user-selected local folder.
            batch_root = Path(save_dir_text).expanduser()
            try:
                batch_root.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                st.error(f"تعذر إنشاء/فتح مجلد الحفظ: {e}")
                st.stop()
            batch_id = DB.create_batch(vins, cats, batch_root, force)
            st.session_state.output_dir = str(batch_root.resolve())
            launch_batch(batch_id, batch_root)
            st.rerun()

with c2:
    if st.button("⏸ Pause", use_container_width=True):
        if st.session_state.scraper:
            st.session_state.scraper.pause()

with c3:
    if st.button("⏯ Resume", use_container_width=True):
        if st.session_state.scraper:
            st.session_state.scraper.resume()

with c4:
    if st.button("⏹ Stop Safely", use_container_width=True):
        if st.session_state.scraper:
            st.session_state.scraper.stop()

with c5:
    if st.button("↩ Resume Previous", use_container_width=True):
        row = DB.latest_unfinished_batch()
        if not row:
            st.warning("لا توجد Batch غير مكتملة.")
        else:
            out = Path(row["output_dir"])
            out.mkdir(parents=True, exist_ok=True)
            st.session_state.output_dir = str(out)
            launch_batch(row["id"], out)
            st.rerun()

@st.fragment(run_every=1.0)
def live_panel():
    drain_events()

    a, b, c, d = st.columns(4)
    a.metric("Status", st.session_state.status)
    b.metric("VIN Progress", st.session_state.vin_progress)
    c.metric("Current VIN", st.session_state.current_vin)
    d.metric("Sub Progress", st.session_state.sub_progress)

    st.write("**Current Category:**", st.session_state.current_category)

    worker = st.session_state.worker
    if worker and not worker.is_alive() and str(st.session_state.status).startswith("WORKING"):
        st.session_state.status = "FINISHED"

    st.text_area(
        "Execution Log",
        value="\n".join(st.session_state.logs[-400:]),
        height=380,
        disabled=True,
    )

    out = Path(st.session_state.output_dir)
    archive = zip_output_folder(out)
    if archive:
        st.download_button(
            "⬇ Download Results ZIP",
            data=archive,
            file_name=f"{out.name}.zip",
            mime="application/zip",
            use_container_width=True,
        )

    if os.name == "nt" and out.exists():
        if st.button("📂 Open Output Folder", use_container_width=True):
            os.startfile(str(out))

live_panel()

with st.expander("ملاحظات التشغيل"):
    st.markdown(
        """
- الواجهة لا تحتاج Tkinter ولا Jupyter.
- في التشغيل المحلي، Chrome يفتح بشكل مرئي على نفس الكمبيوتر حتى يتصرف الموقع كما في نسخة سطح المكتب.\n- عند الاستضافة لاحقًا، سنستخدم Chrome على الخادم ونضبط وضع التشغيل بما يناسب الاستضافة.
- أي جهاز يملك رابط الصفحة يستطيع استخدام الواجهة من المتصفح.
- لأسباب الموارد، هذه النسخة تسمح بعملية Scraping ثقيلة واحدة في الوقت نفسه على نفس الخادم.
- ملفات النتائج تبقى على الخادم، ويمكن تنزيلها كملف ZIP من الواجهة.
        """
    )
