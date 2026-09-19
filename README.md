# All4Car Streamlit Web Edition

## Local Windows setup

```bat
cd C:\Users\ASUS\Desktop\SsgAsia\all4car_streamlit_v1

py -3.11 -m venv .venv
.venv\Scripts\activate

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

streamlit run app.py
```

Then open:

`http://localhost:8501`

Chrome must be installed on Windows. Selenium Manager resolves the driver automatically.

### Important V2 behavior
- Local Windows mode opens Chrome visibly (`HEADLESS = False`) to match the successful desktop notebook behavior.
- Main-category discovery waits specifically for current All4Car `#/schemas?` links after VIN search.
- Results are saved directly into the Output Folder entered in the Streamlit page, exactly like the desktop workflow.
- The **Open Output Folder** button is available on local Windows.
- When hosted publicly, a browser cannot write directly into a visitor's arbitrary local folder. Hosted mode saves on the server and the visitor downloads the result ZIP.

## Python dependencies

Only these external Python packages are required:

- streamlit
- selenium
- requests
- openpyxl

Removed from the web edition:
- tkinter (desktop GUI; built into some Python installations, not needed)
- webdriver-manager (Selenium Manager / system chromedriver is used)
- jupyter and ipykernel (not needed to run the web app)
- pyinstaller (not needed because this is no longer an EXE)

## Docker test

```bat
docker build -t all4car-streamlit .
docker run --rm -p 8501:8501 all4car-streamlit
```

Open `http://localhost:8501`.

## Render deployment

Push this folder to a GitHub repository. In Render:

1. New > Web Service.
2. Connect the GitHub repository.
3. Runtime/Language: Docker.
4. Render detects the Dockerfile.
5. Deploy.
6. Open the generated `onrender.com` URL.

For persistent SQLite checkpoints/results across restarts, attach a persistent disk and point the app workspace to it in a later production configuration.

## V3 fix
Fixed the Main Category lookup return contract. The V2 finder returned one Selenium
WebElement while the proven V9 navigation code expects `(actual_name, element)`.
This caused `cannot unpack non-iterable WebElement object` immediately after the
category became visible. V3 returns the expected tuple in exact, partial, and fallback matches.

## V4 expansion fix
The previous web build still carried the obsolete V8 expansion selector
`button._710xV-kIMAg-`, so it reported zero expansion clicks and collected only the
24 cards initially rendered by All4Car.

V4 expands against the current stable sub-group marker
`a[data-test-id="parts-link"]`. It repeatedly scrolls to the last card/page bottom,
clicks a visible Show more/Load more control when present, and waits for the actual
card count to increase. Collection starts only after three stable rounds.

## V5
- Retry WinError 32 / PermissionError up to 10 times before failing.
- Close all openpyxl handles before atomic replacement.
- Clean stale `.tmp.xlsx`.
- Force Re-scrape deletes the old VIN/category workbook first, so it is truly fresh.
- If Excel remains locked, close the workbook and use Resume Previous Batch.
