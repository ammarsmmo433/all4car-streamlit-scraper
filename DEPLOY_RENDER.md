# All4Car Streamlit V6 — Render Deployment

This version keeps the proven V5 scraping logic and prepares it for Linux/Docker hosting.

## Local Windows test
```bat
py -3.11 -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
streamlit run app.py
```

On Windows the browser remains visible by default.

## Render deployment
1. Create a new GitHub repository.
2. Upload all files from this folder to the repository root.
3. In Render choose **New > Web Service** and connect the GitHub repository.
4. Render should detect the Dockerfile. Set Language/Runtime to **Docker** if needed.
5. Deploy.
6. Environment variables used by V6:
   - `ALL4CAR_HEADLESS=1`
   - `ALL4CAR_DATA_DIR=/var/data/all4car_batch_scraper`
7. Health check path: `/_stcore/health`.

## Persistence
The first deployment can be tested without a persistent disk. However, Render's normal
filesystem is ephemeral. For reliable Resume/checkpoints and retained Excel outputs across
service restarts/redeploys, use a paid Render web service with a persistent disk mounted at:

`/var/data`

and keep:

`ALL4CAR_DATA_DIR=/var/data/all4car_batch_scraper`

## Downloads
On hosted mode, users cannot choose an arbitrary local Windows folder for server-side writes.
The scraper stores results on the server and the Streamlit page provides the ZIP download.
