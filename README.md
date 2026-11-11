# FMP IT Ticket System — Synology Docker

## Run with Synology Container Manager
1. Extract this ZIP into a new project folder on the NAS.
2. Open `docker-compose.yml` and make sure port `8000` is available. If it is already used, change the left side of `8000:8000` (for example `8001:8000`).
3. In Container Manager → Project, create a project using this folder and build/run it.
4. Open `http://NAS-IP:8000/` (or the changed external port).
5. Health check: `http://NAS-IP:8000/health`.

## Default credentials for a brand-new database
- Username: `admin`
- Password: `Ab1234`

The password in `.env` is only used when creating the initial database/admin user. Changing it later does not change an existing admin account.

## Data and upgrades
The app stores its SQLite database at `./data/tickets.db` on the NAS, mounted as `/data/tickets.db` in the container. Back up this file before replacing the project. Do not delete or overwrite `data/tickets.db` if it contains tickets you need.

## Troubleshooting
- If the container cannot start, open Container Manager → Project → Log and inspect the first Python traceback.
- If port 8000 is already in use, change the host-side port, e.g. `8001:8000`, then browse to `http://NAS-IP:8001/`.
- This compose file does not mount the project directory over `/app`; only `./data` is mounted. This avoids masking `/app/main.py` inside the image.
- `.env` is included for convenient initial testing. Change `SECRET_KEY` and `ADMIN_PASS` before production use; keep `.env` private.
