# gunicorn reads this file AUTOMATICALLY when it starts in the project folder,
# even if the Start Command on Render is just "gunicorn app:app".
# This makes sure many users can scrape at the same time.
import os

bind = "0.0.0.0:" + os.environ.get("PORT", "10000")

# ONE process only: JOBS is kept in memory, so all requests must hit the
# same process. Concurrency comes from threads instead.
workers = 1
worker_class = "gthread"
threads = 32

# Live-progress streams stay open for the whole scrape; don't kill them.
timeout = 0
graceful_timeout = 30
keepalive = 5
