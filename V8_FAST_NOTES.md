# V8 Fast Mode

Built from the working V7 Xvfb/headful Render version.

Optimizations:
- ALL4CAR_FAST_MODE=1 by default.
- Removed Google connectivity request before every Selenium navigation.
- Replaced fixed 2-second VIN wait with an explicit category-link wait.
- Faster WebDriver polling.
- Reduced safe click settling delay.
- Preserved V4 subgroup expansion logic, but shortened its polling/settling windows.
- Reduced subgroup table wait from 12s max to 5s in Fast Mode.
- Reduced post-render fixed sleeps.
- Shorter retry backoff and page-load limits in Fast Mode.
- Added page-load and subgroup elapsed-time log messages.

Set ALL4CAR_FAST_MODE=0 to restore conservative timings.
