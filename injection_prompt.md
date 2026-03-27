# AI Injection & Execution Prompt

If you are starting on a fresh computer where the code does **not** have the internet pause/resume logic yet, you can use this prompt. It tells the AI agent exactly how to upgrade the code and then launch the scan.

***

### Copy & Paste this into the AI IDE:

> "I need you to upgrade and run the `twitter_scanner_v2.py` script in this folder. Currently, it lacks a fail-safe for internet disconnections. 
> 
> Please perform these steps in order:
> 
> 1. **Inject Internet Resiliency:** 
>    - Add `import socket` to the top of `twitter_scanner_v2.py`.
>    - Create a helper function `wait_for_internet()` that attempts to connect to `1.1.1.1` on port `53` with a 3-second timeout. If it fails, it should print '[!] Internet disconnected! Pausing...' and loop every 5 seconds until the connection returns.
>    - Call this `wait_for_internet()` function inside the `run_scan` loop **immediately before** every call to `scrape_profile`.
>    - Additionally, if `scrape_profile` returns a 'Timeout' or 'Error' status, call `wait_for_internet()` again to ensure we don't waste retries while offline.
> 
> 2. **Environment Setup:** 
>    - Install the dependencies from `requirements.txt`.
>    - Run `playwright install chromium`.
> 
> 3. **Launch Background Scan:** 
>    - Run the upgraded script in `--headless` mode.
>    - Target the file: `SHEETS/verified results cleaned.xlsx`.
>    - Redirect all terminal output to `scan_v2.log` so I can monitor progress later.
> 
> Please confirm once the code has been updated and the background scan has started."
