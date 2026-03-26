# AI Prompt Guide (v2 Scanner)

If you are sharing this project folder with someone else, they can copy and paste the prompt below directly into their AI IDE (like Cursor, Windsurf, or Antigravity) to automatically set up the environment and start the enriched scan. The `twitter_scanner_v2.py` script inside this folder *already* contains the internet pause/resume logic we added today, so their AI just needs to set up Python and launch it.

***

### Copy & Paste this into the AI IDE:

> "Hello! Please analyze the files in this folder. I need you to run the `twitter_scanner_v2.py` script on the Excel file located at `SHEETS/verified results cleaned.xlsx`. 
> 
> Please perform the following steps:
> 1. **Set up the Environment:** Provide or configure a proper Python environment, and install the dependencies listed in `requirements.txt` (`pip install -r requirements.txt`).
> 2. **Install Playwright Browsers:** Run `playwright install chromium`. 
> 3. **Launch the Scanner:** Run `twitter_scanner_v2.py` in `--headless` mode on the active sheet. Please run this process in the background, redirecting the output to a log file (e.g., `scan_v2.log`) so I can monitor the progress without keeping my terminal open. 
> 
> Note: The script already contains a robust internet disconnect fail-safe (`wait_for_internet()`), so you only need to run it. Let me know once the background process has started successfully!"
