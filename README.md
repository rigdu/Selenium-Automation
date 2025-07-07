# Web Automation GUI

A Python Tkinter GUI tool to automate web browser actions (login, navigation, and data processing) using Selenium.

## Features

- GUI for easy start/stop and logging.
- Credentials read from `pass.txt`.
- Keeps track of processed rows in `processed_rows.csv`.
- Highlights elements being operated on.
- Robust logging of actions and errors.
- Designed for batch processing of web UI elements.

## Requirements

- Python 3.x
- Google Chrome browser
- ChromeDriver (place `chromedriver.exe` in the script directory)
- Python packages: `selenium`, `pyautogui`, `tk`

Install with:

```bash
pip install -r requirements.txt
```

## Usage

1. Place your username and password in a file called `pass.txt` (first line: username, second line: password).
2. Place `chromedriver.exe` in the script directory.
3. Run the script:

   ```bash
   python web_automation_gui.py
   ```

4. Enter the starting index and click **Start**.
5. Click **Stop** to halt automation (browser window remains open).

## Output

- Processed indices and timestamps are saved in `processed_rows.csv`.
- Logs are shown in the GUI.

---

**Note**: This script is tailored for a specific web interface at `http://test.exactllyweb.com/Home`.  
Update XPaths and logic if your UI differs.
