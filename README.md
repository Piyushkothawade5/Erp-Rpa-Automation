# ERP Data Entry Automation (Python RPA)

> **Status:** Live in production at Shubhada Polymers Pvt. Ltd., Nashik  
> **Result:** 50–60% reduction in data entry time · Operator freed for parallel tasks · ~150% effective productivity gain

A production-grade Python RPA (Robotic Process Automation) script that automates bulk data entry into a legacy ERP system — handling invalid items, popup interruptions, batch saves, crash recovery, and parallel processing without any human intervention.

---

## The Problem

Operators at a manufacturing company had to manually enter hundreds of item records into a legacy ERP system every day:

- Open a new transaction form
- Type item code → press Enter × 4 → enter location → enter quantity
- Handle random popups that interrupted the flow
- Save and approve every batch of records
- Repeat — for hours

The ERP had no bulk import. No API. No automation support.  
Just a screen, a keyboard, and an operator doing repetitive clicks.

---

## What the Script Does

```
items.txt (item_code, location, qty)
        ↓
  Resume from last saved position
        ↓
  Background popup watcher (thread)
        ↓
  For each item:
    → Paste item code via clipboard
    → Detect invalid item popup (image recognition)
    → If invalid → log as INVALID → skip
    → If valid → enter location + qty
    → Every 20 records → Save + Approve + New transaction
    → ESC pressed → save progress → exit safely
        ↓
  item_validation_log.xlsx (full audit log)
```

---

## Key Engineering Features

### 1. Crash Recovery — Resume from Any Point
Every successfully processed record writes its index to `progress_items.txt`.  
If the script crashes or is stopped, it resumes from exactly where it left off — no re-processing, no duplicates.

```python
def save_progress(index):
    with open(PROGRESS_FILE, "w") as f:
        f.write(str(index))

def load_progress():
    try:
        with open(PROGRESS_FILE, "r") as f:
            return int(f.read().strip())
    except:
        return 0
```

### 2. Background Popup Watcher (Threading)
Legacy ERP systems throw random confirmation/notification popups at unpredictable intervals.  
A dedicated background thread runs continuously, detecting and dismissing these popups using image recognition — so the main thread never gets stuck.

```python
def popup_watcher():
    while not popup_stop_event.is_set():
        cleared = clear_general_popup_once()
        time.sleep(0.15 if cleared else POPUP_WATCH_INTERVAL)

popup_thread = threading.Thread(target=popup_watcher, daemon=True)
popup_thread.start()
```

### 3. Invalid Item Detection (Image Recognition)
Not all item codes in the input file are valid in the ERP. When an invalid code is entered, the ERP shows a specific error popup.  
The script detects this popup using `pyautogui.locateOnScreen` with confidence matching, dismisses it, logs the item as INVALID, and continues — without interrupting the batch.

```python
def invalid_popup_present():
    location = pyautogui.locateOnScreen('invalid_item.png', confidence=0.8, grayscale=True)
    return bool(location)
```

### 4. Safe Action Wrapper
Every screen interaction is wrapped in `safe_action()` — which clears popups before the action, executes it, then checks for new popups after. This makes every click popup-safe without repeating the logic everywhere.

```python
def safe_action(action, after_delay=0.0, post_popup_checks=2):
    clear_general_popups()
    action()
    if after_delay:
        time.sleep(after_delay)
    for _ in range(post_popup_checks):
        clear_general_popups(max_rounds=1)
```

### 5. Batch Save and Approve
The ERP requires records to be saved and approved in batches. Every `BATCH_SIZE` records (default: 20), the script automatically:
- Clicks Save → confirms the save dialog
- Clicks Approve → selects Entered By → confirms
- Clicks Approve again → selects Approved By → confirms
- Opens a new transaction form with the target date

All without human intervention.

### 6. Excel Audit Log
Every processed item is logged to `item_validation_log.xlsx` with:

| Input_SrNo | Valid_SrNo | ItemCode | Location | Qty | Status | Time |
|---|---|---|---|---|---|---|
| 1 | 1 | ITEM001 | WH-A | 10 | VALID | 2026-04-01 09:15:32 |
| 2 | — | ITEM999 | WH-B | 5 | INVALID | 2026-04-01 09:15:38 |

The log persists across runs and is used to recover the valid serial number counter on resume.

### 7. Emergency Stop (ESC Key)
Pressing ESC at any point:
- Saves current progress index
- Saves the Excel log
- Signals the popup watcher thread to stop
- Exits cleanly

---

## Tech Stack

| Component | Technology |
|---|---|
| Screen automation | pyautogui |
| Image recognition | pyautogui.locateOnScreen |
| Clipboard handling | pyperclip |
| Parallel popup handling | threading |
| Keyboard monitoring | keyboard |
| Audit logging | openpyxl |
| Input data | Plain text file (CSV format) |

---

## Input Format

```
# items.txt
ITEM001,WH-A,10
ITEM002,WH-B,25
ITEM003,WH-A,5
```

---

## Configuration

All settings in one place at the top of the script:

```python
TARGET_DATE = "01/04/2026"   # Transaction date
BATCH_SIZE = 20              # Records per save-approve cycle
START_DELAY = 5              # Seconds before script starts (time to focus ERP window)
```

---

## What's in This Repo

```
/
├── README.md
├── 1.Run.py                  ← Main automation script
├── items.txt                 ← Sample input file (dummy data)

```

> **Note:** `thanks_btn.png` and `invalid_item.png` are image templates used for popup detection. These are captured from the ERP screen and are not included as they contain proprietary UI elements.

---

## Business Impact

| Metric | Before | After |
|---|---|---|
| Data entry time | ~4 hrs/day | ~1.5 hrs/day |
| Time reduction | — | 50–60% |
| Invalid items | Caught manually | Auto-detected and logged |
| Audit trail | None | Full Excel log with timestamps |
| Crash recovery | Start over | Resume from last position |
| Operator availability | Tied to screen | Free for parallel tasks |

**Effective productivity gain: ~150%** — time saved + operator available for other work simultaneously.

---

## Why This Is RPA, Not Just a Script

Most automation scripts assume a perfect environment — no popups, no invalid data, no crashes.  
This script was built for a real, messy, legacy ERP:

- Popups appear randomly and must be dismissed without breaking the flow
- Input data contains invalid items that must be detected and skipped — not just errored out
- The process runs for hours — crash recovery was non-negotiable
- The ERP has no API — screen coordinates and image recognition are the only interface

This is the same problem that enterprise RPA tools (UiPath, Automation Anywhere, Blue Prism) solve — built from scratch in Python.

---

## About the Developer

**Piyush Ramesh Kothawade**
Data Analyst / AI Automation Analyst
Shubhada Polymers Products Pvt. Ltd., Nashik [LinkedIn](https://www.linkedin.com/in/piyush-kothawade/) · [Portfolio](https://codebasics.io/portfolio/Piyush-Kothawade)

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?logo=linkedin)](https://www.linkedin.com/in/piyush-kothawade/)

*Other projects: OEE Production Web App (React + Supabase) · XML-to-PDF Batch Automation · Excel TI Management System*
