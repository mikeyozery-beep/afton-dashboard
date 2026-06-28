# Script Building Guidelines & Lessons Learned

## Overview
This document consolidates lessons learned from failed scripts in this session. The goal is to minimize script failures through systematic improvements and edge-case handling.

---

## Critical Script Failures & Fixes

### 1. **PowerShell Parameter Errors**
**Failure:** `-RunWithHighestPrivilege` parameter doesn't exist
```powershell
# WRONG
$settings = New-ScheduledTaskSettingsSet -RunWithHighestPrivilege

# RIGHT
$task = Register-ScheduledTask -RunLevel Highest
```
**Lesson:** Use `-RunLevel Highest` at task registration, not in settings.

---

### 2. **Logging Function Calls**
**Failure:** `logger.info()` with no arguments crashes Python
```python
# WRONG
logger.info()

# RIGHT
logger.info("")
```
**Lesson:** All logging calls require at least one argument (message string).

---

### 3. **Return Type Mismatches**
**Failure:** Function returns empty list `[]` but caller expects dict `.values()`
```python
# WRONG
if not items:
    return [], []  # Second item should be dict

# RIGHT
if not items:
    return [], {}  # Returns dict for second value
```
**Lesson:** Document return types clearly and validate at call site.

---

### 4. **Unicode/Encoding in Logging**
**Failure:** Special characters in file paths crash logging
```python
# Issue: logger.info(f"Extracted: {path_with_unicode_chars}")

# Solution: Encode carefully or use plain strings
logger.info(f"Extracted: {filepath.name}")
```
**Lesson:** Be explicit about encoding in logging statements.

---

### 5. **Zip File Extraction**
**Failure:** PowerShell `System.IO.Compression.ZipFile` not available
```powershell
# WRONG - PowerShell zip extraction
[System.IO.Compression.ZipFile]::OpenRead($path)

# RIGHT - Use Python for zip extraction
```
**Lesson:** Use Python's `zipfile` module for zip operations; don't rely on PowerShell libraries.

---

### 6. **Multi-Step Script Problems**
**Failure:** Second script in batch file doesn't run or fails silently
**Root Causes:**
- Using `call` command but second script has errors
- Second script doesn't handle edge cases (no files)
- No error propagation between scripts

**Solution:** Consolidate into single script that:
- Handles ALL steps atomically
- Has comprehensive error handling
- Provides detailed logging
- Doesn't require manual sequencing

---

## Best Practices for Future Scripts

### 1. **Single-Step Consolidation**
**Rule:** If possible, put everything in ONE script file
- Download → Extract → Organize → Report
- No need for user to run multiple commands
- Easier error tracking and recovery

### 2. **Edge Case Handling**
**Always handle:**
- Empty input (no files found)
- Missing directories
- File access/permission errors
- Encoding errors
- Type mismatches on return values

### 3. **Logging Standards**
- Use proper logging format with timestamps
- Include context (what step, which file, etc.)
- Always include try/except blocks with error details
- Log to BOTH console AND file

### 4. **Return Type Consistency**
```python
# Bad: Sometimes returns list, sometimes dict
def process_files():
    if not files:
        return []
    return {}

# Good: Always returns same types
def process_files():
    if not files:
        return [], {}  # Always returns tuple
    return files_list, files_dict
```

### 5. **File Operations**
- Always validate paths exist before processing
- Use pathlib.Path() for cross-platform compatibility
- Handle special characters and Unicode properly
- Log file operations for debugging

### 6. **Python Zip Handling**
```python
import zipfile

# Correct pattern
with zipfile.ZipFile(zip_path, 'r') as zip_ref:
    for file_info in zip_ref.filelist:
        if file_info.filename.lower().endswith(('.xlsx', '.xls')):
            with zip_ref.open(file_info) as source, open(dest, 'wb') as target:
                target.write(source.read())
```

### 7. **Batch File Chaining**
**If you must use multiple scripts:**
```batch
@echo off
REM Use && to chain; second runs only if first succeeds
call script1.bat && call script2.bat

REM Better: consolidate into one script
python consolidated_script.py
```

### 8. **Excel Operations**
- Always read A1 cell for report names
- Handle multiple worksheets explicitly
- Validate cell values exist before use
- Use `data_only=True` for calculated values

---

## Testing Checklist Before Deployment

- [ ] Run with NO files (edge case)
- [ ] Run with mix of file types (.xlsx, .xls, .zip)
- [ ] Run with special characters in filenames
- [ ] Check logging for errors
- [ ] Verify all files actually process (count them)
- [ ] Check return types match expectations
- [ ] Test error scenarios (missing files, access denied)

---

## Script Template (Consolidated Pattern)

```python
#!/usr/bin/env python3
"""
Consolidated automation script
Single file handles: download → extract → organize
No manual step sequencing required
"""

import logging
from pathlib import Path
from datetime import datetime

# Setup - do once at top
SCRIPT_DIR = Path(__file__).parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"process_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Main steps as functions with error handling
def step_1_download():
    try:
        logger.info("STEP 1: Download")
        # ... implementation
        return True
    except Exception as e:
        logger.error(f"Step 1 failed: {e}", exc_info=True)
        return False

def step_2_extract():
    try:
        logger.info("STEP 2: Extract")
        # ... implementation
        return True
    except Exception as e:
        logger.error(f"Step 2 failed: {e}", exc_info=True)
        return False

def step_3_organize():
    try:
        logger.info("STEP 3: Organize")
        # ... implementation
        return True
    except Exception as e:
        logger.error(f"Step 3 failed: {e}", exc_info=True)
        return False

def main():
    logger.info("=" * 70)
    logger.info("AUTOMATED PROCESS START")
    logger.info("=" * 70)
    
    # Execute steps in sequence
    if not step_1_download():
        logger.error("Process stopped at step 1")
        return False
    
    if not step_2_extract():
        logger.error("Process stopped at step 2")
        return False
    
    if not step_3_organize():
        logger.error("Process stopped at step 3")
        return False
    
    logger.info("=" * 70)
    logger.info("PROCESS COMPLETE - SUCCESS")
    logger.info("=" * 70)
    return True

if __name__ == '__main__':
    success = main()
    exit(0 if success else 1)
```

---

## Key Takeaway

**Never require users to run multiple scripts in sequence.** Consolidate everything into a single atomic operation. This reduces failure points, improves error handling, and makes automation reliable.

