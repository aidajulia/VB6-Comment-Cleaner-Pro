---
name: 🐛 Bug Report
about: Report crashes, errors, or incorrect comment cleaning
title: "[BUG] "
labels: bug
assignees: ''

---

## Description  
Briefly describe the unexpected behavior.

**Example**:  
"Mode 2 deletes text comments containing code-like keywords (e.g., 'Note: Dim x = 5')"

---

## Steps to Reproduce  
1. Open VB6 Comment Cleaner Pro  
2. Select Mode: [1/2]  
3. Process file: [filename.bas/cls]  
4. Error occurs at: [specific step]  

**VB6 Example**:  
1. Create `test.cls` with:  
   ```vb
   ' Troubleshooting code (commented)
   ' Dim conn As New ADODB.Connection
   ```
2. Run Mode 2
3. Both lines get deleted (should preserve text comment)

## Expected vs Actual Behavior
**Expected**:
"Commented code removed, text comments preserved"

**Actual**:
"All comments deleted regardless of content"

## Attachments
Problematic VB6 file (.bas/.cls)

Generated .garbage.log

Error screenshot

Crash report (if applicable)

🔒 Remove sensitive data before attaching

## Environment
VB6 IDE Version: 6.0 SP6

Cleaner Version: v1.0.0

OS: Windows 10/11

Cleaning Mode: Mode 1/Mode 2

File Types Processed: .bas/.cls

## Additional Context
Add special scenarios:

Files with mixed encoding

Projects using third-party controls

Large files (>10,000 lines)
