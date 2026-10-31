# Testing — 09 Eval and pilot

Manual test cases. Run each case, then tick the Pass column. Use test
data and a test database throughout.

| Setup | Action | Expected | Pass |
|---|---|---|---|
| Empty eval directory | Run 09-create-golden-set.py, add two items, then a blank question | golden.json contains exactly the two items with question, expected answer and doc IDs | ☐ |
| Golden set with two items, answers file with one correct and one abstained answer | Run 09-run-evaluation.py | Report shows hit rate, citation accuracy and abstain rate matching the hand calculation, failures listed | ☐ |
| Admin workstation joined to the test domain | Run 09-create-pilot-group.sh with a test group and two test users | getent group shows the group, or the manual ADUC steps complete, members listed | ☐ |
| Test audit database with feedback rows from the last two weeks | Run 09-pilot-feedback-report.py with report weeks set to 2 | Report shows total feedback, helpful rate, per-week rows and recent comments | ☐ |
