# Catcher Reports

Automated catcher reports built from TrackMan data appended with independently tracked information. The script can create single game or cumulative reports for any catcher who appears in the file in a matter of minutes. Provides information about framing, blocking, and throwing using run value and frame rate above average.

![Sample catcher report](images/sample-report.png)

## What it measures

- **Framing:** Actual called-strike rate vs. expected from a location-based model trained on 100,000 called pitches from 2026. Shadow-zone pitches are labeled with their expected-strike probability.
- **Blocking:** Block success rate by side and pitch group, plus run-value credit for each successful block.
- **Throwing:** Recorded pop times to 2B and 3B, throw velocity, and run-value credit for caught stealings.

Run values calculated from an MLB play by play dataset from the 2024 season.

## How to run

```
python catcher_report.py <TrackMan game file> [--manual tracking.xlsx]
```

1. During the game(s) keep a manual log of block attempts, steal attempts, and passed balls along with the pitch number and base state in which they occur in a separate csv file.
2. Open a new Colab notebook, paste in the script, and run the cell.
3. The script will automatically append the manual log file and match on pitch number.
4. When finished the PDF downloads


**Requires:** reportlab, openpyxl, Pillow

*No team data is included in this repository. Mid-major division one averages gathered from 2026 TrackMan data from the Big West, Mountain West, and West Coast Conferences. Since base states are not available in TrackMan files, run value had to be calculated from MLB data.*
