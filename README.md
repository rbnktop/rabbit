📂 Fuzzy Rabbit - Global Responsive Edition

Fuzzy Rabbit is a high-performance desktop utility designed to reconcile messy, inconsistent text lists with structured Excel data. Using advanced fuzzy string matching algorithms, it identifies the most likely matches, allows for human review, and extracts associated temporal data (dates) from spreadsheet headers.
✨ Key Features

    Smart Fuzzy Matching: Uses the difflib SequenceMatcher to find similarities even with typos or different formatting.

    Excel Integration: Directly reads .xlsx files and treats column headers as date references.

    Interactive Review: A clean, centralized review window where candidates are color-coded by confidence level.

    Non-Blocking UI: Entirely multithreaded; the application remains responsive even when processing thousands of rows.

    Date Extraction: Automatically identifies the "latest" date associated with a matched item across the spreadsheet.

    Clipboard Ready: One-click copying of final results for easy pasting into other workflows.

🐰 To the final User:

Welcome to Fuzzy Rabbit! This tool is designed to save you hours of manual "VLOOKUP" or "Ctrl+F" work. Here is how to use it:
1. Load Your Data

    Spreadsheet: Click "Select file.xlsx" to load your master data source.

    List: You can either click "Select file.txt" to load a list of items or simply paste your list directly into the large text box in the middle of the app.

2. Check Similarities

Click the "Check Similarities" button. The rabbit will begin searching through your Excel file to find the best possible matches.
3. Review Your Matches

A new window will open. To ensure accuracy, the app asks you to confirm its "guesses." We use a color-coding system to help you work faster:

    Green (90%+): Very high confidence. These are usually perfect matches.

    Gold (75%-89%): High confidence, but worth a quick glance.

    Orange (50%-74%): Moderate confidence. Check these carefully.

    Red (<50%): Low confidence. The app might not have found a good match.

Note: If the app is very confident (69% or higher), it will pre-select the best match for you. If it's not sure, it will default to "Skip".
4. Get Your Results

Once you click "Confirm Selections", the app will look through your spreadsheet headers to find the most recent date associated with those items.

    You will see a final table with your matches and dates.

    Click "Copy Dates" to put the entire column of dates onto your clipboard, ready to paste into Excel or any other program.

📝 License

Distributed under the MIT License. See LICENSE for more information.

Developed by: Cabelo LTDA
