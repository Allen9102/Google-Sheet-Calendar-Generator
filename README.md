# Google Sheet Calendar Generator

A Google Apps Script project that automatically generates a monthly calendar in Google Sheets based on a user-specified year and month.

The script handles date generation, time slots, cross-month weeks, cell merging, and spreadsheet formatting automatically.

## Project Structure

```text
Google-Sheets-Calendar-Generator/
├── .gitignore
├── Calendar generator
├── README.md
└── LICENSE
```

## Features

- Generate a monthly calendar by entering a year and month
- Automatically determine the weekday and dates
- Generate hourly time slots from 8:00 to 23:00
- Handle weeks spanning two months
- Merge unused cells for dates outside the selected month
- Automatically apply formatting, borders, alignment, and frozen headers

## How to Use

Paste the code in "Apps Script," and run the code using "Macros."

<img width="1417" height="288" alt="截圖 2026-09-10 15 48 53" src="https://github.com/user-attachments/assets/dd1c1bd8-0f0c-4dec-83cf-f850e81efaec" />

### 1. Run the Script

Open the Google Sheets spreadsheet and run `generateCalendar()` through Google Apps Script.

Enter the year you want to generate.


<img width="255" height="236" alt="截圖 2026-09-10 15 45 59" src="https://github.com/user-attachments/assets/7d1913cf-9b1b-4882-8d8f-23ae87b3361d" />


Then enter the month.


<img width="255" height="236" alt="截圖 2026-09-10 15 46 31" src="https://github.com/user-attachments/assets/2962afa3-4edf-455d-b3a9-f7f35dae2b3b" />


### 2. Generated Calendar

The script automatically creates a new worksheet containing the monthly calendar.

The weekdays start from Monday (the character for 星期一).

<img width="1417" height="562" alt="截圖 2026-09-10 15 47 17" src="https://github.com/user-attachments/assets/fe3eef9a-e41c-40e7-873c-923bc4270d80" />


## Main Functions

| Function | Description |
|---|---|
| `generateCalendar()` | Generates the monthly calendar and applies formatting |
| `mergeCellsAndFill()` | Merges unused cells and applies background formatting |

## Technologies

- Google Apps Script
- JavaScript
- Google Sheets Spreadsheet Service

