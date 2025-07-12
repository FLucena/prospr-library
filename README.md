# PROSPR Financial Planning System

A Google Apps Script that automatically analyzes budget data and generates a compact report.

## What It Does

- Reads budget data from the 'Monthly Budget' tab
- Identifies spending patterns and deviations
- Generates a professional report
- Creates email drafts ready to send

## Quick Setup

1. Open your Google Spreadsheet
2. Go to `Extensions > Apps Script`
3. Copy all `.gs` files into your Apps Script project
4. Save and refresh your spreadsheet
5. You'll see a new "PROSPR Tools" menu

## Data Format

Your spreadsheet must have:
- Sheet named "Monthly Budget"
- Data starting at row 4
- Columns with "Budget" and "Actual" headers
- Category totals starting with "Total"

Example:
```
Row 4: [Headers]     | Budget | Actual
Row 5: [Food]        |        |
Row 6: [Dining out]  | 300    | 450
Row 7: [Groceries]   | 500    | 480
Row 8: [Total Food]  | 800    | 930
```

## Usage

1. Set up your budget data in the "Monthly Budget" sheet
2. Use `PROSPR Tools > Generate Report`
3. Use `PROSPR Tools > Send Email`

## Security

**Important**: Change the default password in `env.gs` before using:
1. Open `env.gs`
2. Change `ADMIN_CODE: 'PROSPR2025'` to a secure password
3. Use at least 8 characters

## Configuration

Key settings in `Config.gs`:
- `DEVIATION_THRESHOLD`: 15% (minimum deviation to flag)
- `DATA_START_ROW`: 4 (where data begins)
- `MINIMUM_AMOUNT`: 10 (minimum amount to analyze)

## Troubleshooting

**"Required columns not found"**
- Check that you have "Budget" and "Actual" column headers

**"No budget data found"**
- Ensure sheet is named "Monthly Budget"
- Check that data starts at row 4
- Verify you have totals starting with "Total"

## File Structure

```
├── Code.gs              # Main menu setup
├── Config.gs            # Configuration
├── env.gs               # Passwords (change this!)
├── Auth.gs              # Security
├── DataProcessor.gs     # Data analysis
├── Report.gs            # Report generation
├── UI.gs                # User interface
├── Deployment.gs        # Bulk deployment
└── README.md            # This file
```

## Support

For issues:
1. Check the troubleshooting section above
2. Verify your data format matches the example
3. Ensure you changed the default password 