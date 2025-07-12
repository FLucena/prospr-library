# PROSPR Financial Planning Library

A Google Apps Script library that automatically analyzes budget data and generates compact reports.

## What It Does

- Reads budget data from the 'Monthly Budget' tab
- Identifies spending patterns and deviations
- Generates a professional report
- Creates email drafts ready to send

## Quick Setup

1. Create a new Google Apps Script project
2. Copy all `.gs` files into your project
3. Save the project
4. Add the library to your spreadsheet's Apps Script
5. You'll see a new "Admin" menu

## Data Format

The target spreadsheet must have:
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

1. Set up budget data in the target spreadsheet's "Monthly Budget" sheet
2. Use `Admin > Monthly Comparative Report`
3. Use `Admin > Create Email Draft`

## Security

**Important**: Change the default password in `env.gs` before using:
1. Open or create `env.gs`
2. Change `ADMIN_CODE` to a secure password
3. Use at least 8 characters

## Configuration

Key settings in `Config.gs`:
- `DEVIATION_THRESHOLD`: 15% (minimum deviation to flag)
- `DATA_START_ROW`: 4 (where data begins)

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
2. Verify the target spreadsheet's data format matches the example
3. Ensure you changed the default password 