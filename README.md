# sociometry_visualizationscript

### Summary
xls VBA script, which creates the input file for graphviz visualization tool
Excel sheets should follow a specific format
VBA script needs to be customized

### "Installing"
1. Follow the format of the sheet
2. Copy-paste the script into an Excel macro and customize it
3. Install graphviz (https://graphviz.gitlab.io/download/)

### Format of Excel sheet
- Rows: list of those who provided the referral
- Columns: list of those who received the referral
- Row 1: Number of the individual (used it for verifiing if everyone is on the list)
- Row 2: Any binary parameter
- Row 3: Project code
- Row 4: Gender
- Row 5: Team lead (without space lastname, e.g. JDoe)
- Row 6: Name of the person received referral (without space between first and lastname)
- Row 7: Total referrals received
- Column 1: Number of the individual (used it for verifiing if everyone is on the list)
- Column 2: Name of the one who provided referral (without space between first and lastname)
- Column 3: Project code
- Column 4: Gender
- Column 5: Any binary parameter
- Column 6: 1 if the person is a teamleader
- Column 7: Team lead of the person
- Column 8: Total

