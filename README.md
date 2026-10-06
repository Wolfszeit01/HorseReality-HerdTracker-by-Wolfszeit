# Horse Reality Herd Tracker — Public Edition

The current public Apps Script source for the Herd Tracker by Wolfszeit.

## Install or update

Use a separate copy of the matching current spreadsheet template. This repository contains code, not a spreadsheet template; replacing code alone does not migrate an older workbook's columns or formulas.

1. Back up your spreadsheet before updating.
2. Open **Extensions → Apps Script** in the copy.
3. Replace the old project code with **all `.gs` and `.html` files in this repository**. Remove obsolete script files so duplicate functions are not left behind. Also include `appsscript.json` (enable the manifest editor under Project Settings).
4. Save, reload the spreadsheet and authorize the script when requested.
5. Under **Herd Management → Appearance**, apply the current theme and set your branding.

Download the repository using **Code → Download ZIP** to get all files together. Images, this README, and `version.txt` do not need to be pasted into Apps Script.

For bulk uploads with clasp, connect a separate checkout to your intended Apps Script project and upload only the code, HTML and manifest. This repository deliberately contains no connected `.clasp.json` or credentials.

## Themes and branding

- Ocean Blue is the public default.
- Forest & Gold and Midnight & Silver are also included.
- Add your own colours through the Theme Editor.
- A bundled horseshoe logo and **HERD TRACKER** name provide neutral defaults. Configure your stud name, signature and logo under Appearance.
- Leave the logo URL empty when saving to restore the bundled logo. Custom logos require a direct public HTTPS image URL.

## Sharing your own template

Remove personal horse records from every sheet, including Archive, Pedigree, breeding/team/market tabs, result columns and logs. Review hidden sheets, notes, links, balances and settings while preserving template formulas and structure.

**Archive All Horses is not a full cleanup:** it preserves archived records and does not clear every related tab.

## Help and history

Use the built-in **Help & Guide** and contextual tooltips. The screenshot assets here remain available to those dialogs.

See [version.txt](version.txt) for the changelog. Superseded script files remain recoverable through Git history.

The public source has passed local syntax and configuration checks. Test imports and formatting in your separate spreadsheet copy before using it with important data.
