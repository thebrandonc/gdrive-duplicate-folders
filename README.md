<!-- DESCRIPTION -->
## Script for Duplicating Folders in Google Drive

This repo contains a script for duplicating Google Drive folders in bulk. The script adds a new menu item to your Google Sheet with an option to run the duplication process. The script will parse the contents of your spreadsheet and duplicate any folders that don't contain a timestamp in column `A`. All folders will be duplicated into the same parent folder as the source folder.

Check out the YouTube tutorial here: [Check out the Tutorial](https://youtu.be/9n0SCQFaZyw)

## Getting Started

The easiest way to get started is by copying the following spreadsheet template to your Google Drive: [Spreadsheet Template](https://docs.google.com/spreadsheets/d/1psxqXe_9PG5ITPYeU5Sk8Fm0PhjihfsphqN-TAkeEio/edit?usp=sharing). Once copied, run the script by clicking the 'Duplicator' option in the menu bar and selecting 'Duplicate Folders'. Google will then ask that you authorize some permissions, and once they're authorized you're ready to start duplicating folders.

Alternatively, you can get started by manually inserting the script into your Apps Script editor. Start by copying the contents of `duplicateFolders.js`, then create a new spreadsheet in Google Sheets, and open the Apps Script editor `Extensions > Apps Script`. Once the editor is open, delete the `myFunction()` function and paste the script copied from `duplicateFolder.js`. Next, run the `main()` function and Google will prompt you with a request to authorize some permissions so the script can access your spreadsheet and google drive account. 

Once authorization is granted, you are ready to set up the spreadsheet to start duplicating Google Drive folders. Please follow the format of the template spreadsheet provided here: [Spreadsheet Template](https://docs.google.com/spreadsheets/d/1psxqXe_9PG5ITPYeU5Sk8Fm0PhjihfsphqN-TAkeEio/edit?usp=sharing)

The header of your spreadsheet should begin with the following (starting with cell `A1`): 

| date_created   | parent_folder_link             | new_folder_name  |
|:---------------|:-------------------------------|:-----------------|
| [leave blank]  | https://link_to_parent_folder  | new-folder-name  |

Please leave the date created column (column `A`) empty, as the script will use it to determine which folders need to be copied. Once a folder has been copied, a timestemp will appear confirming the date and time it was copied.

The contents of the following two columns should be as follows:\
`parent_folder_link`: link to the folder you would like to copy\
`new_folder_name`: new name for the duplicated folder

Once your ready to run the script, select the `Duplicate Folders` option under the `Duplicator` menu at the top of the spreadsheet.

<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.md` for more information.