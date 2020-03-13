# InDesign utility Scripts

This is a repository for scripts that increase productivity in InDesign, related to my work.

- `add_metadata_from_csv.js`: With this script, a user can put data into a CSV, and run the script to import that data into the file info of the InDesign file.

- `premedia_export_covers.js`: With this script, a user can export a high-resolution PDF from their export presets. Export a low-resolution PDF from their export presets. Add images to a dynamic list for cover image export. User settings are stored for reuse. All files exported are moved and renamed to a organized folder structure. This script is designed to be run before `premedia_export_interior.js`

- `premedia_export_interior.js`: With this script, a user can export a high-resolution PDF from their export presets. Export a low-resolution PDF from their export presets. Designate a low-resolution cover file attach to the low-resolution PDF upon export. Designate a table of contents image and add interior spread images to a dynamic list for image export. User settings are stored for reuse. All files exported are moved and renamed to a organized folder structure. This script is designed to be run after `premedia_export_covers.js`

- `glossary_hyperlink_builder`: With this script, a user can target a text frame containing glossart terms. A hyperlink desination is created for the first word of each paragraph. An array of all matching words is built and the user is prompted to select a matching word to link to. Once confirmed, hyperlinks and hyperlink destinations are created with the hyperlink_glossary style applied.

- `grab_font_info.js`: With this script, a user can run this script to create a text file containing: font name, font type, and number of styles font is applied to. Text file is created to the same directory as the opened file. This is designed to be added to indesigns start-up scripts to auto generate these logs.
