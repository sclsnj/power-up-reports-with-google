# Power Up Reports With Google
__Google Apps Scripts covered in the RIPL "Power Up Reports with Google" webinar (https://ripl.lrs.org/webinars/)__

These sample Google Apps Scripts (GAS), created by staff at the Somerset County Library System of New Jersey (SCLSNJ), are intended for use with Google Sheets and are customized for SCLSNJ's implementation of the TLC/CARL integrated library system product. Feel free to copy and modify; just make sure to customize for your own data sources.

---
***Coming Soon!*** 
More links to live Google Sheets files and updated code. (9/24/2019)

---
__People Count Trends__
https://docs.google.com/spreadsheets/d/1mgWnI4pYLN-OMiks8Bv9_ySjoKLEGwiFzQCyUCTr5pY/edit?usp=sharing
 * Dump and Format method
 * Unbound script can be saved as a standalone file
 * Add a time-based trigger to schedule the script to run automatically (ours runs daily)
 * Looks for a file matching a particular year + name combination
   * If it finds the file, it appends the data to what's already there
   * If if doesn't find the file, it creates it and then adds the data
 * Looks for Gmail messages received yesterday with a particular label and extracts the .CSV attachment
 * The fancy heat map is all Google Sheets formulas and conditional formatting -- see https://github.com/sclsnj/power-up-reports-with-google/blob/master/Heat%20Map%20Instructions.md for more info

__Circ Transaction Trends__
 * Dump and Format method
 * Unbound script can be saved as a standalone file
 * Add a time-based trigger to schedule the script to run automatically (ours runs daily)
 * Moderate query
 * Looks for a file matching a particular year + name combination
   * If it finds the file, it appends the data to what's already there
   * If if doesn't find the file, it creates it and then adds the data
 * The fancy heat map is all Google Sheets formulas and conditional formatting -- see https://github.com/sclsnj/power-up-reports-with-google/blob/master/Heat%20Map%20Instructions.md for more info

__Monthly Statistics__
 * Parse and Update method
 * Unbound script can be saved as a standalone file
 * Add a time-based trigger to schedule the script to run automatically (ours runs monthly)
 * Moderate query
 * Looks for a file matching a particular year + name combination
 * This Google Sheets file also uses a bunch of other approaches covered in the webinar

__Long In Transit__
 * Parse and Update method
 * Bound script needs to be added to the Script Editor for a specific Google Sheet in order to run
 * Add a time-based trigger to schedule the script to run automatically (our runs weekly)
 * Moderate query
 * Includes lots of script logic to slice and dice data by branch
 * Replaces all existing data in every tab with new data

__High Holds__
 * Parse and Update method
 * Bound script needs to be added to the Script Editor for a specific Google Sheet in order to run
 * Runs on demand via menu option
 * VERY complicated query - you'll want to do extensive testing before you embed in the script
   * See https://github.com/sclsnj/power-up-reports-with-google/blob/master/High%20Holds%20annotated.sql for more in-depth info
 * Script includes some added logic to fine-tune hold ratios
 * Inserts new data into a new tab, instead of overwriting existing data

