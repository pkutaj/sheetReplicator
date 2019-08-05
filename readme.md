# The VBA Sheet Replicator 

## PURPOSE
The aim of this code is to copy a template sheet with masterdata into all of the sheets in the same folder:

* Using single master data sheet for 30+ excel sheets for easy maintenance
* I know there is not enough abstraction such as the name of the excel file
* The pattern is to create just a single instance of a sheet pointing to the master data sheet and then replicate this to N sheets in question
* The context is that these sheets are used as inputs for report creation and broadcast

## TOC

STEP | ACTION
-----|----------------------------------------------------------------
0    | Declare global bindings
1.   | Get file-names from a current folder
1.1  | Declare dynamic array for filenames
1.2  | Make sure you are in a current folder and assign the first file
1.3  | Loop through the files in the current folder to fill an array
2.   | Loop through the array and move the data in
2.1  | Copy the master data sheet
2.2  | Copy the headers range
2.3  | Save and close without asking

## SIGNATURE 
mrPaul, 2019-08-05 20:00:58, early August evening in Brno