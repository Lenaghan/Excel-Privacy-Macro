# Excel-Projects

### pii_scrubber.bas
Exported VBA module with a few functions.
1. GenerateNameCombinations
     Needs a table with two columns in Sheet1 (FirstName, LastName).
     Iterates through all combinations of the two columns.
     Outputs name combinations to Sheet1 starting in column E.
2. ReplaceNamesWithFakes
     The first variable construcion hold the column where the names are (future update will include pop-up input instead).
     Script iterates through the cells in the column, replacing with random names from the output of GenerateNameCombinations.
     Starts at the second row that is not empty to skip header.
     Allows for 15% chance of duplication of names.
     Does not populate any cells that were empty
     Call ResetCell to clear existing formatting before replacing any name.
3. ResetCell
     Sets all cell formats to False, 0 or xlAutomatic
     Sets font to 11pt Arial 
6. ReplacePIIWithRegex - Remove sensitive and identifiable information.
     Add VBA reference to "Microsoft VBScript Regular Expressions 5.5".
     In "Microsoft Visual Basic for Applications" window select "Tools" from the top menu. Select "References".
     Check the box next to "Microsoft VBScript Regular Expressions 5.5" to include in your workbook.
     Click "OK".

     Additional regex that can be inserted, as needed:
   
       International Phone Numbers:
         ^(?:(?:\(?(?:00|\+)([1-4]\d\d|[1-9]\d?)\)?)?[\-\.\ \\\/]?)?((?:\(?\d{1,}\)?[\-\.\ \\\/]?){0,})(?:[\-\.\ \\\/]?(?:#|ext\.?|extension|x)[\-\.\ \\\/]?(\d+))?$
        
       SSN - either hypen(-) space( ) none as separator:
         ^d{3}-\d{2}-\d{4}$ | ^\d{9}$ | d{3}\w\d{2}\w\d{4}$$
        
       Passport:
         ^[A-PR-WY][1-9]\d\s?\d{4}[1-9]$
        
       emails:
         ^[a-zA-Z0-9._%-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}
         ^[\w\.=-]+@[\w\.-]+\.[\w]{2,3}$
        
       URLs:
         (?i)\b((?:[a-z][\w-]+:(?:\/{1,3}|[a-z0-9%])|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}\/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'".,<>?«»“”‘’]))
        
       File paths:
         \\[^\\]+$
         ("^d{3}-\d{2}-\d{4}$ | ^\d{9}$ | d{3}\w\d{2}\w\d{4}$$", "^[A-PR-WY][1-9]\d\s?\d{4}[1-9]$","^[a-zA-Z0-9._%-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}", "^[\w\.=-]+@[\w\.-]+\.[\w]{2,3}$") 

### Scrubber.xlsm
Excel macro-enabled spreadsheet with name generator already executed.
Includes pii_scrubber.bas in VB editor.
