# item-checker
Python GUI to check for banned keywords and basic language checks in an Excel spreadsheet

This is for specific language patterns and Excel spreadsheet format from ED output. Any issues which are found are 
saved to a separate error summary Excel spreadsheet in the same directory as the input file.

Custom dictionary: optional text file with each word on a separate line. The text file can be created/updated 
separately in a text editor.

Lists of banned words are defined towards the start of the item-checker.py file. The lists can be modified ad-hoc 
whilst running the GUI by modifying the lists in the text boxes.
