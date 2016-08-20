# party-namelist
Match names from a .txt, sort them and put them into a Excel document.

This program reads trough a .txt document line by line, so names have to have its own line, but must not necessarily stand alone. The regex filters out any numbers and matches the regex pattern for fore- and surename. [Forename] [Surename]

Then it shows the user how many names are found and compares it to the actual linecount in the file. If line and namecount is not equal the program will inform you.
Note: The program only works for an amount of 200 names which can be printed properly any more than that and printing might not be as easy.

This program is free to use and critic is always welcome.
