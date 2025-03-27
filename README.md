# Text2Excel

Text2Excel is a GUI desktop application that can extract data from a text file and put them in an Excel or csv file using regular expression (regex) patterns. It uses python re module.

You should right click on the patterns widget and use the options in the context menu to add the patterns.

It has an option called 'Exact Order'. If you have not enabled this option, it starts placing the data from the last row or column of the Excel file in which there is data, and does not check the last cell of each column or row, but if you enable this option, it checks each column or row and puts data exactly along the previous data.

If you want to put only a part of pattern in the file you should make a pattern with a group named 'item'. like this:

```
\w{5}(?P<item>\d)
```

The above pattern finds a text that has 5 word chracters and a digit after it. It only puts the digit in the excel file. But if you don't make this group it puts both word characters and digit in the the excel file.

You can choose to put the data in columns or rows and in which sheet.

If you want to save data in a csv file you just need to give the output file `.csv` format

---
It uses openpyxl module so you need to install it with this commad:

```
$ pip install openpyxl
```
