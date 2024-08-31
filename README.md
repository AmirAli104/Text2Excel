# Text2Excel

Text2Excel is a GUI desktop application that can extract data from a text file and put them in an Excel file using regular expression (re) patterns. It uses python re module.

You should right click on the patterns widget and use the options in the context menu to add the patterns.

It has an option called 'Exact Order'. If you have not enabled this option, it starts placing the data from the last row or column of the Excel file in which there is data, and does not check the last cell of each column or row, but if you enable this option, it checks each column or row and puts data exactly along the previous data.

You can choose to put the data in columns or rows and in which sheet

It uses openpyxl module so you need to install it with this commad:

```
$ pip install openpyxl
```
