
# Automate filling word documents with Python

I have done this project to learn how to fill in a predefined Word template based on Excel input with the help of python language.



## Acknowledgements
I took help from this page. 

Link:-https://betterprogramming.pub/automate-filling-templates-with-python-1ff6c6fd595e

## Deployment

Libraries used os,openpyxl,docx and docxtp.

os- Used for changing the current working directory path and also for making folders in the end.
openpyxl- Used for getting the data from the excel data source. Load_workbook is the package used to get the excel data.
docxtpl - To load the word templates to be filled automatically. 
DocxTemplate is the package used for loading the templates.

## Stages

1)Import the Libraries.

2)Define the paths for all the template files and Excel data.

3)Load the workbook and all the template files.

4)Make a dictonary where the keys are field names drawn from the excel worksheet.

5)Use while loop for iteration.

6)Make new folders and place the completed files in them.

