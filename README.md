# excel-diff
## This repository will provide you Python3 code to find differences between two workbooks.

This will check for each column and row value for each Sheet of two Workbooks. 

## Packages Required :
1. OPENPYXL {can be installed by pip : pip install openpyxl}
2. PYTHON 3.6 {python.org}

## Command :
```console
py@bar : python exceldiff.py <workbookname1> <workbookname2>
```

## Working :
1. Make sure both Workbooks contains same number of sheets with same name. 
2. All the differences will be stored in a file whose name will be displayed on terminal. The Report will contain all the required information classified by sheet name and row-column indices.
