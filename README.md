# gendoc
Generate Word (Docx) document from template filled with provided data in excel file.


## For DEV
To create the standalone application, type:
```
pyinstaller --onefile --noconsole main.py --hidden-import="openpyxl.cell._writer" 
```

NB Make sure openpyxl version is 3.0.9