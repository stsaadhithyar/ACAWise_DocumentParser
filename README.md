# ACA_DocumentParser

There are mainly 5 iles which are present in the Repo and this code is mainly used to extract data from form 1095-c and then load them in the excel shee twhich is given as the Template.xlsx

Inorder to acces it First install the necessary packages which is present in the requirements.txt inorder to install it use the below command in the Terminal

```
pip install -r requirements.txt
```

Firstly inorder to start the process the "Searchable_and_checkbox.py" must be executed which takes in 2 parameters which are the scanned PDF path and the degree of rotation. After executing
the code it generates a rotated and blacked out checkboxes incorporated PDF.

Iorder to generate a excel it take 2 files which are
* Form_field1.txt
* Template.xlsx

The Form_field1.txt has a key names so only for that keys th values are extracted. The data is loaded in the Tempalate.xlsx format and is generated as another excel file.
