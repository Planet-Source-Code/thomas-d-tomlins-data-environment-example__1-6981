<div align="center">

## Data Environment Example


</div>

### Description

DataEnvironment is one item that is hard to find

Detail information about how to use it. I truly

Believe VB's DataEnvironment is the way to go

But using it takes time. This program will go

over some way's to make your data-environment more Flexable during run-time operations that is not usually covered in the majority books available to users.
 
### More Info
 
Must be able to get the Data Environment

None Found


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Thomas D\. Tomlins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thomas-d-tomlins.md)
**Level**          |Intermediate
**User Rating**    |3.8 (23 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/thomas-d-tomlins-data-environment-example__1-6981/archive/master.zip)





### Source Code

```
'IF ANYONE IMPROVES OR ADDS TO THIS CODE PLEASE FORWARD _
  A COPY TO ME SO I CAN UPDATE MY RECORDS AND INTERNITE SITES _
  E-MAIL: TDTOMLINS@YAHOO.COM
'DataEnvironment is one item that is hard to find _
  Detail information about how to use it. I truly _
  Believe VB's DataEnvironment is the way to go _
  But using it takes time. This program will go _
  over some way's to make your data-environment more _
  Flexable during run-time operations that is not _
  usually covered in the majority books available to users.
'When making changes be sure the Table,Field,Record is within _
  the database.
'Open a dataproject if you already have a form _
  open then you will have to add a _
  DataEnvironment to your project
' within data environment make a connection to _
  Biblio.mdb (comes with VB usually in dir _
  C:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb
'Create a command Add an SQL statement: Select * from Authors
'Create another command add a Data object-Database as TABLE _
  Object will be TITLES.
'Create a another command add a SQL statement: _
  SELECT Titles.* FROM Titles WHERE (`Year Published` = ?) _
  In the Paramaters Tab set DATA TYPE as SMLINT and _
  set HOST DATA TYPE as INTEGER.
'ON THE FORM ADD THE FOLLOWING
'Add To the from a DataGrid, Three CommandButtons, _
  Three Labels with TextBox for each
Option Explicit
Private Sub Command1_Click()
On Error GoTo errorhandler
' To use this routine you MUST have your command _
  as a SQL statement and have a valid statement _
  within it.
DataEnvironment1.Commands.Item("Command1").CommandText = Text1.Text
'You must manually rebind your datagrid to activate the _
  Required commands
With DataGrid1
  .DataMember = "Command1"
  Set .DataSource = DataEnvironment1
End With
' You must close the recordset between commands
DataEnvironment1.rsCommand1.Close
Exit Sub
errorhandler:
  Call errorRoutine
  Resume Next
End Sub
Private Sub Command2_Click()
'Valad Tables: Titles, Publishers, Authors, 'Title Author'
'NOTE: you must put single ' around Title Author.
On Error GoTo errorhandler
' To use this routine you MUST have your command _
  as a DataObject statement and have a valid Object and _
  Object name within it.
DataEnvironment1.Commands.Item(2).CommandText = Text2.Text
'You must manually rebind your datagrid to activate the _
 Required commands
With DataGrid1
  .DataMember = "Command2"
  Set .DataSource = DataEnvironment1
End With
' You must close the recordset between commands
DataEnvironment1.rsCommand2.Close
Exit Sub
errorhandler:
  Call errorRoutine
  Resume Next
End Sub
Private Sub Command3_Click()
On Error GoTo errorhandler
' To use this routine you MUST have your command _
  as a SQL statement and have a valid statement _
  within it. Use the ? to indicate the Paramater. _
  Make sure your Parameter settings are correct.
DataEnvironment1.Command3 Text3.Text
'You must manually rebind your datagrid to activate the _
 Required commands
With DataGrid1
  .DataMember = "Command3"
  Set .DataSource = DataEnvironment1
End With
' You must close the recordset between commands
DataEnvironment1.rsCommand3.Close
Exit Sub
errorhandler:
  Call errorRoutine
  Resume Next
End Sub
Private Sub errorRoutine()
MsgBox ("You must have appropriate commands in the textbox")
End Sub
Private Sub Command4_Click()
  DataReport1.Show
End Sub
Private Sub Form_Load()
 MsgBox "Valid Tables: Titles, Publishers, Authors, 'Title Author'" _
    'NOTE: you must put single ' around Title Author."
Label1.Caption = " Enter SQL statement"
Text1.Text = "Select * From Titles"
Command1.Caption = "Run SQL statement"
Label2.Caption = "Enter Table Name"
Text2.Text = "Authors"
Command2.Caption = "Run Table Statement"
Label3.Caption = "Enter Year to search Publisher"
Text3.Text = "1985"
Command3.Caption = "Run Paramater Statement"
End Sub
```

