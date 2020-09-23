Attribute VB_Name = "modSample"
Option Explicit

'The data access wrapper offers the following functions:
'-------------------------------------------------------
'moData.GetArray - Get data in two dimensional array
'moData.GetRecordSet - Get data in recordset
'moData.GetString - Get data in delimited string
'moData.Execute - Update/Delete/Insert data
'moData.PutRecordset - Write disconnected recordset back to database
'moData.ConvertToCSV - Convert recordset to comma delimited
'moData.ConvertToHTML - Convert recordset to HTML table

'I have put some articles about data access wrappers, segmenting your
'code and n-teir development my site www.vbdatabase.com.

Sub Main()

   'Variables
   Dim oCategory As clsDbCategories
   Dim RS As ADODB.Recordset
   
   'Create category object
   Set oCategory = New clsDbCategories
   
   'Get data from database
   Set RS = oCategory.Populate(1)
   
   'Display data
   MsgBox "CategoryID: " & RS("CategoryID") & vbCrLf & "CategoryName: " & RS("CategoryName"), vbOKOnly + vbInformation, "Data Access Wrapper"
   
   'Clean up
   Set RS = Nothing
   Set oCategory = Nothing
   
   MsgBox "You can use this data access layer to generate HTML tables, CSV files etc for your vb programmes and Web components.", vbOKOnly + vbInformation, "Data Access Wrapper"
   
End Sub
