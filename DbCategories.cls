VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDbCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "CodeGenLevel" ,"1-V1EMC"
Attribute VB_Ext_KEY = "CodeGenDate" ,"15-Jul-2000 08:44:24"
Option Explicit

'This class shows you how you might go about using the wrapper..

'Class constants (data access layer object)
Const C_CLASSNAME = "clsDbCategories"

'Date format for database inserts
Const C_DATE_FORMAT As String = "dd-mmm-yyyy"

'Our data access wrapper
Private moData As clsDataAccess

Private Sub Class_Initialize()

   'Create data access wrapper
   Set moData = New clsDataAccess

End Sub

Private Sub Class_Terminate()

   'Clean up data access wrapper
   Set moData = Nothing

End Sub

'Example function you might have calling the data
'access wrapper instead of ADO directly. Updates data..
Public Function Update(ByVal lCategoryID As Long, ByVal sCategoryName As String, ByVal sDescription As Variant)

   'Error handling
   On Error GoTo ErrorHandler

   'Variables
   Dim sQuery As String

   'Build query for update
   sQuery = "UPDATE Categories " & _
               "SET CategoryName = ?, " & _
                   "Description = ? " & _
             "WHERE CategoryID = ? "

   'Update database
   moData.Execute sQuery, "", QO_UseText, _
            Array("@CategoryName", adVarWChar, 15, sCategoryName), _
            Array("@Description", adLongVarWChar, 0, sDescription), _
            Array("@CategoryID", adInteger, 0, lCategoryID)
   
   'Commit/exit
   Exit Function

ErrorHandler:

   'Clean up
   Err.Raise 0, C_CLASSNAME, "Update"

End Function

'Example function to select data
Public Function Populate(ByVal lCategoryID As Long) As ADODB.Recordset

   'Error handling
   On Error GoTo ErrorHandler

   'Variables
   Dim sQuery As String

   'Build Query
   sQuery = "SELECT CategoryID, " & _
                   "CategoryName, " & _
                   "Description " & _
              "FROM Categories " & _
             "WHERE CategoryID = ? "

   'Return records from database
   Set Populate = moData.GetRecordSet(sQuery, "", QO_UseText, _
            Array("@CategoryID", adInteger, 0, lCategoryID))
   
   'Commit/exit
   Exit Function

ErrorHandler:

   'Clean up
   Err.Raise 0, C_CLASSNAME, "Update"

End Function

