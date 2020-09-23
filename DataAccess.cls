VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "CodeGenLevel" ,"1-V1EMC"
Attribute VB_Ext_KEY = "CodeGenDate" ,"15-Jul-2000 08:44:24"
Option Explicit

'This data access wrapper illustrates the best techniques for accessing your
'data.  To use; create an instance of this class an call the methods.  You
'could do this from anywhere but preferably you have created data layer
'class(es) which do this..

'Class constants
Const C_CLASSNAME = "clsDataAccess"

'Wrap ADO CommandTypeEnum with our own to future proof
Public Enum QueryOptions
   QO_UseTable = ADODB.adCmdTable
   QO_UseText = ADODB.adCmdText
   QO_UseStoredProc = ADODB.adCmdStoredProc
   QO_UseUnknown = ADODB.adCmdUnknown
   QO_UseDefault = -1
End Enum

'String handling options
Public Enum StringHandling
   SH_GetCSV = 1
   SH_GetHTML = 2
End Enum

Private Function GetConnectionString() As String

   'This function returns a hard coded connection string to your database.  This method
   'is one of the faster styles available; although it does introduce some inflexibility.

   'There are alternate methods you can use to do this including using IObjectConstructor
   'in COM+, storing the connection string in the registry or an ini-files

   GetConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Northwind2.mdb;Persist Security Info=False"

End Function

Public Function GetRecordSet(ByVal sQuery As String, ByVal sConnect As String, eType As QueryOptions, ParamArray aParams() As Variant) As Recordset

   'Error handling
   On Error GoTo ErrorHandler

   'Create the ADO objects
   Dim RS As ADODB.Recordset, Cmd As ADODB.Command
   Set RS = New ADODB.Recordset
   Set Cmd = New ADODB.Command

   'Use helper function to build parameters for command object
   CollectParams Cmd, aParams

   'Determine whether passed in, or hard coded connection
   If sConnect = vbNullString Then
      Cmd.ActiveConnection = GetConnectionString()
   Else
      Cmd.ActiveConnection = sConnect
   End If

   'Init the ADO objects & the query parameters (if any)
   Cmd.CommandText = sQuery
   Cmd.CommandType = eType

   'Execute the query for readonly
   RS.CursorLocation = adUseClient
   RS.CursorType = adOpenForwardOnly
   RS.LockType = adLockBatchOptimistic
   RS.Open Cmd

   'Disconnect the recordset
   Set Cmd.ActiveConnection = Nothing
   Set Cmd = Nothing
   Set RS.ActiveConnection = Nothing

   'Return the resultant recordset
   Set GetRecordSet = RS
   Exit Function

ErrorHandler:

   'Clean up
   Set RS = Nothing
   Set Cmd = Nothing
   Set GetRecordSet = Nothing

   'Raise error
   Err.Raise 0, C_CLASSNAME, "GetRecordSet"

End Function

Public Sub PutRecordset(RS As ADODB.Recordset, Optional ByVal sConnect As String = "")

   'Error handling
   On Error GoTo ErrorHandler

   'Reconnect to database
   If sConnect = vbNullString Then
      RS.ActiveConnection = GetConnectionString()
   Else
      RS.ActiveConnection = sConnect
   End If

   'Update records
   RS.UpdateBatch

   'Disconnect the recordset again and exit
   Set RS.ActiveConnection = Nothing
   Exit Sub

ErrorHandler:

   'Raise error
   Err.Raise 0, C_CLASSNAME, "PutRecordSet"

End Sub

Public Function Execute(ByVal sQuery As String, ByVal sConnect As String, eType As QueryOptions, ParamArray aParams() As Variant) As Boolean

   Dim Cmd As ADODB.Command

   'Enable error handler
   On Error GoTo ErrorHandler

   'Assume failure
   Execute = False

   'Create the ADO objects
   Set Cmd = New ADODB.Command

   'Use helper function to build parameters for command object
   CollectParams Cmd, aParams

   'Determine whether passed in, or hard coded connection
   If sConnect = vbNullString Then
      Cmd.ActiveConnection = GetConnectionString()
   Else
      Cmd.ActiveConnection = sConnect
   End If

   'Init the ADO objects
   Cmd.CommandText = sQuery
   Cmd.CommandType = eType

   'Execute the query without returning a recordset
   Cmd.Execute , , ADODB.adExecuteNoRecords
   
   'Cleanup and return nothing
   Set Cmd.ActiveConnection = Nothing
   Set Cmd = Nothing

   'Return
   Execute = True
   Exit Function

ErrorHandler:

   'Cleanup and return nothing
   Set Cmd.ActiveConnection = Nothing
   Set Cmd = Nothing

   'Raise error
   Err.Raise 0, C_CLASSNAME, "Execute"

End Function

Public Function GetArray(ByVal sQuery As String, ByVal sConnect As String, eType As QueryOptions, ParamArray aParams() As Variant) As Variant()

   'Error handling
   On Error GoTo ErrorHandler

   'Create the ADO objects
   Dim RS As ADODB.Recordset, Cmd As ADODB.Command
   Set RS = New ADODB.Recordset
   Set Cmd = New ADODB.Command

   'Use helper function to build parameters for command object
   CollectParams Cmd, aParams

   'Determine whether passed in, or hard coded connection
   If sConnect = vbNullString Then
      Cmd.ActiveConnection = GetConnectionString()
   Else
      Cmd.ActiveConnection = sConnect
   End If

   'Init the ADO objects & the query parameters (if any)
   Cmd.CommandText = sQuery
   Cmd.CommandType = eType

   'Execute the query for readonly
   RS.CursorLocation = adUseClient
   RS.CursorType = adOpenForwardOnly
   RS.LockType = adLockBatchOptimistic
   RS.Open Cmd

   'Convert the recordset to an array
   If RS.EOF = False And RS.BOF = False Then
      GetArray = RS.GetRows
   Else
      ReDim GetArray(-1 To -1, -1 To -1)
   End If

   'Clean up and exit
   Set RS = Nothing
   Set Cmd = Nothing
   Exit Function

ErrorHandler:

   'Clean up (array left empty)
   Set RS = Nothing
   Set Cmd = Nothing

   'Raise error
   Err.Raise 0, C_CLASSNAME, "GetArray"

End Function

Public Function GetString(ByVal sQuery As String, ByVal sConnect As String, eType As QueryOptions, eStringType As StringHandling, ParamArray aParams() As Variant) As String

   'Error handling
   On Error GoTo ErrorHandler

   'Create the ADO objects
   Dim RS As ADODB.Recordset, Cmd As ADODB.Command
   Set RS = New ADODB.Recordset
   Set Cmd = New ADODB.Command

   'Use helper function to build parameters for command object
   CollectParams Cmd, aParams

   'Determine whether passed in, or hard coded connection
   If sConnect = vbNullString Then
      Cmd.ActiveConnection = GetConnectionString()
   Else
      Cmd.ActiveConnection = sConnect
   End If

   'Init the ADO objects & the query parameters (if any)
   Cmd.CommandText = sQuery
   Cmd.CommandType = eType

   'Execute the query for readonly
   RS.CursorLocation = adUseClient
   RS.CursorType = adOpenForwardOnly
   RS.LockType = adLockBatchOptimistic
   RS.Open Cmd

   'Convert the recordset to a string
   If RS.EOF = False And RS.BOF = False Then

      'Dermine formatting
      If eStringType = SH_GetHTML Then
         GetString = ConvertToHTML(RS)
      Else
         GetString = ConvertToCSV(RS)
      End If

   Else
      GetString = ""
   End If

   'Clean up and exit
   Set RS = Nothing
   Set Cmd = Nothing
   Exit Function

ErrorHandler:

   'Clean up (string left empty)
   Set RS = Nothing
   Set Cmd = Nothing

   'Raise error
   Err.Raise 0, C_CLASSNAME, "GetString"

End Function

Public Function ConvertToCSV(RS As ADODB.Recordset) As String

   'Check for closed or empty recordset
   If RS.EOF = True Or RS.BOF = True Then
      ConvertToCSV = ""
      Exit Function
   End If

   'Convert recordset to comma seperated values
   ConvertToCSV = RS.GetString(adClipString, -1, ",", vbCrLf, "(NULL)")

End Function

Public Function ConvertToHTML(RS As ADODB.Recordset) As String

   'Check for closed or empty recordset
   If RS.EOF = True Or RS.BOF = True Then
      ConvertToHTML = ""
      Exit Function
   End If

   'Convert recordset to HTML table syntax
   ConvertToHTML = RS.GetString(adClipString, -1, "</TD><TD>", "</TD></TR>" & vbCrLf & "<TR><TD>", "(NULL)")
   ConvertToHTML = "<TR><TD>" & Left(ConvertToHTML, Len(ConvertToHTML) - 8)

End Function

Private Sub CollectParams(ByRef Cmd As ADODB.Command, ParamArray ArgParams() As Variant)

   Dim vParams As Variant
   Dim i As Integer, v As Variant

   'Point to first element in array
   vParams = ArgParams(0)

   'For each array element
   For i = LBound(vParams) To UBound(vParams)
      Cmd.Parameters.Append Cmd.CreateParameter(vParams(i)(0), vParams(i)(1), adParamInput, vParams(i)(2), vParams(i)(3))
   Next i

End Sub

