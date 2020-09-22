Attribute VB_Name = "DbFields"
Option Private Module
Public DbPass As String         'Stores Database Password
Public DbFileName As String     'Stores Database FileName
Public rsName As String         'Stores Table Name
Public Db As Database
Public rs As Recordset
Public DbOpen As Boolean
Public RecOpen As Boolean
Public RecrdPos As Integer      'Indicates Current Record
Public RecNum As Integer        'Stores Total #Records
Public FldNum As Integer        'Stores Total #Fields
Public FldData As Variant       'Stores Field Data
Public CurrFldName As String    'Stores Field Name
Public CurrFld As Integer       'Indicates Current Field
Public Sub OpenDbase()
'On Error GoTo errH:
Set Db = OpenDatabase(DbFileName, False, False, ";pwd=" & DbPass)
    DbOpen = True
Exit Sub
'errH:
'MsgBox err.Description, vbCritical, "Open Database."
'DbOpen = False
'frmMain.Show
'err.Clear
End Sub
Public Sub RsOpen()
On Error GoTo errHandler:
Set rs = Db.OpenRecordset(rsName, dbOpenTable)
    RecOpen = True
    RecNum = rs.RecordCount
    FldNum = rs.Fields.Count - 1
    If Not rs.RecordCount = 0 Then
        RecrdPos = 1
    Else
        RecrdPos = 0
    End If
    CurrFld = 0
Exit Sub
errHandler:
    err.Clear
End Sub
Public Sub NextRecord()
On Error GoTo errHandler:
    If Not rs.EOF Then
        RecrdPos = RecrdPos + 1
        rs.MoveNext
        GetRecordData
        GetFieldName
    End If
Exit Sub
errHandler:
    err.Clear
    If Not rs.RecordCount = 0 Then
    rs.MoveLast
    GetRecordData
    GetFieldName
    RecrdPos = RecNum
    Else
    RecrdPos = 0
    RecNum = 0
    End If
End Sub
Public Sub GetRecordData()
If Not rs.RecordCount = 0 Then
    If Not IsNull(rs.Fields(CurrFld)) Then
        FldData = rs.Fields(CurrFld)
    Else
        FldData = ""
    End If
Else
    FldData = ""
End If
End Sub
Public Sub GetFieldName()
    CurrFldName = rs.Fields(CurrFld).Name
End Sub
Public Sub PrevRecord()
On Error GoTo errHandler:
    If Not rs.BOF Then
        RecrdPos = RecrdPos - 1
        rs.MovePrevious
        GetRecordData
        GetFieldName
    End If
Exit Sub
errHandler:
    err.Clear
    If Not rs.RecordCount = 0 Then
    rs.MoveFirst
    GetRecordData
    GetFieldName
    RecrdPos = 1
    Else
    RecrdPos = 0
    RecNum = 0
    End If
End Sub
Public Sub NextFld()
    If Not (CurrFld + 1) > FldNum Then
        CurrFld = CurrFld + 1
        GetFieldName
        GetRecordData
    End If
End Sub
Public Sub PrevFld()
    If Not (CurrFld - 1) < 0 Then
        CurrFld = CurrFld - 1
        GetFieldName
        GetRecordData
    End If
End Sub
Public Sub DelRec()
On Error GoTo errHandler:
    rs.Delete
    RecNum = RecNum - 1
    If RecrdPos = RecNum Then
        RecrdPos = RecrdPos + 1
        PrevRecord
    ElseIf RecrdPos = 1 Then
        RecrdPos = RecrdPos - 1
        NextRecord
    Else
        RecrdPos = RecrdPos - 1
        NextRecord
    End If
Exit Sub
errHandler:
err.Clear
End Sub
