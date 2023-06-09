VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' we can use this class to return an ADO recordset from a closed Excel file
Private pConnection As ADODB.Connection
Private pRecordSet As ADODB.Recordset
Private pSQLFields As String
Private pSQLExtra As String
Private pSQL As String
Private pDataSource As String
Private pTable As String
Private pDset As cDataSet
Private pWhere As Range
Private pClearSheet As Boolean
Private pCreateDset As Boolean
Private peAdoConnection As eAdoConnections
Public Enum eAdoConnections
    eAdoAuto
    eAdoExcel2007
    eAdoAccess2007
    eAdoUnknown
End Enum
Public Property Get sql() As String
    sql = pSQL
End Property
Public Property Get where() As Range
    Set where = pWhere
End Property
Public Property Get dset() As cDataSet
    Set dset = pDset
End Property
Public Function init(Optional rOutRange As Range = Nothing, _
                      Optional sDataSource As String = vbNullString, _
                      Optional bClearsheet As Boolean = True, _
                      Optional bCreateDset As Boolean = True, _
                      Optional eConnection As eAdoConnections = eAdoAuto, _
                      Optional complain As Boolean = True) As cADO
                      
    Set pWhere = rOutRange
    pCreateDset = bCreateDset
    pClearSheet = bClearsheet
    Set pConnection = New ADODB.Connection
    Set pRecordSet = New ADODB.Recordset
    pDataSource = sDataSource
    peAdoConnection = eConnection
    Set pDset = Nothing
    If pDataSource = vbNullString Then pDataSource = ThisWorkbook.path & "\" & ThisWorkbook.Name
    If peAdoConnection = eAdoAuto Then
        peAdoConnection = tryToGetConnectionType
        If peAdoConnection = eAdoUnknown Then
            If complain Then MsgBox ("Dont know how to connect to " & pDataSource)
            Set init = Nothing
            Exit Function
        End If
    End If
    Set init = Me
End Function
Public Function kill()
    
    With pRecordSet
        .Close
    End With
    
    With pConnection
        .Close
    End With
    
    Set pRecordSet = Nothing
    Set pConnection = Nothing
    
End Function
Private Function tryToGetConnectionType() As eAdoConnections
    Dim p As Long
    tryToGetConnectionType = eAdoUnknown
    p = InStrRev(pDataSource, ".")
    If p <> 0 Then
        Select Case Mid(pDataSource, p + 1)
            Case "xlsm", "xlsx", "xlsb"
                tryToGetConnectionType = eAdoExcel2007
            
            Case "accdb"
                tryToGetConnectionType = eAdoAccess2007
            
        End Select
    End If
    
End Function
Public Function execute(Optional sTable As String = vbNullString, _
                        Optional sSqlFields As String = "*", _
                        Optional sSqlExtra As String = vbNullString) As cADO
    Dim fCol As ADODB.Field, r As Range, c As Long, w As Worksheet, cString As String

    ' CONNECT TO target datasource and execute sql
    Set pConnection = New ADODB.Connection
    pTable = sTable
    If pTable = vbNullString Then pTable = ActiveSheet.Name
    
    Select Case peAdoConnection
        Case eAdoExcel2007
            cString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & pDataSource & ";" & _
                      "Extended Properties=""Excel 12.0;HDR=Yes"";"
            pSQL = Trim("select " & thisOrThat(sSqlFields, pSQLFields) & " from [" & _
            thisOrThat(sTable, pTable) & "$] " & thisOrThat(sSqlExtra, pSQLExtra))
            
        Case eAdoAccess2007
            cString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & pDataSource & ";" & _
                      "Persist Security Info=False;"
            pSQL = Trim("select " & thisOrThat(sSqlFields, pSQLFields) & " from [" & _
            thisOrThat(sTable, pTable) & "] " & thisOrThat(sSqlExtra, pSQLExtra))
            
        Case Else
            Debug.Assert False
    
    End Select
    
    With pConnection
        .Open cString
    End With
    
    Set pRecordSet = New ADODB.Recordset
    With pRecordSet
        .Open pSQL, pConnection, adOpenStatic, adLockOptimistic
        ' headings
        If pWhere Is Nothing Then
            Set w = Sheets.add
            Set pWhere = w.Cells(1, 1)
        End If
        If pClearSheet Then pWhere.Worksheet.Cells.ClearContents
        Set r = pWhere.Resize(1, 1)
        For Each fCol In .Fields
            r.value = fCol.Name
            Set r = r.Offset(, 1)
        Next fCol
        Set r = pWhere.Resize(1, 1).Offset(1)
        While Not .EOF
            c = 0
            For Each fCol In .Fields
                r.Offset(, c).value = fCol.value
                c = c + 1
            Next fCol
            Set r = r.Offset(1)
            .MoveNext
        Wend
        ' reset size of created data
        Set pWhere = pWhere.Resize(r.row - pWhere.row, .Fields.count)

    End With

    ' now let's create a new cDataSet
    If pCreateDset Then
        Set pDset = New cDataSet
        With pDset
            .populateData pWhere, , pTable, , , , , , , False
        End With
    End If
    Set execute = Me
End Function
Private Function thisOrThat(sThis As String, sThat As String) As String
    If sThis = vbNullString Then
        thisOrThat = sThat
    Else
        thisOrThat = sThis
    End If
End Function
Private Function createTable(tableName As String, cj As cJobject)
    ' drop existing version
    ' this is the only time i've ever used resume next
    On Error Resume Next
    pConnection.execute "DROP TABLE " & tableName
    On Error GoTo 0
    
    Dim jo As cJobject, u As String

    u = vbNullString
    For Each jo In cj.children
        If (Len(u) > 0) Then u = u + ","
        u = u & jo.child("name").toString & "," & jo.child("type").toString
    Next jo
    u = "CREATE TABLE " & tableName & "(" & u & ")"
End Function
Private Function insertIntoTable(ds) As cJobject
    Dim c As String, dr As cDataRow, dc As cCell
    For Each dr In ds.rows
        c = vbNullString
        For Each dc In dr.columns
            If Len(c) > 0 Then c = c + ","
            
        Next dc
    Next dr
    

End Function

