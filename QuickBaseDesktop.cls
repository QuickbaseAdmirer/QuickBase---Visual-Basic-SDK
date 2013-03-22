VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QBDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Binary
Option Explicit
Option Base 0

'© 2001 Intuit Inc. All rights reserved.
'Use is subject to the IP Rights Notice and Restrictions available at
'http://developer.intuit.com/legal/IPRNotice_021201.html


Public QDBHost As String
Public useHTTPS As Boolean

Public qdb As QuickBaseClient
Public qdbResponse As New MSXML.DOMDocument

Const UPDATE_ID_FIELDNAME = "_qdb_update_id_"
Public MinutesGMTOffset As Integer

Sub displayEditingForm(strDBID As String, mode As Integer)


DoCmd.OpenForm "frm_qdb_" + strDBID, mode
End Sub

Function downloadAttachedFile(DBID As String, strFileName As String)
    Dim fs As Variant
    Dim f As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(Application.CurrentProject.Path + "\" + DBID + "\" + strFileName) Then
        fs.deleteFile Application.CurrentProject.Path + "\" + DBID + "\" + strFileName, True
    End If
    Call qdb.downloadAttachedFile(DBID, Application.CurrentProject.Path + "\" + DBID, strFileName)
    Set f = fs.GetFile(Application.CurrentProject.Path + "\" + DBID + "\" + strFileName)
    If Not (f.Attributes = f.Attributes And 1) Then
        f.Attributes = f.Attributes + 1
    End If
End Function

Sub updateViewsTable(strDBID As String)
    Dim rst As ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    
    Dim strFID As String
    Dim strCriteria As String
    Dim strFIDs() As String
    Dim strQuery As String
    Dim strOptions As String
    Dim strCriteriaParts() As String
    Dim strFieldName As String
    Dim strViewName As String
    
    Dim strQID As String
    Dim strSQL As String
    Dim ViewNodeList As MSXML.IXMLDOMNodeList
    Dim FieldNodeList As MSXML.IXMLDOMNodeList
    
    Dim strWHERE As String
    Dim strORDERBY As String
    
    
    Set ViewNodeList = qdbResponse.documentElement.selectNodes("/*/table/queries/query")
    'get the List All field list from the field node list
    Set FieldNodeList = qdbResponse.documentElement.selectNodes("/*/table/fields/field[appears_by_default=1]")

'Clean out the view descriptions in the qdbViews table
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL ("DELETE FROM qdbViews WHERE (((qdbViews.DBID)='" + strDBID + "'));")
    DoCmd.SetWarnings (True)
    
    Set rst = New ADODB.Recordset
    rst.Open "qdbViews", CurrentProject.Connection, adOpenForwardOnly, adLockOptimistic

    For i = 0 To ViewNodeList.length - 1
        If ViewNodeList(i).selectNodes("qyopts").length > 0 Then
            strOptions = ViewNodeList(i).selectSingleNode("qyopts").nodeTypedValue
            If InStr(strOptions, "xst.") = 0 And InStr(strOptions, "xfl.") = 0 Then
                rst.AddNew
                rst("DBID") = strDBID
                rst("Name") = makeMSAccessObjectName(ViewNodeList(i).selectSingleNode("qyname").nodeTypedValue)
                rst("QID") = CInt(ViewNodeList(i).selectSingleNode("@id").nodeTypedValue)
                strSQL = "SELECT "
                If ViewNodeList(i).selectNodes("qyclst").length > 0 Then
                    strFIDs = Split(ViewNodeList(i).selectSingleNode("qyclst").nodeTypedValue, ".")
                Else
                    'get the List All clist from the field node list
                    ReDim strFIDs(FieldNodeList.length - 1)
                    For j = 0 To FieldNodeList.length - 1
                        strFIDs(j) = FieldNodeList(j).selectSingleNode("@id").Text
                    Next j
                End If
                For j = 0 To UBound(strFIDs)
                    If strFIDs(j) <> "-1" Then
                        strFieldName = makeMSAccessObjectName(qdbResponse.documentElement.selectSingleNode("/*/table/fields/field[@id=" + strFIDs(j) + "]/label").Text)
                        strSQL = strSQL + "qdb_" + strDBID + ".[" + strFIDs(j) + "] AS [" + strFieldName + "], "
                    End If
                Next j
                strSQL = Trim(strSQL)
                If Right(strSQL, 1) = "," Then
                    strSQL = Left(strSQL, Len(strSQL) - 1)
                End If
                strSQL = strSQL + " FROM qdb_" + strDBID + " "
                'parse the query for the where clause
                strQuery = ViewNodeList(i).selectSingleNode("qycrit").nodeTypedValue
                If strQuery <> "{'0'.CT.''}" Then
                    strSQL = strSQL + " WHERE "
                    strWHERE = ""
                    Do While InStr(strQuery, "{") > 0
                        strCriteria = Mid(strQuery, InStr(strQuery, "{") + 1, InStr(strQuery, "}") - InStr(strQuery, "{") - 1)
                        strQuery = Mid(strQuery, InStr(strQuery, "}") + 1)
                        strCriteriaParts = Split(strCriteria, ".")
                        strFID = Mid(strCriteriaParts(0), 2, Len(strCriteriaParts(0)) - 2)
                        strWHERE = strWHERE + "qdb_" + strDBID + ".[" + strFID + "] "
                        'If we have a checkbox field we need to set the criteria to true or false
                        strCriteriaParts(2) = Mid(strCriteriaParts(2), 2, Len(strCriteriaParts(2)) - 2)
                        If qdbResponse.documentElement.selectSingleNode("/*/table/fields/field[@id=" + strFID + "]/@base_type").Text = "bool" Then
                            If UCase(strCriteriaParts(2)) = "YES" Or UCase(strCriteriaParts(2)) = 1 Or UCase(strCriteriaParts(2)) = "true" Then
                                strCriteriaParts(2) = "_True_"
                            Else
                                strCriteriaParts(2) = "_False_"
                            End If
                        End If
                        Select Case strCriteriaParts(1)
                            Case "CT" 'Contains
                                strWHERE = strWHERE + "like '*" + strCriteriaParts(2) + "*' AND "
                            Case "XCT" 'Does Not contain
                                 strWHERE = strWHERE + "not like '*" + strCriteriaParts(2) + "*' AND "
                            Case "EX" 'Is
                                strWHERE = strWHERE + "= '" + strCriteriaParts(2) + "' AND "
                            Case "XEX" 'Is not
                                strWHERE = strWHERE + "!= '" + strCriteriaParts(2) + "' AND "
                            Case "SW" 'Starts with
                                strWHERE = strWHERE + "like '" + strCriteriaParts(2) + "*' AND "
                            Case "XSW" 'Does not start with
                                 strWHERE = strWHERE + "not like '" + strCriteriaParts(2) + "*' AND "
                            Case "BF" 'Is before
                                strWHERE = strWHERE + "< #" + strCriteriaParts(2) + "# AND "
                            Case "OBF" 'Is on or before
                                strWHERE = strWHERE + "<= #" + strCriteriaParts(2) + "# AND "
                            Case "AF" 'Is after
                                strWHERE = strWHERE + "> #" + strCriteriaParts(2) + "# AND "
                            Case "OAF" 'Is on or after
                                strWHERE = strWHERE + ">= #" + strCriteriaParts(2) + "# AND "
                            Case "LT" 'Is less than
                                strWHERE = strWHERE + "< " + strCriteriaParts(2) + " AND "
                            Case "LTE" 'Is less than or equal to
                                strWHERE = strWHERE + "<= " + strCriteriaParts(2) + " AND "
                            Case "GT" 'Is greater than
                                strWHERE = strWHERE + "> " + strCriteriaParts(2) + " AND "
                            Case "GTE" 'Is greater than or equal to
                                strWHERE = strWHERE + ">= " + strCriteriaParts(2) + " AND "
                        End Select
                    Loop
                    strWHERE = Replace(strWHERE, "'_True_'", "True")
                    strWHERE = Replace(strWHERE, "'_False_'", "False")
                    'remove the trailing AND
                    strWHERE = Left(strWHERE, Len(strWHERE) - 4)
                    strSQL = strSQL + strWHERE
                    rst("WHERE") = strWHERE
                End If
                If ViewNodeList(i).selectNodes("qyslst").length > 0 Then
                    strSQL = strSQL + "ORDER BY "
                    strORDERBY = ""
                    'parse the slist in a loop
                    strFIDs = Split(ViewNodeList(i).selectSingleNode("qyslst").nodeTypedValue, ".")
                    For j = 0 To UBound(strFIDs)
                        strORDERBY = strORDERBY + "qdb_" + strDBID + ".[" + strFIDs(j) + "] "
                        If Mid(strOptions, j + InStr(strOptions, "sortorder-") + 10, 1) = "A" Then
                            strORDERBY = strORDERBY + "ASC, "
                        Else
                            strORDERBY = strORDERBY + "DESC, "
                        End If
                    Next j
                    strORDERBY = Trim(strORDERBY)
                    If Right(strORDERBY, 1) = "," Then
                        strORDERBY = Left(strORDERBY, Len(strORDERBY) - 1)
                    End If
                End If
                rst("ORDERBY") = strORDERBY
                rst("SQL") = Trim(strSQL + strORDERBY) + ";"
                rst.Update
            End If
        End If
    Next i
End Sub

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
   String, vValue As Variant) As Long
       Dim cch As Long
       Dim lrc As Long
       Dim lType As Long
       Dim lValue As Long
       Dim sValue As String

       On Error GoTo QueryValueExError

       ' Determine the size and type of data to be read
       lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
       If lrc <> ERROR_NONE Then Error 5

       Select Case lType
           ' For strings
           Case REG_SZ:
               sValue = String(cch, 0)

   lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
   sValue, cch)
               If lrc = ERROR_NONE Then
                   vValue = Left$(sValue, cch - 1)
               Else
                   vValue = Empty
               End If
           ' For DWORDS
           Case REG_DWORD:
   lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
   lValue, cch)
               If lrc = ERROR_NONE Then vValue = lValue
           Case Else
               'all other data types not supported
               lrc = -1
       End Select

QueryValueExExit:
       QueryValueEx = lrc
Exit Function

QueryValueExError:
       Resume QueryValueExExit
Exit Function
End Function

Function QueryValue(sKeyName As String, sValueName As String) As Variant
       Dim lRetVal As Long         'result of the API functions
       Dim hKey As Long         'handle of opened key
       Dim vValue As Variant      'setting of queried value

       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, _
   KEY_QUERY_VALUE, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       QueryValue = vValue
       RegCloseKey (hKey)

End Function
Function getCompleteCList(DBID As String) As String
Dim clist As String
Dim fieldcount As Integer
Dim strSQL As String
Dim rstFields As ADODB.Recordset

clist = ""
strSQL = "SELECT qdbFields.*, qdbFields.DBID FROM qdbFields WHERE ((qdbFields.DBID)='" + DBID + "');"
fieldcount = 0
Set rstFields = New ADODB.Recordset
rstFields.Open strSQL, CurrentProject.Connection, adOpenForwardOnly
Do While Not rstFields.EOF
    clist = clist + rstFields("FID").Value + "."
    fieldcount = fieldcount + 1
    rstFields.MoveNext
Loop
rstFields.Close
getCompleteCList = clist
End Function

Function getDateModifiedFID(DBID As String) As String
Dim strSQL As String
Dim rstRecords As ADODB.Recordset

strSQL = "SELECT qdbFields.FID FROM qdbFields WHERE (((qdbFields.DBID)='" + DBID + "') AND ((qdbFields.Role)='modified'));"
Set rstRecords = New ADODB.Recordset
rstRecords.Open strSQL, CurrentProject.Connection, adOpenForwardOnly
rstRecords.MoveFirst
getDateModifiedFID = rstRecords("FID")
rstRecords.Close
End Function

Function getServerChanges(DBID As String, qdb As QuickBaseClient, localDateModifiedBeforeSync As Date, DateModifiedOfLastEditedRecord As Date, maxRidBeforeSync As Integer, ridOfFirstNewRecord As Integer, strDateModifiedLabel As String, strRecordIDFID As String) As String

Dim clist As String
Dim fieldcounter As Integer
Dim recordcounter As Integer
Dim newRecordCounter As Integer
Dim editedRecordCounter As Integer
Dim RecordNodeList As MSXML.IXMLDOMNodeList
Dim FieldNodeList As MSXML.IXMLDOMNodeList
Dim CellNodeList As MSXML.IXMLDOMNodeList
Dim strFIDs() As String
Dim rst As ADODB.Recordset
Dim query As String
Dim strStart As String
Dim dateModified As Date
Dim currentRID As String

newRecordCounter = 0
editedRecordCounter = 0

clist = getCompleteCList(DBID)
Set rst = New ADODB.Recordset
rst.Open "qdb_" + DBID, CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTableDirect
rst.Index = strRecordIDFID
strStart = Format(localDateModifiedBeforeSync, "mm/dd/yyyy")
'Get all the records that changed since the last sync
'They're sorted from oldest to newest
query = "{'" + strDateModifiedLabel + "'.OAF.'" + strStart + "'}"
Set qdbResponse = qdb.DoQuery(DBID, query, clist, strDateModifiedLabel, "sortorder-A")
Set FieldNodeList = qdbResponse.documentElement.selectNodes("/*/table/fields/field")
ReDim strFIDs(FieldNodeList.length)
For fieldcounter = 0 To FieldNodeList.length - 1
    strFIDs(fieldcounter) = FieldNodeList(fieldcounter).selectSingleNode("@id").Text
Next fieldcounter
Set RecordNodeList = qdbResponse.documentElement.selectNodes("/*/table/records/record")
For recordcounter = 0 To RecordNodeList.length - 1
    Set CellNodeList = RecordNodeList(recordcounter).selectNodes("f")
    dateModified = qdb.int64ToDate(RecordNodeList(recordcounter).selectSingleNode("f[@id='" + strDateModifiedLabel + "']").Text)
    currentRID = RecordNodeList(recordcounter).selectSingleNode("f[@id='" + strRecordIDFID + "']").Text
    If dateModified <= DateModifiedOfLastEditedRecord And _
        (CInt(currentRID) < ridOfFirstNewRecord Or ridOfFirstNewRecord = 0) Then
        'We need to make this edit
        rst.Seek CInt(currentRID)
        If rst.EOF Then 'This is a new record so we need to add it
            rst.AddNew
            newRecordCounter = newRecordCounter + 1
        End If
        updateLocalRecord DBID, rst, CellNodeList, strFIDs(), FieldNodeList, RecordNodeList(recordcounter).selectSingleNode("f[@id='" + strDateModifiedLabel + "']").Text
        rst.Update
    End If
Next recordcounter
rst.Close
getServerChanges = "Local DB: " + CStr(editedRecordCounter) + " records edited, " + CStr(newRecordCounter) + " records added."
End Function

Function HideConflictScreen(DBID As String)
    Access.Forms("conflict_qdb_" + DBID).Visible = False
End Function

Function CloseConflictScreen(DBID As String)
    DoCmd.Close acForm, "conflict_qdb_" + DBID
End Function

Sub DisplayError()
Select Case Err.Number - vbObjectError
        Case 3
            MsgBox "We're sorry that we could not access that QuickBase." + vbCrLf + _
               "You do not have sufficent permissions to perform that operation." + vbCrLf + _
               "Please try again after requesting additional permission from the" + vbCrLf + _
               "owner or using a different e-mail address or screen name and password.", vbOKOnly, "QuickBase Desktop Insufficent Permission"

        Case 4
            MsgBox "We're sorry that we could not log you into QuickBase with the e-mail address or screen name you provided." + vbCrLf + _
               "Or it could be that you've mistyped your password." + vbCrLf + _
               "Please try again with a different e-mail address, screen name or password.", vbOKOnly, "QuickBase Desktop Login Problem"
        Case 32
            MsgBox "We're sorry that we could not access that QuickBase database." + vbCrLf + _
               "It could be that someone has deleted this database." + vbCrLf + _
               "Please try again with a different database name or click on the 'Refresh' button and try again.", vbOKOnly, "QuickBase Database not Found"
        Case Else
            MsgBox "QuickBase returned the following error:" + vbCrLf + Err.Description + vbCrLf + "Error Number: " + CStr(Err.Number - vbObjectError), vbOKOnly, "QuickBase Desktop Error " + CStr(Err.Number - vbObjectError)
    End Select
End Sub

Function DeleteLocalDB(strDBID As String) As Boolean

Dim cat As ADOX.Catalog
Dim QDBTable As ADOX.Table
Dim QDBView As ADOX.View
Dim rstRecords As ADODB.Recordset
Dim strSQL As String
Dim strQueryName As String

    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection
    
    strSQL = "SELECT qdbTables.Label FROM qdbTables WHERE (((qdbTables.DBID)='" + strDBID + "'));"
    Set rstRecords = New ADODB.Recordset
    rstRecords.Open strSQL, CurrentProject.Connection, adOpenForwardOnly
        strQueryName = rstRecords("Label")
    rstRecords.Close
    
    On Error Resume Next
        DoCmd.Close acForm, "frm_qdb_" + strDBID
        DoCmd.DeleteObject acForm, "frm_qdb_" + strDBID
        DoCmd.Close acForm, "conflict_qdb_" + strDBID
        DoCmd.DeleteObject acForm, "conflict_qdb_" + strDBID
        cat.Tables.Delete "qdb_" + strDBID
        cat.Tables.Delete "orig_" + strDBID
        cat.Tables.Delete strQueryName
    On Error GoTo 0

    'Need to walk through the forms collection to make sure the form is gone
    'Need to walk the tables collection also
    
    'Clean out the field descriptions in the qdbFields table
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL ("DELETE FROM qdbFields WHERE (((qdbFields.DBID)='" + strDBID + "'));")
    DoCmd.SetWarnings (True)
    
    
    'Clean out the view descriptions in the qdbViews table
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL ("DELETE FROM qdbViews WHERE (((qdbViews.DBID)='" + strDBID + "'));")
    DoCmd.SetWarnings (True)
    
    
    'Delete the querys
    
    For Each QDBView In cat.Views
        If strQueryName = QDBView.Name Then
            cat.Views.Delete strQueryName
        End If
    Next QDBView
    
    DeleteLocalDB = True
End Function
Sub GetGrantedDBs(strUsername As String, strPassword As String)
    Dim DBNodeList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim rst As ADODB.Recordset
    Dim i As Integer
    Dim strDBID As String
    Dim cat As ADOX.Catalog
    Dim QDBTable As ADOX.Table
    
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection
    
    i = qdb.Authenticate(strUsername, strPassword)
 
    Set qdbResponse = qdb.GetGrantedDBs()
    Set DBNodeList = qdbResponse.documentElement.selectNodes("/*/databases/dbinfo")
    
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL ("DELETE FROM qdbTables WHERE qdbTables.isLocal=False;")
    DoCmd.SetWarnings (True)
    Set rst = New ADODB.Recordset
    rst.Open "qdbTables", CurrentProject.Connection, adOpenForwardOnly, adLockOptimistic
    For Each xmlNode In DBNodeList
        rst.AddNew
        rst("Label") = xmlNode.selectSingleNode("dbname").nodeTypedValue
        rst("DBID") = xmlNode.selectSingleNode("dbid").nodeTypedValue
        For Each QDBTable In cat.Tables
            If "qdb_" + rst("DBID") = QDBTable.Name Then
                rst("isLocal") = True
                Call setTableDescription(QDBTable.Name, rst("Label"))
            End If
        Next QDBTable
        On Error Resume Next
        Err.Clear
        rst.Update
        If Err.Number <> 0 Then
            rst.CancelUpdate
        End If
    Next xmlNode
    Exit Sub
End Sub
Sub makeQueryForTable(DBID As String, DBName As String, strSQL As String)
Dim db As DAO.Database
Dim qdf As DAO.QueryDef
Dim prpDescription As DAO.Property
Dim QueryTitle As String


QueryTitle = makeMSAccessObjectName(DBName)
Set db = CurrentDb()
On Error Resume Next
    db.QueryDefs.Delete QueryTitle
On Error GoTo 0
Set qdf = db.CreateQueryDef(QueryTitle, strSQL)
End Sub

Sub prototype()




    DoCmd.OpenQuery "All Field Types", acNormal, acEdit




End Sub

Function RefreshFields(DBID As String, QDBRecord As Variant) As Boolean
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim frmDataSheet As Form
    Dim QDBRecordCopy() As Variant

    'We have to lock the deleted field on the data manipulation form

    Call DoCmd.OpenForm("frm_qdb_" + DBID, acDesign, "", "", acFormPropertySettings, acHidden)

    Set frmDataSheet = Access.Application.Forms("frm_qdb_" + DBID)
    Set qdbResponse = qdb.GetSchema(DBID)
    Set rst = New ADODB.Recordset
    strSQL = "SELECT qdbFields.* FROM qdbFields WHERE (((qdbFields.DBID)='" + DBID + "'));"
    
    rst.Open strSQL, CurrentProject.Connection, adOpenForwardOnly, adLockOptimistic

    Do While Not rst.EOF
        If qdbResponse.documentElement.selectNodes("/*/table/fields/field[@id='" + rst("FID") + "']/@id").length = 0 Then
            'This field does not exist any longer
            rst("Deleted") = True
            rst.Update
            frmDataSheet.Controls(rst("Label")).Locked = True
        End If
        rst.MoveNext
    Loop
    rst.Close
    DoCmd.Close acForm, "frm_qdb_" + DBID, acSaveYes
    'Now we need to remove the field from the array so we can successfully make the edit
    
    Dim lastfield As Integer
    Dim firstfield As Integer
    Dim j As Integer
        
    
    lastfield = UBound(QDBRecord, 2)
    firstfield = LBound(QDBRecord, 2)
    ReDim QDBRecordCopy(1, lastfield)
    j = 0
    For i = firstfield To lastfield
        If qdbResponse.documentElement.selectNodes("/*/table/fields/field[@id='" + CStr(QDBRecord(0, i)) + "']/@id").length <> 0 Then
            QDBRecordCopy(0, j) = QDBRecord(0, i)
            QDBRecordCopy(1, j) = QDBRecord(1, i)
            j = j + 1
        End If
    Next i
    If j = 0 Then
        RefreshFields = False
        Exit Function
    Else
        ReDim Preserve QDBRecordCopy(1, j - 1)
        QDBRecord = QDBRecordCopy
        RefreshFields = True
    End If
    
End Function

Sub RegisterTable(strDBID As String, DBName As String)
    Dim rst As ADODB.Recordset
    Dim cat As ADOX.Catalog
    Dim QDBTable As ADOX.Table
    
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection
    Set rst = New ADODB.Recordset
    rst.Open "qdbTables", CurrentProject.Connection, adOpenForwardOnly, adLockOptimistic
    rst.Find ("DBID = '" + strDBID + "'")
    If rst.EOF Then
        rst.AddNew
        rst("Label") = DBName
        rst("DBID") = strDBID
        For Each QDBTable In cat.Tables
            If "qdb_" + rst("DBID") = QDBTable.Name Then
                rst("isLocal") = True
            End If
        Next QDBTable
        On Error Resume Next
        Err.Clear
        rst.Update
        If Err.Number <> 0 Then
            rst.CancelUpdate
        End If
    Else
        rst("isLocal") = True
    End If
    Call setTableDescription("qdb_" + rst("DBID"), DBName)
    rst.Update
    rst.Close
End Sub
Function getDBIDbyName(strUsername As String, strPassword As String, strDBName As String)
    Dim i As Integer
    
    i = qdb.Authenticate(strUsername, strPassword)
    getDBIDbyName = qdb.FindDBByName(strDBName)
End Function
Function mirrorQDBTable(strDBID As String, strPrefix As String, strUsername As String, strPassword As String) As String
    
    Dim NewTable As ADOX.Table
    Dim strFIDs() As String

    Dim fldField As ADOX.Column
    Dim clist As String
    Dim RecordNodeList As MSXML.IXMLDOMNodeList
    Dim FieldNodeList As MSXML.IXMLDOMNodeList
    Dim RecordIDFID As String
    Dim rst As ADODB.Recordset
   
    
    Dim i As Integer
    Dim j As Integer
    Dim strFID As String
    Dim DBName As String
    Dim cat As ADOX.Catalog
    
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection
    
    i = qdb.Authenticate(strUsername, strPassword)
    
    For Each NewTable In cat.Tables
        If strPrefix + strDBID = NewTable.Name Then
            cat.Tables.Delete strPrefix + strDBID
        End If
    Next NewTable
    
    Set NewTable = New ADOX.Table
    
    NewTable.Name = strPrefix + strDBID
    
    On Error GoTo noschema
    Set qdbResponse = qdb.GetSchema(strDBID)
    On Error GoTo 0
    DBName = qdbResponse.documentElement.selectSingleNode("/*/table/name").Text
    Set FieldNodeList = qdbResponse.documentElement.selectNodes("/*/table/fields/field")
    clist = UpdateFieldsTable(strDBID, FieldNodeList, makeMSAccessObjectName(DBName))
    ReDim strFIDs(FieldNodeList.length)
    For i = 0 To FieldNodeList.length - 1
        Set fldField = New ADOX.Column
        strFIDs(i) = FieldNodeList(i).selectSingleNode("@id").Text
        fldField.Name = strFIDs(i)
        fldField.Attributes = adColNullable
        Select Case FieldNodeList(i).selectSingleNode("@base_type").Text
            Case "float"
                fldField.Type = adDouble
            Case "text"
                fldField.Type = adLongVarWChar
            Case "bool"
                fldField.Attributes = fldField.Attributes And Not adColNullable
                fldField.Type = adBoolean
            Case "int64"
                fldField.Type = adDate
            Case "int32"
                Select Case FieldNodeList(i).selectSingleNode("@field_type").Text
                    Case "recordid"
                        fldField.Type = adInteger
                        RecordIDFID = strFIDs(i)
                    Case Else
                        fldField.Type = adInteger
                End Select
        End Select
        NewTable.Columns.Append fldField
    Next i
    Set fldField = New ADOX.Column
    fldField.Name = UPDATE_ID_FIELDNAME
    fldField.Type = adVarWChar
    fldField.Attributes = adColNullable
    NewTable.Columns.Append fldField
    'Below describes bug
    'http://support.microsoft.com/directory/article.asp?ID=KB;EN-US;Q272001
    cat.Tables.Append NewTable
    For i = 0 To FieldNodeList.length - 1
        Select Case FieldNodeList(i).selectSingleNode("@base_type").Text
            Case "text"
                Select Case FieldNodeList(i).selectSingleNode("@field_type").Text
                    Case "file", "url"
                        setHyperLinkField "qdb_" + strDBID, FieldNodeList(i).selectSingleNode("@id").Text
                    Case Else
                End Select
                NewTable.Columns.Item(FieldNodeList(i).selectSingleNode("@id").Text).Properties("Jet OLEDB:Allow Zero Length").Value = True
        End Select
        NewTable.Columns.Item(FieldNodeList(i).selectSingleNode("@id").Text).Properties("Description").Value = FieldNodeList(i).selectSingleNode("label").Text
    Next i
    updateViewsTable strDBID
    Set rst = New ADODB.Recordset
    rst.Open strPrefix + strDBID, CurrentProject.Connection, adOpenForwardOnly, adLockPessimistic
    Set qdbResponse = qdb.DoQuery(strDBID, "{'0'.CT.''}", clist, "", "")
    Set RecordNodeList = qdbResponse.documentElement.selectNodes("/*/table/records/record")
    Call SysCmd(acSysCmdInitMeter, "Downloading " + strDBID, RecordNodeList.length)
    For i = 0 To RecordNodeList.length - 1
        rst.AddNew
            Call updateLocalRecord(strDBID, rst, RecordNodeList(i).selectNodes("f"), strFIDs(), FieldNodeList, RecordNodeList(i).selectSingleNode("f[@id=/*/table/fields/field[@role='modified']/@id]").Text)
        rst.Update
        Call SysCmd(acSysCmdUpdateMeter, i)
    Next i
    rst.Close
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL ("CREATE UNIQUE INDEX " + RecordIDFID + " ON " + strPrefix + strDBID + " (" + RecordIDFID + ") WITH IGNORE NULL;")
    DoCmd.SetWarnings (True)
    Call RegisterTable(strDBID, DBName)
    mirrorQDBTable = strPrefix + strDBID
    Call SysCmd(acSysCmdRemoveMeter)
    Exit Function
noschema:


If Err.Number = vbObjectError + 1 Then
    Err.Raise vbObjectError + 1, "QuickBase Desktop", "Sorry we could not make a local copy of your database. You probably don't have full access to restriced fields."
Else
    Err.Raise vbObjectError + 1, "QuickBase Desktop", Err.Description
End If
    Exit Function
End Function

Sub copyTable(tblName As String, newTableName As String)

Dim cat As ADOX.Catalog
Set cat = New ADOX.Catalog
cat.ActiveConnection = CurrentProject.Connection
    

    On Error Resume Next
        cat.Tables.Delete newTableName
    On Error GoTo 0
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL ("SELECT [" + tblName + "].* INTO [" + newTableName + "] FROM " + tblName + ";")
    DoCmd.SetWarnings (True)
End Sub
Function getRecordIDFID(DBID As String) As String
Dim rstRecords As ADODB.Recordset
Dim strSQL As String

Set rstRecords = New ADODB.Recordset
strSQL = "SELECT qdbFields.*, qdbFields.DBID FROM qdbFields WHERE (((qdbFields.DBID)='" + DBID + "'));"
rstRecords.Open strSQL, CurrentProject.Connection, adOpenForwardOnly
Do While Not rstRecords.EOF
    If rstRecords("Field Type") = "recordid" Then
        getRecordIDFID = rstRecords("FID")
        rstRecords.Close
        Set rstRecords = Nothing
        Exit Function
    End If
    rstRecords.MoveNext
Loop
rstRecords.Close
Set rstRecords = Nothing
End Function

Sub DeleteRecords(DBID As String, strUsername As String, strPassword As String)
Dim strSQL As String
Dim rstRecords As ADODB.Recordset
Dim strRID As String
Dim strRecordIDFID As String
Dim i As Integer
Dim deletedRecords As Integer



Set rstRecords = New ADODB.Recordset
strRecordIDFID = getRecordIDFID(DBID)
i = qdb.Authenticate(strUsername, strPassword)


strSQL = "SELECT  [orig_" + DBID + "].[" + strRecordIDFID + "] FROM orig_" + DBID + " LEFT JOIN qdb_" + DBID + " ON [orig_" + DBID + "].[" + strRecordIDFID + "] = [qdb_" + DBID + "].[" + strRecordIDFID + "] WHERE ((([qdb_" + DBID + "].[" + strRecordIDFID + "]) Is Null) AND (([orig_" + DBID + "].[" + strRecordIDFID + "]) Is Not Null));"
rstRecords.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockPessimistic
deletedRecords = rstRecords.RecordCount
If deletedRecords > 0 Then
    If MsgBox("Are you sure you want to delete " + CStr(deletedRecords) + " records from www.quickbase.com?", vbOKCancel + vbQuestion, "QuickBase Delete Confirmation") <> vbOK Then
        rstRecords.Close
        Exit Sub
    End If
End If
deletedRecords = 0
Do While Not rstRecords.EOF
    'Delete the record from QuickBase
    On Error Resume Next
    Err.Clear
    strRID = qdb.DeleteRecord(DBID, rstRecords(strRecordIDFID))
    If Err.Number = 0 Then
        deletedRecords = deletedRecords + 1
    ElseIf Err.Number = 30 + vbObjectError Then
        'The record is already gone
    End If
    On Error GoTo 0
    rstRecords.Delete
    rstRecords.MoveNext
Loop
rstRecords.Close
End Sub
Function ResolveConflict(qdb As QuickBaseClient, DBID As String, strRecordIDFID As String, rstOriginalRecords As ADODB.Recordset, rstModifiedRecords As ADODB.Recordset, QDBRecord() As Variant, ByRef QDBCurrentRecord() As Variant)

Dim fieldcounter As Integer
Dim resolvingConflict As Boolean
Dim lngLoop As Long
Dim strFieldName As String
Dim clist As String
Dim strCurrentUpdateID As String
Dim CurrentRecord As Collection
Dim ConflictingFIDs As Collection



DoEvents
resolvingConflict = True
lngLoop = 0

clist = getCompleteCList(DBID)


'Let's get the current record

strCurrentUpdateID = qdb.getRecordAsArray(DBID, clist, strRecordIDFID, rstModifiedRecords(strRecordIDFID), QDBCurrentRecord())

'Let's make a collection out of it
Set CurrentRecord = New Collection

For fieldcounter = 0 To UBound(QDBCurrentRecord, 2)
    CurrentRecord.Add Key:=QDBCurrentRecord(0, fieldcounter), Item:=QDBCurrentRecord(1, fieldcounter)
Next fieldcounter

'Check to see if there is a conflict first
Set ConflictingFIDs = New Collection
For fieldcounter = 0 To UBound(QDBRecord, 2)
    'Check all the pairs
    If rstOriginalRecords(CStr(QDBRecord(0, fieldcounter))) = CurrentRecord(CStr(QDBRecord(0, fieldcounter))) _
        Or (IsNull(rstOriginalRecords(CStr(QDBRecord(0, fieldcounter)))) _
            And IsNull(CurrentRecord(CStr(QDBRecord(0, fieldcounter))))) Then
        'If we're here we do not have a record conflict!
    ElseIf rstModifiedRecords(CStr(QDBRecord(0, fieldcounter))) = CurrentRecord(CStr(QDBRecord(0, fieldcounter))) _
        Or (IsNull(rstModifiedRecords(CStr(QDBRecord(0, fieldcounter)))) _
            And IsNull(CurrentRecord(CStr(QDBRecord(0, fieldcounter))))) Then
            'If we're here we do not have a record conflict!
    ElseIf rstModifiedRecords(CStr(QDBRecord(0, fieldcounter))) = rstOriginalRecords(CStr(QDBRecord(0, fieldcounter))) _
        Or (IsNull(rstModifiedRecords(CStr(QDBRecord(0, fieldcounter)))) _
            And IsNull(rstOriginalRecords(CStr(QDBRecord(0, fieldcounter))))) Then
            'If we're here we do not have a record conflict!
    Else
        'We have a record conflict
        'So let's record the FIDs in a collection
        ConflictingFIDs.Add Item:=CStr(QDBRecord(0, fieldcounter)), Key:=CStr(QDBRecord(0, fieldcounter))
    End If
Next fieldcounter

If ConflictingFIDs.Count = 0 Then 'There is no conflict!
    ResolveConflict = strCurrentUpdateID
    Exit Function
End If
MsgBox "We'll show you what you had when you last synched," + vbCrLf + _
        "your changes and what was changed up at www.quickbase.com", vbOKOnly, "Replication Conflict"
Call DoCmd.OpenForm("conflict_qdb_" + DBID, acNormal, "", "", acFormPropertySettings, acHidden)


'Set all the control values here


For fieldcounter = 0 To rstOriginalRecords.Fields.Count - 1
    strFieldName = rstOriginalRecords(fieldcounter).Name
    strFieldName = Mid(strFieldName, InStr(strFieldName, ".") + 1)
    If strFieldName <> UPDATE_ID_FIELDNAME Then
        Access.Forms("conflict_qdb_" + DBID).Controls("orig_" + strFieldName).Value = rstOriginalRecords(fieldcounter)
    End If
Next fieldcounter
For fieldcounter = 0 To rstModifiedRecords.Fields.Count - 1
    strFieldName = rstModifiedRecords(fieldcounter).Name
    strFieldName = Mid(strFieldName, InStr(strFieldName, ".") + 1)
    If strFieldName <> UPDATE_ID_FIELDNAME Then
        Access.Forms("conflict_qdb_" + DBID).Controls("proposed_" + strFieldName).Value = rstModifiedRecords(fieldcounter)
    End If
Next fieldcounter
For fieldcounter = 0 To UBound(QDBCurrentRecord, 2)
    Access.Forms("conflict_qdb_" + DBID).Controls("current_" + QDBCurrentRecord(0, fieldcounter)).Value = QDBCurrentRecord(1, fieldcounter)
Next fieldcounter
For fieldcounter = 1 To ConflictingFIDs.Count
        Access.Forms("conflict_qdb_" + DBID).Controls("proposed_" + ConflictingFIDs(fieldcounter)).Locked = False
        Access.Forms("conflict_qdb_" + DBID).Controls("proposed_" + ConflictingFIDs(fieldcounter)).BackColor = 16777215
Next fieldcounter

On Error GoTo 0
'Now make the form visible
DoCmd.Hourglass False
Access.Forms("conflict_qdb_" + DBID).Visible = True
'loop and wait until the form is closed or hidden

Do While resolvingConflict
    DoEvents
    lngLoop = lngLoop + 1
    If lngLoop > 1000 Then
        lngLoop = 0
        If SysCmd(acSysCmdGetObjectState, acForm, "conflict_qdb_" + DBID) = 0 Then
            'form is no longer loaded
            strCurrentUpdateID = "-" + strCurrentUpdateID
            resolvingConflict = False
        ElseIf Access.Forms("conflict_qdb_" + DBID).Visible = False Then
            'form is no longer visible
            resolvingConflict = False
            'Need to pull out the changes before we close the form
            For fieldcounter = 0 To UBound(QDBRecord, 2)
                QDBRecord(1, fieldcounter) = Access.Forms("conflict_qdb_" + DBID).Controls("proposed_" + CStr(QDBRecord(0, fieldcounter))).Value
            Next fieldcounter
            
            DoCmd.Close acForm, "conflict_qdb_" + DBID
        End If
    End If
Loop
DoCmd.Hourglass True
ResolveConflict = strCurrentUpdateID
End Function

Sub downloadAllAttachedFiles(DBID As String, strUsername As String, strPassword As String)
Dim rstFields As ADODB.Recordset
Dim rstModifiedRecords As ADODB.Recordset
Dim strSQL As String
Dim strFieldName As String
Dim strFileName As String
Dim strTempFilename As String
Dim i As Integer

Call qdb.Authenticate(strUsername, strPassword)
Call qdb.GetDBInfo(DBID)

Set rstFields = New ADODB.Recordset
strSQL = "SELECT qdbFields.FID FROM qdbFields WHERE (((qdbFields.DBID)='" + DBID + "') AND ((qdbFields.Deleted)=False) AND ((qdbFields.[Field Type]) = 'File'));"
rstFields.Open strSQL, CurrentProject.Connection, adOpenKeyset

Set rstModifiedRecords = New ADODB.Recordset
strSQL = "SELECT [qdb_" + DBID + "].* FROM qdb_" + DBID + ";"
rstModifiedRecords.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
i = SysCmd(acSysCmdInitMeter, "Downloading files from " + DBID, rstModifiedRecords.RecordCount)
On Error Resume Next
    MkDir Application.CurrentProject.Path + "\" + DBID
On Error GoTo 0
i = 0
Do While Not rstModifiedRecords.EOF
    rstFields.Requery
    rstFields.MoveFirst
    Do While Not rstFields.EOF
        strFieldName = rstFields(0)
        If IsNull(rstModifiedRecords(strFieldName)) Then
            
        Else
            strFileName = rstModifiedRecords(strFieldName)
            If InStr(strFileName, "file://") Then
                strFileName = Left(strFileName, InStr(strFileName, "#") - 1)
                downloadAttachedFile DBID, strFileName
            End If
        End If
        rstFields.MoveNext
    Loop
    rstModifiedRecords.MoveNext
    Call SysCmd(acSysCmdUpdateMeter, i)
    i = i + 1
Loop
Call SysCmd(acSysCmdRemoveMeter)
End Sub

Sub updateLocalRecord(strDBID As String, rst As ADODB.Recordset, CellNodeList As MSXML.IXMLDOMNodeList, strFIDs() As String, FieldNodeList As MSXML.IXMLDOMNodeList, UpdateID As String)

Dim j As Integer
Dim strFileName As String

For j = 0 To CellNodeList.length - 1
            Select Case rst(strFIDs(j)).Type
                Case adDate
                    If CellNodeList(j).nodeTypedValue <> "" Then
                        rst(strFIDs(j)) = qdb.int64ToDate(CellNodeList(j).nodeTypedValue) '+ (MinutesGMTOffset / (60 * 24))
                    End If
                Case adDouble
                    If CellNodeList(j).nodeTypedValue = "" Then
                        rst(strFIDs(j)) = Null
                    Else
                        rst(strFIDs(j)) = CDbl(CellNodeList(j).nodeTypedValue)
                    End If
                Case Else
                    'Somehow we have to determine if this is a URL of file field
                    Select Case FieldNodeList(j).selectSingleNode("@field_type").Text
                        Case "file"
                            If CellNodeList(j).nodeTypedValue <> "" Then
                                If InStr(CellNodeList(j).nodeTypedValue, ":") > 0 Then
                                    'We have a URL
                                    rst(strFIDs(j)) = CellNodeList(j).nodeTypedValue + "#" + CellNodeList(j).nodeTypedValue + "##"
                                Else
                                    'We have a file
                                    rst(strFIDs(j)) = CellNodeList(j).nodeTypedValue + "#file://" + Application.CurrentProject.Path + "\" + strDBID + "\" + CellNodeList(j).nodeTypedValue + "##"
                                    'Need to download the file just in case it has changed
                                    downloadAttachedFile strDBID, CellNodeList(j).nodeTypedValue
                                End If
                            End If
                        Case "url"
                            rst(strFIDs(j)) = CellNodeList(j).nodeTypedValue + "#" + CellNodeList(j).nodeTypedValue + "##"
                        Case Else
                            rst(strFIDs(j)) = Replace(CellNodeList(j).nodeTypedValue, Chr(10), vbCrLf)
                    End Select
            End Select
        Next j
        rst(UPDATE_ID_FIELDNAME) = UpdateID
End Sub

Sub UpdateRecords(DBID As String, strUsername As String, strPassword As String)

Dim strSQL As String
Dim rstOriginalRecords As ADODB.Recordset
Dim rstModifiedRecords As ADODB.Recordset
Dim rstFields As ADODB.Recordset
Dim fldLoop As ADODB.Field
Dim QDBRecord() As Variant
Dim QDBCurrentRecord() As Variant

Dim strRID As String
Dim strUpdateID As String
Dim strFieldName As String
Dim varOriginalValue As Variant
Dim varProposedValue As Variant

Dim strRecordIDFID As String
Dim strDateModifiedLabel As String
Dim i As Integer
Dim isModified As Boolean
Dim newDateModified As Date

Dim localDBLastModified As Date
Dim remoteLastModified As Date
Dim localDateModifiedBeforeSync As Date
Dim maxRidBeforeSync As Integer
Dim ridOfFirstNewRecord As Integer
Dim DateModifiedOfLastEditedRecord As Date
Dim dirtyDB As Boolean
Dim countRecordsAdded As Integer
Dim countRecordsEdited As Integer
Dim countCellsEdited As Integer
Dim countRecordsDeleted As Integer


Call SysCmd(acSysCmdInitMeter, "Syncing " + DBID, 100)
            
ridOfFirstNewRecord = 0
DateModifiedOfLastEditedRecord = 0
strRecordIDFID = getRecordIDFID(DBID)
strDateModifiedLabel = getDateModifiedFID(DBID)

Call SysCmd(acSysCmdUpdateMeter, 10)

i = qdb.Authenticate(strUsername, strPassword)


'Get the last record modified time of the database.
remoteLastModified = qdb.getDBLastModified(DBID)

maxRidBeforeSync = getLocalDBMaxRecordID(DBID)
localDateModifiedBeforeSync = getLocalDBLastModifiedDate(DBID)


If remoteLastModified <= localDateModifiedBeforeSync Then
    'The database at www.quickbase.com has not been modified since we last synched
    'so we don't have to look for new or modified records.
    dirtyDB = False
Else
    'The table at www.quickbase.com has been modified
    'We need to remember the two times here.
    'We can use this information later to retrieve all the changes that were made
    'up on www.quickbase.com since we last synched
    'but first we're going to push the changes from the local copy to the
    'quickbase server
    dirtyDB = True
End If

Set rstFields = New ADODB.Recordset
strSQL = "SELECT qdbFields.FID FROM qdbFields WHERE (((qdbFields.DBID)='" + DBID + "') AND ((qdbFields.Deleted)=False) AND ((qdbFields.Role) IS Null) AND ((qdbFields.Mode) IS Null));"
rstFields.Open strSQL, CurrentProject.Connection, adOpenKeyset


Set rstOriginalRecords = New ADODB.Recordset
Set rstModifiedRecords = New ADODB.Recordset
'Let's join the snapshot table we've kept around since we last synched with the working version
'This is an outer join that will display all the working table records
strSQL = "SELECT [orig_" + DBID + "].* FROM orig_" + DBID + " RIGHT JOIN qdb_" + DBID + " ON [orig_" + DBID + "].[" + strRecordIDFID + "] = [qdb_" + DBID + "].[" + strRecordIDFID + "] ORDER BY [qdb_" + DBID + "].[" + strRecordIDFID + "];"
rstOriginalRecords.Open strSQL, CurrentProject.Connection, adOpenForwardOnly, adLockPessimistic

strSQL = "SELECT [qdb_" + DBID + "].* FROM orig_" + DBID + " RIGHT JOIN qdb_" + DBID + " ON [orig_" + DBID + "].[" + strRecordIDFID + "] = [qdb_" + DBID + "].[" + strRecordIDFID + "] ORDER BY [qdb_" + DBID + "].[" + strRecordIDFID + "];"
rstModifiedRecords.Open strSQL, CurrentProject.Connection, adOpenForwardOnly, adLockPessimistic

Call SysCmd(acSysCmdUpdateMeter, 30)
Do While Not rstOriginalRecords.EOF
    If IsNull(rstOriginalRecords(strRecordIDFID)) And IsNull(rstModifiedRecords(strRecordIDFID)) Then
        'This is a new record and has to be added!
        i = 0
        
        ReDim QDBRecord(0, 0)
        ReDim QDBRecord(1, rstFields.RecordCount)
        rstFields.Requery
        rstFields.MoveFirst
        Do While Not rstFields.EOF
            strFieldName = rstFields(0)
            QDBRecord(0, i) = CInt(strFieldName)
            On Error Resume Next
            QDBRecord(1, i) = rstModifiedRecords(strFieldName)
            rstOriginalRecords(strFieldName) = QDBRecord(1, i)
            On Error GoTo 0
            rstFields.MoveNext
            i = i + 1
        Loop
        ReDim Preserve QDBRecord(1, i - 1)
        'Add the record to the QuickBase
        Err.Clear
        On Error Resume Next
           strRID = qdb.AddRecordByArray(DBID, strUpdateID, QDBRecord())
        If Err.Number = 0 Then
            On Error GoTo 0
            countRecordsAdded = countRecordsAdded + 1
            rstModifiedRecords(UPDATE_ID_FIELDNAME) = strUpdateID
            rstModifiedRecords(strRecordIDFID) = strRID
            rstModifiedRecords.Update
            If ridOfFirstNewRecord = 0 Then
                ridOfFirstNewRecord = CInt(strRID)
            End If
        Else
            Select Case Err.Number
                Case 2 + vbObjectError 'Invalid input caused by a choice set
                                       'on a multiple choice field that is not in that QuickBase's drop down list
                    MsgBox "Sorry your record was not added." + vbCrLf + _
                    "You added a record with a multiple choice field set to a value that is no longer valid in this particular QuickBase. Please select a value in the local copy of your database that is valid or change the field properties of your QuickBase to accomodate the new choice. You can eliminate this problem by setting 'Allow users to add new choices' on all you multiple choice fields to true.", _
                    vbOKOnly, "QuickBase Desktop"
                Case Else
                    MsgBox "Sorry your record was not added." + vbCrLf + qdb.errortext, vbOKOnly, "QuickBase Desktop"
            End Select
            On Error GoTo 0
            rstModifiedRecords.CancelUpdate
            rstOriginalRecords.CancelUpdate
        End If
    ElseIf Not IsNull(rstOriginalRecords(strRecordIDFID)) Then
        'Check to see if the record has been modified
        isModified = False
        i = 0
        ReDim QDBRecord(0, 0)
        ReDim QDBRecord(1, rstFields.RecordCount)
        rstFields.Requery
        rstFields.MoveFirst
        Do While Not rstFields.EOF
            strFieldName = rstFields(0)
            varOriginalValue = rstOriginalRecords(strFieldName)
            varProposedValue = rstModifiedRecords(strFieldName)
            If IsNull(varProposedValue) And IsNull(varOriginalValue) Then
                'No change
            ElseIf IsNull(varProposedValue) Or IsNull(varOriginalValue) _
            Or (varProposedValue <> varOriginalValue) Then
                isModified = True
                QDBRecord(0, i) = CInt(strFieldName)
                If IsNull(varProposedValue) Then
                    QDBRecord(1, i) = ""
                Else
                    QDBRecord(1, i) = varProposedValue
                End If
                'change the original copy to indicate that we don't need to do that edit again
                rstOriginalRecords(strFieldName) = varProposedValue
                countCellsEdited = countCellsEdited + 1
                i = i + 1
            End If
            rstFields.MoveNext
        Loop
        If isModified Then
            'Edit the record in QuickBase
            ReDim Preserve QDBRecord(1, i - 1)
            Err.Number = -1
            Do While Err.Number <> 0
                Err.Clear
                On Error Resume Next
                If isModified Then
                    strUpdateID = qdb.EditRecordByArray(DBID, rstOriginalRecords(strRecordIDFID), rstOriginalRecords(UPDATE_ID_FIELDNAME), QDBRecord())
                    Select Case Err.Number
                        Case 0
                            On Error GoTo 0
                            'Modify the record in the orig table to match the modified record
                            DateModifiedOfLastEditedRecord = qdb.int64ToDate(strUpdateID)
                            rstOriginalRecords(UPDATE_ID_FIELDNAME) = strUpdateID
                            rstModifiedRecords(UPDATE_ID_FIELDNAME) = strUpdateID
                            newDateModified = qdb.int64ToDate(strUpdateID)
                            rstOriginalRecords(strDateModifiedLabel) = newDateModified
                            rstModifiedRecords(strDateModifiedLabel) = newDateModified
                            countRecordsEdited = countRecordsEdited + 1
                            rstModifiedRecords.Update
                            rstOriginalRecords.Update
                        Case 31 + vbObjectError 'No such field
                            'We need to refresh the schema!
                            isModified = RefreshFields(DBID, QDBRecord())
                            Err.Number = -1
                        Case 60 + vbObjectError
                            On Error GoTo 0
                            'We have a replication conflict here!
                            'Undo the local changes
                            rstOriginalRecords.CancelUpdate
                            rstModifiedRecords.CancelUpdate
                            strUpdateID = ""
                            Do While strUpdateID = ""
                                strUpdateID = ResolveConflict(qdb, DBID, strRecordIDFID, rstOriginalRecords, rstModifiedRecords, QDBRecord(), QDBCurrentRecord())
                                
                                If Left(strUpdateID, 1) = "-" Then Exit Do
                                On Error Resume Next
                                Err.Clear
                                strUpdateID = qdb.EditRecordByArray(DBID, rstOriginalRecords(strRecordIDFID), strUpdateID, QDBRecord())
                                If Err.Number <> 0 Then
                                    strUpdateID = ""
                                End If
                                On Error GoTo 0
                            Loop
                            For i = 0 To UBound(QDBCurrentRecord, 2)
                                rstOriginalRecords(CStr(QDBCurrentRecord(0, i))) = QDBCurrentRecord(1, i)
                                rstModifiedRecords(CStr(QDBCurrentRecord(0, i))) = QDBCurrentRecord(1, i)
                            Next i
                            If Left(strUpdateID, 1) <> "-" Then
                                'The user did not cancel out the changes
                                For i = 0 To UBound(QDBRecord, 2)
                                    rstOriginalRecords(CStr(QDBRecord(0, i))) = QDBRecord(1, i)
                                    rstModifiedRecords(CStr(QDBRecord(0, i))) = QDBRecord(1, i)
                                Next i
                                strUpdateID = Mid(strUpdateID, 2)
                            End If
                            rstOriginalRecords(UPDATE_ID_FIELDNAME) = strUpdateID
                            rstModifiedRecords(UPDATE_ID_FIELDNAME) = strUpdateID
                            newDateModified = qdb.int64ToDate(strUpdateID)
                            rstOriginalRecords(strDateModifiedLabel) = newDateModified
                            rstModifiedRecords(strDateModifiedLabel) = newDateModified
                            countRecordsEdited = countRecordsEdited + 1
                            rstOriginalRecords.Update
                            rstModifiedRecords.Update
                        Case Else
                            On Error GoTo 0
                            'Don't modify the original table to indicate that we did not successfully edit the www.quickbase.com copy
                            rstOriginalRecords.CancelUpdate
                            rstModifiedRecords.CancelUpdate
                    End Select
                End If
            Loop
        End If
    End If
    rstOriginalRecords.MoveNext
    rstModifiedRecords.MoveNext
Loop
rstOriginalRecords.Close
rstModifiedRecords.Close
If dirtyDB Then
    If DateModifiedOfLastEditedRecord = 0 Then
        DateModifiedOfLastEditedRecord = remoteLastModified
    End If
    Dim strLocalChanges As String
    strLocalChanges = getServerChanges(DBID, qdb, localDateModifiedBeforeSync, DateModifiedOfLastEditedRecord, maxRidBeforeSync, ridOfFirstNewRecord, strDateModifiedLabel, strRecordIDFID)
End If
Call SysCmd(acSysCmdUpdateMeter, 100)
Call copyTable("qdb_" + DBID, "orig_" + DBID)
Call SysCmd(acSysCmdRemoveMeter)
Dim SysCmdMessage As String
If dirtyDB Then
    SysCmdMessage = "Server DB was modified since you last synched, "
Else
    SysCmdMessage = "Server DB was not modified since you last synched, "
End If
SysCmdMessage = strLocalChanges + " " + SysCmdMessage + CStr(countRecordsAdded) + " records were added, " + CStr(countCellsEdited) + " cells were edited, " + CStr(countRecordsEdited) + " records were edited."
Call SysCmd(acSysCmdSetStatus, SysCmdMessage)
MsgBox SysCmdMessage, vbOKOnly, "QuickBase Desktop Synch Complete"
Exit Sub
End Sub

Sub printUniqueASCIICharacters(strVariant As Variant)
Dim i As Integer
For i = 0 To 255
If InStr(strVariant, Chr(i)) > 0 Then
    Debug.Print i
End If
Next i
End Sub

Sub BuildConflictSheet(DBID As String)
Dim frmConflictSheet As Form
Dim rstRecords As ADODB.Recordset
Dim ctl As Control
Dim ctlOrig As Control
Dim ctlCurrent As Control
Dim ctlLabel As Control
Dim i As Integer
Dim strSQL As String
Dim strFormName As String
Dim columnNumber As Integer

Const TWIPS_CHARACTER_WIDTH = 144
Const TWIPS_ROW_HEIGHT = 350
Const TWIPS_CONTROL_HEIGHT = 288
Const TWIPS_LABEL_COLUMN_WIDTH = 1440
Const TWIPS_GUTTER_COLUMN_WIDTH = 180
Const TWIPS_CONTROL_COLUMN_WIDTH = 1440 * 5.5
Const TWIPS_DETAIL_HEIGHT = 1440 * 4
Const TWIPS_FORM_WIDTH = 1440 * 9
Const TWIPS_COLUMNWIDTH = 1440 * 7


DoCmd.Echo (False)
DoCmd.SetWarnings (False)
On Error Resume Next
DoCmd.Close acForm, "conflict_qdb_" + DBID
DoCmd.DeleteObject acForm, "conflict_qdb_" + DBID
On Error GoTo 0
DoCmd.SetWarnings (True)

Set frmConflictSheet = Access.Application.CreateForm()
frmConflictSheet.DefaultView = 0
frmConflictSheet.Section(acDetail).Height = TWIPS_DETAIL_HEIGHT
frmConflictSheet.Width = TWIPS_FORM_WIDTH
frmConflictSheet.RecordSelectors = False
frmConflictSheet.NavigationButtons = False



strSQL = "SELECT qdbFields.*, qdbFields.DBID FROM qdbFields WHERE (((qdbFields.DBID)='" + DBID + "'));"

Set rstRecords = New ADODB.Recordset
rstRecords.Open strSQL, CurrentProject.Connection, adOpenForwardOnly
i = 1
Set ctl = CreateControl(frmConflictSheet.Name, acCommandButton, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (TWIPS_CONTROL_COLUMN_WIDTH * (1 / 3)), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT * 2)
ctl.Name = "cmdOK"
ctl.Caption = "Update QuickBase and Local Copy" + vbCrLf + "with Changes Below"
ctl.OnClick = "=HideConflictScreen(""" + DBID + """)"

Set ctl = CreateControl(frmConflictSheet.Name, acCommandButton, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (TWIPS_CONTROL_COLUMN_WIDTH * (2 / 3)), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT * 2)
ctl.Name = "cmdCancel"
ctl.Caption = "Ignore My Changes and Update" + vbCrLf + "Local Copy with Column Below"
ctl.OnClick = "=CloseConflictScreen(""" + DBID + """)"
i = 3

Set ctlOrig = CreateControl(frmConflictSheet.Name, acLabel, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
Set ctl = CreateControl(frmConflictSheet.Name, acLabel, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (TWIPS_CONTROL_COLUMN_WIDTH / 3) + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
Set ctlCurrent = CreateControl(frmConflictSheet.Name, acLabel, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (2 * TWIPS_CONTROL_COLUMN_WIDTH / 3) + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
ctlOrig.Name = "lblOriginal"
ctl.Name = "lblProposed"
ctlCurrent.Name = "lblCurrent"
ctlOrig.Caption = "When you last synched"
ctl.Caption = "Your changes"
ctlCurrent.Caption = "Currently on www.quickbase.com"

i = 4
Do While Not rstRecords.EOF
    If rstRecords("Num Lines") > 0 Then
        If ((i * TWIPS_ROW_HEIGHT) + (TWIPS_CONTROL_HEIGHT * rstRecords("Num Lines"))) \ (22 * 1440) >= 1 Then
            columnNumber = columnNumber + 1
            i = 0
        End If
    Else
        If ((i * TWIPS_ROW_HEIGHT) + TWIPS_CONTROL_HEIGHT) \ (22 * 1440) >= 1 Then
            columnNumber = columnNumber + 1
            i = 0
        End If
    End If

    'Place a control on the form
    If IsNull(rstRecords("Choices")) Or rstRecords("Choices") = "" Then
        If rstRecords("Base Type") = "bool" Then
            Set ctlOrig = CreateControl(frmConflictSheet.Name, acCheckBox, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (TWIPS_CONTROL_COLUMN_WIDTH / 6) + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
            Set ctl = CreateControl(frmConflictSheet.Name, acCheckBox, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (TWIPS_CONTROL_COLUMN_WIDTH / 2) + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
            Set ctlCurrent = CreateControl(frmConflictSheet.Name, acCheckBox, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (TWIPS_CONTROL_COLUMN_WIDTH * (5 / 6)) + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
        Else
            Set ctlOrig = CreateControl(frmConflictSheet.Name, acTextBox, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
            Set ctl = CreateControl(frmConflictSheet.Name, acTextBox, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (TWIPS_CONTROL_COLUMN_WIDTH / 3) + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
            Set ctlCurrent = CreateControl(frmConflictSheet.Name, acTextBox, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (2 * TWIPS_CONTROL_COLUMN_WIDTH / 3) + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
            ctl.EnterKeyBehavior = True
        End If
    Else
        Set ctlOrig = CreateControl(frmConflictSheet.Name, acTextBox, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
        Set ctl = CreateControl(frmConflictSheet.Name, acComboBox, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (TWIPS_CONTROL_COLUMN_WIDTH / 3) + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3, TWIPS_CONTROL_HEIGHT)
        Set ctlCurrent = CreateControl(frmConflictSheet.Name, acTextBox, acDetail, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (2 * TWIPS_CONTROL_COLUMN_WIDTH / 3), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH / 3 + (columnNumber * TWIPS_COLUMNWIDTH), TWIPS_CONTROL_HEIGHT)
        ctl.RowSourceType = "Value List"
        ctl.RowSource = rstRecords("Choices")
        ctl.LimitToList = rstRecords("Allow New Choices")
    End If
    ctl.Name = "proposed_" + rstRecords("FID")
    ctlOrig.Name = "orig_" + rstRecords("FID")
    ctlCurrent.Name = "current_" + rstRecords("FID")
    Set ctlLabel = CreateControl(frmConflictSheet.Name, acLabel, acDetail, "", "", (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_LABEL_COLUMN_WIDTH, TWIPS_CONTROL_HEIGHT)
    ctlLabel.Caption = rstRecords("Label")
    ctlLabel.Name = "lbl_" + rstRecords("FID")
    ctlLabel.TextAlign = 3
    ctlLabel.Width = TWIPS_LABEL_COLUMN_WIDTH
    If rstRecords("Num Lines") > 0 Then
        ctlOrig.Height = TWIPS_CONTROL_HEIGHT * rstRecords("Num Lines")
        ctlCurrent.Height = TWIPS_CONTROL_HEIGHT * rstRecords("Num Lines")
        ctl.Height = TWIPS_CONTROL_HEIGHT * rstRecords("Num Lines")
        i = i + rstRecords("Num Lines")
    Else
        i = i + 1
    End If
    ctl.Locked = True
    ctlOrig.Locked = True
    ctlCurrent.Locked = True
    If rstRecords("Base Type") <> "bool" Then
        ctl.BackColor = 12632256
        ctlOrig.BackColor = 12632256
        ctlCurrent.BackColor = 12632256
    End If
    rstRecords.MoveNext
Loop
rstRecords.Close
strSQL = "SELECT qdbTables.Label FROM qdbTables WHERE (((qdbTables.DBID)='" + DBID + "'));"
rstRecords.Open strSQL, CurrentProject.Connection, adOpenForwardOnly
frmConflictSheet.Caption = "Conflict Resolution for: " + rstRecords("Label")
rstRecords.Close
frmConflictSheet.HasModule = False
frmConflictSheet.Modal = True
frmConflictSheet.PopUp = True
strFormName = frmConflictSheet.Name
DoCmd.Close acForm, strFormName, acSaveYes
DoCmd.Rename "conflict_qdb_" + DBID, acForm, strFormName
DoCmd.Echo (True)
End Sub
Sub BuildDataSheet(DBID As String)
Dim frmDataSheet As Form
Dim rstFields As ADODB.Recordset
Dim ctl As Control
Dim ctlLabel As Control
Dim i As Integer
Dim strSQL As String
Dim strFormName As String
Dim columnNumber As Integer


Const TWIPS_CHARACTER_WIDTH = 144
Const TWIPS_ROW_HEIGHT = 350
Const TWIPS_CONTROL_HEIGHT = 288
Const TWIPS_LABEL_COLUMN_WIDTH = 2000
Const TWIPS_GUTTER_COLUMN_WIDTH = 180
Const TWIPS_CONTROL_COLUMN_WIDTH = 1440 * 3
Const TWIPS_HEADER_HEIGHT = 144 * 3
Const TWIPS_DETAIL_HEIGHT = 1440 * 4
Const TWIPS_FORM_WIDTH = 1440 * 7
Const TWIPS_COLUMNWIDTH = 1440 * 7


DoCmd.Echo (False)
DoCmd.SetWarnings (False)
On Error Resume Next
DoCmd.Close acForm, "frm_qdb_" + DBID
DoCmd.DeleteObject acForm, "frm_qdb_" + DBID
On Error GoTo 0
DoCmd.SetWarnings (True)

Set frmDataSheet = Access.Application.CreateForm(, "frmTemplate")
frmDataSheet.OnUnload = "Save Confirmation"
frmDataSheet.DefaultView = 2
frmDataSheet.Section(acDetail).Height = TWIPS_DETAIL_HEIGHT
frmDataSheet.Section(acHeader).Visible = True
frmDataSheet.Section(acHeader).BackColor = 16776960
frmDataSheet.Section(acHeader).Height = TWIPS_HEADER_HEIGHT
frmDataSheet.Width = TWIPS_FORM_WIDTH
frmDataSheet.RecordSelectors = False


Set ctlLabel = CreateControl(frmDataSheet.Name, acLabel, acHeader, "", "", 0, 0, TWIPS_LABEL_COLUMN_WIDTH, TWIPS_CONTROL_HEIGHT)
ctlLabel.Name = "lblViews"
ctlLabel.Caption = "QuickBase Views"
ctlLabel.TextAlign = 3
ctlLabel.Width = TWIPS_LABEL_COLUMN_WIDTH
Set ctl = CreateControl(frmDataSheet.Name, acComboBox, acHeader, "", "", TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH, TWIPS_CONTROL_HEIGHT)
ctl.RowSourceType = "Table/Query"
ctl.RowSource = "SELECT [qdbViews].[Name], [qdbViews].[qid] FROM qdbViews WHERE ((([qdbViews].[DBID])='" + DBID + "'));"

strSQL = "SELECT qdbFields.*, qdbFields.DBID FROM qdbFields WHERE (((qdbFields.DBID)='" + DBID + "'));"

Set rstFields = New ADODB.Recordset
rstFields.Open strSQL, CurrentProject.Connection, adOpenForwardOnly
i = 0
columnNumber = 0

Do While Not rstFields.EOF
    If rstFields("Num Lines") > 0 Then
        If ((i * TWIPS_ROW_HEIGHT) + (TWIPS_CONTROL_HEIGHT * rstFields("Num Lines"))) \ (22 * 1440) >= 1 Then
            columnNumber = columnNumber + 1
            i = 0
        End If
    Else
        If ((i * TWIPS_ROW_HEIGHT) + TWIPS_CONTROL_HEIGHT) \ (22 * 1440) >= 1 Then
            columnNumber = columnNumber + 1
            i = 0
        End If
    End If
    
    'Place a control on the form
    If IsNull(rstFields("Choices")) Or rstFields("Choices") = "" Then
        If rstFields("Base Type") = "bool" Then
            Set ctl = CreateControl(frmDataSheet.Name, acCheckBox, acDetail, "", rstFields("FID"), TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH, TWIPS_CONTROL_HEIGHT)
        Else
            Set ctl = CreateControl(frmDataSheet.Name, acTextBox, acDetail, "", rstFields("FID"), TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH, TWIPS_CONTROL_HEIGHT)
            ctl.EnterKeyBehavior = True
        End If
    Else
        Set ctl = CreateControl(frmDataSheet.Name, acComboBox, acDetail, "", rstFields("FID"), TWIPS_LABEL_COLUMN_WIDTH + TWIPS_GUTTER_COLUMN_WIDTH + (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_CONTROL_COLUMN_WIDTH, TWIPS_CONTROL_HEIGHT)
        ctl.RowSourceType = "Value List"
        ctl.RowSource = rstFields("Choices")
        ctl.LimitToList = rstFields("Allow New Choices")
    End If
    ctl.Name = rstFields("Label")
    If rstFields("Deleted") Then
        ctl.Locked = True
    End If
    Set ctlLabel = CreateControl(frmDataSheet.Name, acLabel, acDetail, "", "", (columnNumber * TWIPS_COLUMNWIDTH), i * TWIPS_ROW_HEIGHT, TWIPS_LABEL_COLUMN_WIDTH, TWIPS_CONTROL_HEIGHT)
    ctlLabel.Name = "lbl_" + rstFields("FID")
    ctlLabel.Caption = rstFields("Label")
    ctlLabel.TextAlign = 3
    If rstFields("Width") > 0 Then
        ctl.Width = TWIPS_CHARACTER_WIDTH * rstFields("Width")
    End If
    ctlLabel.Width = TWIPS_LABEL_COLUMN_WIDTH
    If rstFields("Num Lines") > 0 Then
        ctl.Height = TWIPS_CONTROL_HEIGHT * rstFields("Num Lines")
        i = i + rstFields("Num Lines")
    Else
        i = i + 1
    End If
    If IsNull(rstFields("Mode")) And IsNull(rstFields("Role")) Then
        ctl.Enabled = True
    Else
        ctl.Enabled = False
    End If
    rstFields.MoveNext
Loop
rstFields.Close
strSQL = "SELECT qdbTables.Label FROM qdbTables WHERE (((qdbTables.DBID)='" + DBID + "'));"
rstFields.Open strSQL, CurrentProject.Connection, adOpenForwardOnly
frmDataSheet.RecordSource = "qdb_" + DBID
frmDataSheet.Caption = rstFields("Label")
rstFields.Close
frmDataSheet.HasModule = False
strFormName = frmDataSheet.Name
DoCmd.Close acForm, strFormName, acSaveYes
DoCmd.Rename "frm_qdb_" + DBID, acForm, strFormName
DoCmd.Echo (True)
End Sub

Function makeMSAccessObjectName(strQDBFieldName As String)
Dim strFieldName As String
    strFieldName = Replace(strQDBFieldName, ".", ",")
    strFieldName = Replace(strFieldName, "[", "{")
    strFieldName = Replace(strFieldName, "]", "}")
    strFieldName = Replace(strFieldName, "!", ",")
    strFieldName = Replace(strFieldName, "`", "'")
    strFieldName = Trim(strFieldName)
    makeMSAccessObjectName = Left(strFieldName, 64)
End Function

Function getLocalDBLastModifiedDate(strDBID As String) As Date
Dim strSQL As String
Dim rstRecords As ADODB.Recordset
Dim strLastModifiedFID As String


strLastModifiedFID = getDateModifiedFID(strDBID)
strSQL = "SELECT orig_" + strDBID + ".[" + strLastModifiedFID + "] FROM orig_" + strDBID + " ORDER BY orig_" + strDBID + ".[" + strLastModifiedFID + "] DESC;"
Set rstRecords = New ADODB.Recordset
rstRecords.Open strSQL, CurrentProject.Connection, adOpenForwardOnly
If rstRecords.EOF Then
    getLocalDBLastModifiedDate = #1/1/1970#
Else
    rstRecords.MoveFirst
    getLocalDBLastModifiedDate = rstRecords(strLastModifiedFID)
End If
rstRecords.Close
End Function

Function getLocalDBMaxRecordID(strDBID As String) As Integer
Dim strSQL As String
Dim rstRecords As ADODB.Recordset
Dim strRecordIDFID As String


strRecordIDFID = getRecordIDFID(strDBID)
strSQL = "SELECT orig_" + strDBID + ".[" + strRecordIDFID + "] FROM orig_" + strDBID + " ORDER BY orig_" + strDBID + ".[" + strRecordIDFID + "] DESC;"
Set rstRecords = New ADODB.Recordset
rstRecords.Open strSQL, CurrentProject.Connection, adOpenForwardOnly
If rstRecords.EOF Then
    getLocalDBMaxRecordID = 0
Else
    rstRecords.MoveFirst
    getLocalDBMaxRecordID = rstRecords(strRecordIDFID)
End If
rstRecords.Close
End Function
Function UpdateFieldsTable(strDBID As String, FieldNodeList As MSXML.IXMLDOMNodeList, DBName As String) As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    Dim clist As String
    Dim strFieldName As String
    Dim strFID As String
    Dim strChoice As String
    Dim strChoices As String
    Dim ChoiceNodeList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strSQL As String
    
    strSQL = "SELECT "
    
    'Clean out the field descriptions in the qdbFields table
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL ("DELETE FROM qdbFields WHERE (((qdbFields.DBID)='" + strDBID + "'));")
    DoCmd.SetWarnings (True)
    Set rst = New ADODB.Recordset
    rst.Open "qdbFields", CurrentProject.Connection, adOpenForwardOnly, adLockOptimistic

    For i = 0 To FieldNodeList.length - 1
        rst.AddNew
        rst("DBID") = strDBID
        strFieldName = makeMSAccessObjectName(FieldNodeList(i).selectSingleNode("label").nodeTypedValue)
        rst("Label") = strFieldName
        strFID = FieldNodeList(i).selectSingleNode("@id").Text
        rst("FID") = strFID
        strSQL = strSQL + "qdb_" + strDBID + ".[" + strFID + "] AS [" + strFieldName + "], "
        rst("Field Type") = FieldNodeList(i).selectSingleNode("@field_type").Text
        rst("Base Type") = FieldNodeList(i).selectSingleNode("@base_type").Text
        On Error Resume Next
            rst("Mode") = FieldNodeList(i).selectSingleNode("@mode").Text
            rst("Role") = FieldNodeList(i).selectSingleNode("@role").Text
            rst("Width") = FieldNodeList(i).selectSingleNode("width").Text
            rst("Num Lines") = FieldNodeList(i).selectSingleNode("num_lines").Text
            rst("Required") = CBool(FieldNodeList(i).selectSingleNode("required").Text)
            rst("Appears By Default") = CBool(FieldNodeList(i).selectSingleNode("appears_by_default").Text)
            rst("Allow New Choices") = CBool(FieldNodeList(i).selectSingleNode("allow_new_choices").Text)
            Set ChoiceNodeList = FieldNodeList(i).selectNodes("choices/choice")
            strChoices = ""
            For Each xmlNode In ChoiceNodeList
                strChoice = xmlNode.Text
                strChoice = Replace(strChoice, ";", ",")
                If Len(strChoices + strChoice) < 2048 Then
                    strChoices = strChoices + strChoice + ";"
                Else
                    rst("Allow New Choices") = True
                    strChoices = "Sorry QuickBase Desktop could not accomodate all the choices."
                    Exit For
                End If
            Next xmlNode
            rst("Choices") = strChoices
        On Error GoTo 0
        clist = clist + strFID + "."
        rst.Update
    Next i
    rst.Close
    strSQL = Trim(strSQL)
    If Right(strSQL, 1) = "," Then
        strSQL = Left(strSQL, Len(strSQL) - 1)
    End If
    strSQL = strSQL + " FROM qdb_" + strDBID + ";"
    makeQueryForTable strDBID, DBName, strSQL
    UpdateFieldsTable = clist
End Function


Sub setTableDescription(QDBTableName As String, strDescription As String)
Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim prpDescription As DAO.Property

Set db = CurrentDb()
Set tdf = db.TableDefs(QDBTableName)
Set prpDescription = tdf.CreateProperty("Description", dbText, strDescription)
On Error Resume Next
    tdf.Properties.Append prpDescription
On Error GoTo 0
tdf.Properties("Description") = strDescription
End Sub

Sub setHyperLinkField(QDBTableName As String, strFieldName As String)
Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim fld As DAO.Field

Set db = CurrentDb()
Set tdf = db.TableDefs(QDBTableName)
tdf.Fields.Delete strFieldName

Set fld = tdf.CreateField(strFieldName, dbMemo)
fld.Attributes = fld.Attributes + dbHyperlinkField
tdf.Fields.Append fld
tdf.Fields.Refresh
End Sub

