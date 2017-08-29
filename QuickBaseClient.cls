VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QuickBaseClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'© 2001 Intuit Inc. All rights reserved.
'Use is subject to the IP Rights Notice and Restrictions available at
'http://developer.intuit.com/legal/IPRNotice_021201.html


Private Password As String
Private UserName As String
Private UserToken As String
Private strProxyPassword As String
Private strProxyUsername As String
Private ticket As String
Private QDBHost As String
Private useHTTPS As Boolean

Public errorcode As Long
Public errortext As String
Public errordetail As String
Public httpContentLengthProgress  As Long
Public httpContentLength As Long

Private Const OB32CHARACTERS = "abcdefghijkmnpqrstuvwxyz23456789"
Private Const Map64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Const MILLISECONDS_IN_A_DAY = 86400000#
Private Const DAYS_BETWEEN_JAVASCRIPT_AND_MICROSOFT_DATE_REFERENCES = 25569#


Function encode32(ByVal strDecimal As String) As String

Dim ob32 As String
Dim intDecimal As Long
intDecimal = CLng(strDecimal)
Dim remainder As Long

Do While (intDecimal > 0)
   remainder = intDecimal Mod 32
   ob32 = Mid(OB32CHARACTERS, remainder + 1, 1) & ob32
   intDecimal = intDecimal \ 32
Loop
encode32 = ob32

End Function

Public Function FieldAddChoices(dbid As String, fid As String, ParamArray NameValues()) As Long
    Dim xmlQDBRequest As New MSXML.DOMDocument
    Dim firstfield As Long
    Dim lastfield As Long
    Dim i As Long
    
    Set xmlQDBRequest = InitXMLRequest()
    addParameter xmlQDBRequest, "fid", fid
    lastfield = UBound(NameValues)
    firstfield = LBound(NameValues)
    For i = firstfield To lastfield
        addParameter xmlQDBRequest, "choice", NameValues(i)
    Next i
    
    Set xmlQDBRequest = APIXMLPost(dbid, "API_FieldAddChoices", xmlQDBRequest, useHTTPS)
    FieldAddChoices = CInt(xmlQDBRequest.documentElement.selectSingleNode("/*/numadded").Text)
End Function

Public Function CreateDatabase(ByVal dbname As String, ByVal dbdesc As String) As String
Dim xmlQDBRequest As New MSXML.DOMDocument
    
    Set xmlQDBRequest = InitXMLRequest()
    addParameter xmlQDBRequest, "dbname", dbname
    addParameter xmlQDBRequest, "dbdesc", dbdesc
    CreateDatabase = ""
    CreateDatabase = APIXMLPost("main", "API_CreateDatabase", xmlQDBRequest, useHTTPS).documentElement.selectSingleNode("/*/dbid").Text
End Function
Public Sub DeleteDatabase(ByVal dbid As String)
Dim xmlQDBRequest As New MSXML.DOMDocument
    
    Set xmlQDBRequest = InitXMLRequest()
    addParameter xmlQDBRequest, "dbid", dbid
    Call APIXMLPost(dbid, "API_DeleteDatabase", xmlQDBRequest, useHTTPS)
End Sub
Public Function AddField(ByVal dbid As String, ByVal label As String, ByVal fieldtype As String, ByVal Formula As Boolean) As String
Dim xmlQDBRequest As New MSXML.DOMDocument
    
    Set xmlQDBRequest = InitXMLRequest()
    addParameter xmlQDBRequest, "label", label
    addParameter xmlQDBRequest, "type", fieldtype
    If Formula Then
        addParameter xmlQDBRequest, "mode", "virtual"
    End If
    AddField = ""
    AddField = APIXMLPost(dbid, "API_AddField", xmlQDBRequest, useHTTPS).documentElement.selectSingleNode("/*/fid").Text
End Function
Public Sub SetFieldProperties(ByVal dbid As String, ByVal fid As String, ParamArray NameValues())
    Dim xmlQDBRequest As New MSXML.DOMDocument
    Dim lastfield As Long
    Dim firstfield As Long
    Dim i As Long
    
    Set xmlQDBRequest = InitXMLRequest()
    addParameter xmlQDBRequest, "fid", fid
    lastfield = UBound(NameValues)
    firstfield = LBound(NameValues)
    i = 0
    For i = firstfield To lastfield Step 2
        addParameter xmlQDBRequest, NameValues(i), NameValues(i + 1)
    Next i
    Call APIXMLPost(dbid, "API_SetFieldProperties", xmlQDBRequest, useHTTPS)
End Sub

Public Function Authenticate(Optional ByVal strUsername As String, Optional ByVal strPassword As String, Optional ByVal strUserToken As String) As Long
    If strUsername <> "" Then
        UserName = strUsername
        Password = strPassword
    Else
        UserToken = strUserToken
    End If
    ticket = ""
    Authenticate = 0
End Function
Public Function proxyAuthenticate(strUsername As String, strPassword As String) As Integer
    strProxyUsername = strUsername
    strProxyPassword = strPassword
    proxyAuthenticate = 0
End Function
Function downloadAttachedFile(ByVal dbid As String, ByVal rid As String, ByVal fid As String, ByVal DownloadDirectory As String, ByVal Filename As String) As String
    Filename = makeValidFilename(Filename)
    downloadAttachedFile = HTTPPost("", True, "/up/" + dbid + "/g/r" & encode32(rid) & "/e" & encode32(fid) & "/?ticket=" + getTicket(), "text/html", "", DownloadDirectory + "\" + Filename)
End Function
Public Function FindDBByName(ByVal dbname As String) As String
    Dim xmlQDBRequest As New MSXML.DOMDocument
    
        Set xmlQDBRequest = InitXMLRequest()
        addParameter xmlQDBRequest, "dbname", dbname
        FindDBByName = ""
        FindDBByName = APIXMLPost("main", "API_FindDBByName", xmlQDBRequest, useHTTPS).documentElement.selectSingleNode("/*/dbid").Text
End Function
Public Function CloneDatabase(ByVal sourcedbid As String, ByVal Name As String, ByVal Description As String) As String
    Dim xmlQDBRequest As New MSXML.DOMDocument
    
        Set xmlQDBRequest = InitXMLRequest()
        addParameter xmlQDBRequest, "newdbname", Name
        addParameter xmlQDBRequest, "newdbdesc", Description
        CloneDatabase = ""
        CloneDatabase = APIXMLPost(sourcedbid, "API_CloneDatabase", xmlQDBRequest, useHTTPS).documentElement.selectSingleNode("/*/newdbid").Text
End Function
Public Function ImportFromCSV(ByVal dbid As String, ByVal CSV As String, ByVal clist As String, rids() As Long, skipfirst As Boolean) As Long
    Dim xmlQDBRequest As New MSXML.DOMDocument
    Dim RidNodeList As MSXML.IXMLDOMNodeList

    Set xmlQDBRequest = InitXMLRequest()
    addParameter xmlQDBRequest, "clist", clist
    If skipfirst Then
        addParameter xmlQDBRequest, "skipfirst", "1"
    End If
    addCDATAParameter xmlQDBRequest, "records_csv", CSV
    Set xmlQDBRequest = APIXMLPost(dbid, "API_ImportFromCSV", xmlQDBRequest, useHTTPS)
    Set RidNodeList = xmlQDBRequest.selectNodes("/*/rids/rid")
    Dim ridListLength As Long
    Dim i As Long
    ridListLength = RidNodeList.length
    If ridListLength > 0 Then
        ReDim rids(ridListLength - 1)
        For i = 0 To ridListLength - 1
            rids(i) = RidNodeList(i).Text
        Next i
    End If
    On Error Resume Next
    ImportFromCSV = CLng(xmlQDBRequest.documentElement.selectSingleNode("/*/num_recs_added").Text)
    ImportFromCSV = CLng(xmlQDBRequest.documentElement.selectSingleNode("/*/num_recs_updated").Text)
    xmlQDBRequest = Nothing
End Function
Public Function AddRecordByArray(ByVal dbid As String, ByRef update_id As String, ByRef NameValues()) As String
    Dim xmlQDBRequest As New MSXML.DOMDocument
    Dim firstfield As Long
    Dim lastfield As Long
    Dim i As Long
    AddRecordByArray = ""
    
    Set xmlQDBRequest = InitXMLRequest()
    lastfield = UBound(NameValues, 2)
    firstfield = LBound(NameValues, 2)
      
    For i = firstfield To lastfield
        If IsNull(NameValues(0, i)) Then
            Err.Raise vbObjectError + 2, "QuickBase.QuickBaseClient", "AddRecordByArray: Please do use null for field names or fids"
            Exit Function
        End If
        If IsNull(NameValues(1, i)) Then
            NameValues(1, i) = ""
        End If
        
        If (IsNumeric(NameValues(0, i)) And Not IsDate(NameValues(0, i))) Then
          addFieldParameter xmlQDBRequest, "fid", CStr(NameValues(0, i)), NameValues(1, i)
        Else
          addFieldParameter xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(0, i))), NameValues(1, i)
        End If
    Next i
    Set xmlQDBRequest = APIXMLPost(dbid, "API_AddRecord", xmlQDBRequest, useHTTPS)
    update_id = xmlQDBRequest.documentElement.selectSingleNode("/*/update_id").Text
    AddRecordByArray = xmlQDBRequest.documentElement.selectSingleNode("/*/rid").Text
End Function

Public Function EditRecordByArray(ByVal dbid As String, ByVal rid As String, ByRef update_id As String, ByRef NameValues()) As String
    Dim xmlQDBRequest As New MSXML.DOMDocument
    Dim firstfield As Long
    Dim lastfield As Long
    Dim i As Long
    EditRecordByArray = ""
    
    Set xmlQDBRequest = InitXMLRequest()
    lastfield = UBound(NameValues, 2)
    firstfield = LBound(NameValues, 2)
    
    
   
    For i = firstfield To lastfield
          If (IsNumeric(NameValues(0, i)) And Not IsDate(NameValues(0, i))) Then
            addFieldParameter xmlQDBRequest, "fid", CStr(NameValues(0, i)), NameValues(1, i)
          Else
            addFieldParameter xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(0, i))), NameValues(1, i)
          End If
    Next i
    addParameter xmlQDBRequest, "rid", rid
    If update_id <> "" Then
        addParameter xmlQDBRequest, "update_id", update_id
    End If
    EditRecordByArray = APIXMLPost(dbid, "API_EditRecord", xmlQDBRequest, useHTTPS).documentElement.selectSingleNode("/*/update_id").Text
End Function
Public Function AddRecord(ByVal dbid As String, ByRef update_id As String, ParamArray NameValues()) As String
    Dim xmlQDBRequest As New MSXML.DOMDocument
    Dim firstfield As Long
    Dim lastfield As Long
    Dim i As Long
    AddRecord = ""
    
    Set xmlQDBRequest = InitXMLRequest()
    lastfield = UBound(NameValues)
    firstfield = LBound(NameValues)
    If ((lastfield - firstfield + 1) Mod 2) <> 0 Then
        Err.Raise vbObjectError + 3, "QuickBase.QuickBaseClient", "AddRecord: Please use an even number of arguements after the DBID"
        Exit Function
    End If
    
   
    For i = firstfield To lastfield Step 2
          If (IsNumeric(NameValues(i)) And Not IsDate(NameValues(i))) Then
            addFieldParameter xmlQDBRequest, "fid", CStr(NameValues(i)), NameValues(i + 1)
          Else
            addFieldParameter xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(i))), NameValues(i + 1)
          End If
    Next i
    Set xmlQDBRequest = APIXMLPost(dbid, "API_AddRecord", xmlQDBRequest, useHTTPS)
    update_id = xmlQDBRequest.documentElement.selectSingleNode("/*/update_id").Text
    AddRecord = xmlQDBRequest.documentElement.selectSingleNode("/*/rid").Text
End Function
Public Function EditRecord(ByVal dbid As String, ByVal rid As String, ByRef update_id As String, ParamArray NameValues()) As String
    Dim xmlQDBRequest As New MSXML.DOMDocument
    Dim firstfield As Long
    Dim lastfield As Long
    Dim i As Long
    EditRecord = ""
    
    Set xmlQDBRequest = InitXMLRequest()
    lastfield = UBound(NameValues)
    firstfield = LBound(NameValues)
    If ((lastfield - firstfield + 1) Mod 2) <> 0 Then
        Err.Raise vbObjectError + 4, "QuickBase.QuickBaseClient", "EditRecord: Please use an even number of arguements."
        Exit Function
    End If
    
   
    For i = firstfield To lastfield Step 2
          If (IsNumeric(NameValues(i)) And Not IsDate(NameValues(i))) Then
            addFieldParameter xmlQDBRequest, "fid", CStr(NameValues(i)), NameValues(i + 1)
          Else
            addFieldParameter xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(i))), NameValues(i + 1)
          End If
    Next i
    addParameter xmlQDBRequest, "rid", rid
    If update_id <> "" Then
        addParameter xmlQDBRequest, "update_id", update_id
    End If
    EditRecord = APIXMLPost(dbid, "API_EditRecord", xmlQDBRequest, useHTTPS).documentElement.selectSingleNode("/*/update_id").Text
       
End Function
Public Function DeleteRecord(ByVal dbid As String, ByVal rid As Variant) As String
    Dim xmlQDBRequest As New MSXML.DOMDocument
    DeleteRecord = ""
    
    Set xmlQDBRequest = InitXMLRequest()
    addParameter xmlQDBRequest, "rid", CStr(rid)
    
    DeleteRecord = APIXMLPost(dbid, "API_DeleteRecord", xmlQDBRequest, useHTTPS).documentElement.selectSingleNode("/*/rid").Text
      
End Function
Public Function GetSchema(ByVal dbid As String) As MSXML.DOMDocument
    Dim xmlQDBRequest As New MSXML.DOMDocument
        
    Set xmlQDBRequest = InitXMLRequest()
    Set GetSchema = APIXMLPost(dbid, "API_GetSchema", xmlQDBRequest, useHTTPS)
End Function

Public Function GetGrantedDBs(withEmbeddedTables As Boolean, excludeParents As Boolean, adminOnly As Boolean) As MSXML.DOMDocument
    Dim xmlQDBRequest As New MSXML.DOMDocument
        
    Set xmlQDBRequest = InitXMLRequest()
    If withEmbeddedTables Then
        addParameter xmlQDBRequest, "withEmbeddedTables", "1"
    Else
        addParameter xmlQDBRequest, "withEmbeddedTables", "0"
    End If
    If excludeParents Then
        addParameter xmlQDBRequest, "excludeParents", "1"
    Else
        addParameter xmlQDBRequest, "excludeParents", "0"
    End If
    If adminOnly Then
        addParameter xmlQDBRequest, "adminOnly", "1"
    End If
    Set GetGrantedDBs = APIXMLPost("main", "API_GrantedDBs", xmlQDBRequest, useHTTPS)
End Function
Public Function GetDBInfo(ByVal dbid As String) As MSXML.DOMDocument
    Dim xmlQDBRequest As New MSXML.DOMDocument
        
    Set xmlQDBRequest = InitXMLRequest()
    Set GetDBInfo = APIXMLPost(dbid, "API_GetDBInfo", xmlQDBRequest, useHTTPS)
End Function

Public Function ChangeRecordOwner(ByVal dbid As String, ByVal rid As Variant, ByVal Owner As String) As Boolean
    Dim xmlQDBRequest As New MSXML.DOMDocument
        
    Set xmlQDBRequest = InitXMLRequest()
    addParameter xmlQDBRequest, "rid", CStr(rid)
    addParameter xmlQDBRequest, "newowner", Owner
    On Error GoTo noChange
    Call APIXMLPost(dbid, "API_ChangeRecordOwner", xmlQDBRequest, useHTTPS)
    ChangeRecordOwner = True
    Exit Function
    
noChange:
    ChangeRecordOwner = False
    Exit Function
End Function
Public Function DoQuery(ByVal dbid As String, ByVal query As String, ByVal clist As String, ByVal slist As String, ByVal options As String) As MSXML.DOMDocument
    Dim xmlQDBRequest As New MSXML.DOMDocument
        
    Set xmlQDBRequest = InitXMLRequest()
    
    If Left(query, 1) = "{" And Right(query, 1) = "}" Then
        addParameter xmlQDBRequest, "query", query
    ElseIf CStr(Val(query)) = query Then
        addParameter xmlQDBRequest, "qid", query
    Else
        addParameter xmlQDBRequest, "qname", query
    End If
    addParameter xmlQDBRequest, "clist", clist
    addParameter xmlQDBRequest, "slist", slist
    addParameter xmlQDBRequest, "options", options
    addParameter xmlQDBRequest, "fmt", "structured"
    Set DoQuery = APIXMLPost(dbid, "API_DoQuery", xmlQDBRequest, useHTTPS)
End Function
Public Function GenResultsTable(ByVal dbid As String, ByVal query As String, ByVal clist As String, ByVal slist As String, ByVal options As String) As String
    Dim xmlQDBRequest As New MSXML.DOMDocument
        
    Set xmlQDBRequest = InitXMLRequest()
    
    If Left(query, 1) = "{" And Right(query, 1) = "}" Then
        addParameter xmlQDBRequest, "query", query
    ElseIf CStr(Val(query)) = query Then
        addParameter xmlQDBRequest, "qid", query
    Else
        addParameter xmlQDBRequest, "qname", query
    End If
    addParameter xmlQDBRequest, "clist", clist
    addParameter xmlQDBRequest, "slist", slist
    addParameter xmlQDBRequest, "options", options
    GenResultsTable = APIHTMLPost(dbid, "API_GenResultsTable", xmlQDBRequest, useHTTPS)
End Function

Public Function DoQueryAsArray(ByVal dbid As String, ByVal query As String, ByVal clist As String, ByVal slist As String, ByVal options As String) As Variant()
    Dim xmlQDBResponse As New MSXML.DOMDocument
    Set xmlQDBResponse = DoQuery(dbid, query, clist, slist, options)
    Dim QDBRecord() As Variant

    Dim i As Long
    Dim j As Long
    Dim FieldNodeList As MSXML.IXMLDOMNodeList
    Dim RecordNodeList As MSXML.IXMLDOMNodeList
    Dim intFields As Long
    Dim strFieldValue As String
        
    Set FieldNodeList = xmlQDBResponse.documentElement.selectNodes("/*/table/fields/field")
    intFields = FieldNodeList.length
    
    Set RecordNodeList = xmlQDBResponse.documentElement.selectNodes("/*/table/records/record")
    
    ReDim QDBRecord(RecordNodeList.length + 1, intFields)
    
    For i = 0 To intFields - 1
        QDBRecord(0, i) = FieldNodeList(i).selectSingleNode("label").nodeTypedValue
    Next i
    
    For i = 1 To RecordNodeList.length
        For j = 0 To intFields - 1
            strFieldValue = RecordNodeList(i - 1).selectSingleNode("f[" + CStr(j) + "]").nodeTypedValue
            Select Case FieldNodeList(j).selectSingleNode("@base_type").Text
                Case "float"
                    If strFieldValue <> "" Then
                        QDBRecord(i, j) = makeDouble(strFieldValue)
                    End If
                Case "text"
                    QDBRecord(i, j) = Replace(strFieldValue, Chr(10), vbCrLf)
                Case "bool"
                    QDBRecord(i, j) = CBool(strFieldValue)
                Case "int64"
                    If strFieldValue <> "" Then
                        If FieldNodeList(j).selectSingleNode("@field_type").Text <> "date" Then
                            QDBRecord(i, j) = CDbl(strFieldValue) / MILLISECONDS_IN_A_DAY
                        Else
                            QDBRecord(i, j) = int64ToDate(strFieldValue)
                        End If
                    End If
                Case "int32"
                    If FieldNodeList(j).selectSingleNode("@field_type").Text = "userid" Then
                        On Error Resume Next
                            Dim tempLong As Long
                            tempLong = CLng(strFieldValue)
                            If Err.Number = 0 Then
                                QDBRecord(i, j) = xmlQDBResponse.selectSingleNode("/*/table/lusers/luser[@id='" & strFieldValue & "']").Text
                            Else
                                QDBRecord(i, j) = strFieldValue
                            End If
                        On Error GoTo 0
                    Else
                        QDBRecord(i, j) = CLng(strFieldValue)
                    End If
            End Select
        Next j
    Next i
    DoQueryAsArray = QDBRecord()
End Function

Public Function APIXMLPost(ByVal dbid As String, ByVal action As String, ByRef xmlQDBRequest As MSXML.DOMDocument, useHTTPS As Boolean) As MSXML.DOMDocument

  Dim script As String
  Dim hInternetOpen As Long
  Dim hInternetConnect As Long
  Dim hHttpOpenRequest As Long
  Dim bRet As Boolean
  Dim httpPort As Long
  Dim fFlags As Long
  Dim xmlDoc As New MSXML.DOMDocument
  Dim content As String
  
  
  script = "/db/" + dbid + "?act=" + action
  content = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & xmlQDBRequest.xml
  xmlDoc.async = False
  xmlDoc.loadXML HTTPPost(QDBHost, useHTTPS, script, "text/xml", content, "")
  On Error Resume Next
  ticket = xmlDoc.documentElement.selectSingleNode("/*/ticket").Text
  If xmlDoc.documentElement.selectSingleNode("/*/errdetail") Is Nothing Then
    errordetail = xmlDoc.documentElement.selectSingleNode("/*/errtext").Text
  Else
    errordetail = xmlDoc.documentElement.selectSingleNode("/*/errdetail").Text
  End If
  On Error GoTo 0
  If errorcode <> 0 Then
      Err.Raise vbObjectError + CStr(errorcode), "QuickBase.QuickBaseClient", script + ": " + errordetail
  End If

  Set APIXMLPost = xmlDoc
End Function
Public Function APIHTMLPost(ByVal dbid As String, ByVal action As String, ByRef xmlQDBRequest As MSXML.DOMDocument, ByVal useHTTPS As Boolean) As String

  Dim script As String
  Dim hInternetOpen As Long
  Dim hInternetConnect As Long
  Dim hHttpOpenRequest As Long
  Dim bRet As Boolean
  Dim httpPort As Long
  Dim fFlags As Long
  
  
  script = "/db/" + dbid + "?act=" + action
  APIHTMLPost = HTTPPost(QDBHost, useHTTPS, script, "text/xml", xmlQDBRequest.xml, "")
  
End Function
Public Function getServer() As String
getServer = QDBHost
End Function

Public Function getTicket()
    If ticket = "" Then
        Dim xmlQDBRequest As New MSXML.DOMDocument
        
        Set xmlQDBRequest = InitXMLRequest()
        Call APIXMLPost("main", "API_Authenticate", xmlQDBRequest, useHTTPS)
    End If
        getTicket = ticket
End Function

Public Function InitXMLRequest() As MSXML.DOMDocument
    Dim xmlQDBRequest As New MSXML.DOMDocument
    Dim Root As IXMLDOMElement
    Dim ElementNode As IXMLDOMNode
    Dim TextNode As IXMLDOMNode
    
    Set Root = xmlQDBRequest.createElement("qdbapi")
    Set xmlQDBRequest.documentElement = Root
    If Len(ticket) <> 0 Then
        addParameter xmlQDBRequest, "ticket", ticket
    ElseIf UserName <> "" Then
        addParameter xmlQDBRequest, "username", UserName
        addParameter xmlQDBRequest, "password", Password
    ElseIf UserToken <> "" Then
        addParameter xmlQDBRequest, "usertoken", UserToken
    End If
    Set InitXMLRequest = xmlQDBRequest
End Function

Public Sub addParameter(ByRef xmlQDBRequest As MSXML.DOMDocument, ByVal Name As String, ByVal Value As String)
    Dim Root As IXMLDOMElement
    Dim ElementNode As IXMLDOMNode
    Dim TextNode As IXMLDOMNode

    Set Root = xmlQDBRequest.documentElement
    Set ElementNode = xmlQDBRequest.createNode(NODE_ELEMENT, Name, "")
    Set TextNode = xmlQDBRequest.createNode(NODE_TEXT, "", "")
    TextNode.nodeTypedValue = Value
    ElementNode.appendChild TextNode
    Root.appendChild ElementNode
    Set Root = Nothing
    Set ElementNode = Nothing
    Set TextNode = Nothing
End Sub
Public Sub addCDATAParameter(ByRef xmlQDBRequest As MSXML.DOMDocument, ByVal Name As String, ByVal Value As String)
    Dim Root As IXMLDOMElement
    Dim ElementNode As IXMLDOMNode
    Dim CDATANode As IXMLDOMNode

    Set Root = xmlQDBRequest.documentElement
    Set ElementNode = xmlQDBRequest.createNode(NODE_ELEMENT, Name, "")
    Set CDATANode = xmlQDBRequest.createNode(NODE_CDATA_SECTION, "", "")
    CDATANode.nodeTypedValue = Value
    ElementNode.appendChild CDATANode
    Root.appendChild ElementNode
    Set Root = Nothing
    Set ElementNode = Nothing
    Set CDATANode = Nothing
End Sub

Public Sub addFieldParameter(ByRef xmlQDBRequest As MSXML.DOMDocument, ByVal attrName As String, ByVal Name As String, ByVal Value As Variant)
    Dim Root As IXMLDOMElement
    Dim ElementNode As IXMLDOMNode
    Dim TextNode As IXMLDOMNode
    Dim attrField As IXMLDOMAttribute
    Dim attrFileName As IXMLDOMAttribute

    
    Set Root = xmlQDBRequest.documentElement
    Set ElementNode = xmlQDBRequest.createNode(NODE_ELEMENT, "field", "")
    Set attrField = xmlQDBRequest.createAttribute(attrName)
    attrField.Value = Name
    Call ElementNode.Attributes.setNamedItem(attrField)
    
            
    If TypeName(Value) = "File" Then
        Set attrFileName = xmlQDBRequest.createAttribute("filename")
        attrFileName.Value = Value.Name
        Call ElementNode.Attributes.setNamedItem(attrFileName)
    End If
        
    Set TextNode = xmlQDBRequest.createNode(NODE_TEXT, "", "")
    If TypeName(Value) = "File" Then
        TextNode.nodeTypedValue = fileEncode64(Value.Path)
    Else
        TextNode.nodeTypedValue = CStr(Value)
    End If
    ElementNode.appendChild TextNode
    
    Root.appendChild ElementNode
    Set Root = Nothing
    Set ElementNode = Nothing
    Set attrField = Nothing
    Set TextNode = Nothing
End Sub
Function int64ToDate(ByVal int64 As String) As Date
    If int64 = "" Then
        Exit Function
    End If
    Dim dblTemp As Double
    dblTemp = CDbl(int64)
    If dblTemp <= -59011459200001# Then
        int64ToDate = CDate("1/1/100")
    ElseIf dblTemp > 255611376000000# Then
        int64ToDate = CDate("12/31/9999")
    Else
        int64ToDate = DAYS_BETWEEN_JAVASCRIPT_AND_MICROSOFT_DATE_REFERENCES + (dblTemp / MILLISECONDS_IN_A_DAY)
    End If
End Function


Function makeAlphaNumLowerCase(ByVal strString As String) As String
    Dim i As Long
    Dim chrString As String
    Dim byteChar As Byte
    
    For i = 1 To Len(strString)
        byteChar = Asc(Mid(strString, i, 1))
        If IsCharAlphaNumeric(byteChar) Then
            makeAlphaNumLowerCase = makeAlphaNumLowerCase + Chr(byteChar)
        Else
            makeAlphaNumLowerCase = makeAlphaNumLowerCase + "_"
        End If
    Next i
    makeAlphaNumLowerCase = LCase(makeAlphaNumLowerCase)
End Function

Public Sub setServer(ByVal strHost As String, ByVal HTTPS As Boolean)
If strHost <> "" Then
    QDBHost = strHost
    useHTTPS = HTTPS
Else
    QDBHost = "www.quickbase.com"
    useHTTPS = True
End If
End Sub

Public Function getDBLastModified(ByVal dbid As String) As Date
    Dim qdbResponse As New MSXML.DOMDocument
    Dim strInt64Time As String
        
    Set qdbResponse = GetDBInfo(dbid)
    strInt64Time = qdbResponse.documentElement.selectSingleNode("/*/lastRecModTime").nodeTypedValue
    If Left(strInt64Time, 1) = "-" Then
        getDBLastModified = #1/1/1970#
    Else
        getDBLastModified = int64ToDate(qdbResponse.documentElement.selectSingleNode("/*/lastRecModTime").nodeTypedValue)
    End If
End Function

Function getCompleteCSVSnapshot(ByVal dbid As String) As String
    Dim FieldNodeList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlQDBRequest As New MSXML.DOMDocument
    Dim qdbResponse As New MSXML.DOMDocument
    Dim clist As String
        
    Set qdbResponse = GetSchema(dbid)
    Set FieldNodeList = qdbResponse.documentElement.selectNodes("/*/table/fields/field/@id")
    
    For Each xmlNode In FieldNodeList
        clist = clist + xmlNode.Text + "."
    Next xmlNode
    getCompleteCSVSnapshot = GenResultsTable(dbid, "{'0'.CT.''}", clist, "", "csv")
End Function
Function getRecordAsArray(ByVal dbid As String, ByVal clist As String, ByVal ridFID As String, ByVal rid As String, ByRef QDBRecord() As Variant) As String
    Dim xmlQDBResponse As New MSXML.DOMDocument
    Set xmlQDBResponse = DoQuery(dbid, "{'" + ridFID + "'.EX.'" + rid + "'}", clist, "", "")
    Dim strFieldValue As String
    Dim i As Long
    Dim FieldNodeList As MSXML.IXMLDOMNodeList
    Dim FieldDefNodeList As MSXML.IXMLDOMNodeList
    
    Set FieldDefNodeList = xmlQDBResponse.documentElement.selectNodes("/*/table/fields/field")
    Set FieldNodeList = xmlQDBResponse.documentElement.selectNodes("/*/table/records/record/f")
    
    ReDim QDBRecord(1, FieldNodeList.length - 1)
    For i = 0 To FieldNodeList.length - 1
        QDBRecord(0, i) = FieldNodeList(i).selectSingleNode("@id").nodeTypedValue
        strFieldValue = FieldNodeList(i).selectSingleNode(".").nodeTypedValue

        Select Case FieldDefNodeList(i).selectSingleNode("@base_type").Text
            Case "float"
                If strFieldValue <> "" Then
                    QDBRecord(1, i) = makeDouble(strFieldValue)
                End If
            Case "text"
                QDBRecord(1, i) = Replace(strFieldValue, Chr(10), vbCrLf)
            Case "bool"
                QDBRecord(1, i) = CBool(strFieldValue)
            Case "int64"
                If strFieldValue <> "" Then
                    If FieldDefNodeList(i).selectSingleNode("@field_type").Text <> "date" Then
                        QDBRecord(1, i) = CDbl(strFieldValue) / MILLISECONDS_IN_A_DAY
                    Else
                        QDBRecord(1, i) = int64ToDate(strFieldValue)
                    End If
                End If
            Case "int32"
                If FieldDefNodeList(i).selectSingleNode("@field_type").Text = "userid" Then
                    On Error Resume Next
                        Dim tempLong As Long
                        tempLong = CLng(strFieldValue)
                        If Err.Number = 0 Then
                            QDBRecord(1, i) = xmlQDBResponse.selectSingleNode("/*/table/lusers/luser[@id='" & strFieldValue & "']").Text
                        Else
                            QDBRecord(1, i) = strFieldValue
                        End If
                    On Error GoTo 0
                Else
                    QDBRecord(1, i) = CLng(strFieldValue)
                End If
        End Select
    Next i
    getRecordAsArray = xmlQDBResponse.documentElement.selectSingleNode("/*/table/records/record/f[@id=/*/table/fields/field[@role='modified']/@id]").nodeTypedValue
End Function
Private Function HTTPPost(QDBHost As String, useHTTPS As Boolean, script As String, content_type As String, content As String, returnFileName As String) As String
  If QDBHost = "" Then
    QDBHost = "www.quickbase.com"
    useHTTPS = True
  End If
  Const BLOCKSIZE = 65526
  Dim sReadBuffer         As String * BLOCKSIZE
  Dim lNumberOfBytesRead  As Long
  
  #If WININET Then
  
  Dim hInternetOpen As Long
  Dim hInternetConnect As Long
  Dim hHttpOpenRequest As Long
  Dim bRet As Boolean
  Dim httpPort As Long
  Dim fFlags As Long
  
  hInternetOpen = 0
  hInternetConnect = 0
  hHttpOpenRequest = 0
  
  
  'Use registry access settings.
  
  hInternetOpen = InternetOpen("VBQuickBaseAPI 1.0", _
                  INTERNET_OPEN_TYPE_PRECONFIG, _
                  vbNullString, _
                  vbNullString, _
                  0)
  
  If hInternetOpen <> 0 Then

     'Change the server to your server name
    If useHTTPS Then
        httpPort = INTERNET_DEFAULT_HTTPS_PORT
    Else
        httpPort = INTERNET_DEFAULT_HTTP_PORT
    End If
     hInternetConnect = InternetConnect(hInternetOpen, _
                        QDBHost, _
                        httpPort, _
                        vbNullString, _
                        "HTTP/1.0", _
                        INTERNET_SERVICE_HTTP, _
                        0, _
                        0)
  
     If hInternetConnect <> 0 Then
      'Brings the data across the wire even if it locally cached.
      fFlags = INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_COOKIES
      If useHTTPS Then
        fFlags = fFlags Or INTERNET_FLAG_SECURE Or INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or INTERNET_FLAG_IGNORE_CERT_DATE_INVALID Or INTERNET_FLAG_NO_COOKIES
      End If
      hHttpOpenRequest = HttpOpenRequest(hInternetConnect, _
                           "POST", _
                           script$, _
                           "HTTP/1.0", _
                           vbNullString, _
                           0, _
                           fFlags, _
                           0)
  
        If hHttpOpenRequest <> 0 Then
           Dim sHeader As String
           sHeader = "Content-Type: " + content_type & vbCrLf
           
           bRet = HttpAddRequestHeaders(hHttpOpenRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
           If Len(ticket) <> 0 Then
                sHeader = "Cookie: TICKET=" + ticket & vbCrLf
                bRet = HttpAddRequestHeaders(hHttpOpenRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
           End If
           Dim lpszPostData As String
           Dim lPostDataLen As Long
  
           lpszPostData = content
           lPostDataLen = Len(lpszPostData)
           bRet = HttpSendRequest(hHttpOpenRequest, _
                  vbNullString, _
                  0, _
                  lpszPostData, _
                  lPostDataLen)
  
           Dim bDoLoop As Boolean
           Dim sbuffer As String
           Dim contentLength As Long
           Dim lOffset As Long
           
           lNumberOfBytesRead = Len(sReadBuffer)
           sReadBuffer = "QUICKBASE-ERRCODE" + Chr(0)
           bRet = HttpQueryInfo(hHttpOpenRequest, HTTP_QUERY_CUSTOM, ByVal sReadBuffer, lNumberOfBytesRead, 0)
           If bRet Then
            errorcode = CInt(Left(sReadBuffer, lNumberOfBytesRead))
           Else
            errorcode = -1
           End If
           lNumberOfBytesRead = Len(sReadBuffer)
           sReadBuffer = "QUICKBASE-ERRTEXT" + Chr(0)
           bRet = HttpQueryInfo(hHttpOpenRequest, HTTP_QUERY_CUSTOM, ByVal sReadBuffer, lNumberOfBytesRead, 0)
           errortext = Left(sReadBuffer, lNumberOfBytesRead)
           
           lNumberOfBytesRead = Len(sReadBuffer)
           bRet = HttpQueryInfo(hHttpOpenRequest, HTTP_QUERY_CONTENT_LENGTH, ByVal sReadBuffer, lNumberOfBytesRead, 0)
           contentLength = Left(sReadBuffer, lNumberOfBytesRead)
           httpContentLength = contentLength
           
           If returnFileName <> "" Then
            Dim hFile As Long
            
            
            hFile = FreeFile
            Open returnFileName For Binary Access Write As hFile
            Seek hFile, 1
           Else
            sbuffer = String(contentLength, 0)
            lOffset = 0
            httpContentLengthProgress = 0
           End If
           bDoLoop = True
           Do While bDoLoop
            sReadBuffer = vbNullString
            bDoLoop = InternetReadFile(hHttpOpenRequest, _
               sReadBuffer, BLOCKSIZE, lNumberOfBytesRead)
            If Not CBool(lNumberOfBytesRead) Then Exit Do
            If returnFileName <> "" Then
                If lNumberOfBytesRead < BLOCKSIZE Then
                    Dim sWriteBuffer As String
                    sWriteBuffer = Left(sReadBuffer, lNumberOfBytesRead)
                    Put #hFile, , sWriteBuffer
                Else
                    Put #hFile, , sReadBuffer
                End If
            Else
                If lNumberOfBytesRead < BLOCKSIZE Then
                    Mid(sbuffer, lOffset + 1, lNumberOfBytesRead) = Left(sReadBuffer, lNumberOfBytesRead)
                Else
                    Mid(sbuffer, lOffset + 1, lNumberOfBytesRead) = sReadBuffer
                End If
            End If
            lOffset = lOffset + lNumberOfBytesRead
            httpContentLengthProgress = lOffset
            DoEvents
           Loop
           If returnFileName <> "" Then
            Close hFile
           End If
           bRet = InternetCloseHandle(hHttpOpenRequest)
        End If
        bRet = InternetCloseHandle(hInternetConnect)
     End If
     bRet = InternetCloseHandle(hInternetOpen)
  End If
  If returnFileName <> "" Then
    HTTPPost = returnFileName
    httpContentLengthProgress = 0
    httpContentLength = 0
    Exit Function
  Else
    HTTPPost = Left(sbuffer, contentLength)
    httpContentLengthProgress = 0
    httpContentLength = 0
  End If
  If errorcode = -1 Then
    On Error GoTo 0
    Err.Raise vbObjectError + 1000, "QuickBase.QuickBaseClient", script + ": Communications with " + QDBHost + " was interrupted or timed out or the SSL certificate on " + QDBHost + " expired or is invalid."
  End If
  #Else
    Dim URL As String
    Dim HttpReq As WinHttpRequest
    Set HttpReq = New WinHttpRequest
    HttpReq.Option(WinHttpRequestOption_UserAgentString) = "VBQuickBaseWinHTTPAPI 1.0"
    HttpReq.Option(WinHttpRequestOption_SSlErrorIgnoreFlags) = &H100 Or &H200 Or &H1000 Or &H2000    'Unknown certification authority (CA) or untrusted root, Wrong usage, Invalid common name (CN), Invalid date or certificate expired

    If useHTTPS Then
        URL = "https://" + QDBHost + script
    Else
        URL = "http://" + QDBHost + script
    End If
    HttpReq.Open "POST", URL, False
    HttpReq.setRequestHeader "Content-Type", content_type
    HttpReq.send vbCrLf + vbCrLf + content
    If returnFileName <> "" Then
        Dim hFile As Long
        hFile = FreeFile
        Open returnFileName For Binary Access Write As hFile
        Seek hFile, 1
        Put #hFile, , HttpReq.responseText
        Close hFile
        HTTPPost = returnFileName
    Else
        errorcode = CInt(HttpReq.getResponseHeader("QUICKBASE-ERRCODE"))
        errortext = HttpReq.getResponseHeader("QUICKBASE-ERRTEXT")
        HTTPPost = HttpReq.responseText
    End If
    Set HttpReq = Nothing
  #End If
  httpContentLengthProgress = 0
  httpContentLength = 0
End Function

Public Function makeDouble(strString As String) As Double
    Dim i As Long
    Dim chrString As String
    Dim byteChar As Byte
    Dim resultString As String
    
    On Error Resume Next
    makeDouble = CDbl(strString)
    If Err.Number = 0 Then
        Exit Function
    End If
    On Error GoTo 0
    resultString = ""
    For i = 1 To Len(strString)
        byteChar = Asc(Mid(strString, i, 1))
        If (((Not IsCharAlpha(byteChar)) And IsCharAlphaNumeric(byteChar)) Or byteChar = Asc(".") Or byteChar = Asc("-")) Then
            resultString = resultString + Chr(byteChar)
        End If
    Next i
    On Error Resume Next
    makeDouble = CDbl(resultString)
    Exit Function
End Function


Function fileEncode64(Path As String) As String
Dim triplicate As Long
Dim i As Long
Dim outputText As String
Dim hFile As Integer
Dim fileLength As Long
Dim fileTriads As Long
Dim firstByte As Byte
Dim secondByte As Byte
Dim thirdByte As Byte
Dim fileRemainder As Integer

hFile = FreeFile
Open Path For Binary Access Read As hFile
Seek hFile, 1

fileLength = LOF(hFile)
fileRemainder = fileLength Mod 3
fileTriads = fileLength \ 3
If fileRemainder > 0 Then
    outputText = Space((fileTriads + 1) * 4)
Else
    outputText = Space(fileTriads * 4)
End If


For i = 0 To fileTriads - 1             ' loop through octets
' build 24 bit triplicate
    Get #hFile, , firstByte
    Get #hFile, , secondByte
    Get #hFile, , thirdByte

    triplicate = (firstByte * 65536) + (secondByte * CLng(256)) + thirdByte
    ' extract four 6 bit quartets from triplicate
    Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & Mid(Map64, ((triplicate And 4032) \ 64) + 1, 1) & Mid(Map64, (triplicate And 63) + 1, 1)
Next                                                    ' next octet
Select Case fileRemainder
  Case 1
    Get #hFile, , firstByte
    triplicate = (firstByte * 65536)
    Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & "=="
  Case 2
    Get #hFile, , firstByte
    Get #hFile, , secondByte
    triplicate = (firstByte * 65536) + (secondByte * CLng(256))
    Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & Mid(Map64, ((triplicate And 4032) \ 64) + 1, 1) & "="
End Select
Close hFile
fileEncode64 = outputText
End Function

Function makeValidFilename(ByVal strString As String) As String
    Dim i As Integer
    Dim chrString As String
    Dim byteChar As String
    
    For i = 1 To Len(strString)
        byteChar = Mid(strString, i, 1)
        If byteChar = "\" Or byteChar = "/" Or _
           byteChar = ":" Or byteChar = "*" Or _
           Asc(byteChar) = 63 Or byteChar = """" Or _
           byteChar = "<" Or byteChar = ">" Or _
           byteChar = "|" Or byteChar = "'" _
        Then
            makeValidFilename = makeValidFilename + "_"
        Else
            makeValidFilename = makeValidFilename + byteChar
        End If
    Next i
End Function
Public Function getServerStatus() As MSXML.DOMDocument
        Dim xmlQDBRequest As New MSXML.DOMDocument
        Set xmlQDBRequest = InitXMLRequest()
        Set getServerStatus = APIXMLPost("main", "API_OBStatus", xmlQDBRequest, useHTTPS)
End Function
Public Function GetNumRecords(ByVal dbid As String) As Long
    Dim xmlQDBRequest As New MSXML.DOMDocument
        
    Set xmlQDBRequest = InitXMLRequest()
    GetNumRecords = CLng(APIXMLPost(dbid, "API_GetNumRecords", xmlQDBRequest, useHTTPS).selectSingleNode("/*/num_records").Text)
End Function

Public Function GetDBPage(ByVal dbid As String, ByVal page As String) As MSXML.DOMDocument
    Dim xmlQDBRequest As New MSXML.DOMDocument
        
    Set xmlQDBRequest = InitXMLRequest()
    If CStr(Val(page)) = page Then
        addParameter xmlQDBRequest, "pageid", page
    Else
        addParameter xmlQDBRequest, "pagename", page
    End If
    Set GetDBPage = APIXMLPost(dbid, "API_GetDBPage", xmlQDBRequest, useHTTPS)
End Function
 
Public Function AddReplaceDBPage(ByVal dbid As String, ByVal page As String, ByVal pagetype As String, ByVal pagebody As String) As MSXML.DOMDocument
    Dim xmlQDBRequest As New MSXML.DOMDocument
        
    Set xmlQDBRequest = InitXMLRequest()
    If CStr(Val(page)) = page Then
        addParameter xmlQDBRequest, "pageid", page
    Else
        addParameter xmlQDBRequest, "pagename", page
    End If
    addParameter xmlQDBRequest, "pagetype", pagetype
    addParameter xmlQDBRequest, "pagebody", pagebody
    Set AddReplaceDBPage = APIXMLPost(dbid, "API_AddReplaceDBPage", xmlQDBRequest, useHTTPS)
End Function

Public Function ListDBPages(ByVal dbid As String) As MSXML.DOMDocument
    Dim xmlQDBRequest As New MSXML.DOMDocument
        
    Set xmlQDBRequest = InitXMLRequest()
    Set ListDBPages = APIXMLPost(dbid, "API_ListDBPages", xmlQDBRequest, useHTTPS)
End Function











