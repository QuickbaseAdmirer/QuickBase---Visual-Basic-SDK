VERSION 5.00
Begin VB.Form TestPanel 
   Caption         =   "QuickBaseClient Test Panel"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   2880
      TabIndex        =   22
      Top             =   2160
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   840
      TabIndex        =   21
      Top             =   3240
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   840
      TabIndex        =   20
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtServer 
      Height          =   288
      Left            =   2280
      TabIndex        =   18
      Text            =   "www.quickbase.com"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtErrorText 
      Height          =   288
      Left            =   5280
      TabIndex        =   15
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox txtErrorCode 
      Height          =   288
      Left            =   7680
      TabIndex        =   14
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtRid 
      Height          =   288
      Left            =   2280
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtDBName 
      Height          =   288
      Left            =   5520
      TabIndex        =   8
      Text            =   "QuickBase VB API Demo"
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtdbid 
      Height          =   288
      Left            =   5520
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtPassword 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "Password"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtUsername 
      Height          =   288
      Left            =   2280
      TabIndex        =   3
      Text            =   "Depositor"
      Top             =   210
      Width           =   1695
   End
   Begin VB.TextBox txtResult 
      Height          =   1815
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3360
      Width           =   3375
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   372
      Left            =   6360
      TabIndex        =   1
      Top             =   1680
      Width           =   1092
   End
   Begin VB.ComboBox cmbAction 
      Height          =   315
      ItemData        =   "TestPanel.frx":0000
      Left            =   5520
      List            =   "TestPanel.frx":0013
      TabIndex        =   0
      Text            =   "Choose an API Call!"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label UserTokenTip 
      Caption         =   "Omit for User Token"
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblResult 
      Caption         =   "Result of API Call"
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblFile 
      Caption         =   "Pick a file to Upload"
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblServer 
      Caption         =   "Host or IP Address"
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblErrorText 
      Caption         =   "Error Text"
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblErrcode 
      Caption         =   "Error Code"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblDBID 
      Caption         =   "Database DBID"
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblDBName 
      Caption         =   "Database Name"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblAction 
      Caption         =   "Action"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblRid 
      Caption         =   "Record ID#"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      Caption         =   "Username/UserToken"
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "TestPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'© 2001 Intuit Inc. All rights reserved.
'Use is subject to the IP Rights Notice and Restrictions available at
'http://developer.intuit.com/legal/IPRNotice_021201.html

Dim QDB As QuickBase.QuickBaseClient

Private Sub cmdSubmit_Click()
Dim objFs As Object
Dim objFile As Object
Dim RecordArray(2, 4)
Dim xmlQDBResponse As New MSXML.DOMDocument
Dim strUpdateID As String
Set QDB = New QuickBase.QuickBaseClient
QDB.setServer txtServer.Text, True
If txtPassword.Text = "" Then
    txtResult.Text = QDB.Authenticate(, , txtUsername.Text)
Else
    txtResult.Text = QDB.Authenticate(txtUsername.Text, txtPassword.Text)
End If

DoEvents
Select Case cmbAction.Text

    Case "API_FindDBByName"
       'I don't recommend using this call.
       'It's much better practice to either hard code the database identifier by finding it manually.
       'Please read https://www.quickbase.com/db/6mztyxu8?a=dr&r=w to learn how to manually find the
       'database identifier of a QuickBase table.
       txtdbid.Text = QDB.FindDBByName(txtDBName.Text)
    Case "API_CloneDatabase"
       txtResult.Text = QDB.CloneDatabase(txtdbid.Text, "Delete Please VB API testing", "Please Delete")
       txtdbid.Text = txtResult.Text
    Case "API_AddRecord"
        'First we'll add a record with the function call that takes a variable number of arguments
        txtResult.Text = QDB.AddRecord(txtdbid.Text, "", "Assigned to", "Barney", 1006, "set by a fid!")
        'Now we'll use the call that takes a two dimensional array
        RecordArray(0, 0) = "Assigned to"
        RecordArray(1, 0) = "Barney"
        RecordArray(0, 1) = 1006
        RecordArray(1, 1) = "set by a fid!"
        RecordArray(0, 2) = "Status"
        RecordArray(1, 2) = "Completed"
        RecordArray(0, 3) = "Analysis"
        RecordArray(1, 3) = "set by a field name!"
        RecordArray(0, 4) = 1017
        
        Set objFs = CreateObject("Scripting.FileSystemObject")
        If File1.FileName <> "" Then
            If Mid(File1.Path, Len(File1.Path)) = "\" Then
                Set objFile = objFs.GetFile(File1.Path & File1.FileName)
            Else
                Set objFile = objFs.GetFile(File1.Path & "\" & File1.FileName)
            End If
            Set RecordArray(1, 4) = objFile
        End If

        
        
       txtResult.Text = QDB.AddRecordByArray(txtdbid.Text, strUpdateID, RecordArray())
       txtRid.Text = txtResult.Text
    Case "API_EditRecord"
       txtResult.Text = QDB.EditRecord(txtdbid.Text, txtRid.Text, "", "Assigned to", "Fred", 1006, "Edited by Edit Record")
    Case "API_DoQuery"
        Set xmlQDBResponse = QDB.DoQuery(txtdbid.Text, "1", "", "", "")
        Dim i As Integer
        Dim j As Integer
        Dim FieldNodeList
        Dim RecordNodeList
        Dim RecordNode
        Set FieldNodeList = xmlQDBResponse.documentElement.selectNodes("/*/table/fields/field")
        Set RecordNodeList = xmlQDBResponse.documentElement.selectNodes("/*/table/records/record")
        For i = 0 To RecordNodeList.length - 1
            Set RecordNode = RecordNodeList.nextNode()
            Dim FieldValues
            Set FieldValues = RecordNode.selectNodes("f")
            Dim FieldNode
            For j = 0 To FieldValues.length - 1
                Set FieldNode = FieldValues.nextNode()
                txtResult.Text = txtResult.Text + "Field Name: " + FieldNodeList(j).selectSingleNode("label").nodeTypedValue
                txtResult.Text = txtResult.Text + " Field Value: " + FieldNode.selectSingleNode(".").nodeTypedValue + vbCrLf
            Next j
        Next i
        Dim ResultArray() As Variant
        ResultArray = QDB.DoQueryAsArray(txtdbid.Text, "{'0'.CT.''}", "", "", "")
        For i = 0 To UBound(ResultArray, 1)
            For j = 0 To UBound(ResultArray, 2)
                txtResult.Text = txtResult.Text + " " + CStr(ResultArray(i, j))
            Next j
            txtResult.Text = txtResult.Text + vbCrLf
        Next i
    Case Else
        MsgBox "Please Choose an API Call!"
    End Select

txtErrorCode.Text = Format(QDB.errorcode)
txtErrorText.Text = Format(QDB.errortext)
Set QDB = Nothing
End Sub


