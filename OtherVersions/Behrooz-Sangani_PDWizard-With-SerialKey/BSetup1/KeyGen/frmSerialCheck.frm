VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSerialCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BSetup! KeyGen"
   ClientHeight    =   6330
   ClientLeft      =   4590
   ClientTop       =   3075
   ClientWidth     =   6420
   Icon            =   "frmSerialCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4560
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSaveTxt 
      Caption         =   "Save To File"
      Height          =   375
      Left            =   2040
      TabIndex        =   32
      Top             =   5880
      Width           =   2535
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   375
      Left            =   5760
      TabIndex        =   31
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   5400
      TabIndex        =   30
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   375
      Left            =   5040
      TabIndex        =   29
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   375
      Left            =   4680
      TabIndex        =   28
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Users"
      Height          =   375
      Left            =   3360
      TabIndex        =   27
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add User"
      Height          =   375
      Left            =   2040
      TabIndex        =   25
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtComments 
      Height          =   615
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox txtAddress 
      Height          =   855
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   2880
      TabIndex        =   18
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   390
      Left            =   4920
      TabIndex        =   7
      Top             =   2520
      Width           =   1260
   End
   Begin VB.TextBox txtOwnerKey 
      Height          =   345
      Index           =   0
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   2
      Top             =   2040
      Width           =   675
   End
   Begin VB.TextBox txtOwnerKey 
      Height          =   345
      Index           =   1
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2040
      Width           =   675
   End
   Begin VB.TextBox txtOwnerKey 
      Height          =   345
      Index           =   2
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2040
      Width           =   675
   End
   Begin VB.TextBox txtOwnerKey 
      Height          =   345
      Index           =   3
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2040
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   0
      ScaleHeight     =   6495
      ScaleWidth      =   1455
      TabIndex        =   8
      Top             =   0
      Width           =   1455
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Picture         =   "frmSerialCheck.frx":0442
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "KeyGen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BSetup!     ©Behrooz Sangani"
         ForeColor       =   &H80000011&
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   5880
         Width           =   1455
      End
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   3000
      TabIndex        =   1
      Top             =   1560
      Width           =   3165
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   390
      Left            =   3600
      TabIndex        =   6
      Top             =   2520
      Width           =   1260
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   390
      Left            =   5160
      TabIndex        =   33
      Top             =   5880
      Width           =   1020
   End
   Begin VB.TextBox txtsChars 
      Height          =   345
      Left            =   3000
      TabIndex        =   0
      ToolTipText     =   "You have to copy this code and include it in Registration Form"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   26
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "Comments:"
      Height          =   255
      Left            =   1680
      TabIndex        =   24
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Address:"
      Height          =   255
      Left            =   1680
      TabIndex        =   23
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "User Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   22
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Email:"
      Height          =   255
      Left            =   1680
      TabIndex        =   21
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblSep 
      Caption         =   "_"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblSep 
      Caption         =   "_"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblSep 
      Caption         =   "_"
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   15
      Top             =   2040
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   1560
      X2              =   1560
      Y1              =   0
      Y2              =   6600
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Name:"
      Height          =   270
      Index           =   0
      Left            =   1680
      TabIndex        =   14
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Registration Key:"
      Height          =   270
      Index           =   1
      Left            =   1680
      TabIndex        =   13
      Top             =   2040
      Width           =   1320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000011&
      X1              =   1560
      X2              =   6720
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "PC ID:"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblDesc 
      Height          =   975
      Left            =   1680
      TabIndex        =   11
      Top             =   120
      Width           =   4455
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   1560
      X2              =   6600
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   -360
      Y2              =   6480
   End
End
Attribute VB_Name = "frmSerialCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  Registration Key Generator
'  Makes unique registration key depended on ID and Name
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 05/06/2002
'  WebSite: http://www.geocities.com/bs20014/
'  Legal Copyright: Behrooz Sangani © 05/06/2002
'=========================================================================================
'This is the Reg Key maker. You can simply manage registered users
'with database utilities provided.

'This application only works when the encryption sub
'in this app is the same in setup project. If ever you
'decided to change the encryption algorithm you must
'modify this project too.

Dim Reg As New clsValidator         'Key validator
Dim Success As Boolean              'Validation success
Dim sKey As String                  'Reg Key
'database
Dim db As Database
Dim rs As Recordset
Dim strSQL As String
'=========================================================================================
Private Sub cmdAdd_Click()
    Dim RegKey As String
    Dim aRs As Recordset

    RegKey = UCase(txtOwnerKey(0).Text & _
        txtOwnerKey(1).Text & _
        txtOwnerKey(2).Text & _
        txtOwnerKey(3).Text)

    If txtsChars.Text = "" Or txtUserName.Text = "" _
        Or Len(RegKey) <> 16 Or txtEmail.Text = "" Then
        'Required fields
        MsgBox "Incomplete data. Cannot add information to database..." & vbCrLf & _
            "Required fields: PC ID, User Name, Reg Key and Email", vbCritical, App.Title
        Exit Sub
    End If
 
    'Check for duplicate PC ID. If False add data...
    If Not DupFound(txtsChars.Text) Then
        Set aRs = db.OpenRecordset("Users", dbOpenTable)

        aRs.AddNew

        aRs.Fields("PCID") = txtsChars.Text
        aRs.Fields("User Name") = txtUserName.Text
        aRs.Fields("Reg Key") = RegKey
        aRs.Fields("Email") = txtEmail.Text
        aRs.Fields("Address") = txtAddress.Text
        aRs.Fields("Comments") = txtComments.Text
        aRs.Fields("Reg Date") = lblDate.Caption

        aRs.Update
        
        Set aRs = Nothing
    Else
        MsgBox "PC ID already exists in the database. Duplicate found!", vbCritical, App.Title
    End If
End Sub 'cmdAdd_Click()
'=========================================================================================
Private Sub cmdBack_Click()
    If Not rs.BOF Then
        rs.MovePrevious
        PopulateFields
    End If
End Sub 'cmdBack_Click()
'=========================================================================================
Private Sub cmdClose_Click()
    Unload Me
    End
End Sub 'cmdClose_Click()
'=========================================================================================
Private Sub cmdFirst_Click()
    rs.MoveFirst
    PopulateFields
End Sub 'cmdFirst_Click()
'=========================================================================================
Private Sub cmdGenerate_Click()
    Dim oReg As New COwnerRegistration
    Dim oVal As New clsValidator
    Dim sKey As String
    
    'Generate Reg Key with the encrypted PC ID
    sKey = oReg.GenerateKey(txtUserName.Text, oVal.Encrypt(txtsChars.Text))

    txtOwnerKey(0).Text = Left(sKey, 4)
    txtOwnerKey(1).Text = Mid(sKey, 5, 4)
    txtOwnerKey(2).Text = Mid(sKey, 9, 4)
    txtOwnerKey(3).Text = Mid(sKey, 13, 4)


    Set oReg = Nothing
    Set oVal = Nothing

End Sub 'cmdGenerate_Click()
'=========================================================================================
Private Sub cmdLast_Click()
    rs.MoveLast
    PopulateFields
End Sub 'cmdLast_Click()
'=========================================================================================
Private Sub cmdLoad_Click()
    'Load/Unload database recordsets
    If rs.RecordCount <> 0 Then
        If cmdLoad.Caption = "Load Users" Then
            EnableFields False
            cmdLoad.Caption = "Close"
            rs.MoveLast
            rs.MoveFirst
            PopulateFields
        ElseIf cmdLoad.Caption = "Close" Then
            EnableFields True
            cmdLoad.Caption = "Load Users"
            lblDate.Caption = Format(Date, "dd/mm/yy")
            For Each Control In Me
                If TypeOf Control Is TextBox Then Control.Text = ""
            Next
            rs.MoveLast
            rs.MoveFirst
        End If
    Else
        MsgBox "No user found in database!", vbExclamation, App.Title
    End If
End Sub 'cmdLoad_Click()
'=========================================================================================
Private Sub cmdNext_Click()
    If Not rs.EOF Then
        rs.MoveNext
        PopulateFields
    End If
End Sub 'cmdNext_Click()
'=========================================================================================
Private Sub cmdSaveTxt_Click()
    On Error GoTo error
    With CD1
        .CancelError = True
        .DialogTitle = "Save Registration Info To..."
        .Filter = "Text Files (*.txt) |*.txt"
        .Flags = &H2
        .InitDir = AppPath
        .ShowSave
        If Len(.FileName) <> 0 Then
            DoSave .FileName
        End If
    End With
error:
End Sub 'cmdSaveTxt_Click()
'=========================================================================================
Sub DoSave(sPath As String)
    'This is just to simplify things
    rgk = txtOwnerKey(0).Text & "-" & _
        txtOwnerKey(1).Text & "-" & _
        txtOwnerKey(2).Text & "-" & _
        txtOwnerKey(3).Text

    F = FreeFile
    Open sPath For Output As F
    Print #F, "=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
    Print #F, "BSetup! Registration Info"
    Print #F, "=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
    Print #F,
    Print #F, lblDate.Caption
    Print #F,
    Print #F, "PC ID: " & txtsChars.Text
    Print #F, "User Name: " & txtUserName.Text
    Print #F, "Registration Key: " & rgk
    Print #F, "Email: " & txtEmail.Text
    Print #F,
    Print #F, "Address: " & vbCrLf & txtAddress.Text
    Print #F,
    Print #F, "Comments: " & txtComments.Text
    Close F

End Sub 'DoSave(sPath As String)
'=========================================================================================
Private Sub cmdTest_Click()
    On Error Resume Next
    'Combine text boxes to get the key
    sKey = UCase(txtOwnerKey(0).Text & _
        txtOwnerKey(1).Text & _
        txtOwnerKey(2).Text & _
        txtOwnerKey(3).Text)

    'check for correct key
    Dim oReg As COwnerRegistration
    Dim oVal As New clsValidator

    Set oReg = New COwnerRegistration

    'check for correct key
    If oReg.IsKeyOK(sKey, txtUserName.Text, oVal.Encrypt(txtsChars.Text)) Then
        Success = True
    Else
        Success = False
    End If

    Set oReg = Nothing

    MsgBox "Validation: " & Success

End Sub 'cmdTest_Click()
'=========================================================================================
Private Sub Form_Load()

    lblDesc.Caption = "Enter PC ID and username to generate a Registration Key. Please note that the encryption method must be exactly the method used in the setup program." & _
        vbCrLf & "Then add more information and add user to users database."
    lblDate.Caption = Format(Date, "dd/mm/yy")
    
    EnableFields True
    
    OpenDB

End Sub 'Form_Load()
'=========================================================================================
Sub OpenDB()
    On Error Resume Next
    'I set the database password for security. open exclusive
    'the database and set your own password.
    Set db = OpenDatabase(AppPath & "RegUsers.mdb", False, False, ";pwd=BSetup!")
    strSQL = "select * from users"
    Set rs = db.OpenRecordset(strSQL)
    rs.MoveLast
    rs.MoveFirst
End Sub 'OpenDB()
'=========================================================================================
Sub CloseDB()
    On Error Resume Next
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
End Sub 'CloseDB()
'=========================================================================================
Sub EnableFields(YesNo As Boolean)
    'Lets disable and enable controls on loading database
    Select Case YesNo
        Case False
            For Each Control In Me
                If TypeOf Control Is TextBox Then Control.Locked = True
            Next
            cmdAdd.Enabled = False
            cmdNext.Enabled = True
            cmdBack.Enabled = True
            cmdLast.Enabled = True
            cmdFirst.Enabled = True
        Case True
            For Each Control In Me
                If TypeOf Control Is TextBox Then Control.Locked = False
            Next
            cmdAdd.Enabled = True
            cmdNext.Enabled = False
            cmdBack.Enabled = False
            cmdLast.Enabled = False
            cmdFirst.Enabled = False
    End Select
End Sub 'EnableFields(YesNo As Boolean)
'=========================================================================================
Function DupFound(sID As String) As Boolean
    'Check database for duplicate PC ID to see if user is already registered
    Dim dRs As Recordset
    Set dRs = db.OpenRecordset("select * from users where " & "PCID" & " like '" & sID & "'")
    If dRs.RecordCount <> 0 Then
        DupFound = True
    Else
        DupFound = False
    End If
    Set dRs = Nothing
End Function 'DupFound(sID As String) As Boolean
'=========================================================================================
Sub PopulateFields()
    'show recordset fields
    On Error Resume Next
    Dim RegKey As String
    txtsChars.Text = rs.Fields("PCID")
    txtUserName.Text = rs.Fields("User Name")
    RegKey = rs.Fields("Reg Key")
    txtOwnerKey(0).Text = Left(RegKey, 4)
    txtOwnerKey(1).Text = Mid(RegKey, 5, 4)
    txtOwnerKey(2).Text = Mid(RegKey, 9, 4)
    txtOwnerKey(3).Text = Mid(RegKey, 13, 4)
    txtEmail.Text = rs.Fields("Email")
    txtAddress.Text = rs.Fields("Address")
    txtComments.Text = rs.Fields("Comments")
    lblDate.Caption = "Date Saved: " & rs.Fields("Reg Date")
End Sub 'PopulateFields()
'=========================================================================================
Private Sub Form_Unload(Cancel As Integer)
    CloseDB
End Sub 'Form_Unload(Cancel As Integer)
'=========================================================================================
