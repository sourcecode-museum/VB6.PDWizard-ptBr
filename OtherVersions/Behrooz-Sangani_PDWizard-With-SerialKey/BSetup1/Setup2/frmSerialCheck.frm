VERSION 5.00
Begin VB.Form frmSerialCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "%AppTitle%"
   ClientHeight    =   4650
   ClientLeft      =   4590
   ClientTop       =   3075
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOwnerKey 
      Height          =   345
      Index           =   0
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   1
      Top             =   3240
      Width           =   675
   End
   Begin VB.TextBox txtOwnerKey 
      Height          =   345
      Index           =   1
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   2
      Top             =   3240
      Width           =   675
   End
   Begin VB.TextBox txtOwnerKey 
      Height          =   345
      Index           =   2
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   3
      Top             =   3240
      Width           =   675
   End
   Begin VB.TextBox txtOwnerKey 
      Height          =   345
      Index           =   3
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   4
      Top             =   3240
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   1455
      TabIndex        =   8
      Top             =   0
      Width           =   1455
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Registration"
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
      Begin VB.Image imgWelcome 
         Height          =   480
         Left            =   480
         Picture         =   "frmSerialCheck.frx":0000
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BSetup!     ©Behrooz Sangani"
         ForeColor       =   &H80000011&
         Height          =   495
         Left            =   0
         MouseIcon       =   "frmSerialCheck.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   4200
         Width           =   1455
      End
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   3000
      TabIndex        =   0
      Top             =   2760
      Width           =   3165
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   390
      Left            =   5040
      TabIndex        =   5
      Top             =   4080
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3720
      TabIndex        =   6
      Top             =   4080
      Width           =   1140
   End
   Begin VB.TextBox txtsChars 
      BackColor       =   &H8000000F&
      Height          =   345
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "You have to copy this code and include it in Registration Form"
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label lblSep 
      Caption         =   "_"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   17
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblSep 
      Caption         =   "_"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   16
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblSep 
      Caption         =   "_"
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   15
      Top             =   3240
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   1560
      X2              =   1560
      Y1              =   0
      Y2              =   5160
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Name:"
      Height          =   270
      Index           =   0
      Left            =   1680
      TabIndex        =   14
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Registration Key:"
      Height          =   270
      Index           =   1
      Left            =   1680
      TabIndex        =   13
      Top             =   3240
      Width           =   1320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000011&
      X1              =   1560
      X2              =   6720
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label1 
      Caption         =   "PC ID:"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblDesc 
      Height          =   1815
      Left            =   1680
      TabIndex        =   11
      Top             =   360
      Width           =   4455
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   1560
      X2              =   6600
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   -360
      Y2              =   4680
   End
End
Attribute VB_Name = "frmSerialCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  Registration Form
'  Adds custom registration to VB default setup program (Package & Deployment)
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 05/06/2002
'  WebSite: http://www.geocities.com/bs20014/
'  Legal Copyright: Behrooz Sangani © 05/06/2002
'=========================================================================================
'Freeware under only one condition:
'   Leave the credit label and it's reference to web

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim Reg As New clsValidator         'Key validator
Dim Success As Boolean              'Validation success
Dim sKey As String                  'Reg Key
'=========================================================================================
Private Sub cmdCancel_Click()
    'On cancel use setup unload process
    Success = False
    ExitSetup Me, gintRET_EXIT
End Sub 'cmdCancel_Click()
'=========================================================================================
Private Sub cmdOK_Click()
    On Error Resume Next
    'Combine text boxes to get the key
    sKey = UCase(txtOwnerKey(0).Text & _
        txtOwnerKey(1).Text & _
        txtOwnerKey(2).Text & _
        txtOwnerKey(3).Text)

    'check for correct key
    If Reg.SerialValidation(txtUserName.Text, sKey) Then
        'on success add information to registry and load setup welcome form
        AddToReg
        Success = True
        Unload Me
        frmWelcome.Show vbModal
    Else
        MsgBox "Invalid Registration Key, please try again!" & vbCrLf & "If you have trouble with registration please contact support team.", , "Invalid key"
        txtOwnerKey(0).SetFocus
        SendKeys "{Home}+{End}"
        Success = False
    End If

    Set Reg = Nothing

End Sub 'cmdOK_Click()
'=========================================================================================
Private Sub Form_Load()
    On Error Resume Next
    'Our caption same as AppTitle
    Caption = gstrAppName
    'Give the user his unique PC Code
    txtsChars = Reg.SpecificChars
    lblDesc.Caption = "In order to continue setup you have to provide your Registration Key. Please enter your name and registration key exactly as it is provided." & _
        vbCrLf & vbCrLf & "If you do not have a registration key yet copy the PC ID in the box below and include it in your Registration Form. Then exit setup and complete your registration. Once you obtained a key you can run setup and continue installation."
    'setup method to center forms
    CenterForm Me
End Sub 'Form_Load()
'=========================================================================================
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCredit.ForeColor = &H80000011
End Sub 'Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=========================================================================================
Private Sub lblCredit_Click()
    'Please leave this line as it is.
    Call ShellExecute(0&, vbNullString, "http://www.geocities.com/bs20014/", vbNullString, vbNullString, vbNormalFocus)
End Sub 'lblCredit_Click()
'=========================================================================================
Private Sub lblCredit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCredit.ForeColor = vbRed
End Sub 'lblCredit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=========================================================================================
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCredit.ForeColor = &H80000011
End Sub 'Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=========================================================================================
Private Sub txtOwnerKey_Change(Index As Integer)
    On Error Resume Next
    EnableOK
    'If we have more than 4 characters set focus on the next text box
    If Len(txtOwnerKey(Index)) = 4 Then
        If Index = 3 Then
            EnableOK
        Else
            txtOwnerKey(Index + 1).SetFocus
            SendKeys "{Home}+{End}"
        End If
    ElseIf Len(txtOwnerKey(Index)) < 4 Then
        cmdOK.Enabled = False
    End If
End Sub 'txtOwnerKey_Change(Index As Integer)
'=========================================================================================
Sub EnableOK()
    On Error Resume Next
    'check all textboxes to see if we must enable OK button
    If Len(txtUserName) <> 0 And Len(txtOwnerKey(0)) = 4 _
        And Len(txtOwnerKey(1)) = 4 _
        And Len(txtOwnerKey(2)) = 4 _
        And Len(txtOwnerKey(3)) = 4 Then
        
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub 'EnableOK()
'=========================================================================================
Private Sub txtUserName_Change()
    EnableOK
End Sub 'txtUserName_Change()
'=========================================================================================
Private Sub AddToReg()
    On Error GoTo error
    
    Dim Ret As Long
    Dim bRet As Boolean
    
    'Registry handling exists in the setup app, so we use them
    'We set three values in the registry under the path
    'HKEY_CURRENT_USER\Software\%AppName%
    'UserName   &   RegKey      &   PCID
    'PC ID is a unique number generated from hard disk serial number
    'as explained in the class files
    bRet = RegCreateKey(HKEY_CURRENT_USER, "Software", gstrAppName, Ret)
    If bRet Then
        'If you wish to remain registered after app removal you must
        'use False else set the values to True to remove registry entries
        'after uninstall
        RegSetStringValue Ret, "UserName", txtUserName.Text, True
        RegSetStringValue Ret, "RegKey", sKey, True
        RegSetStringValue Ret, "PCID", txtsChars.Text, True
    Else
        GoTo error
    End If

    RegCloseKey Ret

    Exit Sub
error:
    'Give error message and exit setup
    MsgBox "Error adding info to registry!", vbCritical, gstrAppName
    ExitSetup Me, gintRET_FATAL
End Sub 'AddToReg()
'=========================================================================================
