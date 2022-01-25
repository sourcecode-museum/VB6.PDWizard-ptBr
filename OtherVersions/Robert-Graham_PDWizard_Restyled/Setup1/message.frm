VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   900
   ClientLeft      =   1065
   ClientTop       =   1995
   ClientWidth     =   5340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "message.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   5340
   Begin VB.PictureBox picMessage 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H80000005&
      ForeColor       =   &H80000002&
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   705
      TabIndex        =   1
      Top             =   0
      Width           =   705
      Begin VB.Image imgMsg 
         Height          =   480
         Left            =   80
         Picture         =   "message.frx":0442
         Top             =   160
         Width           =   480
      End
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "*"
      Height          =   195
      Left            =   945
      TabIndex        =   0
      Top             =   360
      Width           =   4110
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
' frmMessage
Private Sub Form_Load()
    SetFormFont Me
    SetBoldCaptions
End Sub
Private Sub SetBoldCaptions()
Dim otmpFont As StdFont
    Set otmpFont = New StdFont
    With otmpFont
        .Name = Me.lblMsg.Font.Name
        .Charset = Me.lblMsg.Font.Charset
        .Size = Me.lblMsg.Font.Size
        .Bold = True
    End With
    Set Me.lblMsg.Font = otmpFont

End Sub



