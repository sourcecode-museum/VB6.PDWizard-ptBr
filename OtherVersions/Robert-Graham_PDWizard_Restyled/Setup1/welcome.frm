VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   4620
   ClientLeft      =   1740
   ClientTop       =   1410
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "welcome.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7305
   Begin VB.PictureBox picNavigate 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7305
      TabIndex        =   3
      Top             =   4005
      Width           =   7305
      Begin VB.CommandButton cmdExit 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   4920
         MaskColor       =   &H00000000&
         TabIndex        =   6
         Tag             =   "103"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   6120
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
   End
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7335
      TabIndex        =   1
      Top             =   0
      Width           =   7335
      Begin VB.Label lblWelcome 
         BackColor       =   &H80000009&
         Caption         =   "lblTitle"
         Height          =   615
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   6135
      End
      Begin VB.Image imgTitle 
         Height          =   615
         Left            =   120
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame fraNavigate 
      Caption         =   "Frame1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   -30
      TabIndex        =   4
      Top             =   3800
      Width           =   7395
   End
   Begin VB.Label lblWelcome2 
      Height          =   1335
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   6645
   End
   Begin VB.Label lblRunning 
      AutoSize        =   -1  'True
      Caption         =   "*"
      Height          =   795
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   6645
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'frmWelcome
Private Sub cmdExit_Click()
    ExitSetup Me, gintRET_EXIT
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim intWidth As Integer
    Me.imgTitle.Picture = LoadResPicture(101, vbResBitmap)
    fraNavigate.Caption = gstrTitle

    SetFormFont Me
    cmdExit.Caption = ResolveResString(resBTNEXIT)
    cmdOK.Caption = ResolveResString(resBTNSTART)
    lblRunning.Caption = ResolveResString(resLBLRUNNING)
    
    Caption = gstrTitle
    intWidth = TextWidth(Caption) + cmdOK.Width * 2
    If intWidth > Width Then
        Width = intWidth
    End If

    lblWelcome.Caption = ResolveResString(resWELCOME, "|1", gstrAppName)
    lblWelcome2.Caption = ResolveResString(201)
    SetBoldCaptions
    CenterForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        ExitSetup Me, gintRET_EXIT
        Cancel = 1
    End If
End Sub

Private Sub SetBoldCaptions()
Dim otmpFont As StdFont
    Set otmpFont = New StdFont
    With otmpFont
        .Name = Me.lblRunning.Font.Name
        .Charset = Me.lblRunning.Font.Charset
        .Size = Me.lblRunning.Font.Size
        .Bold = True
    End With

    Set Me.lblRunning.Font = otmpFont
    Set Me.lblWelcome.Font = otmpFont
    Set Me.lblWelcome2.Font = otmpFont
End Sub

