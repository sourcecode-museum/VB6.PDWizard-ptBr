VERSION 5.00
Begin VB.Form frmCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   4620
   ClientLeft      =   870
   ClientTop       =   1530
   ClientWidth     =   7305
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Copy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7305
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3000
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
      TabIndex        =   5
      Top             =   0
      Width           =   7335
      Begin VB.Image imgTitle 
         Height          =   615
         Left            =   120
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblCopy 
         BackColor       =   &H80000009&
         Caption         =   "lblTitle"
         Height          =   615
         Left            =   960
         TabIndex        =   6
         Top             =   120
         Width           =   6135
      End
   End
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
         Caption         =   "&Finish"
         Height          =   312
         Left            =   6120
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "104"
         Top             =   120
         Width           =   1092
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
      Left            =   -40
      TabIndex        =   2
      Top             =   3800
      Width           =   7395
   End
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      FillColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   885
      ScaleHeight     =   330
      ScaleWidth      =   5535
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1905
      Width           =   5592
   End
   Begin VB.Label lblDestFile 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   885
      TabIndex        =   0
      Top             =   1500
      Width           =   5640
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
' frmCopy

Private bPause As Boolean
Private Sub cmdExit_Click()
    ExitSetup Me, gintRET_EXIT
End Sub

Private Sub Form_Load()
    SetFormFont Me
    Me.imgTitle.Picture = LoadResPicture(101, vbResBitmap)
    fraNavigate.Caption = gstrTitle

    cmdExit.Caption = ResolveResString(resBTNCANCEL)
    lblCopy.Caption = ResolveResString(resLBLDESTFILE)
    lblDestFile.Caption = vbNullString
    
    frmCopy.Caption = gstrTitle
    SetBoldCaptions
    Me.Refresh
    

    bPause = True
    Me.Timer1.Interval = 1000

    Do While bPause
        DoEvents
    Loop

End Sub

Private Sub Timer1_Timer()
    bPause = False
    Timer1.Interval = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        ExitSetup Me, gintRET_EXIT
        Cancel = 1
    End If

    bPause = True
    Timer1.Interval = 1000

    Do While bPause
        DoEvents
    Loop
End Sub

Private Sub SetBoldCaptions()
Dim otmpFont As StdFont
    Set otmpFont = New StdFont
    With otmpFont
        .Name = Me.lblCopy.Font.Name
        .Charset = Me.lblCopy.Font.Charset
        .Size = Me.lblCopy.Font.Size
        .Bold = True
    End With

    Set Me.lblCopy.Font = otmpFont

End Sub



