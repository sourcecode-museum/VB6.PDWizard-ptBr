VERSION 5.00
Begin VB.Form frmBegin 
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
   Icon            =   "begin.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7305
   Begin VB.Frame fraDir 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   7095
      Begin VB.CommandButton cmdChDir 
         Caption         =   ". . ."
         Height          =   315
         Left            =   6240
         TabIndex        =   9
         Top             =   360
         Width           =   500
      End
      Begin VB.Label lblDestDir 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*"
         Height          =   315
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   5775
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
      Begin VB.CommandButton cmdInstall 
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
      Begin VB.Label lblBegin 
         BackColor       =   &H80000009&
         Caption         =   "lblTitle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Left            =   -40
      TabIndex        =   4
      Top             =   3800
      Width           =   7395
   End
   Begin VB.Label lblInstallMsg 
      AutoSize        =   -1  'True
      Caption         =   "*"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   6645
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub cmdChDir_Click()
    ShowPathDialog gstrDIR_DEST

    If gfRetVal = gintRET_CONT Then
        lblDestDir.Caption = gstrDestDir
        cmdInstall.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    ExitSetup Me, gintRET_EXIT
End Sub


Private Sub cmdInstall_Click()
    If IsValidDestDir(gstrDestDir) = True Then
        Unload Me
        DoEvents
    End If
End Sub

Private Sub Form_Load()
    SetFormFont Me
    Me.imgTitle.Picture = LoadResPicture(101, vbResBitmap)
    fraNavigate.Caption = gstrTitle
    fraDir.Caption = ResolveResString(resFRMDIRECTORY)
    cmdChDir.Caption = ResolveResString(resBTNCHGDIR)
    cmdExit.Caption = ResolveResString(resBTNEXIT)
    cmdInstall.Caption = ResolveResString(resBTNOK)
    lblBegin.Caption = ResolveResString(resLBLBEGIN)
    cmdInstall.ToolTipText = ResolveResString(resBTNTOOLTIPBEGIN)
    
    Caption = gstrTitle
    lblInstallMsg.Caption = ResolveResString(IIf(gfForceUseDefDest, resSPECNODEST, resSPECDEST), "|1", gstrAppName)
    lblDestDir.Caption = gstrDestDir

    If gfForceUseDefDest Then
        'We are forced to use the default destination directory, so the user
        '  will not be able to change it.
        cmdChDir.Visible = False
    End If
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
        .Name = Me.lblBegin.Font.Name
        .Charset = Me.lblBegin.Font.Charset
        .Size = Me.lblBegin.Font.Size
        .Bold = True
    End With

    Set Me.lblBegin.Font = otmpFont
    Set Me.lblInstallMsg.Font = otmpFont

End Sub


