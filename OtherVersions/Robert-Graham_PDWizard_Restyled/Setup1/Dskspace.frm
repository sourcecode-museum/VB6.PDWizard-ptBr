VERSION 5.00
Begin VB.Form frmDskSpace 
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
   Icon            =   "Dskspace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7305
   Begin VB.Frame fraNoSpace 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   6975
      Begin VB.Shape shpHeading 
         BorderColor     =   &H00000000&
         Height          =   480
         Left            =   1080
         Top             =   120
         Width           =   4980
      End
      Begin VB.Label lblReqH 
         Alignment       =   1  'Right Justify
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
         Height          =   405
         Left            =   1695
         TabIndex        =   15
         Top             =   180
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNeedH 
         Alignment       =   1  'Right Justify
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
         Height          =   405
         Left            =   4770
         TabIndex        =   14
         Top             =   180
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAvailH 
         Alignment       =   1  'Right Justify
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
         Height          =   405
         Left            =   3240
         TabIndex        =   13
         Top             =   180
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDiskH 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1125
         TabIndex        =   12
         Top             =   375
         Width           =   105
      End
      Begin VB.Shape shpSpace 
         BorderColor     =   &H00000000&
         Height          =   390
         Left            =   1080
         Top             =   600
         Width           =   4980
      End
      Begin VB.Label lblReq 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1695
         TabIndex        =   11
         Top             =   690
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblNeed 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   4770
         TabIndex        =   10
         Top             =   690
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblAvail 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   3225
         TabIndex        =   9
         Top             =   690
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblDisk 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1125
         TabIndex        =   8
         Top             =   690
         Visible         =   0   'False
         Width           =   510
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
      TabIndex        =   3
      Top             =   0
      Width           =   7335
      Begin VB.Image imgTitle 
         Height          =   615
         Left            =   120
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblNoSpace 
         BackColor       =   &H80000009&
         Caption         =   "lblTitle"
         Height          =   615
         Left            =   960
         TabIndex        =   4
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
      TabIndex        =   1
      Top             =   4005
      Width           =   7305
      Begin VB.CommandButton cmdChgDrv 
         Caption         =   "#"
         Height          =   312
         Left            =   3720
         MaskColor       =   &H00000000&
         TabIndex        =   6
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdInstall 
         Caption         =   "#"
         Height          =   312
         Left            =   4920
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "#"
         Height          =   312
         Left            =   6120
         MaskColor       =   &H00000000&
         TabIndex        =   2
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
      TabIndex        =   0
      Top             =   3800
      Width           =   7395
   End
End
Attribute VB_Name = "frmDskSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'
' Form/Module Constants
'
Const strFMT$ = "######0 K"

Private Sub cmdChgDrv_Click()
    gfRetVal = gintRET_CANCEL
    Unload Me
End Sub

Private Sub cmdExit_Click()
    ExitSetup Me, gintRET_EXIT
End Sub

Private Sub cmdInstall_Click()
    gfRetVal = gintRET_CONT
    Unload Me
End Sub

Private Sub Form_Load()
    Const ONE_K& = 1024

    Dim intIdx As Integer
    Dim lAvail As Long
    Dim lReq As Long
    Dim intHeight As Integer
    Dim intTop As Integer

    SetFormFont Me
    Me.imgTitle.Picture = LoadResPicture(101, vbResBitmap)
    fraNavigate.Caption = gstrTitle
    SetBoldCaptions

    cmdExit.Caption = ResolveResString(resBTNEXIT)
    cmdInstall.Caption = ResolveResString(resBTNINSTALLNOW)
    cmdChgDrv.Caption = ResolveResString(resBTNCHGDRV)
    lblDiskH.Caption = ResolveResString(resLBLDRIVE)
    lblAvailH.Caption = ResolveResString(resLBLAVAIL)
    lblNeedH.Caption = ResolveResString(resLBLNEEDED)
    lblReqH.Caption = ResolveResString(resLBLREQUIRED)
    lblNoSpace.Caption = ResolveResString(resLBLNOSPACE)
    frmDskSpace.Caption = gstrTitle

    intHeight = lblDisk(0).Height * 1.6
    intTop = lblDisk(0).Top

    '
    'borders are for design mode only...
    '
    lblDisk(0).BorderStyle = 0
    lblReq(0).BorderStyle = 0
    lblAvail(0).BorderStyle = 0
    lblNeed(0).BorderStyle = 0

    For intIdx = 1 To Len(gstrDrivesUsed)
        Load lblDisk(intIdx)
        Load lblReq(intIdx)
        Load lblAvail(intIdx)
        Load lblNeed(intIdx)

        lAvail = gsDiskSpace(intIdx).lAvail
        lReq = gsDiskSpace(intIdx).lReq

        lblDisk(intIdx).Caption = Mid$(gstrDrivesUsed, intIdx, 1) & gstrCOLON
        lblReq(intIdx).Caption = Format$(lReq / ONE_K, strFMT)
        lblAvail(intIdx).Caption = Format$(lAvail / ONE_K, strFMT)
        lblNeed(intIdx).Caption = Format$(IIf(lReq > lAvail, lReq - lAvail, 0) / ONE_K, strFMT)

        lblDisk(intIdx).Top = intTop
        lblReq(intIdx).Top = intTop
        lblAvail(intIdx).Top = intTop
        lblNeed(intIdx).Top = intTop

        intTop = intTop + intHeight

        lblDisk(intIdx).Visible = True
        lblReq(intIdx).Visible = True
        lblAvail(intIdx).Visible = True
        lblNeed(intIdx).Visible = True
    Next

    shpSpace.Height = intHeight * (intIdx - 1)
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
        .Name = Me.lblNoSpace.Font.Name
        .Charset = Me.lblNoSpace.Font.Charset
        .Size = Me.lblNoSpace.Font.Size
        .Bold = True
    End With

    Set Me.lblNoSpace.Font = otmpFont

End Sub

