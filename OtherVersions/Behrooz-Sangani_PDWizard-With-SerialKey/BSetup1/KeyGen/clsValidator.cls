VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=========================================================================================
'  Key Validator Class
'  Validates the key based on the pc chars
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 05/06/2002
'  WebSite: http://www.geocities.com/bs20014/
'  Legal Copyright: Behrooz Sangani © 05/06/2002
'=========================================================================================
'Key Generation adopted from: Phil Fresle as described in classes

Private Const VALID_CHARS As String = "0123456789ABCDEFGHJKLMNPQRTUVWXY"
'=========================================================================================
Public Function SerialValidation(sName As String, sKey As String) As Boolean
    
    Dim oReg As COwnerRegistration

    Set oReg = New COwnerRegistration

    'check for correct key
    If oReg.IsKeyOK(sKey, sName, Encrypt(SpecificChars)) Then
        SerialValidation = True
    Else
        SerialValidation = False
    End If

    Set oReg = Nothing

End Function 'SerialValidation(sName As String, sKey As String) As Boolean
'=========================================================================================
Public Function SpecificChars() As String
    'avoid simple crackings
    'You may change this to your own secure chars.
    'The longer the characters the hradrer the cracking

    'This method takes Hard Disk Serial Number so there
    'is no way to setup the program on two PCs with the
    'same reg key

    Dim HD As String
    Dim SC As String
    Dim SH As String
    Dim cn As Integer
    Dim MD As New CMD5
    Dim sMD As String
    
    'Get the hard disk serial number
    HD = CStr(HardSerial)
    
    'Specific Characters:  AjbSreG2Le54Takk90sE3cFF
    'don't simply put your characters as a string. The disassembling
    'is then very easy.
    SC = Chr(65) & Chr(106) & Chr(98) & Chr(83) & _
        Chr(114) & Chr(101) & Chr(71) & Chr(50) & _
        Chr(76) & Chr(101) & Chr(53) & Chr(52) & _
        Chr(84) & Chr(97) & Chr(107) & Chr(107) & _
        Chr(57) & Chr(48) & Chr(115) & Chr(69) & _
        Chr(51) & Chr(99) & Chr(70) & Chr(70)
    
    'MD5 our chars and the HD serial
    sMD = MD.MD5(HD & SC)
    
    Set MD = Nothing
    
    SH = ""
    For lCount = 1 To 16
        lChar = CLng("&H" & Mid(sMD, (lCount * 2) - 1, 2)) Mod 32
        SH = SH & Mid(VALID_CHARS, lChar + 1, 1)
    Next
    
    SpecificChars = SH
End Function 'SpecificChars() As String
'=========================================================================================

'THIS FUCTION IS MADE PUBLIC FOR THE KEYGEN ONLY
'IN THE SETUP APP IT IS PRIVATE
Public Function Encrypt(sChars As String) As String
    'This is the encryption method
    'You can use your own method and you
    'better change this for every application so that
    'your reg keys won't match other apps
    'The method below encrypts every character combined with
    'the encryption of the whole string and returns a very
    'long string that makes cracking harder

    Dim MD2 As New CMD5
    Dim tmp As String
    Dim ln As Integer

    For ln = 1 To Len(sChars)
        tmp = tmp & MD2.MD5(Mid(sChars, ln, 1) & MD2.MD5(sChars))
    Next ln
    
    Encrypt = tmp
    
End Function

