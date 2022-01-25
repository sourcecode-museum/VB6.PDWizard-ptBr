Attribute VB_Name = "ModHard"
'=========================================================================================
'  Main Module
'  API and AppPath
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 07/06/2002
'  WebSite: http://www.geocities.com/bs20014/
'  Legal Copyright: Behrooz Sangani © 07/06/2002
'=========================================================================================


'=========================================================================================
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Function HardSerial() As Long
    Dim Buf$, Name$, Flags&, Length&
    Dim Serial As Long
    GetVolumeInformation "C:\", Buf$, 255, Serial, Length, Flags, Name$, 255
    HardSerial = Serial
End Function 'HardSerial() As Long
'=========================================================================================
Public Function AppPath() As String

    Dim sAns As String
    sAns = App.Path
    If Right(App.Path, 1) <> "\" Then sAns = sAns & "\"
    AppPath = sAns

End Function 'AppPath() As String
'=========================================================================================

