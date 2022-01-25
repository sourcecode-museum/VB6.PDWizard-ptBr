Attribute VB_Name = "ModHard"

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Function HardSerial() As Long
    Dim Buf$, Name$, Flags&, Length&
    Dim Serial As Long
    GetVolumeInformation "C:\", Buf$, 255, Serial, Length, Flags, Name$, 255
    HardSerial = Serial
End Function

