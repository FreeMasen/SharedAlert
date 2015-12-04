Attribute VB_Name = "modPlaySound"
Public Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" _
            (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub playSound()
Call sndPlaySound32("C:\Windows\Media\Windows User Account Control.wav", 0)
End Sub
