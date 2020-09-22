Attribute VB_Name = "modPlayWAVFromRes"
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long 'sound play declaration
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4

Private snd() As Byte
Dim snd1 As Integer
Sub PlayFromRes(ResID As String)
snd = LoadResData(101, "CUSTOM")
snd1 = sndPlaySound(snd(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY) ' plays the sound
End Sub


