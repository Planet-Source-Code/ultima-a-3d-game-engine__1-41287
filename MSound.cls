VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Buf As DirectSoundSecondaryBuffer8

Public Function Create(FPath As String, Volume As Long)
    Dim DSBufDesc As DSBUFFERDESC
    
    Set Buf = Nothing
    
    DSBufDesc.lFlags = DSBCAPS_CTRLVOLUME
    
    Set Buf = DS.CreateSoundBufferFromFile(FPath, DSBufDesc)
    Buf.SetVolume Volume
End Function

Public Function PlaySound(Infinite As Boolean)
    Buf.Stop
    Buf.SetCurrentPosition 0
    If Infinite = True Then
        Buf.Play DSBPLAY_LOOPING
    Else
        Buf.Play DSBPLAY_DEFAULT
    End If
End Function

Public Function StopSound()
    Buf.Stop
    Buf.SetCurrentPosition 0
End Function

