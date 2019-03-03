Attribute VB_Name = "DSound"
Option Explicit

'Almost all of this code (for direct sound) is furnised from www.directx4vb.com

Public Ds As DirectSound

'My cheap sounds!! Some are loaded as an array to make them
'polyphonic.

Public DsGoal(5) As DirectSoundBuffer
Public dsGameOver As DirectSoundBuffer
Public dsBonus(7) As DirectSoundBuffer
Public dsBip(5) As DirectSoundBuffer
Public dsMiss As DirectSoundBuffer
Public dsPaddle As DirectSoundBuffer
Private dsMeter As DirectSoundBuffer
Private Bzzz As DirectSoundBuffer
Private ExtraBall As DirectSoundBuffer
Private Vocoder As DirectSoundBuffer
Private BeatLoop As DirectSoundBuffer
Private Filter As DirectSoundBuffer

Public DsDesc As DSBUFFERDESC
Public DsWave As WAVEFORMATEX

Public CurrFreq As Long

Public Sub InitSound()

   Randomize
   Set Ds = dx.DirectSoundCreate("")
   
   'It is best to check for errors before continuing
   If Err.Number <> 0 Then
       MsgBox "Unable to Continue, Error creating Directsound object."
       Exit Sub
   End If
   
   Ds.SetCooperativeLevel frmDX.hwnd, DSSCL_NORMAL
   
   'These settings are quite important; what you set here reflects
   'in the quality that the sound plays.
   '8bit + 1 channel = Poor quality, but less processing time and less memory
   '16bit + 2 channels = Good quality, more memory and more processing time
   'The comments about processing time and memory will only affect you
   'when you start putting a heavy load on the sound card; and/or it is
   'playing/storing lots of sound files.
   DsDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
   DsWave.nFormatTag = WAVE_FORMAT_PCM 'Sound Must be PCM otherwise we get errors
   DsWave.nChannels = 1    '1= Mono, 2 = Stereo
   DsWave.lSamplesPerSec = 22050
   DsWave.nBitsPerSample = 16 '16 =16bit, 8=8bit
   DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
   DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign
   
   'This line creates a sound buffer based on the information you have just
   'put in the two structures
   
   
   'Here are my cheap sounds being loaded into a buffer!
   Dim i As Long, j As Long
   
   
   'A polyphonic sound will not cut itself off when a new instance of the sound is
   'played. In other words, if the sound was half a second long and it tried
   'to play twice before 1 second elapsed it would cut itself off. Unless it is
   'polyphonic.
   'So, to avoid this cutoff from happening, I cheated a bit and put some of
   'my individual sounds into an array. These were sounds that I knew (once the game
   'started going faster) would cut themselves off due to frequent instances.
   'So now the first subscript of the array will sound, and then the next until it
   'reaches the end of the array, where it will start over.
   
   'To understand the difference, try changing one of the arrays to have only one
   'index and then check out the sounds (particularly the 'Bip' sound) You'll notice
   'what's going on and what I've been jammering about here.
     
   For i = 0 To 5
      Set DsGoal(i) = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\Block" & i + 1 & ".Wav", DsDesc, DsWave)
   Next i
    
   For i = LBound(dsBip()) To UBound(dsBip())
      Set dsBip(i) = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\Bip.Wav", DsDesc, DsWave)
   Next i
   
   'these are monophonic
   Set dsGameOver = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\GameOver1.Wav", DsDesc, DsWave)
   Set dsMiss = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\Miss.Wav", DsDesc, DsWave)
   Set dsPaddle = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\Paddle.Wav", DsDesc, DsWave)
   Set dsMeter = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\Meter.wav", DsDesc, DsWave)
   Set Bzzz = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\Bzzz.wav", DsDesc, DsWave)
   Set ExtraBall = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\Extra.wav", DsDesc, DsWave)
   Set Vocoder = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\vocoder.wav", DsDesc, DsWave)
   Set BeatLoop = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\beatloop.wav", DsDesc, DsWave)
   Set Filter = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\filter.wav", DsDesc, DsWave)
   

End Sub

Public Sub PlayGoalSound()

   Static i As Long
   
   DsGoal(i).Play DSBPLAY_DEFAULT
   
   i = i + 1
   
   If i > 5 Then
      i = 0
   End If

End Sub

Public Sub PlayBonusSound()

Static i As Integer

dsBonus(i).Play DSBPLAY_DEFAULT
i = i + 1
If i > UBound(dsBonus()) Then
   i = 0
End If

End Sub

Public Sub PlayGameOverSound()

dsGameOver.Play DSBPLAY_DEFAULT

End Sub

Public Sub PLayBipSound()

Static i As Integer

dsBip(i).Play DSBPLAY_DEFAULT
i = i + 1
If i > UBound(dsBip()) Then
   i = 0
End If

End Sub

Public Sub PlayMissSound()

dsMiss.Play DSBPLAY_DEFAULT

End Sub

Public Sub PlayPaddleSound()

dsPaddle.Play DSBPLAY_DEFAULT

End Sub

Public Sub PlayMeterSound()

   dsMeter.Play DSBPLAY_DEFAULT
   
End Sub

Public Sub PlayBzzz()

   Bzzz.Play DSBPLAY_DEFAULT
   
End Sub

Public Sub PlayExtraBall()

   ExtraBall.Play DSBPLAY_DEFAULT
   
End Sub

Public Sub PlayVocoder()

   Vocoder.Play DSBPLAY_DEFAULT
   
End Sub

Public Sub PlayBeatz()

   BeatLoop.Play DSBPLAY_LOOPING
   
End Sub

Public Sub StopBeatz()

   BeatLoop.Stop
   BeatLoop.SetCurrentPosition 0
   
End Sub

Public Sub PlayFilter()

   Filter.Play DSBPLAY_DEFAULT
   
End Sub

Sub StopSound()
'Stopping it effectively pauses it
'DsBuffer.Stop
'Only when you set the current position to 0 will
'it go back to the beginning
'DsBuffer.SetCurrentPosition 0
'You can use the SetCurrentPosition to jump around
'a sound file as you like; which can be useful
End Sub

