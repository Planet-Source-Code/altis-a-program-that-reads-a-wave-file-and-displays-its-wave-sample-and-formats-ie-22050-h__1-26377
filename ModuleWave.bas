Attribute VB_Name = "ModuleWave"
Option Explicit

' API Declarations
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Enum sndConst
    SND_ASYNC = &H1 ' play asynchronously
    snd_loop = &H8 ' loop the sound until Next sndPlaySound
    SND_MEMORY = &H4 ' lpszSoundName points To a memory file
    SND_NODEFAULT = &H2 ' silence Not default, If sound not found
    SND_NOSTOP = &H10 ' don't stop any currently playing sound
    SND_SYNC = &H0 ' play synchronously (default), halts prog use till done playing
End Enum

' Declare Color Constants
Const cBlue = &HFF0000
Const cRed = &HFF
Const cGreen = &HFF00
Const cYellow = &HFFFF
Const cTeal = &H64FF64
Const cPurple = &HFF007D
Const cGray = &H7D7D7D
Const cOrange = &H96FF

' Define File Header
Public Type FileHeader
    lRiff As String * 4
    lFileSize As Long
    lWave As Long
    lFormat As Long
    lFormatLength As Long
    ' Note: The values in this user defined type MUST follow the above order.
End Type
' Define First Chunk
Public Type FormatChunk
'    ChunkID As String * 4       ' Chunk ID
    ' Chunk ID is always "fmt "
    wFormatTag As Integer       ' wFormatTag
    ' This code does not support compression.  Therefore, wFormatTag is always one
    
    wChannels As Integer        ' wChannels
    ' wChannels contains the number of audio channels for the sound.
    ' 1 = mono, 2 = Stereo, etc.
        
    dwSamplesPerSec As Long     ' dwSamplesPerSec
    ' dwSamplesPerSec is the sample rate at which the sound is to be
    ' played back in sample frames per second (ie, Hertz).
    
    dwAvgBytesPerSec As Long    ' dwAvgBytesPerSec
    ' dwAvgBytesPerSec field indicates how many bytes play every second
    
    wBlockAlign As Integer      ' wBlockAlign
    ' wBlockAlign is the size of a sample frame, in terms of bytes
    
    wBitsPerSample As Integer    ' wBitsPerSample
    ' wBitsPerSample field indicates the bit resolution of a sample point
    ' (ie, a 16-bit waveform would have wBitsPerSample = 16).
    
    
    ' Note: there may be additional fields here, depending upon wFormatTag.
    ' Note: The values in this user defined type MUST follow the above order.
End Type
' Define Data Chunks
Public Type ChunkHeader
    ChunkID As String * 4       ' chunkID
    ' Chunk ID is always "data"
    ChunkSize As Long           ' chunkSize
    ' Chunk Size is the number of bytes in the chunk
End Type

' Public Declarations
Public wChunk As ChunkHeader
Public wFileHeader As FileHeader
Public wFormat As FormatChunk

Public B_Data() As Byte, LB_Data() As Byte
Public I_Data() As Integer, LI_Data() As Integer

Public SStart As Single, SLength As Single  ' Selection values

Public dBitsPerTwip As Single    ' For display only

Public FileLoaded As Boolean

Public Peak As Single, Dip As Single    ' The max and min for the wave

Private Temp As Single  ' Dummy value



Public Sub Wave_Display()
    ' Local variable declaration
    Dim CurrentTrack As Integer
        
    Dim I As Single
    
    Dim LoopsCount As Integer
    
    Dim MRatio As Single    ' We will use this value to scale the crest of the wave
                                ' in WaveForm Window
    Dim GRatio As Single    ' We will use this value to scale the crest of the wave
                                ' in General Window
    Dim X As Single, Y As Single
    
    ' Lock out controls in form
    With frmMain
        ' Lock out text boxes
        .lblSelStart.Enabled = False
        .txtSelStart.Enabled = False
        .lblSelLength.Enabled = False
        .txtSelLength.Enabled = False
        
        ' Reset text in text boxes
        .txtSelStart.Text = 0
        .txtSelLength.Text = 0
        
        ' Lock out menus
        .mnuFile.Enabled = False
        .mnuHelp.Enabled = False
        
        ' Disable command buttons
        .cmdZoomIn.Enabled = False
        .cmdZoomOut.Enabled = False
        .cmdPlay.Enabled = False
        
        ' Disable check box
        .chkLoop.Enabled = False
        
        ' Disable select functions
        .picGeneralView.Enabled = False
           
        ' Clear .picture box
        .PicWaveForm.Cls
        
        ' Let user know the program is drawing
        .PicWaveForm.Print "Displaying wave sample..."
        DoEvents
    End With
    
    ' Count how many bits are there in a quarter twips
    ' so we can skip them if necessary.
    dBitsPerTwip = (wChunk.ChunkSize / wFormat.wChannels) / (frmMain.PicWaveForm.Width)
    
    ' Make sure dBitsPerTwip is not zero
    Temp = dBitsPerTwip
    If dBitsPerTwip < 1 Then Temp = 1
    
    ' Get ratio (We need the negative so that the wave may draw inverse)
    MRatio = -(frmMain.PicWaveForm.Height / (wFormat.wChannels + 1)) / (Peak - Dip)
    GRatio = -(frmMain.picGeneralView.Height / (wFormat.wChannels + 1)) / (Peak - Dip)
    
    ' Prepare to draw wave form and skip some if necessary.
    With wFormat
        For I = 1 To wChunk.ChunkSize - .wChannels Step .wChannels * Format(dBitsPerTwip, "##000")
            X = X + (frmMain.PicWaveForm.ScaleWidth / (wChunk.ChunkSize / .wBlockAlign) * Format(dBitsPerTwip, "##000"))
            ' Resize Temporary value
            ReDim LI_Data(.wChannels)
            For CurrentTrack = 1 To .wChannels
                ' Set color
                Select Case CurrentTrack
                    Case 1
                        frmMain.PicWaveForm.ForeColor = cBlue
                    Case 2
                        frmMain.PicWaveForm.ForeColor = cRed
                    Case 3
                        frmMain.PicWaveForm.ForeColor = cGreen
                    Case 4
                        frmMain.PicWaveForm.ForeColor = cYellow
                    Case 5
                        frmMain.PicWaveForm.ForeColor = cTeal
                    Case 6
                        frmMain.PicWaveForm.ForeColor = cPurple
                    Case 7
                        frmMain.PicWaveForm.ForeColor = cGray
                    Case 8
                        frmMain.PicWaveForm.ForeColor = cOrange
                End Select
                ' Draw wave form and invert it so it would display correctly
                ' Display differently according to its size
                If dBitsPerTwip > 0.5 Then
                    ' Connect all dots
                    Y = (frmMain.PicWaveForm.Height / (.wChannels + 1) * CurrentTrack)
                    frmMain.PicWaveForm.Line (X, Y + LI_Data(CurrentTrack) * MRatio)- _
                        (X, Y + I_Data(I + CurrentTrack - 1) * MRatio)
                    ' Draw center line
                    If I = 1 Then frmMain.PicWaveForm.Line (0, Y)-(frmMain.PicWaveForm.Width, Y)
                Else
                    ' No connection between dots
                    Y = (frmMain.PicWaveForm.Height / (.wChannels + 1) * CurrentTrack)
                    frmMain.PicWaveForm.Line (X, Y + I_Data(I + CurrentTrack - 1) * MRatio)- _
                        (X + (frmMain.PicWaveForm.ScaleWidth / wChunk.ChunkSize), _
                        Y + I_Data(I + CurrentTrack - 1) * MRatio), , BF
                    ' Draw center line
                    If I = 1 Then frmMain.PicWaveForm.Line (0, Y)-(frmMain.PicWaveForm.Width, Y)
                
                End If
                
                ' Draw wave form in General Window
                    frmMain.picGeneralView.ForeColor = frmMain.PicWaveForm.ForeColor
                    Y = (frmMain.picGeneralView.Height / (.wChannels + 1) * CurrentTrack)
                    frmMain.picGeneralView.Line (X, Y + LI_Data(CurrentTrack) * GRatio)- _
                        (X, Y + I_Data(I + CurrentTrack - 1) * GRatio)
                    
                    ' Draw center line
                    If I = 1 Then frmMain.picGeneralView.Line (0, Y)-(frmMain.picGeneralView.Width, Y)
                
                ' Update LI_Data
                LI_Data(CurrentTrack) = I_Data(I + CurrentTrack - 1)
            Next CurrentTrack
            
            ' Update display every two hundred fifty loop
            If LoopsCount >= 250 Then
                DoEvents
                LoopsCount = 0
            Else
                LoopsCount = LoopsCount + 1
            End If
        
            ' Regardless of outcome, stop drawing any further if X is greater than
            ' the width of the window
            If X > frmMain.picGeneralView.ScaleWidth Then Exit For
        Next I
    End With
    
    ' Enable previously disabled controls
    With frmMain
        ' Make frmmain.picGeneralView's wave sample ineffect to Cls
        .picGeneralView.Picture = .picGeneralView.Image
                        
        ' Unlock text boxes
        .lblSelStart.Enabled = True
        .txtSelStart.Enabled = True
        .lblSelLength.Enabled = True
        .txtSelLength.Enabled = True
        
        ' Disable check box
        .chkLoop.Enabled = True
        
        ' Unlock menus
        .mnuFile.Enabled = True
        .mnuHelp.Enabled = True
        
        ' Enable command button
        .cmdPlay.Enabled = True
        
        ' Enable zoom function
        .picGeneralView.Enabled = True
        
        ' Loading complete.  Cover previous message.
        .PicWaveForm.ForeColor = .PicWaveForm.BackColor
        .PicWaveForm.CurrentX = 0: .PicWaveForm.CurrentY = 0
        .PicWaveForm.Print "Displaying wave sample..."
        
        ' Reset forecolor
        .PicWaveForm.ForeColor = vbBlack
    End With
End Sub

Public Sub Wave_EnlargeView(BStart As Single, BLength As Single)
    ' Local variable declaration
    Dim bLengthPerTwip As Single
    
    Dim CurrentTrack As Integer
        
    Dim I As Single
    
    Dim LoopsCount As Integer
    
    Dim MRatio As Single    ' We will use this value to scale the crest of the wave
                                ' in WaveForm Window

    Dim X As Single, Y As Single
    
                    
    ' Reset values
    ReDim LB_Data(wFormat.wChannels)
    ReDim LI_Data(wFormat.wChannels)
    
    
    ' Lock out controls
    With frmMain
        ' Lock out menus
        .mnuFile.Enabled = False
        .mnuHelp.Enabled = False
        
        ' Clear .picture box
        .PicWaveForm.Cls
        
        ' Let user know the program is drawing
        .PicWaveForm.Print "Displaying wave sample..."
        DoEvents
    End With
    
    ' Count how many bits are there in a quarter twips
    ' so we can skip them if necessary.
    bLengthPerTwip = (BLength * wFormat.wChannels) / frmMain.PicWaveForm.Width
    
    ' Make sure dBitsPerTwip is not zero
    Temp = bLengthPerTwip
    If Temp < 1 Then Temp = 1
    
    ' Get ratio
    MRatio = -(frmMain.PicWaveForm.Height / (wFormat.wChannels + 1)) / (Peak - Dip)
    
    ' Prepare to draw wave form
    With wFormat
        For I = Int(BStart * .wChannels) To Int(BStart + BLength) * .wChannels - .wChannels _
            Step Int(Temp * .wChannels)
            ' Increase X
            X = X + frmMain.PicWaveForm.ScaleWidth / (BLength / Temp)
            For CurrentTrack = 1 To .wChannels
                ' Set color
                Select Case CurrentTrack
                    Case 1
                        frmMain.PicWaveForm.ForeColor = cBlue
                    Case 2
                        frmMain.PicWaveForm.ForeColor = cRed
                    Case 3
                        frmMain.PicWaveForm.ForeColor = cGreen
                    Case 4
                        frmMain.PicWaveForm.ForeColor = cYellow
                    Case 5
                        frmMain.PicWaveForm.ForeColor = cTeal
                    Case 6
                        frmMain.PicWaveForm.ForeColor = cPurple
                    Case 7
                        frmMain.PicWaveForm.ForeColor = cGray
                    Case 8
                        frmMain.PicWaveForm.ForeColor = cOrange
                End Select
                ' Draw wave form and invert it so it would display correctly
                If bLengthPerTwip > 0.25 Then
                    ' Connect all dots
                    Y = (frmMain.PicWaveForm.Height / (.wChannels + 1) * CurrentTrack)
                    frmMain.PicWaveForm.Line (X, Y + LI_Data(CurrentTrack) * MRatio)- _
                        (X, Y + I_Data(I + CurrentTrack - 1) * MRatio)
                    ' Draw center line
                    frmMain.PicWaveForm.ForeColor = vbBlack
                    If I = (BStart * .wChannels) Then frmMain.PicWaveForm.Line (0, Y)-(frmMain.PicWaveForm.Width, Y)
                Else
                    ' No connection between dots
                    Y = (frmMain.PicWaveForm.Height / (.wChannels + 1) * CurrentTrack)
                    frmMain.PicWaveForm.Line (X, Y + I_Data(I + CurrentTrack - 1) * MRatio)- _
                        (X + (frmMain.PicWaveForm.ScaleWidth / BLength), _
                        Y + I_Data(I + CurrentTrack - 1) * MRatio), , BF
                    ' Draw center line
                    frmMain.PicWaveForm.ForeColor = vbBlack
                    If I = (BStart * .wChannels) Then frmMain.PicWaveForm.Line (0, Y)-(frmMain.PicWaveForm.Width, Y)
                
                End If
                            
                ' Update LI_Data
                LI_Data(CurrentTrack) = I_Data(I + CurrentTrack - 1)
            Next CurrentTrack
            
            ' Update display every two hundred fifty loop
            If LoopsCount >= 250 Then
                DoEvents
                LoopsCount = 0
            Else
                LoopsCount = LoopsCount + 1
            End If
        
            ' Regardless of outcome, stop drawing any further if X is greater than
            ' the width of the window
            If X > frmMain.picGeneralView.ScaleWidth Then Exit For
        Next I
    End With
                    
    ' Enable previously disabled controls
    With frmMain
        ' Unlock text boxes
        .lblSelStart.Enabled = True
        .txtSelStart.Enabled = True
        .lblSelLength.Enabled = True
        .txtSelLength.Enabled = True
        
        ' Unlock menus
        .mnuFile.Enabled = True
        .mnuHelp.Enabled = True
        
        ' Enable command button
        .cmdPlay.Enabled = True
        
        ' Loading complete.  Cover previous message.
        .PicWaveForm.ForeColor = .PicWaveForm.BackColor
        .PicWaveForm.CurrentX = 0: .PicWaveForm.CurrentY = 0
        .PicWaveForm.Print "Displaying wave sample..."
        
        ' Reset forecolor
        .PicWaveForm.ForeColor = vbBlack
    End With
End Sub

