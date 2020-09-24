VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Wave"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGeneralView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1740
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   1620
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2100
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicWaveForm 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1740
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox picOptions 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   -900
      ScaleHeight     =   495
      ScaleWidth      =   9015
      TabIndex        =   2
      Top             =   2220
      Width           =   9015
      Begin VB.CheckBox chkLoop 
         Caption         =   "Loop"
         Height          =   315
         Left            =   8000
         TabIndex        =   10
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7000
         TabIndex        =   9
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdZoomOut 
         Caption         =   "Zoom Out"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5000
         TabIndex        =   8
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton cmdZoomIn 
         Caption         =   "Zoom In"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4020
         TabIndex        =   7
         Top             =   0
         Width           =   1000
      End
      Begin VB.TextBox txtSelLength 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   6
         Text            =   "0"
         Top             =   0
         Width           =   1000
      End
      Begin VB.TextBox txtSelStart 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1000
         TabIndex        =   4
         Text            =   "0"
         Top             =   0
         Width           =   1000
      End
      Begin VB.Label lblSelLength 
         Caption         =   "Select Length"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2000
         TabIndex        =   5
         Top             =   0
         Width           =   1000
      End
      Begin VB.Label lblSelStart 
         Caption         =   "Select Start"
         Enabled         =   0   'False
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1000
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBreak01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBreak02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About this program..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SoundPlay As Boolean

Private Sub cmdPlay_Click()
    Dim I As Single
    Dim lSecond As Single
    
    Dim X As Single, Y As Single
    
    ' Disable Menus
    mnuFile.Enabled = False
    mnuHelp.Enabled = False
    
    ' Disable picture box
    picGeneralView.Enabled = False

    ' Lock out text boxes
    lblSelStart.Enabled = False
    txtSelStart.Enabled = False
    lblSelLength.Enabled = False
    txtSelLength.Enabled = False
    
    ' Disable check box
    chkLoop.Enabled = False
    
    ' Do different action according to the Button's caption
    Select Case cmdPlay.Caption
        Case "Play"
            ' Change its caption
            cmdPlay.Caption = "Stop"
            
            ' Play sound
            If chkLoop.Value = 1 Then
                ' Play loop
                sndPlaySound Caption, SND_ASYNC + snd_loop + SND_NODEFAULT
            Else
                ' Play once
                sndPlaySound Caption, SND_ASYNC + SND_NODEFAULT
            End If
            SoundPlay = True
            
            ' Approximate the current location at which the wave is at
            With wFormat
                Do
                    X = frmMain.picGeneralView.Width / ((wChunk.ChunkSize / .wBlockAlign) / .dwSamplesPerSec) * 0.1
                    Do
                        lSecond = Timer
                        ' Wait for 0.1 seconds
                        Do
                        Loop Until Timer - lSecond >= 0.1
                        ' Refresh
                        frmMain.picGeneralView.Cls
                        ' Divide picture box to the appropiate length of time and
                        ' draw line at appropiate position
                        If chkLoop.Value = 0 Then
                            ' Draw only if there is no loop.  Otherwise, there will be a
                            ' 90 percents error when there are more than five loops
                            X = X + frmMain.picGeneralView.Width / ((wChunk.ChunkSize / .wBlockAlign) / .dwSamplesPerSec) * 0.11
                            frmMain.picGeneralView.Line (X, 0)-(X, picGeneralView.Height), vbBlack
                        End If
                        ' Update display
                        DoEvents
                        
                        ' Exit if user stopped the sound
                        If SoundPlay = False Then Exit Do
    
                    Loop Until X > frmMain.picGeneralView.ScaleWidth
                Loop Until chkLoop.Value = 0 Or cmdPlay.Caption = "Play"
            End With
            ' Stop
            cmdPlay.Caption = "Play"
            SoundPlay = False
        Case "Stop"
            ' Change its Caption
            cmdPlay.Caption = "Play"
            
            ' Stop any playing sound
            sndPlaySound "", SND_ASYNC + SND_NODEFAULT
            SoundPlay = False
            
            ' Clear picture box
            picGeneralView.Cls
    End Select
    
    ' Enable Menus
    mnuFile.Enabled = True
    mnuHelp.Enabled = True
    
    ' Enable picture box
    picGeneralView.Enabled = True
    
    ' Unlock text boxes
    lblSelStart.Enabled = True
    txtSelStart.Enabled = True
    lblSelLength.Enabled = True
    txtSelLength.Enabled = True

    ' Enable check box
    chkLoop.Enabled = True
End Sub

Private Sub cmdZoomIn_Click()
    ' Enable Zoom out
    cmdZoomOut.Enabled = True
    
    ' Call Procedures
    Call txtSelLength_KeyDown(13, 0)
    Call Wave_EnlargeView(txtSelStart.Text, txtSelLength.Text)
End Sub

Private Sub cmdZoomOut_Click()
    ' Disable Zoom out
    cmdZoomOut.Enabled = False
    
    ' Clear picture box
    picGeneralView.Cls
    
    ' Call Procedure
    Wave_Display
End Sub

Private Sub Form_Load()
    ' Enlarge window
    Width = Width * 2
    Height = Height * 2
End Sub

Private Sub Form_Resize()
    ' Ignore Error
    On Error Resume Next
    ' Resize Main Viewer
    PicWaveForm.Move 0, 0, ScaleWidth, ScaleHeight / 5 * 4 - 700
    ' Resize General Viewer
    picGeneralView.Move 0, PicWaveForm.Height + 750, _
        ScaleWidth, ScaleHeight / 5 - 100
        
    ' Relocate Options box
    picOptions.Move 0, PicWaveForm.Height + 30, ScaleWidth, 700
    
    ' If a file has been loaded and the form resizes, redraw wave samples
    If FileLoaded = True Then
        ' Clear Gneral View because Cls will not have any effect
        picGeneralView.Picture = LoadPicture()
        ' Call Procedure
        Call Wave_Display
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' End program
    End
End Sub

Private Sub mnuFileExit_Click()
    ' End Program
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    ' Local Declaration
    Dim FileNumber As Integer
    Dim I As Single
    Dim wChannelType As String


    ' Ignore Errors
'    On Error Resume Next

    ' Get Free File Number
    FileNumber = FreeFile
    
    ' Open Common Dialog box
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Wav(*.wav)|*.wav"
    CommonDialog1.ShowOpen
    
    
    
    ' Check file name
    If CommonDialog1.FileName <> "" Then
        ' Open file
        Open CommonDialog1.FileName For Binary As #FileNumber
            ' Read data from the file
            Get #FileNumber, , wFileHeader
            Get #FileNumber, , wFormat
            ' NOTE: Although we did not use some of the value in
            ' user define type, it is necessary to define them all
            ' inorder to read the information properly.
            
            ' Proceed only if the header is 'Riff' or 'Wave'.
            ' Otherwise, End Procedure.
            If UCase(wFileHeader.lRiff) <> "RIFF" And _
                UCase(wFileHeader.lRiff) <> "WAVE" Then Exit Sub
            
            ' Skip all unnecessary Chunks until we find 'Data'.
            ' We only need Format and Data.
            Select Case wFormat.wBitsPerSample
                Case Is < 16    ' Eight Bits wave format
                    ' Display message
                    I = MsgBox("You have requested to view a " & wFormat.wBitsPerSample & _
                        "-bits wave file.  This program only supports 16-bits (Standard PCM) wave format.", _
                        vbCritical, "Unable to Comply")
                Case 16 ' Sixteen Bits wave format
                    ' Change Caption
                    Caption = CommonDialog1.FileName
                    
                    ' See how many tracks are there in the wave file
                    Select Case wFormat.wChannels
                        Case 1
                            wChannelType = "Mono"
                        Case 2
                            wChannelType = "Stereo"
                        Case 3
                            wChannelType = "3 Channels"
                        Case 4
                            wChannelType = "quad"
                        Case Is >= 5
                            wChannelType = wFormat.wChannels & " Channels"
                    End Select
                
                    ' Clear picture boxes
                    PicWaveForm.Cls
                    picGeneralView.Cls
                    picGeneralView.Picture = LoadPicture()
                    picOptions.Cls
                    
                    ' Let user know program is working
                    PicWaveForm.Print "Loading wave file.  This may takes a while..."
                    DoEvents
                    
                    ' Seek to next chunk by discarding any format bytes.
                    ReDim B_Data(0)
                    For I = 1 To wFileHeader.lFormatLength - 16
                        Get #FileNumber, , B_Data(0)    ' Ignore BYTES
                    Next
                   
                    ' Reset peak and dip so they become zero
                    Peak = 0: Dip = 0
                    
                    Do
                        ' Read chunk's header
                        Get #FileNumber, , wChunk
                        ' Resize array
                        ReDim I_Data(wChunk.ChunkSize)
                        ' Read chunk's data
                        For I = 1 To wChunk.ChunkSize
                            Get #FileNumber, , I_Data(I)
                            If I_Data(I) < Dip Then Dip = CSng(I_Data(I))
                            If I_Data(I) > Peak Then Peak = CSng(I_Data(I))
                        Next I
                    Loop While UCase(wChunk.ChunkID) <> "DATA"

                    ' Display file informations
                    picOptions.CurrentX = lblSelStart.Left
                    picOptions.CurrentY = 450
                    picOptions.Print wFormat.dwSamplesPerSec & " Hertz, " & _
                        wFormat.wBitsPerSample & "-bits, " & _
                        wChannelType & ", " & _
                        wChunk.ChunkSize & " KBytes"
                    
                    ' Show wave sample
                    Wave_Display
                    
                    ' Note that a file has been open
                    FileLoaded = True
            End Select
        Close #FileNumber
    End If
    
    ' Loading complete.  Cover previous message.
    PicWaveForm.ForeColor = vbWhite
    PicWaveForm.CurrentX = 0: PicWaveForm.CurrentY = 0
    PicWaveForm.Print "Displaying wave sample..."
    
    ' Reset forecolor
    PicWaveForm.ForeColor = vbBlack
End Sub

Private Sub mnuFileProperties_Click()
    Dim wChannelType As String
    Dim IntResponse As Integer

    If FileLoaded = True Then
        ' See how many tracks are there in the wave file
        Select Case wFormat.wChannels
            Case 1
                wChannelType = "1 Channel - Mono"
            Case 2
                wChannelType = "2 Channels - Stereo"
            Case 3
                wChannelType = "3 Channels"
            Case 4
                wChannelType = "4 Channels - quad"
            Case Is >= 5
                wChannelType = wFormat.wChannels & " Channels"
        End Select
        ' Display file's property if it has been loaded
        IntResponse = MsgBox("Sample Rate : " & wFormat.dwSamplesPerSec & " Hertz" & Chr(13) & _
            "Resolution : " & wFormat.wBitsPerSample & "-bits" & Chr(13) & _
            "Channels : " & wChannelType & Chr(13) & _
            "Size : " & wChunk.ChunkSize & " KBytes" & Chr(13) & _
            "Length : " & _
            Format((wChunk.ChunkSize / wFormat.wBlockAlign) / wFormat.dwSamplesPerSec, _
                            "##0.000") & " seconds", vbInformation)
    Else
        ' Display message box
        IntResponse = MsgBox("No Information available, please make sure you have opened a file.", _
            vbExclamation, "Wave file information not available")
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    Dim IntResponse As Integer
    
    ' Show message box
    IntResponse = MsgBox("Get the most out of wave files by Altis ©2001" & Chr(13) & Chr(13) & _
        "For more information on wave formats, go to www.wotsit.org " & _
        "and search for 'Wave Formats'." & Chr(13) & Chr(13) & _
        "   OR" & Chr(13) & Chr(13) & _
        "go to msdn.Microsoft.com and search for 'Reading Wave'.  " & _
        "(Recommend Article Four)" & Chr(13) & Chr(13) & _
        "For submissions by other authors, check out " & Chr(13) & _
        "http://www.planet-source-code.com/xq/ASP/txtCodeId.7897/lngWId.1/qx/vb/scripts/ShowCode.htm" & Chr(13) & Chr(13) & _
        "It is one of the best.", vbInformation, "About this program")
End Sub

Private Sub picGeneralView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' See if the user is holding the left button and
    ' if any wave file is opened
    If Button = 1 And FileLoaded = True Then
        With wFormat
            ' Update text box and value
            SStart = X
            txtSelStart.Text = Int(X * (dBitsPerTwip * .wChannels) / .wBlockAlign)
            ' Call Procedure
            Call picGeneralView_MouseMove(1, 0, X, 0)
        End With
    Else
    If Button = 2 Then
        ' Right click means Cancel
        picGeneralView.Cls
        ' Reset text in text boxes
        txtSelStart.Text = 0
        txtSelLength.Text = 0
    End If
    End If
End Sub

Private Sub picGeneralView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' See if the user is dragging and if any file is opened
    If Button = 1 And FileLoaded = True Then
        With wFormat
            ' Clear old picture
            picGeneralView.Cls
            ' Draw outline
            picGeneralView.ForeColor = vbBlack
            picGeneralView.Line (SStart, 0)- _
                (X, picGeneralView.ScaleHeight), , B
            ' Update text box
            txtSelLength.Text = Int((X * (dBitsPerTwip * .wChannels) / .wBlockAlign) - _
                (SStart * (dBitsPerTwip * .wChannels) / .wBlockAlign))
        End With
    End If
End Sub

Private Sub picGeneralView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Check if user release left button
    If Button = 1 Then
        ' Make sure no value in the text boxes is negative
        If Val(txtSelLength.Text) < 0 Then
            With wFormat
                ' Swarp values
                txtSelLength.Text = Int(Val(txtSelStart.Text) - _
                    (X * (dBitsPerTwip * .wChannels) / .wBlockAlign))
                txtSelStart.Text = Int(X * (dBitsPerTwip * .wChannels) / .wBlockAlign)
            End With
        End If
        If Val(txtSelStart.Text) < 0 Then txtSelStart.Text = 0
    End If
End Sub

Private Sub picOptions_Resize()
    ' Relocate controls within the picture box
    lblSelStart.Move 200, 0, 1000, 375
    txtSelStart.Move 1300, 0, 1000, 375
    lblSelLength.Move 2400, 0, 1000, 375
    txtSelLength.Move 3500, 0, 1000, 375
    cmdZoomIn.Move 4600, 0, 1000, 375
    cmdZoomOut.Move 5700, 0, 1000, 375
    cmdPlay.Move 7800, 0, 1000, 375
    chkLoop.Move 8900, 0, 1000, 375
End Sub

Private Sub txtSelLength_Change()
    ' If the value is no zero then enable zoom
    If Val(txtSelLength.Text) <> 0 Then
        cmdZoomIn.Enabled = True
    Else
        cmdZoomIn.Enabled = False
    End If
End Sub

Private Sub txtSelLength_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Get enter key
    If KeyCode = 13 Then
        With wFormat
            ' Clear old picture
            picGeneralView.Cls
            ' Draw outline
            SStart = Val(txtSelStart.Text) * .wBlockAlign / (dBitsPerTwip * .wChannels)
            picGeneralView.ForeColor = vbBlack
            picGeneralView.Line (SStart, 0)- _
            (SStart + Val(txtSelLength.Text) * .wBlockAlign / (dBitsPerTwip * .wChannels), _
                picGeneralView.ScaleHeight), , B
        End With
    End If
End Sub

Private Sub txtSelLength_LostFocus()
    ' Make sure text are values
    txtSelLength.Text = Val(txtSelLength.Text)
End Sub

Private Sub txtSelStart_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Get enter key
    If KeyCode = 13 Then
        With wFormat
            ' Update value
            SStart = Val(txtSelStart.Text) * .wBlockAlign / (dBitsPerTwip * .wChannels)
            ' Call Procedure
            Call txtSelLength_KeyDown(13, 0)
        End With
    End If
End Sub

Private Sub txtSelStart_LostFocus()
    ' Make sure text are values
    txtSelStart.Text = Val(txtSelStart.Text)
End Sub

