VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WSBF-FM RDS Sender"
   ClientHeight    =   5652
   ClientLeft      =   120
   ClientTop       =   744
   ClientWidth     =   7572
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5652
   ScaleWidth      =   7572
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Radio Text Plus :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   7095
      Begin VB.Label lblRTP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Radio Text :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   7095
      Begin VB.Label lblTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dynamic Program Service Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   7095
      Begin VB.Label lblDPS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6855
      End
   End
   Begin MSComctlLib.ProgressBar progSend 
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11240
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6240
      Top             =   360
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Sending"
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send To RDS"
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Song Playing :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.Label lblCurrentSong 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "The Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      TabIndex        =   6
      Top             =   6240
      Width           =   5295
   End
   Begin VB.Menu Settings 
      Caption         =   "Settings"
      Index           =   3
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Index           =   1
   End
   Begin VB.Menu Info 
      Caption         =   "About"
      Index           =   2
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program was created by Kevin Haag and is licensed under the
'Questionable Priorities Public License v1.0
'The QPPL is included with the original distribution of this source code.
'Some portions of this program were originally written by other people;
'a note above each of these portions denotes the original authors.
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Wait sub written by Kevin Jones a.k.a. "Zorvek"
Public Sub Wait(ByVal Seconds As Single)

' Pauses for the number of seconds specified. Seconds can be specified down to
' 1/100 of a second. The Windows Sleep routine is called during each cycle to
' give other applications
' time because, while DoEvents does the same, it does
' not wait and hence the VB loop code consumes more CPU cycles.

   Const MaxSystemSleepInterval = 25 ' milliseconds
   Const MinSystemSleepInterval = 1 ' milliseconds
   
   Dim ResumeTime As Double
   Dim Factor As Long
   Dim SleepDuration As Double
   
   Factor = CLng(24) * 60 * 60
   
   ResumeTime = Int(Now) + (Timer + Seconds) / Factor
   
   Do
      SleepDuration = (ResumeTime - (Int(Now) + Timer / Factor)) * Factor * 1000
      If SleepDuration > MaxSystemSleepInterval Then SleepDuration = MaxSystemSleepInterval
      If SleepDuration < MinSystemSleepInterval Then SleepDuration = MinSystemSleepInterval
      Sleep SleepDuration
      DoEvents
   Loop Until Int(Now) + Timer / Factor >= ResumeTime
   
End Sub

'I didn't write this, but this functionis widely available on the internet.
Function FileText(ByVal filename As String) As String
     Dim handle As Integer
    ' open in binary mode
    handle = FreeFile
    Open filename$ For Binary As #handle
    ' read the string and close the file
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
End Function

'This function takes out the naughty words and avoids certain Clbuttic Mistakes
Public Function censorship(ByVal dirtyInput As String) As String
            dirtyInput = Replace(dirtyInput, "&", "AND")
            dirtyInput = Replace(dirtyInput, "SHIT", "CRAP")
            dirtyInput = Replace(dirtyInput, "FUCK", "FRAK")
            dirtyInput = Replace(dirtyInput, "CUNT", "RUNT")
            dirtyInput = Replace(dirtyInput, "ASSHOLE", "ASSHAT")
            dirtyInput = Replace(dirtyInput, "ASS HOLE", "ASSHAT")
            dirtyInput = Replace(dirtyInput, "ASS-HOLE", "ASSHAT")
            dirtyInput = Replace(dirtyInput, "TIT ", "TEAT ")
            dirtyInput = Replace(dirtyInput, "TITTIES", "BREASTS")
            dirtyInput = Replace(dirtyInput, "TITS", "TEATS")
            dirtyInput = Replace(dirtyInput, "CLIT", "CLAW")
            dirtyInput = Replace(dirtyInput, "GODDAMN", "GORRAM")
            dirtyInput = Replace(dirtyInput, "GODDAM", "GORRAM")
            dirtyInput = Replace(dirtyInput, "GOD-DAMN", "GORRAM")
            dirtyInput = Replace(dirtyInput, "GOD DAMN", "GORRAM")
            dirtyInput = Replace(dirtyInput, "GOD DAM", "GORRAM")
            dirtyInput = Replace(dirtyInput, "GOD-DAM", "GORRAM")
            dirtyInput = Replace(dirtyInput, "BITCH", "BETCH")
            dirtyInput = Replace(dirtyInput, "COCK", "CAWK")
            dirtyInput = Replace(dirtyInput, "NIGGER", "CANADIAN")
            dirtyInput = Replace(dirtyInput, "NIGGA", "CANADIAN")
            dirtyInput = Replace(dirtyInput, "DICK HEAD", "D-HEAD")
            dirtyInput = Replace(dirtyInput, "DICK-HEAD", "D-HEAD")
            dirtyInput = Replace(dirtyInput, "DICKHEAD", "D-HEAD")
            dirtyInput = Replace(dirtyInput, "PUSSY", "POSSY")
            dirtyInput = Replace(dirtyInput, "POSSYCAT", "PUSSYCAT")
            dirtyInput = Replace(dirtyInput, "POSSY CAT", "PUSSY CAT")
            dirtyInput = Replace(dirtyInput, "POSSY-CAT", "PUSSY-CAT")
            censorship = dirtyInput

End Function
'Here's what happens when you click send
Private Sub cmdSend_Click()
    lblCurrentSong.Visible = True
    lblDPS.Visible = True
    lblTEXT.Visible = True
    lblRTP.Visible = True

    Dim Mod_Time As Date
    Dim Mod_Time_Two As Date
    Dim RTP As String
    Dim Total_Length As Integer
    Dim DPSMake As String
    Dim DPSSong As String
    Dim DPSArtist As String
    Dim Song_Name_Length_Fix As String
    Dim Song_Name_Fix As Integer
    
    'Show and hide the appropriate buttons
    cmdStop.Visible = True
    cmdSend.Visible = False
    'The settings for the COM port
    MSComm1.CommPort = Choix_Port
    MSComm1.Settings = Vitesse(Choix_Vitesse) + "," + Parite(Choix_Parite) + "," + Bit_Donnee(Choix_Bit_Donnee) + "," + Bit_Arret(Choix_Bit_Arret)
    MSComm1.Handshaking = Choix_Flux
    MSComm1.InputLen = 0
    MSComm1.PortOpen = True
    
    'Show the progress bar, set the variable to keep the infinite loop going
    progSend.Visible = True
    Keep_Going = 1
    'Sets the modtime to something before anythign was made
    Mod_Time = #5/12/1989 8:04:00 AM#
    Mod_Time_Two = #5/12/1989 8:04:00 AM#
    'While loop sends whatever is in the box, every time it changes
    While Keep_Going > 0
        
        'If the file has been modified, send the new contents, also
        'Make sure the port is open and that you haven't clicked stop
        If (FileDateTime(Song_Name_File) > Mod_Time Or FileDateTime(Artist_Name_File) > Mod_Time_Two) And (MSComm1.PortOpen) = True Then
            'READ the text from the file with the song name FIRST
            'Whatever you use to write the information to the files,
            'MAKE SURE you WRITE the artist name FIRST
            'Also it needs to be uppercase because most radios don't understand lowercase
            Song_Name_Text = UCase$(FileText(Song_Name_File))
            Artist_Name_Text = UCase$(FileText(Artist_Name_File))
            'Remove Curse Words
            Song_Name_Text = censorship(Song_Name_Text)
            Artist_Name_Text = censorship(Artist_Name_Text)
            
            
            'The output string formatted for the Inovonics 730 which can't be more that 128 characters
            'so get the lengths so you can figure out if it is
            Song_Name_Length = Len(Song_Name_Text)
            Artist_Name_Length = Len(Artist_Name_Text)
            Total_Length = Song_Name_Length + Artist_Name_Length + 12
            'We want to keep the originals intact because we have to use them more than onces
            DPSArtist = Artist_Name_Text
            DPSSong = Song_Name_Text
            
            'Special case used for anything custom, basically the password is IAMADUCK
            While Song_Name_Text = "IAMADUCK" And MSComm1.PortOpen = True
                'Make it less than 128 chars
                DPSArtist = Left$(Artist_Name_Text, 128)
                'send it
                If (MSComm1.PortOpen) = True Then
                    MSComm1.Output = "DPS=" + DPSArtist + Chr(13)
                End If
                lblDPS.Caption = "DPS=" + DPSArtist + Chr(13)
                'Make it less than 64 chars
                DPSArtist = Left$(Artist_Name_Text, 64)
                'Send it
                If (MSComm1.PortOpen) = True Then
                    MSComm1.Output = "TEXT=" + DPSArtist + Chr(13)
                End If
                lblTEXT.Caption = "TEXT=" + DPSArtist + Chr(13)
                'Clear the unused case and set the overall text
                lblCurrentSong.Caption = Artist_Name_Text
                lblRTP.Caption = "NONE"
                
                'Clear Progress Bar
                progSend.Value = 0
                Wait (1)
                'Update the progress bar, we end up waiting 61 seconds because the RDS can send over and
                'over too quickly which causes the text to not show up right if we don't wait
                While progSend.Value < 100
                    progSend.Value = progSend.Value + 1
                    Wait (0.6)
                Wend
                Song_Name_Text = UCase$(FileText(Song_Name_File))
                Artist_Name_Text = UCase$(FileText(Artist_Name_File))
                Song_Name_Text = censorship(Song_Name_Text)
                Artist_Name_Text = censorship(Artist_Name_Text)
                Song_Name_Length = Len(Song_Name_Text)
                Artist_Name_Length = Len(Artist_Name_Text)
                Total_Length = Song_Name_Length + Artist_Name_Length + 12
                DPSArtist = Artist_Name_Text
                DPSSong = Song_Name_Text
            Wend
                
                
            'Special case, when the live sessions show is on, keep sending it out over and over and
            'display it differently so they know it's live
            While (Song_Name_Text = "LIVE SESSION" Or Song_Name_Text = "LIVE SESSIONS") And MSComm1.PortOpen = True
                'Easiest way to make sure our stuff is less than 128 characters
                DPSArtist = Left$(Artist_Name_Text, 102)
                'send it
                If (MSComm1.PortOpen) = True Then
                    MSComm1.Output = "DPS=LIVE SESSION WITH " + DPSArtist + " ON WSBF" + Chr(13)
                End If
                lblDPS.Caption = "DPS=LIVE SESSION WITH " + DPSArtist + " ON WSBF" + Chr(13)
                'Then for the Text, make sure it's less than 64 characters
                DPSArtist = Left$(Artist_Name_Text, 51)
                'Send it as well
                If (MSComm1.PortOpen) = True Then
                    MSComm1.Output = "TEXT=" + DPSArtist + " LIVE ON WSBF" + Chr(13)
                End If
                lblTEXT.Caption = "TEXT=" + DPSArtist + " LIVE ON WSBF"
                'Now change the Artist name length so we can calculate the RT+ value properly
                Artist_Name_Length = Len(DPSArtist) - 1
                'Creates the string to send out and fixes the requirement of 09 instead of just 9
                If (Artist_Name_Length < 10) Then
                    Artist_Name_Length_String = "0" + Str(Artist_Name_Length)
                Else
                    Artist_Name_Length_String = Str(Artist_Name_Length)
                End If
                'I'm using this variable for something it wasn't originally intended to be used for
                'I add 5 here because that will give me the starting position of WSBF in the Text
                Song_Name_Fix = Artist_Name_Length + 10
                'Again format it and make it a string
                If (Song_Name_Fix < 10) Then
                    Song_Name_Length_Fix = "0" + Str(Song_Name_Fix)
                Else
                    Song_Name_Length_Fix = Str(Song_Name_Fix)
                End If
                'Make the RT+ string
                RTP = "RTP=04,00," + Artist_Name_Length_String + ",31," + Song_Name_Length_Fix + ",03" + Chr(13)
                'Make sure no spaces got added anywhere
                RTP = Replace(RTP, " ", "")
                'send it
                If (MSComm1.PortOpen) = True Then
                    MSComm1.Output = RTP
                End If
                lblRTP.Caption = RTP
                'Show it on the program
                lblCurrentSong.Caption = "LIVE SESSION WITH " + DPSArtist + " ON WSBF"
                
                'Clear Progress Bar
                progSend.Value = 0
                Wait (1)
                'Update the progress bar, we end up waiting the usual 1.5 seconds before reading the file again
                While progSend.Value < 100
                    progSend.Value = progSend.Value + 1
                    Wait (0.6)
                Wend
                Song_Name_Text = UCase$(FileText(Song_Name_File))
                Artist_Name_Text = UCase$(FileText(Artist_Name_File))
                Song_Name_Text = censorship(Song_Name_Text)
                Artist_Name_Text = censorship(Artist_Name_Text)
                Song_Name_Length = Len(Song_Name_Text)
                Artist_Name_Length = Len(Artist_Name_Text)
                Total_Length = Song_Name_Length + Artist_Name_Length + 12
                DPSArtist = Artist_Name_Text
                DPSSong = Song_Name_Text
            Wend
            
            'Now we begin typical cases
            'Make sure this sucker is less than 128 Characters
            While Total_Length > 128
                'See which of these is the longer culprit
                If Artist_Name_Length > Song_Name_Length Then
                    'remove the last letter from the trouble maker
                    DPSArtist = Left(DPSArtist, Artist_Name_Length - 1)
                Else: DPSSong = Left(DPSSong, Song_Name_Length - 1)
                End If
                'Calculate the new lengths, if removing that char fixed the problem, it exits the loop and moves on
                'otherwise it sees which is longer and subtracts another char
                Song_Name_Length = Len(DPSSong)
                Artist_Name_Length = Len(DPSArtist)
                Total_Length = Song_Name_Length + Artist_Name_Length + 12
            Wend
            'format the output properly and send it
            DPSMake = "DPS=" + DPSSong + " BY " + DPSArtist + " ON WSBF"
            If (MSComm1.PortOpen) = True Then
                MSComm1.Output = DPSMake + Chr(13)
            End If
            'Display what's playing in the program
            lblCurrentSong.Caption = Song_Name_Text + " BY " + Artist_Name_Text
            lblDPS.Caption = DPSMake + Chr(13)
            'Update the time that the song name file was last modified
            Mod_Time = FileDateTime(Song_Name_File)
            Mod_Time_Two = FileDateTime(Artist_Name_File)
            Song_Name_Length = Len(Song_Name_Text)
            Artist_Name_Length = Len(Artist_Name_Text)
            'total lengths of the string that will get sent
            Total_Length = Song_Name_Length + Artist_Name_Length + 5
            'Make sure this sucker is less than 64 Characters
            While Total_Length > 64
                'Hold onto your butts, same junk as before
                If Artist_Name_Length > Song_Name_Length Then
                    Artist_Name_Text = Left(Artist_Name_Text, Artist_Name_Length - 1)
                Else: Song_Name_Text = Left(Song_Name_Text, Song_Name_Length - 1)
                End If
                Song_Name_Length = Len(Song_Name_Text)
                Artist_Name_Length = Len(Artist_Name_Text)
                Total_Length = Song_Name_Length + Artist_Name_Length + 5
            Wend
            'Send the Radio Text out to the transmitter
            If (MSComm1.PortOpen) = True Then
                MSComm1.Output = "TEXT=" + Song_Name_Text + " BY " + Artist_Name_Text + Chr(13)
            End If
            lblTEXT.Caption = "TEXT=" + Song_Name_Text + " BY " + Artist_Name_Text + Chr(13)
            'Calculate the RT+, refer to the Inovonics 730 manual for how to calculate
            Song_Name_Length = Song_Name_Length - 1 'It starts at zero
            'Fix the need for 09 and not just 9
            If (Song_Name_Length < 10) Then
                Song_Name_Length_String = "0" + Str(Song_Name_Length)
            Else
                Song_Name_Length_String = Str(Song_Name_Length)
            End If
            'Ditto, plus the 5 more characters
            Song_Name_Fix = Song_Name_Length + 5
            If (Song_Name_Fix < 10) Then
                Song_Name_Length_Fix = "0" + Str(Song_Name_Fix)
            Else
                Song_Name_Length_Fix = Str(Song_Name_Fix)
            End If
            'same thing
            Artist_Name_Length = Artist_Name_Length - 1
            If (Artist_Name_Length < 10) Then
                Artist_Name_Length_String = "0" + Str(Artist_Name_Length)
            Else
                Artist_Name_Length_String = Str(Artist_Name_Length)
            End If
            RTP = "RTP=01,00," + Song_Name_Length_String + ",04," + Song_Name_Length_Fix + "," + Artist_Name_Length_String + Chr(13)
            'Make sure no spaces got added anywhere
            RTP = Replace(RTP, " ", "")
            If (MSComm1.PortOpen) = True Then
                MSComm1.Output = RTP
            End If
            lblRTP.Caption = RTP
            'Clearing Progress Bar
            progSend.Value = 0
            Wait (0.5)
            'Update the progress bar, we end up waiting the usual 1.5 seconds before reading the file again
            While progSend.Value < 100
                progSend.Value = progSend.Value + 1
                Wait (0.01)
            Wend
        Else
            'So what if the song isn't a special case and hasn't changed the same?
            progSend.Value = 100
            'Just keep the progress bar full and check the song name again in 1.5 seconds
            Wait (1.5)
        End If

    Wend
End Sub
'Written by sibair (Matthieu Poulain and Jean-Marc Mangin) on 10/03/1999 for
'the Interface serial testing program.
'No idea what this does, it's a remnant of the original program that I don't feel comfortable deleting
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (UnloadMode <> vbFormCode) Then
        Unload frmMain
    End If
End Sub

Private Sub cmdStop_Click()
    'What happens when you click Stop
    cmdSend.Visible = True
    cmdStop.Visible = False
    'Kills the loop
    Keep_Going = 0
    'Hides the progress bar
    progSend.Value = 100
    progSend.Visible = False
    'Closes the port
    MSComm1.PortOpen = False
    lblCurrentSong.Visible = False
    lblDPS.Visible = False
    lblTEXT.Visible = False
    lblRTP.Visible = False
End Sub

Private Sub cmdQuit_Click()
    'Kill everything
    Keep_Going = 0
    progSend.Value = 100
    If (MSComm1.PortOpen) = True Then
        MSComm1.PortOpen = False
    End If
    Unload frmSettings
    Unload frmMain
    'Bad practice, but if you don't click stop first the program won't quit right without it
    End
End Sub

Private Sub Form_Load()
    'Pull all the saved values as soon as the program starts, if none are saved, load the defaults
    'French stuff written by sibair (Matthieu Poulain and Jean-Marc Mangin) on 10/03/1999 for
    'the Interface serial testing program.
    Choix_Vitesse = Val(GetSetting(App.Title, "Settings", "Vitesse", 6))
    Choix_Parite = Val(GetSetting(App.Title, "Settings", "Parite", 2))
    Choix_Bit_Donnee = Val(GetSetting(App.Title, "Settings", "Bit_donnee", 4))
    Choix_Bit_Arret = Val(GetSetting(App.Title, "Settings", "Bit_Arret", 0))
    Choix_Flux = Val(GetSetting(App.Title, "Settings", "Flux", 0))
    Choix_Port = Val(GetSetting(App.Title, "Settings", "Port", 1))
    Time_Out_Debut = Val(GetSetting(App.Title, "Settings", "Time_Out_Debut", 5))
    Time_Out_Fin = Val(GetSetting(App.Title, "Settings", "Time_Out_Fin", 5))
    Song_Name_File = GetSetting(App.Title, "Settings", "Song_Name_File", "C:\song_name.txt")
    Artist_Name_File = GetSetting(App.Title, "Settings", "Artist_Name_File", "C:\artist_name.txt")
    Send_On_Launch = Val(GetSetting(App.Title, "Settings", "Send_On_Launch", 0))
    'If you change the send on launch setting to yes and save it, then as soon as the program starts, it sends
    'If you don't do the Me.Show then the window won't show up when auto send is on
    Me.Show
    If (Send_On_Launch = 1) Then
        Wait (1)
        cmdSend_Click
    End If
End Sub

Private Sub Info_Click(Index As Integer)
    'Displays the help dialog
    Dim Reponse
    Dim About As String
    About = "WSBF-FM RDS Sender v1.6.2" + Chr(13) + Chr(10) + "-----------------------------" + Chr(13) + Chr(10)
    About = About + "This program was made by Chief Engineer Kevin Haag in the Summer of 2010." + Chr(13) + Chr(10)
    About = About + "It was created in anticipation of our new RDS encoder as a way of" + Chr(13) + Chr(10)
    About = About + "sending the song name out to the transmitter. A lot of the code for" + Chr(13) + Chr(10)
    About = About + "this program was borrowed from an example serial send/receive program" + Chr(13) + Chr(10)
    About = About + "called Interface written by Matthieu Poulain and Jean-Marc Mangin" + Chr(13) + Chr(10)
    About = About + "on 10/03/1999 which was used to test the serial link on the CD Link. The original" + Chr(13) + Chr(10)
    About = About + "program was translated from French and altered to suit our needs. Some variables are still in French though." + Chr(13) + Chr(10)
    About = About + "Zach Musgrave, the Computer Engineer and later General Manager was concerned" + Chr(13) + Chr(10)
    About = About + "that adding a VB application to the mix was a bad idea, since we were" + Chr(13) + Chr(10)
    About = About + "trying to keep our languages consistent.  He may have been right." + Chr(13) + Chr(10)
    About = About + "If you ever need some help with this thing, email Kevin Haag." + Chr(13) + Chr(10)
    About = About + "He probably still checks odsquad64@gmail.com pretty regularly." + Chr(13) + Chr(10)
    About = About + "Don't use anything newer than VB6 to try to edit it, it won't work." + Chr(13) + Chr(10)
    Reponse = MsgBox(About, vbInformation + vbOKOnly, "About WSBF-FM RDS Sender")
End Sub

Private Sub Help_Click(Index As Integer)
    'Displays the About dialog
    Dim Reponse
    Dim Chain As String
    Chain = "What you need to know :" + Chr(13) + Chr(10) + "-----------------------------" + Chr(13) + Chr(10)
    Chain = Chain + "This program checks every 1.5 seconds to see if the song name file was modified." + Chr(13) + Chr(10)
    Chain = Chain + "If it was, the program sends the song data to the serial port. The serial port" + Chr(13) + Chr(10)
    Chain = Chain + "of this computer should be connected to the STL. You need to write the name" + Chr(13) + Chr(10)
    Chain = Chain + "of the artist to a file FIRST and write the name of the song to a file SECOND." + Chr(13) + Chr(10)
    Chain = Chain + "Use whatever method you'd like to update these files." + Chr(13) + Chr(10)
    Chain = Chain + "The VB6 source code is included somewhere if you need to edit something." + Chr(13) + Chr(10)
    Chain = Chain + "Don't use anything newer than VB6 to try to edit this, it won't work." + Chr(13) + Chr(10)
    Chain = Chain + "If you ever connect an extra serial device to the CD Link, lower the Baud." + Chr(13) + Chr(10)
    Chain = Chain + "It now formats the output for use with the Invonics 730 RDS. RTFM, ditto for the CDLink" + Chr(13) + Chr(10)
    Chain = Chain + "If the song name is Live Sessions, the program sends the properly formatted output." + Chr(13) + Chr(10)
    Chain = Chain + "If the song name is IAMADUCK, it sends exactly what is in the artist name file." + Chr(13) + Chr(10)
    Chain = Chain + "For these special cases, it sends them over and over so the RDS doesn't time out." + Chr(13) + Chr(10)
    Reponse = MsgBox(Chain, vbQuestion + vbOKOnly, "Help for WSBF-FM RDS Sender")
End Sub

Private Sub Settings_Click(Index As Integer)
    'Displays the settings dialog
    frmSettings.Visible = True
    frmMain.Enabled = False
End Sub
