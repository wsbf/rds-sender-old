VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5055
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   7695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtArtistNameFile 
      Height          =   285
      Left            =   240
      TabIndex        =   26
      Top             =   3960
      Width           =   6975
   End
   Begin VB.CommandButton cmdOpenArtDialog 
      Caption         =   "..."
      Height          =   255
      Left            =   7200
      TabIndex        =   25
      Top             =   3960
      Width           =   255
   End
   Begin VB.OptionButton optNo 
      Caption         =   "No"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton optYes 
      Caption         =   "Yes"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpenDialog 
      Caption         =   "..."
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtSongNameFile 
      Height          =   285
      Left            =   240
      TabIndex        =   23
      Top             =   3240
      Width           =   6975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtStopTimeOut 
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox txtStartTimeOut 
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox txtComPort 
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton cmdResetDefaults 
      Caption         =   "Reset Defaults"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Reset the defaults to those used in 2010"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Close and Use Displayed Settings"
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      ToolTipText     =   "Use the currently displayed settigns."
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton cmdSaveSettings 
      Caption         =   "Save Settings"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      ToolTipText     =   "Save the currently dispalyed settings to the registry to use when the program loads."
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ComboBox cmbFlowControl 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox cmbStopBits 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox cmbDataBits 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ComboBox cmbParity 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.ComboBox cmbBaud 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Artist Name File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Send On Program Launch? "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblSongNameFile 
      Caption         =   "Song Name File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "COM Port to Use :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Stop Time Out in Seconds :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Start Time Out in Seconds :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Flow Control :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "# of Stop Bits :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "# of Data Bits :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Parity :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Baud Rate :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Mostly written by sibair (Matthieu Poulain and Jean-Marc Mangin) on 10/03/1999 for
'the Interface serial testing program.
Private Sub Get_Setting_Values()
    ' This function makes all the variables what is shown in the boxes
    Choix_Vitesse = cmbBaud.ListIndex
    Choix_Parite = cmbParity.ListIndex
    Choix_Bit_Donnee = cmbDataBits.ListIndex
    Choix_Bit_Arret = cmbStopBits.ListIndex
    Choix_Flux = cmbFlowControl.ListIndex
    Choix_Port = Int(Val(txtComPort.Text))
    Time_Out_Debut = Int(Val(txtStartTimeOut.Text))
    Time_Out_Fin = Int(Val(txtStopTimeOut.Text))
    Song_Name_File = txtSongNameFile.Text
    Artist_Name_File = txtArtistNameFile.Text
    If (optYes.Value = True) And (optNo.Value = False) Then
        Send_On_Launch = 1
    Else
        Send_On_Launch = 0
    End If
End Sub

'Mostly written by sibair (Matthieu Poulain and Jean-Marc Mangin) on 10/03/1999 for
'the Interface serial testing program.
Public Sub Display_Settings()
    'This function makes the boxes show what the variables are set to
    ' General Settings
    ' The speed (baud) of the transmission
    cmbBaud.ListIndex = Choix_Vitesse
    ' The parity of the transmission
    cmbParity.ListIndex = Choix_Parite
    ' The number of data bits for the transmission
    cmbDataBits.ListIndex = Choix_Bit_Donnee
    ' The number of stop bits for the transmission
    cmbStopBits.ListIndex = Choix_Bit_Arret
    ' The flow of the transmission
    cmbFlowControl.ListIndex = Choix_Flux
    ' Choose a port
    txtComPort.Text = Mid(Str(Choix_Port), 2)
    ' display information of Time_Out
    txtStartTimeOut.Text = Mid(Str(Time_Out_Debut), 2)
    txtStopTimeOut.Text = Mid(Str(Time_Out_Fin), 2)
    ' display the song name file
    txtSongNameFile.Text = Song_Name_File
    'display the artist name file
    txtArtistNameFile.Text = Artist_Name_File
    'Display the send on launch setting
        If (Send_On_Launch = 1) Then
        optYes.Value = True
        optNo.Value = False
    Else
        optYes.Value = False
        optNo.Value = True
    End If
End Sub

Private Sub cmdOpenArtDialog_Click()
    'Not my code but widely available on the internet, just opens the file selection dialog
    With CommonDialog1
    .InitDir = "C:\"
    .Filter = "*.txt"
    .FilterIndex = 1
    .ShowOpen
    If .filename = "" Then
        Exit Sub
    Else
        myfiletwo = .filename
    End If
    End With
    txtArtistNameFile.Text = myfiletwo
End Sub
'Written by sibair (Matthieu Poulain and Jean-Marc Mangin) on 10/03/1999 for
'the Interface serial testing program.
'No idea what this is
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (UnloadMode <> vbFormCode) Then
        ' l'évenement Unload ne vient pas du code. "Event Unload does not code."
        Unload frmSettings
    End If
End Sub

'When you click Done, it sets the variables to what is in the boxes and shows the main form
Private Sub cmdDone_Click()
    Get_Setting_Values
    frmMain.Enabled = True
    Unload Me
End Sub

'Not my code but widely available on the internet, opens the file select dialog box
Private Sub cmdOpenDialog_Click()
    With CommonDialog1
    .InitDir = "C:\"
    .Filter = "*.txt"
    .FilterIndex = 1
    .ShowOpen
    If .filename = "" Then
        Exit Sub
    Else
        myfile = .filename
    End If
    End With
    txtSongNameFile.Text = myfile
End Sub

'Makes all the values the defaults from the year 2010 and shows them in the box when you click Default
Private Sub cmdResetDefaults_Click()
    Choix_Vitesse = 6
    Choix_Parite = 2
    Choix_Bit_Donnee = 4
    Choix_Bit_Arret = 0
    Choix_Flux = 0
    Choix_Port = 1
    Time_Out_Debut = 5
    Time_Out_Fin = 5
    Song_Name_File = "C:\song_name.txt"
    Artist_Name_File = "C:\artist_name.txt"
    Send_On_Launch = 0
    Display_Settings
End Sub

'Mostly written by sibair (Matthieu Poulain and Jean-Marc Mangin) on 10/03/1999 for
'the Interface serial testing program.
'This saves all the values to the registry
Private Sub cmdSaveSettings_Click()
    Get_Setting_Values
    SaveSetting App.Title, "Settings", "Vitesse", Mid(Str(Choix_Vitesse), 2)
    SaveSetting App.Title, "Settings", "Parite", Mid(Str(Choix_Parite), 2)
    SaveSetting App.Title, "Settings", "Bit_Donnee", Mid(Str(Choix_Bit_Donnee), 2)
    SaveSetting App.Title, "Settings", "Bit_Arret", Mid(Str(Choix_Bit_Arret), 2)
    SaveSetting App.Title, "Settings", "Flux", Mid(Str(Choix_Flux), 2)
    SaveSetting App.Title, "Settings", "Port", Mid(Str(Choix_Port), 2)
    SaveSetting App.Title, "Settings", "Time_Out_Debut", Mid(Str(Time_Out_Debut), 2)
    SaveSetting App.Title, "Settings", "Time_Out_Fin", Mid(Str(Time_Out_Fin), 2)
    SaveSetting App.Title, "Settings", "Song_Name_File", Song_Name_File
    SaveSetting App.Title, "Settings", "Artist_Name_File", Artist_Name_File
    SaveSetting App.Title, "Settings", "Send_On_Launch", Mid(Str(Send_On_Launch), 2)
End Sub

'Mostly written by sibair (Matthieu Poulain and Jean-Marc Mangin) on 10/03/1999 for
'the Interface serial testing program.
'When the form loads it fills the drop down boxes with the available settings
Private Sub Form_Load()
    'The speed of the transmission
    Vitesse(0) = "110"
    frmSettings.cmbBaud.AddItem (Vitesse(0))
    Vitesse(1) = "300"
    frmSettings.cmbBaud.AddItem (Vitesse(1))
    Vitesse(2) = "600"
    frmSettings.cmbBaud.AddItem (Vitesse(2))
    Vitesse(3) = "1200"
    frmSettings.cmbBaud.AddItem (Vitesse(3))
    Vitesse(4) = "2400"
    frmSettings.cmbBaud.AddItem (Vitesse(4))
    Vitesse(5) = "4800"
    frmSettings.cmbBaud.AddItem (Vitesse(5))
    Vitesse(6) = "9600"
    frmSettings.cmbBaud.AddItem (Vitesse(6))
    Vitesse(7) = "14400"
    frmSettings.cmbBaud.AddItem (Vitesse(7))
    Vitesse(8) = "19200"
    frmSettings.cmbBaud.AddItem (Vitesse(8))
    Vitesse(9) = "28800"
    frmSettings.cmbBaud.AddItem (Vitesse(9))
    Vitesse(10) = "38400"
    frmSettings.cmbBaud.AddItem (Vitesse(10))
    Vitesse(11) = "56000"
    frmSettings.cmbBaud.AddItem (Vitesse(11))
    Vitesse(12) = "128000"
    frmSettings.cmbBaud.AddItem (Vitesse(12))
    Vitesse(13) = "256000"
    frmSettings.cmbBaud.AddItem (Vitesse(13))
    
    Rem The parity of the transmission
    Parite(0) = "E"
    frmSettings.cmbParity.AddItem (Parite(0))
    Parite(1) = "M"
    frmSettings.cmbParity.AddItem (Parite(1))
    Parite(2) = "N"
    frmSettings.cmbParity.AddItem (Parite(2))
    Parite(3) = "O"
    frmSettings.cmbParity.AddItem (Parite(3))
    Parite(4) = "S"
    frmSettings.cmbParity.AddItem (Parite(4))
    
    Rem The number of data bits of the transmission
    Bit_Donnee(0) = "4"
    frmSettings.cmbDataBits.AddItem (Bit_Donnee(0))
    Bit_Donnee(1) = "5"
    frmSettings.cmbDataBits.AddItem (Bit_Donnee(1))
    Bit_Donnee(2) = "6"
    frmSettings.cmbDataBits.AddItem (Bit_Donnee(2))
    Bit_Donnee(3) = "7"
    frmSettings.cmbDataBits.AddItem (Bit_Donnee(3))
    Bit_Donnee(4) = "8"
    frmSettings.cmbDataBits.AddItem (Bit_Donnee(4))
    
    Rem The number of stop bits of the transmission
    Bit_Arret(0) = "1"
    frmSettings.cmbStopBits.AddItem (Bit_Arret(0))
    Bit_Arret(1) = "1.5"
    frmSettings.cmbStopBits.AddItem (Bit_Arret(1))
    Bit_Arret(2) = "2"
    frmSettings.cmbStopBits.AddItem (Bit_Arret(2))
    
    Rem The Flow Control
    Flux(0) = "comNone"
    frmSettings.cmbFlowControl.AddItem (Flux(0))
    Flux(1) = "comXOnXOff"
    frmSettings.cmbFlowControl.AddItem (Flux(1))
    Flux(2) = "comRTS"
    frmSettings.cmbFlowControl.AddItem (Flux(2))
    Flux(3) = "comRTSXOnXOff"
    frmSettings.cmbFlowControl.AddItem (Flux(3))
    
    'And then it shows the settings
    Display_Settings
End Sub

