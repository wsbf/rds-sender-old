Attribute VB_Name = "Variable"
'
' Gloabl Variable Definitions
'
'Anything in French written by sibair (Matthieu Poulain and Jean-Marc Mangin)
'on 10/03/1999 for the Interface serial testing program.
' Speed of the transmission (baud)
Public Choix_Vitesse As Integer
Public Vitesse(0 To 13) As String

' The parity
Public Choix_Parite As Integer
Public Parite(0 To 4) As String

' number of data bits
Public Choix_Bit_Donnee As Integer
Public Bit_Donnee(0 To 4) As String

' number of stop bits
Public Choix_Bit_Arret As Integer
Public Bit_Arret(0 To 2) As String

' flow control
Public Choix_Flux As Integer
Public Flux(0 To 3) As String

' Choose the port for communication
Public Choix_Port As Integer

' Definition of time_out_debut to receive file
Public Time_Out_Debut As Integer

' Definition of time_out_fin to receive file
Public Time_Out_Fin As Integer

' For the infinite loop
Public Keep_Going As Integer

'For the song name file
Public Song_Name_File As String
Public myfile As String
Public Song_Name_Text As String
Public Song_Name_Length As Integer
Public Song_Name_Length_String As String

'For the artist name file
Public Artist_Name_File As String
Public Artist_Name_Text As String
Public myfiletwo As String
Public Artist_Name_Length As Integer
Public Artist_Name_Length_String As String

'To see if the program will send on launch
Public Send_On_Launch As Integer
