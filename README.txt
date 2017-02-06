WSBF_RDS_Sender.exe is licensed under the QPPL v1.0, which is included with the source code.

Make sure you have MSCOMM32.OCX in your system32 folder. You'll get an error without it. There might be another OCX missing, but most computers have it. If you get an error that one is missing, just download the other one from somewhere.

The default song name location is C:\song_name.txt and C:\artist_name.txt is the default artist name location. You can set it to where ever you want though. Currently the logbook is set to output the song and artist name to the files that are being used on the webserver every time the Now Playing changes. If we ever move to a new logbook, all you have to do is remember to update these files however you'd like and this program will still work.

You can set this program to send as soon as it launches and make it a startup program on the webserver so that so long as the webserver is on, you will be sending the song name.

If you make the name of a song "Live Session" or "Live Sessions" the program will output "LIVE SESSION WITH Artist_Name ON WSBF" for DPS and "Artist_Name LIVE ON WSBF" for RT. It will also continuously send the name so the RDS encoder doesn't time out to the default text. If someone doing a live session forgets to log out, this will send forever.

Also, if you make the song name "IAMADUCK" the program will send out exactly what is typed in for the artist name. Use this for special events. If we're doing a Townhall Meeting with President Barker, make the song name "IAMADUCK" and make the artist name "TOWNHALL MEETING WITH PRESIDENT BARKER LIVE ON WSBF" and that is what will show up on people's radios. It's nifty and you can do it straight from the logbook.

The serial port of the computer you want to run this on should be connected to one of the Data ports on the CD Link Transmitter using a 9 pin female to 25 pin female serial cable. Out at the transmitter shack, the special looking cable that I made (which is actually just a straight through cable that you can change into another type of cable) should be connecting the CD Link Receiver to the RDS encoder. The settings on the CD Link Transmitter and Receiver need to be the same as the settings in this program. The CD Link manual will tell you how to change these settings. The maximum baud for the CD Link is 9600. If you add another serial device you will have to lower the baud for each Data port so that the total for both devices is 9600 or less. This program and the CD Link support even and odd parity and some other stuff we don't need to use because the RDS encoder doesn't support them.

This program is based off a serial testing program called Interface written by sibair (Matthieu Poulain and Jean-Marc Mangin) and includes some portions of their original source code. Interface is available at http://www.freevbcode.com/ShowCode.asp?ID=2236
