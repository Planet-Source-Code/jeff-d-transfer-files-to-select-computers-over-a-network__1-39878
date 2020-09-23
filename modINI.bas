Attribute VB_Name = "modINI"
Option Explicit

Sub ReadIni()

  Dim InputLine As String
  Dim StringLen As Long
  Dim SectionName As String
  Dim KeyName As String
  Dim IniFile As String
  Dim intErrMsg As Integer
  
  On Error GoTo ErrReadIni
  
  StringLen = 255
  InputLine = ""
  IniFile = App.Path & "\" & App.EXEName & ".cfg"
  If Not FileExists(IniFile) Then   'TO PUT THIS IN OTHER PROGRAMS, ALSO NEED FileExist FUNCTION BELOW
     Beep
     intErrMsg = MsgBox("CONFIGURATION FILE NOT FOUND!" & NL & NL & _
     "SEARCH FOR AND PLACE IN PROGRAM DIRECTORY A FILE CALLED : " & App.EXEName & ".cfg " & NL & _
                   "PROGRAM WILL TERMINATE.", vbExclamation, "ERROR 001 - UNABLE TO LOCATE CONFIGURATION FILE")
     End
  End If

  SectionName = "MAIN"
  
  'MORE THAN LIKELY, THIS NETWORK COMPUTERS PATH DATABASE WOULD BE STORED ON YOUR
  'DEVELOPMENT MACHINE - SO ONLY YOU CAN ACCESS IT (WITH EXCEPTION OF ADMINISTRATORS)
  
  KeyName = "Entry Background Color"
  InputLine = Space(StringLen)
  RetVal = GetPrivateProfileString(SectionName, KeyName, "Value Not Found", InputLine, StringLen, IniFile)
  InputLine = Left(InputLine, RetVal)
  hexTabColor = InputLine
  
 ' KeyName = "Exe File to Transfer"
 ' InputLine = Space(StringLen)
 ' RetVal = GetPrivateProfileString(SectionName, KeyName, "Value Not Found", InputLine, StringLen, IniFile)
 ' InputLine = Left(InputLine, RetVal)
 ' strExeFile2Xfer = InputLine
  
  KeyName = "Initial Dialogue Box Dir"
  InputLine = Space(StringLen)
  RetVal = GetPrivateProfileString(SectionName, KeyName, "Value Not Found", InputLine, StringLen, IniFile)
  InputLine = Left(InputLine, RetVal)
  strInitXFERSourcePath = InputLine
  
  SectionName = "REPORTS"
    
  SectionName = "AUTORUN"
  
  GoTo EndReadIni
  
ErrReadIni:
  MsgBox "Error in Module ReadIni: " & Err.Number & " - " & Err.Description
  End
  Resume
   
EndReadIni:
End Sub

