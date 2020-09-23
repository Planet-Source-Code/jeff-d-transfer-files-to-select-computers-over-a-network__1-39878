Attribute VB_Name = "modPrimary"
Option Explicit

'CREDITS - http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=1886&lngWId=1
'           Eugene -"Another way to close a form - 5/26/1999"
'           VARIOUS FILE FUNCTIONS CONTAINED IN THIS MODULE - NOT SURE WHERE CAME FROM

'Program variables
Public RetVal As Variant
Public ExePath As String
Public NL As String
Public Release As Variant
Public strExeFile2Xfer As String    'This is the EXECUTABLE you want to TRANSFER
Public strUserName As String
Public strInitXFERSourcePath As String
Public intSelectWS As Long       'Lets you select which WORKSTATION to EDIT or DELETE
Public intSelectGroup As Long
Public rsXFER As Recordset            'Recordset for WORKSTATION table
Public rsGroup As Recordset
Public dbXFER As Database
Public hexTabColor As Variant

Dim strComputerName As String

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Const UNKNOWN = "(Value Unknown Because System Call Failed)"
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer$, nSize As Long) As Long

Sub Main()
   
Dim boolPerform As Boolean
Dim intErrMsg As Integer

Dim intXFER As Integer
Dim intCreateTextfile As Integer
Dim strComputerUpdatedName As String
Dim strlblName As String    'SHOWS FILE BEING UPDATED ON SPLASH SCREEN LABEL

  On Error GoTo ERRErrMsg
  
  If UCase(CurDir) Like "*VB*" Then
    ExePath = "C:\Jeff"
    SetCurrentDirectory (ExePath)
  Else
    ExePath = App.Path
    SetCurrentDirectory (ExePath)
  End If
  
  If App.PrevInstance = True Then
    RetVal = MsgBox("Only one instance of Material Tracability System " & _
             "may be run at a time.", , "Multiple Instance Running")
    End
  End If
  Screen.MousePointer = vbHourglass
  NL = Chr$(13) + Chr$(10)
  ReadIni
  strUserName = GetCurrentUserName
  Screen.MousePointer = vbNormal
  ''''''''''''''''''''''''''''''''''''''''frmMain.Show
  frmWSList.Show

  GoTo EndNoError
   
ERRErrMsg:
   If Err.Number = 3055 Or Err.Number = 3024 Then
      Beep
      intErrMsg = MsgBox("DATABASE FILE NOT AVAILABLE AS INDICATED IN CONFIG FILE LOCATED IN :" & NL & _
                  App.Path & "\" & App.EXEName & ".cfg" & NL & NL & _
                   "CHECK YOUR CONFIG FILE AND/OR PATH TO THE DATABASE FILE.", vbExclamation, "ERROR 002 - UNABLE TO LOCATE DATABASE FILE")
      End
   Else
      MsgBox Err.Number & " - " & Err.Description
   End If
   
EndNoError:
End Sub

' ****************************************************************
' *** REMAINING CODE IN modPrimary FOUND ON PSC OR OTHER SITES ***
' ***   NOTE: Some Functions may not be used by this program   ***
' ****************************************************************

Public Function GetCurrentUserName() As String

Dim l As Long
Dim sUser As String

sUser = Space$(255)
l = GetUserName(sUser, 255)

If l <> 0 Then
   GetCurrentUserName = Left(sUser, InStr(sUser, Chr(0)) - 1)
Else
   Err.Raise Err.LastDllError, , _
     "A system call returned an error code of " _
      & Err.LastDllError
End If

End Function

Private Function ComputerName() As String
    Dim cn As String * 255
    Dim result As Long
    Dim nSize As Long
    
    nSize = 256
    result = GetComputerName(cn, nSize)
    ComputerName = Trim(Left(cn, nSize))
End Function

Public Sub PreSel(txt As Control)
  If TypeOf txt Is TextBox Then
     txt.SelStart = 0
     txt.SelLength = Len(txt)
  End If
End Sub

Public Function DAOStr2Field(ByVal strValue As String) As Variant
   If strValue = "" Then
      DAOStr2Field = Null
   Else
      DAOStr2Field = Trim$(strValue)
   End If
End Function

Public Function FileExists(fName As String) As Boolean

If fName = "" Or Right(fName, 1) = "\" Then
  FileExists = False: Exit Function
End If

FileExists = (Dir(fName) <> "")

End Function

Public Function getextension(filename As String) As String   'GETS THE EXTENSION WITHOUT "."
Dim I As Integer
Dim C As String
Dim pos As Integer

For I = Len(filename) To 2 Step -1
  C = Mid(filename, I, 1)
  If C = "." Then
  pos = I + 1
  End If
Next

getextension = Mid(filename, pos, (Len(filename) + 1 - pos))
End Function

Function GetFilepath(Path As String)
    GetFilepath = Mid(Path, 1, InStrRev(Path, "\"))
End Function

Public Function getfiletitle(filename As String) As String 'GETS THE FILE TITLE (NO DRIVE LETTER)
Dim I As Integer
Dim C As String
Dim pos As Integer

For I = Len(filename) To 2 Step -1
  C = Mid(filename, I, 1)
  If C = "\" Then
  pos = I + 1
  End If
Next

getfiletitle = Mid(filename, pos, (Len(filename) + 1 - pos))
End Function

Public Function GetFileName(flname As String) As String   'GETS THE FILE NAME WITHOUT PATH OR EXTENSION
    
    'Get the filename without the path or extension.
    'Input Values:
    '   flname - path and filename of file.
    'Return Value:
    '   GetFileName - name of file without the extension.
    
    Dim posn As Integer, I As Integer
    Dim fName As String
    
    posn = 0
    'find the position of the last "\" character in filename
    For I = 1 To Len(flname)
        If (Mid(flname, I, 1) = "\") Then posn = I
    Next I

    'get filename without path
    fName = Right(flname, Len(flname) - posn)

    'get filename without extension
    posn = InStr(fName, ".")
        If posn <> 0 Then
            fName = Left(fName, posn - 1)
        End If
    GetFileName = fName
End Function

