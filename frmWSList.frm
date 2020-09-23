VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmWSList 
   BackColor       =   &H0000C000&
   Caption         =   "Transfer a File to Computers on a Network Example"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Object.ToolTipText     =   "Add a new computer to this group"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Edit"
            Object.ToolTipText     =   "Edit the selected record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete the Highlighted Record"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print Report"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Transfer"
            Object.ToolTipText     =   "Transfer Selected File to Network Computers shown in the Grid below"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reset"
            Object.ToolTipText     =   "Reset the Listing"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Group"
            Object.ToolTipText     =   "User Group Maintenance"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            Object.ToolTipText     =   "CLOSE the PROGRAM"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "frmWSList.frx":0000
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10200
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   49152
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWSList.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWSList.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWSList.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWSList.frx":0C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWSList.frx":0F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWSList.frx":3734
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWSList.frx":3A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWSList.frx":3D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWSList.frx":4082
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboGroupName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGridWS 
      Bindings        =   "frmWSList.frx":4C54
      Height          =   5055
      Left            =   120
      OleObjectBlob   =   "frmWSList.frx":4C68
      TabIndex        =   5
      Top             =   1440
      Width           =   10335
   End
   Begin VB.CommandButton Command1 
      Height          =   435
      Left            =   10200
      Picture         =   "frmWSList.frx":5B64
      TabIndex        =   4
      ToolTipText     =   "Transfer Selected File to Computers"
      Top             =   1920
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdCLOSE 
      Height          =   435
      Left            =   10320
      Picture         =   "frmWSList.frx":5E6E
      TabIndex        =   3
      ToolTipText     =   "Close Program"
      Top             =   1320
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Data DatWS 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.CommandButton cmdTransferFile 
      Caption         =   "..."
      Height          =   375
      Left            =   9960
      TabIndex        =   0
      ToolTipText     =   "Select File To Transfer"
      Top             =   840
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "File to Transfer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "User Group:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmWSList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

    Select Case Button.Key
        Case "Add"
            cmdADD_Click
        Case "Edit"
            cmdEDIT_Click
        Case "Delete"
            cmdDelete_Click
        Case "Print"
            cmdPrint_Click
        Case "Transfer"
            Command1_Click
        Case "Close"
            cmdCLOSE_Click
        Case "Reset"
            cmdReset_Click
        Case "Group"
            cmdGroup_Click
    End Select
End Sub

Private Sub cmdGroup_Click()
   Me.Enabled = False
   frmGroupList.Show
End Sub

Private Sub cmdReset_Click()    'CLEARS THE GRID AND ALL SELECTION VALUES

   Dim strSQL As String
   
   'ERASE VERSION, DATE, and FILENAME from Database
   rsXFER.MoveFirst
   Do Until rsXFER.EOF
      rsXFER.Edit
      With rsXFER
         !VersionEXE = ""
         !DateEXE = 0
         !OtherFileName = ""
      End With
      rsXFER.Update
      rsXFER.MoveNext
   Loop
   rsXFER.Close
   
   Set rsXFER = Nothing
   Set dbXFER = Nothing
   Set dbXFER = OpenDatabase(App.Path & "\Transfer.mdb")
   intSelectWS = 0     'Erases the intSelectWS value so another can be picked
   Text1.Text = ""
  
  'LOAD THE COMBO BOX SELECTING WHICH GROUP OF COMPUTERS TO TRANSFER TO.
  Call GroupComboLoad(cboGroupName)
  strSQL = "Select * from tblXFER WHERE GroupID = " & 0
  Set rsXFER = dbXFER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  Set DatWS.Recordset = rsXFER
  intSelectWS = 0
  
  'DISENABLE THESE BUTTONS - EXCEPT FOR CLOSE BUTTON
  Toolbar1.Buttons.Item(1).Enabled = False
  Toolbar1.Buttons.Item(2).Enabled = False
  Toolbar1.Buttons.Item(3).Enabled = False
  Toolbar1.Buttons.Item(5).Enabled = False
  Toolbar1.Buttons.Item(6).Enabled = False
  Toolbar1.Buttons.Item(8).Enabled = False
End Sub

Private Sub cboGroupName_Click()
Dim strSQL As String

'  If cboGroupName.Tag <> cboGroupName.Text Then
      cboGroupName.Tag = cboGroupName.Text
      strSQL = "Select * from tblXFER WHERE GroupID = " & cboGroupName.ItemData(cboGroupName.ListIndex)

     Set rsXFER = dbXFER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  
     On Error Resume Next       'So error 3021 NO RECORD FOUND Error won't occur
     rsXFER.MoveLast
     Set DatWS.Recordset = rsXFER
     On Error Resume Next
     DatWS.Recordset.MoveFirst
     
     'IF NO ERROR OCCURS AT THIS POINT, ENABLE THE COMBO BOXES
     Toolbar1.Buttons.Item(1).Enabled = True
     Toolbar1.Buttons.Item(2).Enabled = True
     Toolbar1.Buttons.Item(3).Enabled = True
     Toolbar1.Buttons.Item(5).Enabled = True
     Toolbar1.Buttons.Item(6).Enabled = True
     Toolbar1.Buttons.Item(8).Enabled = True
'  End If
End Sub

Private Sub cmdCLOSE_Click()
   Set rsXFER = Nothing
   Set dbXFER = Nothing
  ' frmMain.Enabled = True
   Close_Up
   Unload Me
   End
End Sub


Private Sub Command1_Click()
  
  Dim strSQL As String
  Dim blnError As Boolean
  Dim fso As New FileSystemObject       'FileSystemObject must be REFERENCED
  Dim Fexe As File                      'Through "MICROSOFT SCRIPTING RUNTIME"
  Dim intCount As Integer
  Dim strGetFileNameWExt As String
  Dim strGetFileExt As String
  Dim strErrDesc As String
  
  
  'THIS IS THE MAIN PROCEDURE FOR THIS PSC SUBMISSION - IT WOULD HELP TO STEP THROUGH THIS CODE
  
  
  If Text1.Text = "" Then
     intCount = MsgBox("You must select a File to Transfer from the Drop-Down List.", vbExclamation, "SELECT FILE TO TRANSFER")
     Exit Sub
  Else
     'STRING HOLDS THE VALUE OF "FILE TO TRANSFER"
     strExeFile2Xfer = Text1.Text
     Set Fexe = fso.GetFile(strExeFile2Xfer)
  End If
  
  'Gets information from File - Needs Microsoft Scripting Runtime Reference
  strGetFileNameWExt = Mid$(strExeFile2Xfer, InStrRev(strExeFile2Xfer, "\") + 1)
  strGetFileExt = getextension(strExeFile2Xfer)
  
  On Error GoTo cmdERROR_CLick
  

  rsXFER.MoveLast
  PB1.Max = rsXFER.RecordCount
  rsXFER.MoveFirst
  
  PB1.Visible = True
  Do Until rsXFER.EOF
     intCount = intCount + 1
     PB1.Value = intCount
     rsXFER.Edit
     
     
     'COPY THE FILE FROM YOUR LOCATION (Text1.text) TO THE NETWORK LOCATION
     'ACCORDING TO Transfer.mdb RECORD
     
     ''''''FileCopy "c:\jeff\TraceSys\MWTrace.exe", rsXFER!ProgramPath & "MWTrace.exe"
     FileCopy strExeFile2Xfer, rsXFER!ProgramPath & "\" & strGetFileNameWExt
     
     'Following IF..Else..EndIf:
     'If a particular user has the EXE file open on their computer, an error 70
     'PERMISSION DENIED (to copy file) will occur.  If this occurs, NO entry will be made
     'to the Transfer.mdb database regarding VERSION and Date.  However, this also
     'error checks for an error 76 indicating that either the computer if off
     'or the path in the Transfer.mdb is not correctly entered (or no longer valid).
    
     If blnError = True Then
         rsXFER!VersionEXE = strErrDesc
         rsXFER!OtherFileName = ""
         rsXFER!DateEXE = 0
     Else
         rsXFER!VersionEXE = fso.GetFileVersion(strExeFile2Xfer)
         rsXFER!DateEXE = Fexe.DateLastModified
         rsXFER!OtherFileName = strGetFileNameWExt
     End If
     
     blnError = False   'RESETS BACK TO FALSE FOR NEXT rsXFER RECORD TO CONSIDER
     rsXFER.Update
     rsXFER.MoveNext
     DatWS.Refresh
  Loop
  
  GoTo EndCmdERROR_CLick
  
cmdERROR_CLick:
  
  If Err.Number = 70 Then   'PERMISSION DENIED ERROR - SOMEONE HAS THE .EXE RUNNING ON THEIR COMPUTER
     strErrDesc = "PERMISSION DENIED"
     blnError = True
     Resume Next
  ElseIf Err.Number = 76 Then 'PATH NOT FOUND - INCORRECTLY ENTERED IN DATABASE OR COMPUTER IS OFF
     strErrDesc = "PATH NOT FOUND"
     blnError = True
     Resume Next
  ElseIf Err.Number = 52 Then 'PATH NOT FOUND - INCORRECTLY ENTERED IN DATABASE OR COMPUTER IS OFF
     strErrDesc = "BAD FILE NAME"
     blnError = True
     Resume Next
  Else
     MsgBox Err.Number & " - " & Err.Description
  End If
  
EndCmdERROR_CLick:
  MsgBox "DONE"
  PB1.Value = 0
  PB1.Visible = False
  
End Sub

Private Sub cmdTransferFile_Click()
   CommonDialog1.InitDir = strInitXFERSourcePath
   'CommonDialog1.DefaultExt = "xls"
   CommonDialog1.Filter = "ALL Files (*.*)|*.*|Word Documents (*.doc)|*.doc|Excel Spreadsheets (*.xls)|Access Databases (*.mdb)"
   
   CommonDialog1.Flags = cdlOFNHideReadOnly + _
      cdlOFNFileMustExist + cdlOFNPathMustExist
   CommonDialog1.ShowOpen
   Text1.Text = CommonDialog1.filename
End Sub

Private Sub cmdADD_Click()
   intSelectWS = 0
   frmWS.cmdOK.Enabled = False
   frmWSList.Enabled = False
   frmWS.Show
   frmWSList.Refresh
End Sub


Private Sub cmdEDIT_Click()
   Dim RecordBookMark As Long
   
   If rsXFER.RecordCount > 0 Then
      RecordBookMark = DatWS.Recordset.AbsolutePosition
      On Error Resume Next
      intSelectWS = DatWS.Recordset("RecNo")
   End If
   frmWS.cmdOK.Enabled = False
   frmWSList.Enabled = False
   frmWS.Show
   frmWS.cmdCANCEL.SetFocus
End Sub

Private Sub Form_Load()
  Dim strSQL As String
  
  'THIS PROGRAM USES DAO - IT COULD HAVE BEEN WRITTEN WITH ADO, BUT THIS PROGRAM
  'WAS WRITTEN BACK IN EARLY 2000 - FEEL FREE TO CHANGE TO ADO IF YOU WISH
  
  Set dbXFER = OpenDatabase(App.Path & "\Transfer.mdb")
  intSelectWS = 0     'Erases the intSelectWS value so another can be picked
  
  'LOAD THE COMBO BOX SELECTING WHICH GROUP OF COMPUTERS TO TRANSFER TO.
  Call GroupComboLoad(cboGroupName)
  
  'TO BEGIN WITH, DISENABLE THESE BUTTONS - EXCEPT FOR CLOSE BUTTON
  Toolbar1.Buttons.Item(1).Enabled = False
  Toolbar1.Buttons.Item(2).Enabled = False
  Toolbar1.Buttons.Item(3).Enabled = False
  Toolbar1.Buttons.Item(5).Enabled = False
  Toolbar1.Buttons.Item(6).Enabled = False
  Toolbar1.Buttons.Item(8).Enabled = False
  
End Sub

''''PRIVATE SUB NEEDED TO SELECT RECORD OFF THE TABLE''''
Private Sub dbGridWS_RowColChange(LastRow As Variant, ByVal LostCol As Integer)
  Dim MyBookMark As Long
  
  On Error Resume Next
  If rsXFER.RecordCount > 0 Then
     MyBookMark = DatWS.Recordset.AbsolutePosition
     On Error Resume Next
     intSelectWS = DatWS.Recordset("RecNo")
  End If
End Sub

Private Sub cmdDelete_Click()
  Dim DelAnswer As Integer
  Dim strSQL As String
   
  DelAnswer = MsgBox("Are you sure you want to DELETE this record?" _
               & NL & NL & DatWS.Recordset("ComputerName"), _
               vbExclamation + vbYesNo, "DELETE RECORD")
  If DelAnswer = vbYes Then
     On Error Resume Next   'Will not cause 3021 No Record error
     rsXFER.Delete
  Else
     Exit Sub 'and do nothing
  End If
  
  'RESET THE DATA CONTROL & GRID
  rsXFER.Close
  strSQL = "Select * from tblXFER WHERE GroupID = " & cboGroupName.ItemData(cboGroupName.ListIndex)
  Set rsXFER = dbXFER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  Set DatWS.Recordset = rsXFER
  intSelectWS = 0
End Sub

Private Sub cmdPrint_Click()
   MsgBox ("PRINT FUNCTION IS NOT CURRENTLY ACTIVE" & NL & "Look for it in version 2.0")
End Sub
Private Sub Close_Up()

Dim x As Long
Dim inc As Long
'inc = 40
inc = 180

   For x = Me.Height To 300 Step -inc
      DoEvents
      Me.Move Me.Left, Me.Top + (inc \ 2), Me.Width, x
   Next x

  'This is the width part of the same sequence above
   For x = Me.Width To 2000 Step -inc
      DoEvents
      Me.Move Me.Left + (inc \ 2), Me.Top, x, Me.Height
   Next x
End Sub

