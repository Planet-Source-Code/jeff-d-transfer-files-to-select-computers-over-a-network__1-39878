VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmGroupList 
   BackColor       =   &H00FF0000&
   Caption         =   "GROUPS LISITNG"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGridGroup 
      Bindings        =   "frmGroupList.frx":0000
      Height          =   5565
      Left            =   180
      OleObjectBlob   =   "frmGroupList.frx":0017
      TabIndex        =   4
      Top             =   270
      Width           =   3345
   End
   Begin VB.Data DatGroup 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   990
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   90
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "&CLOSE"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   6090
      Width           =   825
   End
   Begin VB.CommandButton cmdDELETE 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   1830
      TabIndex        =   2
      Top             =   6090
      Width           =   825
   End
   Begin VB.CommandButton cmdEDIT 
      Caption         =   "&EDIT"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   6090
      Width           =   825
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "&ADD"
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   6090
      Width           =   825
   End
End
Attribute VB_Name = "frmGroupList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdADD_Click()
   intSelectGroup = 0
   frmGroup.cmdOK.Enabled = False
   frmGroupList.Enabled = False
   frmGroup.Show
   frmGroup.txtGroupName.SetFocus
   frmGroupList.Refresh
End Sub

Private Sub cmdCLOSE_Click()    'Look at frmProjects - cmdAddMoreGroup_Click()

   frmWSList.Enabled = True
   Close_Up
   Unload Me
End Sub

Private Sub cmdEDIT_Click()
   Dim RecordBookMark As Long
   
   If rsGroup.RecordCount > 0 Then
      RecordBookMark = DatGroup.Recordset.AbsolutePosition
      On Error Resume Next
      intSelectGroup = DatGroup.Recordset("GroupID")
   End If
   frmGroup.cmdOK.Enabled = False
   frmGroupList.Enabled = False
   frmGroup.Show
   frmGroup.cmdCANCEL.SetFocus
End Sub

Private Sub Form_Load()
  Dim strSQL As String
  
  intSelectGroup = 0     'Erases the intSelectGroup value so another can be picked
  strSQL = "Select * from tblGroup"
  Set rsGroup = dbXFER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  On Error Resume Next
  rsGroup.MoveLast
  Set DatGroup.Recordset = rsGroup
  On Error Resume Next
  DatGroup.Recordset.MoveFirst
End Sub

''''PRIVATE SUB NEEDED TO SELECT RECORD OFF THE TABLE''''
Private Sub dbGridGroup_RowColChange(LastRow As Variant, ByVal LostCol As Integer)
  Dim MyBookMark As Long
  
  If rsGroup.RecordCount > 0 Then
     MyBookMark = DatGroup.Recordset.AbsolutePosition
     On Error Resume Next
     intSelectGroup = DatGroup.Recordset("Group")
  End If
End Sub

Private Sub cmdDelete_Click()
  Dim DelAnswer As Integer
  Dim strSQL As String
   
  DelAnswer = MsgBox("Are you sure you want to DELETE this record?" _
               & NL & NL & DatGroup.Recordset("GroupName"), _
               vbExclamation + vbYesNo, "DELETE RECORD")
  If DelAnswer = vbYes Then
     On Error Resume Next   'Will not cause 3021 No Record error
     rsGroup.Delete
  Else
     Exit Sub 'and do nothing
  End If
  rsGroup.Close
  strSQL = "Select * FROM tblGroup"
  Set rsGroup = dbXFER.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
  Set DatGroup.Recordset = rsGroup
  intSelectGroup = 0
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
