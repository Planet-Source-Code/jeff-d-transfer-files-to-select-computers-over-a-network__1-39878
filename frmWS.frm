VERSION 5.00
Begin VB.Form frmWS 
   BackColor       =   &H0000FFFF&
   Caption         =   "COMPUTER INFORMATION FORM"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProgramPath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   60
      TabIndex        =   1
      Top             =   960
      Width           =   5895
   End
   Begin VB.TextBox txtComputerName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   825
      Left            =   3360
      Picture         =   "frmWS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   915
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "&CANCEL"
      Height          =   825
      Left            =   4260
      Picture         =   "frmWS.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Program Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Computer Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
   
   If intSelectWS = 0 Then
      rsXFER.AddNew
   Else
      rsXFER.Edit
   End If
   With rsXFER
      !GroupID = frmWSList.cboGroupName.ItemData(frmWSList.cboGroupName.ListIndex)
      !ComputerName = DAOStr2Field(txtComputerName.Text)
      !ProgramPath = DAOStr2Field(txtProgramPath.Text)
   End With
   rsXFER.Update
   frmWSList.Enabled = True
   Unload Me
End Sub

Private Sub cmdCANCEL_Click()
   Dim MyBookMark As Long
   
   'IF YOU WANT TO CANCEL AN EDIT & STILL KEEP THE SCREEN UP
   If cmdOK.Enabled = True Then
      If intSelectWS = 0 Then
         rsXFER.AddNew     'TO WORK: go to addnew mode, then can cancel it
         frmWSList.Enabled = True
         Close_Up
         Unload Me
      Else
         rsXFER.Edit       'TO WORK: go to edit mode, then can cancel it
      End If
      rsXFER.CancelUpdate
      If txtProgramPath.Text = "" Then
         frmWSList.Enabled = True
         Close_Up
         Unload Me     'If record is blank with an important field, then do close
      Else
         Call FillTextboxes
      End If
      cmdOK.Enabled = False
      cmdCANCEL.Caption = "CLOSE"
   Else
      frmWSList.Enabled = True
      Close_Up
      Unload Me
   End If
End Sub

Public Function DAOStr2Field(ByVal strValue As String) As Variant
   If strValue = "" Then
      DAOStr2Field = Null
   Else
      DAOStr2Field = Trim$(strValue)
   End If
End Function

Private Sub Form_Load()

  If Not intSelectWS = 0 Then    'If SelectKIN has a string value choice
    cmdOK.Enabled = False
    Call FillTextboxes          'Assign the record field values to textboxes
  End If
  cmdCANCEL.Caption = "CLOSE"
  
End Sub

Private Sub FillTextboxes()     'Assign the record field values to textboxes
     txtComputerName.Text = rsXFER!ComputerName & ""
     txtProgramPath.Text = rsXFER!ProgramPath & ""
End Sub

' *************** FIELD KEY-IN VALIDATION CODE **********************


Private Sub txtComputerName_Gotfocus()
   txtComputerName.BackColor = hexTabColor
End Sub
Private Sub txtComputerName_Keypress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
      SendKeys "{TAB}"
      KeyAscii = 0
   End If
   KeyPChg
End Sub
Private Sub txtComputerName_Lostfocus()
   txtComputerName.BackColor = &HFFFFFF
End Sub

Private Sub txtProgramPath_Gotfocus()
   txtProgramPath.BackColor = hexTabColor
End Sub
Private Sub txtProgramPath_Keypress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
      SendKeys "{TAB}"
      KeyAscii = 0
   End If
   KeyPChg
End Sub
Private Sub txtProgramPath_Lostfocus()
   txtProgramPath.BackColor = &HFFFFFF
End Sub

Private Sub KeyPChg()
   cmdOK.Enabled = True
   cmdCANCEL.Caption = "CANCEL"
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
