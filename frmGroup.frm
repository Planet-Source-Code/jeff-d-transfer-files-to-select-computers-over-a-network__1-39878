VERSION 5.00
Begin VB.Form frmGroup 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Group Entry Form"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGroupName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   2
      Top             =   360
      Width           =   4965
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "&CANCEL"
      Height          =   825
      Left            =   3360
      Picture         =   "frmGroup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   825
      Left            =   2460
      Picture         =   "frmGroup.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Group Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   825
   End
End
Attribute VB_Name = "frmGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
   
   If intSelectGroup = 0 Then
      rsGroup.AddNew
   Else
      rsGroup.Edit
   End If
   With rsGroup
      !GroupName = DAOStr2Field(txtGroupName)
   End With
   rsGroup.Update
   frmGroupList.Enabled = True
   Close_Up
   Unload Me
End Sub

Private Sub cmdCANCEL_Click()
   Dim MyBookMark As Long
   
   'IF YOU WANT TO CANCEL AN EDIT & STILL KEEP THE SCREEN UP
   If cmdOK.Enabled = True Then
      If intSelectGroup = 0 Then
         rsGroup.AddNew     'TO WORK: go to addnew mode, then can cancel it
         frmGroupList.Enabled = True
         Close_Up
         Unload Me
      Else
         rsGroup.Edit       'TO WORK: go to edit mode, then can cancel it
      End If
      rsGroup.CancelUpdate
      If txtGroupName.Text = "" Then
         frmGroupList.Enabled = True
         Close_Up
         Unload Me     'If record is blank with an important field, then do close
      Else
         Call FillTextboxes
      End If
      cmdOK.Enabled = False
      cmdCANCEL.Caption = "CLOSE"
   Else
      frmGroupList.Enabled = True
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
  
  If Not intSelectGroup = 0 Then    'If SelectKIN has a string value choice
    cmdOK.Enabled = False
    Call FillTextboxes          'Assign the record field values to textboxes
  End If
  cmdCANCEL.Caption = "CLOSE"
End Sub

Private Sub FillTextboxes()     'Assign the record field values to textboxes
     txtGroupName = rsGroup!GroupName & ""
End Sub

' *************** FIELD KEY-IN VALIDATION CODE **********************

Private Sub txtGroupName_Gotfocus()
   PreSel txtGroupName
End Sub
Private Sub txtGroupName_Keypress(KeyAscii As Integer)
   KeyPChg
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
