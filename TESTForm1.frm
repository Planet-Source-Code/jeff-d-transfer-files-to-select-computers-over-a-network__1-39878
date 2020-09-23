VERSION 5.00
Begin VB.Form TESTForm1 
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEXIT 
      Caption         =   "&EXIT"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   $"TESTForm1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   $"TESTForm1.frx":0127
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "TESTForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEXIT_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
