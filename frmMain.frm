VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0000C000&
   Caption         =   "GULF MARINE FABRICATORS - MATERIAL TRACABILITY SYSTEM"
   ClientHeight    =   2580
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEXIT 
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "Ends the Material-Weld Tracability System"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Program's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Ends the Material-Weld Tracability System"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Program's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      ToolTipText     =   "Ends the Material-Weld Tracability System"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Program's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Ends the Material-Weld Tracability System"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "! ! ! YOUR APPLICATION HERE ! ! !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   6060
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu mnuWorkstations 
      Caption         =   "&Workstations"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'*******************************************************************************
'HERE YOU HAVE YOUR APPLICATION'S MAIN FORM CODE - PROBABLY CONTROLS AND BUTTONS
'LEADING TO OTHER PARTS OF YOUR PROGRAM.
'*******************************************************************************

Private Sub Form_Load()

    HideMenuFromRegUsers
    
End Sub

Private Sub HideMenuFromRegUsers()
'THIS SUBROUTINE LOOKS FOR A NETWORK USER-ID.  ONLY THIS PERSON, (YOU THE PROGRAMMER),
'WILL BE ABLE TO TRANSFER UPDATED EXECUTABLES TO OTHER USERS MACHINES BECAUSE THE
'COMMAND BUTTON AND PROGRESS BAR WILL BE MADE AVAILABLE TO YOU.  BE SURE THAT
'YOUR NETWORK LOGGON USERID IS REFLECTED IN THE LINE BELOW!
'i.e. Change ------ If strUserName = "agmfjcd" to "YOUR-LOGGON-ID"

'THIS USERID INFORMATION COULD BE OBTAINED BY OTHER METHODS, BUT TO KEEP IT SIMPLE...

    If strUserName = "agmfjcd" Then
       'ALLOW XFER BUTTON TO SHOW - Command1 Button and PB1 - Progress Bar, etc.
    Else
       mnuWorkstations.Visible = False   'So no one else can see WORKSTATIONS menu item
    End If
    
    'YOUR APP INFO BELOW HERE
    lblVersion.Caption = "Material/Weld Tracability System - " & _
    "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Private Sub cmdEXIT_Click()
   
   Unload Me
   End
End Sub

Private Sub mnuExit_Click()
   cmdEXIT_Click
End Sub

Private Sub mnuWorkstations_Click()
   frmMain.Enabled = False
   frmWSList.Show
End Sub
