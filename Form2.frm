VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Om"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "123"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.cdindexer.tk"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1297
      MouseIcon       =   "Form2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2160
      Width           =   1740
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v. 1.2.3 beta"
      Height          =   195
      Left            =   2955
      TabIndex        =   4
      Top             =   405
      Width           =   900
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Om språk osv."
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by: Tim Hegyi"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CD Indexer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2520
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
End Sub

Private Sub Label5_Click()
Call ShellExecute(hwnd, "Open", "http://www.cdindexer.tk", "", "", 1)
End Sub
