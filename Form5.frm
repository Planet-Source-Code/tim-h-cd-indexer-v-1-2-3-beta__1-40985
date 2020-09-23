VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "skapa spegling va enhet"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "steg 3"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   5175
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Text            =   "*.*"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Vilka filtyper som ska registreras (*.*  = alla filer)"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "steg 4"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   5175
      Begin VB.CommandButton Command2 
         Caption         =   "Avbryt"
         Height          =   330
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   330
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "steg 2"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Text            =   "noname"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Namnet som är inskrivet går bra att använda."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   2895
         Visible         =   0   'False
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Namnet som är inskrivet finns redan, välj ett annat namn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   2160
         TabIndex        =   6
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Välj ett namn för speglingen:"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "steg 1"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Enheten är inaktiv, välj en annan enhet."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   2895
         Visible         =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "välj en enhet"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If DrvOK = True And NameOK = True And Text1.Text <> "" Then
    Form4.Starta
Else
    MsgBox lngVars(3)
End If
End Sub

Private Sub Command2_Click()
Form5.Hide
End Sub

Private Sub Drive1_Change()
On Error GoTo visafel
Form4.Dir1.Path = Drive1.Drive
Label5.Visible = False
DrvOK = True
Exit Sub
visafel:
Label5.Visible = True
DrvOK = False
End Sub

Private Sub Form_Load()
DrvOK = True
NameOK = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Cancel = True
'Form5.Hide
End Sub

Private Sub Text1_Change()
If IsNumeric(Text1.Text) = False Then
    If Form1.List1.ListCount = 0 Then
        NameOK = True
        Label3.Visible = False
        Label4.Visible = True
        Exit Sub
    End If
    For a = 0 To Form1.List1.ListCount - 1
        If UCase(Text1.Text) = UCase(Form1.List1.List(a)) Or Text1.Text = "" Then
            Label3.Visible = True
            Label4.Visible = False
            NameOK = False
            Exit Sub
        Else
            NameOK = True
            Label3.Visible = False
            Label4.Visible = True
        End If
    Next
Else
    Label3.Visible = True
    Label4.Visible = False
    NameOK = False
    Exit Sub
End If
Text1.Text = Replace(Text1.Text, "\", "-")
Text1.Text = Replace(Text1.Text, "/", "-")
Text1.SelStart = Len(Text1.Text)
End Sub
