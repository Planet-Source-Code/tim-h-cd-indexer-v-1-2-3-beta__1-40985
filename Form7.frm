VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Konvertera"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stäng"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Starta"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4695
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   285
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VARNING: Om du väjer en fil som redan finns kommer den att skrivas över!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1320
      TabIndex        =   8
      Top             =   1560
      Width           =   2295
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CD1.DialogTitle = lngVars(9)
CD1.Filter = "Drive Mirror file|*.bup"
CD1.ShowOpen
If CD1.FileName <> "" Then
    Text1.Text = CD1.FileName
End If
End Sub

Private Sub Command2_Click()
CD1.DialogTitle = lngVars(10)
CD1.Filter = "CD Indexer file|*.bup"
CD1.FileName = "noname"
CD1.ShowSave
If CD1.FileName <> "" Then
    Text2.Text = CD1.FileName
End If
End Sub

Private Sub Command3_Click()
MsgBox lngVars(26)
dm2file = False
Open Text1.Text For Input As #2
    Line Input #2, tradd
Close #2
If tradd = "/cdi11/" Then

    MsgBox lngVars(12)  '/FIXA!
    Exit Sub
ElseIf tradd = "/dm2/" Then
    dm2file = True
End If
If UCase(Text1.Text) Like UCase("*.bup") And UCase(Text2.Text) Like UCase("*.bup") Then
    If Text1.Text <> "" And Text2.Text <> "" And Text1.Text <> Text2.Text Then
        Form7.MousePointer = 11
        'Form7.Enabled = False
        DoEvents
        Open Text2.Text For Output As #1
            Print #1, "/cdi11/"
            Open Text1.Text For Input As #2
                Do
                    Line Input #2, tradd
                    If tradd = "</end of file/>" Then
                        Print #1, tradd
                    ElseIf tradd Like "<*>" Then
                        If dm2file = False Then
                            prithis = Left(tradd, Len(tradd) - 9) & ">"
                            Print #1, prithis
                        Else
                            Print #1, tradd
                        End If
                    ElseIf tradd = "/dm2/" Then
                        'inget här!
                    Else
                        Print #1, tradd & "/0"
                    End If
                Loop Until EOF(2)
            Close #2
        Close #1
        Form7.MousePointer = 0
        'Form7.Enabled = True
        Form7.Hide
    Else
        MsgBox lngVars(11)
    End If
Else
    MsgBox lngVars(13)
End If
End Sub

Private Sub Command4_Click()
Form7.Hide
End Sub

