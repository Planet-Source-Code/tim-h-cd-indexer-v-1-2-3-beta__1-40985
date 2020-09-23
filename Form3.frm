VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6270
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   5760
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   7095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sök"
      Height          =   930
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton Command3 
         Caption         =   "avbryt"
         Height          =   375
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         Visible         =   0   'False
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sök ändast i markerade speglingar"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   600
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Alla speglingar"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sök"
         Default         =   -1  'True
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   5655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sök i:"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   7095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Antal hittade filer:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If CurFile = "" Then Exit Sub
List1.Clear
Command3.Visible = True
Command2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Dim CurDrive As String
Dim OkSearch As Boolean
Dim aa As Long
If Text1.Text Like "[*]*[*]" Then
Else
    Text1.Text = "*" & Text1.Text & "*"
End If
Form3.MousePointer = 11
Open CurFile For Input As #1
    Do
        Line Input #1, larad
        tmpradd = Split(larad, "/", , vbTextCompare)
        enrad = tmpradd(0)
        If enrad Like "<*>" Then
            If BrakeSearch = True Then GoTo brake
            DoEvents
            CurDrive = Mid(enrad, 2, Len(enrad) - 2)
            For aa = 0 To Form1.List1.ListCount - 1
                If Option1.Value = False Then
                    If Form1.List1.Selected(aa) = True And CurDrive = Form1.List1.List(aa) Then
                        OkSearch = True
                        GoTo stepout
                    Else
                        OkSearch = False
                    End If
                Else
                    OkSearch = True
                End If
            Next
stepout:
        Else
            If OkSearch = True And UCase(enrad) Like UCase(Form3.Text1.Text) And Not enrad = "/cdi11/" Then List1.AddItem CurDrive & enrad
        End If
    Loop Until EOF(1)
brake:
Close #1
Option1.Enabled = True
Option2.Enabled = True
Command2.Enabled = True
Command3.Visible = False
Label3.Caption = List1.ListCount
Form3.MousePointer = 0
End Sub

Private Sub Command2_Click()
Form3.Hide
End Sub

Private Sub Command3_Click()
BrakeSearch = True
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.MousePointer = 0
End Sub

Private Sub Form_GotFocus()
Text1.SetFocus
End Sub

Private Sub Form_Resize()
Frame1.Width = Form3.Width - Frame1.Left - 240
List1.Width = Form3.Width - List1.Left - 240
Form3.Height = 6675
Command2.Left = (Form3.Width / 2) - (Command2.Width / 2)
End Sub

Private Sub List1_DblClick()
If List1.ListIndex < 0 Then Exit Sub
odrv = Split(List1.List(List1.ListIndex), "\", , vbTextCompare)
ExtrTree CStr(odrv(0)), CStr(List1.List(List1.ListIndex))
End Sub
