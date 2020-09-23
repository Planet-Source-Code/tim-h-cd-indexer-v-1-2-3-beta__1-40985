VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Databas info"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Hämta info"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stäng"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form9.Hide
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Dim filer As Long
gb = 0
speglar = 0
filer = 0
Text1.Text = lngVars(27) & CurFile & vbCrLf & lngVars(28) & vbCrLf
DoEvents
Open CurFile For Input As #1
    Do
        Line Input #1, tmprad
        If tmprad = "/cdi11/" Then
            Text1.Text = Text1.Text & lngVars(29) & vbCrLf
            DoEvents
        ElseIf tmprad = "</end of file/>" Then
            Text1.Text = Text1.Text & lngVars(30) & vbCrLf
            DoEvents
        ElseIf tmprad Like "<*>" Then
            speglar = speglar + 1
            Text1.Text = Text1.Text & "."
            DoEvents
        Else
            filer = filer + 1
            strBack = Split(tmprad, "/", , vbTextCompare)
            gb = gb + (strBack(1) / 1024000000)
        End If
    Loop Until EOF(1)
Close #1
Text1.Text = Text1.Text & vbCrLf & "_________________" & vbCrLf & lngVars(31) & vbCrLf & lngVars(32) & speglar & vbCrLf & lngVars(33) & filer & vbCrLf & lngVars(34) & Round(gb, 1)
End Sub

