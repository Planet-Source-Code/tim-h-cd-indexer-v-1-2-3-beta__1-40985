VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form8 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form8"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Sorstera listan alfabetiskt"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2520
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Starta"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stäng"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Make lists from:"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2895
      Begin VB.OptionButton Option2 
         Caption         =   "Only selected mirrors"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All mirrors"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "*.mp3"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dubbel klicka på ett filter för att ta bort det ifrån listan."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Spara som:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
mAdd = True
If Text1.Text <> "" Then
    For a = 0 To List1.ListCount - 1
        If UCase(List1.List(a)) = UCase(Text1.Text) Then
            mAdd = False
            Exit For
        End If
    Next
    If mAdd = True Then
        List1.AddItem Text1.Text
    End If
End If
End Sub

Private Sub Command2_Click()
CD1.DialogTitle = lngVars(23)
CD1.FileName = "noname"
CD1.Filter = "htm file|*.htm"
CD1.DefaultExt = "*.htm"
CD1.ShowSave
If CD1.FileName = "" Then Exit Sub
If CD1.FileName Like "*.htm" Then
    Text2.Text = CD1.FileName
Else
    Text2.Text = CD1.FileName & ".htm"
End If
End Sub

Private Sub Command3_Click()
Form8.Hide
End Sub

Private Sub Command4_Click()
List2.Clear
If CurFile = "" Then Exit Sub
If List1.ListCount <> 0 Then
    If Text2.Text <> "" Then
    Form8.MousePointer = 11
    DoEvents
    Form8.Enabled = False
        'If Text3.Text <> "" Then
              
            Open AppPath & "Listdata\listbase.dat" For Binary As #1
                sInput$ = String(LOF(1), Chr(0))
                Get #1, , sInput$
            Close #1
            
            bParts = Split(sInput$, "[LISTINPUT]", , vbTextCompare)
            
            Open AppPath & "Listdata\put.dat" For Binary As #1
                psInput$ = String(LOF(1), Chr(0))
                Get #1, , psInput$
            Close #1
            bPuts = Split(psInput$, "{|}", , vbTextCompare)
            Open Text2.Text For Output As #2
            Print #2, bParts(0)
            For a = 0 To UBound(bPuts)
                Select Case bPuts(a)
                    Case "file"
                        Print #2, "<b>File:</b>"
                    Case "drive"
                        Print #2, "<b>Drive:</b>"
                    Case "sizemb"
                        Print #2, "<b>Size:</b>"
                    Case Else
                        Print #2, bPuts(a)
                End Select
            Next
            Dim CurDrive As String
            Dim OkSearch As Boolean
            Dim aa As Long
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
                        For aba = 0 To List1.ListCount - 1
                            tebax = Split(enrad, "\", , vbTextCompare)
                            If OkSearch = True And UCase(enrad) Like UCase(List1.List(aba)) And Not enrad = "/cdi11/" Then
                                pFilen = Split(enrad, "\", , vbTextCompare)
                                Filen = pFilen(UBound(pFilen))
                                Storlek = Round(((tmpradd(1) / 1024) / 1000), 2)
                                sListTmp = ""
                                If Check1.Value = 1 Then
                                    For a = 0 To UBound(bPuts)
                                        Select Case bPuts(a)
                                            Case "file"
                                                sListTmp = sListTmp & "F{(/!\)}" & Filen & "{(/!\)}"
                                            Case "drive"
                                                sListTmp = sListTmp & "D{(/!\)}" & CurDrive & "{(/!\)}"
                                            Case "sizemb"
                                                sListTmp = sListTmp & "S{(/!\)}" & Storlek & "Mb" & "{(/!\)}"
                                            Case Else
                                                sListTmp = sListTmp & "T{(/!\)}" & bPuts(a) & "{(/!\)}"
                                        End Select
                                    Next
                                    List2.AddItem sListTmp
                                Else
                                    For a = 0 To UBound(bPuts)
                                        Select Case bPuts(a)
                                            Case "file"
                                                Print #2, Filen
                                            Case "drive"
                                                Print #2, CurDrive
                                            Case "sizemb"
                                                Print #2, Storlek & "Mb"
                                            Case Else
                                                Print #2, bPuts(a)
                                        End Select
                                    Next
                                End If
                            End If
                        Next
                    End If
                Loop Until EOF(1)
brake:
            If Check1.Value = 1 Then
                For ba = 0 To List2.ListCount - 1
                    AtmpListOut = Split(List2.List(ba), "{(/!\)}", , vbTextCompare)
                    For a = 0 To UBound(AtmpListOut)
                        Select Case AtmpListOut(a)
                            Case ""
                                'inget
                            Case "F"
                                Print #2, AtmpListOut(a + 1)
                            Case "S"
                                Print #2, AtmpListOut(a + 1)
                            Case "S"
                                Print #2, AtmpListOut(a + 1)
                            Case Else
                                Print #2, AtmpListOut(a + 1)
                        End Select
                        a = a + 1
                    Next
                Next
            End If
            Close #1
            Print #2, bParts(1)
            Close #2
            List2.Clear
            Form8.Enabled = True
            Form8.MousePointer = 0
    Else
        MsgBox lngVars(22)
    End If

Else
    MsgBox lngVars(21)
End If
End Sub

Private Sub Command5_Click()
MsgBox lngVars(24)
End Sub

Private Sub List1_DblClick()
If List1.ListIndex >= 0 Then
    List1.RemoveItem List1.ListIndex
End If
End Sub
