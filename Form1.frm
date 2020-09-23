VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1(CD Indexer)"
   ClientHeight    =   5565
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7740
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6840
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   3
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5280
      Width           =   7695
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4560
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
   Begin MSComctlLib.TreeView Tree1 
      Height          =   4815
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8493
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   3480
      MousePointer    =   9  'Size W E
      OLEDragMode     =   1  'Automatic
      Top             =   0
      Width           =   105
   End
   Begin VB.Menu Men0 
      Caption         =   "Men0(Arkiv)"
      Begin VB.Menu Men1 
         Caption         =   "Men1(Nytt)"
      End
      Begin VB.Menu Men2 
         Caption         =   "Men2(Öppna)"
         Shortcut        =   ^O
      End
      Begin VB.Menu Men3 
         Caption         =   "Men3(Stäng)"
      End
      Begin VB.Menu Men4 
         Caption         =   "Men4(Avsluta)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Men5 
      Caption         =   "Men5(Verktyg)"
      Begin VB.Menu Men6 
         Caption         =   "Men6(Skanna en enhet)"
      End
      Begin VB.Menu Men7 
         Caption         =   "Men7(Ta bort en spegling)"
      End
      Begin VB.Menu Men15 
         Caption         =   "Men15(Skapa listor)"
      End
      Begin VB.Menu Men14 
         Caption         =   "Men14(konvertera)"
      End
      Begin VB.Menu Men16 
         Caption         =   "Men16(Databas info)"
      End
   End
   Begin VB.Menu Men13 
      Caption         =   "Men13(Sök)"
   End
   Begin VB.Menu Men8 
      Caption         =   "Men8(inställningar)"
      Begin VB.Menu Men9 
         Caption         =   "Men9(Alternativ)"
      End
   End
   Begin VB.Menu Men10 
      Caption         =   "Men10(Hjälp)"
      Begin VB.Menu Men11 
         Caption         =   "Men11(Hjälp)"
      End
      Begin VB.Menu Men12 
         Caption         =   "Men12(Om)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CD Indexer 1.2.3
' Open Source Code
' All code and files is COPYRIGHT (c) 2001-2002
' and belongs to Tim Hegyi
'
' This project has a open source code because
' I don't have any time too work on it anymore.
' Too use this source code you need to accept some stuff.
' 1. I am the Coder of CD Indexer My name is Tim Hegyi
' 2. My name MUST always be witen in the About Box
'    so everyone can see it. Its a hard world out there.
' 3. You can add or delete code to/from the project, but
'    not copy code and use it in other project.
' 4. The name of the program/project must always bee CD Indexer
' 5. You can change the version, but you must wite that its based
'    on CD Indexer 1.2.3
' 6. My e-mail adress is tim_hegyi@hotmail.com and must always bee
'    witen in the about box.
' 7. There must always be a about box
' 8. Please respect this rules.
' 9. Every public release you make of the program must be e-mailed to me.
'    ONLY e-mail me the source code, project files
'10. www.cdindexer.tk is the official CD Indexer adress and it is
'    the only adress you may write in CD Indexer.
'11. If you make anychanges in the source code, it belongs to me, Tim Hegyi
'12. If you can not accept this rules, delete the source code now.
'13. Any questions? e-mail me at: tim_hegyi@hotmail.com
'
'----------------------------------------
'Comments:
' Sorry for the fucked up source code without any comments.

Private Sub Form_Load()
'lägger sökvägen till programet i AppPath
'frmSpl.Show
If Len(App.Path) = 3 Then
    AppPath = App.Path
Else
    AppPath = App.Path & "\"
End If
'On Error Resume Next
resSucc = GetInternetFile(Inet1, "http://www10.brinkster.com/timpa/cver.txt", AppPath)
If resSucc = True Then
    Open AppPath & "cver.txt" For Input As #1
        Line Input #1, tLine
        If IsNumeric(tLine) = True Then
            If CLng(Form2.Label6.Caption) < CLng(tLine) Then
                Do
                    Line Input #1, tLine
                    txtShowNew = txtShowNew & tLine & vbCrLf
                Loop Until EOF(1)
                MsgBox txtShowNew, vbOKOnly, "New version of CD Indexer found!"
            End If
        End If
    Close #1
    Kill AppPath & "cver.txt"
End If
Language "language\" & GetSetting("cd indexer", "settings", "lng", "English") & ".lng", True 'Laddar språk
Form1.Tree1.ImageList = Form1.ImageList1
AppName = Form1.Caption
Form1.List1.ToolTipText = lngVars(14)
Form1.Text1.Text = lngVars(17) & lngVars(18)
'frmSpl.Hide
End Sub
Private Sub Form_Resize()
'Sätter gränser för hur mycke man ska kunna ändra storleken på Form1
'ändrar även tree1, list1 och image1 storlekarna
If Form1.WindowState <> 1 Then
    If Form1.Width < Tree1.Left + 250 Then Form1.Width = Tree1.Left + 250
    Tree1.Left = Image1.Left + Image1.Width
    Image1.Height = Form1.Height - 285
    Tree1.Height = Form1.ScaleHeight - 285
    List1.Height = Form1.ScaleHeight - 285
    Tree1.Width = Form1.ScaleWidth - Tree1.Left
    Text1.Top = Tree1.Height
    Text1.Width = Form1.ScaleWidth
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Tar hand om splitbaren, Image1 är splitbaren
If Button = 1 Then
    Mover = True
    If Image1.Left < 250 Then Image1.Left = 350
    If Image1.Left > Form1.Width - Image1.Width - 250 Then Image1.Left = Form1.Width - Image1.Width - 350
    Image1.Left = Image1.Left - OX + X
    Image1.BorderStyle = 1
    List1.Width = Image1.Left
    Tree1.Left = Image1.Left + Image1.Width
    Tree1.Width = Form1.ScaleWidth - Tree1.Left
Else
    OX = X
    Image1.BorderStyle = 0
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0 'släpper man mus knappen så blir borderstylen 0
Mover = True
End Sub
Private Sub List1_DblClick()
ExtrTree List1.List(List1.ListIndex), ""
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu Men5
End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Mover = True Then
    Image1.Left = X + List1.Left
    
    
    Mover = True
    If Image1.Left < 250 Then Image1.Left = 350
    If Image1.Left > Form1.Width - Image1.Width - 250 Then Image1.Left = Form1.Width - Image1.Width - 350
    'Image1.Left = Image1.Left - OX + X
    Image1.BorderStyle = 1
    List1.Width = Image1.Left
    Tree1.Left = Image1.Left + Image1.Width
    Tree1.Width = Form1.ScaleWidth - Tree1.Left
Else
    Mover = False
End If
End Sub

Private Sub Men1_Click()
CD1.DialogTitle = lngVars(5)
CD1.FileName = "noname"
CD1.Filter = "CD Indexer Database|*.bup"
CD1.DefaultExt = "*.bup"
CD1.ShowSave
If CD1.FileName = "" Then Exit Sub
If CD1.FileTitle = "" Then Exit Sub
If CD1.FileName Like "*.bup" Then
    CurFile = CD1.FileName
Else
    CurFile = CD1.FileName & ".bup"
End If
Open CurFile For Output As #1
    Print #1, "/cdi11/"
    Print #1, "</end of file/>"
Close #1
Form1.Caption = AppName & " (" & CurFile & ")"
List1.Clear
End Sub

Private Sub Men11_Click()
On Error GoTo errhand
Shell "notepad.exe " & AppPath & "help.txt", vbNormalFocus
Exit Sub
errhand:
MsgBox lngVars(15)
End Sub

Private Sub Men12_Click()
Form2.Show vbModal
End Sub

Private Sub Men13_Click()
If CurFile <> "" Then
    Form3.Show
Else
    MsgBox lngVars(4)
End If
End Sub

Private Sub Men14_Click()
Form7.Show vbModal
End Sub

Private Sub Men15_Click()
If CurFile <> "" Then
    Form8.Show
Else
    MsgBox lngVars(20)
End If
End Sub

Private Sub Men16_Click()
If CurFile <> "" Then
    Form9.Show vbModal
Else
    MsgBox lngVars(35)
End If
End Sub

Private Sub Men2_Click()
CD1.DialogTitle = lngVars(0)
CD1.Filter = "CD Indexer Database|*.bup"
CD1.DefaultExt = "*.bup"
CD1.ShowOpen
If CD1.FileName = "" Then Exit Sub
Tree1.Nodes.Clear
List1.Clear
CurFile = CD1.FileName
tid = Timer
On Error GoTo errhandlefile
Open CurFile For Input As #1
    Line Input #1, trad
    If trad <> "/cdi11/" Then
        Close #1
        CurFile = ""
        MsgBox lngVars(1)
        Exit Sub
    End If
    Form1.MousePointer = 11
    Do
        Line Input #1, trad
        Select Case trad
            Case "</end of file/>"
                GoTo sluta
        End Select
        If Right(trad, 1) = ">" And Left(trad, 1) = "<" Then
            'MsgBox "OK"
            List1.AddItem Mid(trad, 2, Len(trad) - 2)
        End If
    Loop Until EOF(1)
sluta:
Close #1
Form1.Caption = AppName & " (" & CurFile & ")"
Form1.MousePointer = 0
Text1.Text = lngVars(17) & List1.ListCount & lngVars(16) & Round(Timer - tid, 5) & " s"
Exit Sub
errhandlefile:
End Sub

Private Sub Men3_Click()
CurFile = ""
List1.Clear
Tree1.Nodes.Clear
Form1.Caption = AppName
End Sub

Private Sub Men4_Click()
End
End Sub

Private Sub Men6_Click()
If CurFile <> "" Then
    Form5.Show vbModal
Else
    MsgBox lngVars(6)
End If
End Sub

Private Sub Men7_Click()
Dim Vrite As Boolean
If CurFile = "" Then Exit Sub
If List1.ListIndex = -1 Then
    MsgBox lngVars(7)
    Exit Sub
End If
notthis = List1.List(List1.ListIndex)
If MsgBox(lngVars(8) & " " & notthis & "?", vbYesNo) = vbYes Then
    tm = Timer
    Form1.MousePointer = 0
    'tmpflname = Left(tm, Len(tm) - 3) & Right(tm, 2)
    notthis = List1.List(List1.ListIndex)
    FileCopy CurFile, AppPath & tm
    Open CurFile For Output As #1
        Open AppPath & tm For Input As #2
            Do
                Line Input #2, ttmpp
                If ttmpp Like "<" & notthis & ">" Then
                    Vrite = False
                ElseIf ttmpp Like "<*>" Then
                    Vrite = True
                    Print #1, ttmpp
                ElseIf Vrite = True Or ttmpp = "</end of file/>" Or ttmpp = "/cdi11/" Then
                    Print #1, ttmpp
                End If
            Loop Until EOF(2)
        Close #2
    Close #1
    List1.RemoveItem (List1.ListIndex)
    Kill AppPath & tm
    Form1.MousePointer = 0
End If
End Sub

Private Sub Men9_Click()
Form6.Show vbModal
End Sub

Private Sub Tree1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Mover = True Then
    Image1.Left = X + Tree1.Left
    
    
    Mover = True
    If Image1.Left < 250 Then Image1.Left = 350
    If Image1.Left > Form1.Width - Image1.Width - 250 Then Image1.Left = Form1.Width - Image1.Width - 350
    'Image1.Left = Image1.Left - OX + X
    Image1.BorderStyle = 1
    List1.Width = Image1.Left
    Tree1.Left = Image1.Left + Image1.Width
    Tree1.Width = Form1.ScaleWidth - Tree1.Left
Else
    Mover = False
End If
End Sub
