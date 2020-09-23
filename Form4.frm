VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   4320
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Hidden          =   -1  'True
      Left            =   2640
      System          =   -1  'True
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   960
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kopierar katalog struktur"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   4950
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Form4.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Public Function DirExplorer(StartDir As String) As Boolean
Dim NoOfDirs, DirCount, OldDir, NoOfFiles

Dir1.Path = StartDir
NoOfDirs = Dir1.ListCount - 1

For DirCount = 0 To NoOfDirs
    OldDir = Dir1.Path
    StopExp = DirExplorer(Dir1.List(DirCount))
Next DirCount

File1.Path = Dir1.Path
NoOfFiles = File1.ListCount - 1

If Len(Dir1.Path) = 3 Then
    dirpath = Dir1.Path
Else
    dirpath = Dir1.Path & "\"
End If

For a = 0 To NoOfFiles
    entry = Right(dirpath, Len(dirpath) - 1) & File1.List(a) & "/" & FileLen(dirpath & File1.List(a))
    Print #20, entry
Next a
Dir1.Path = Dir1.List(-2)
End Function

Public Sub Starta()
List1.Clear
Dir1.Path = Left(Form5.Drive1.Drive, 2) & "\"
Dir1.Refresh
If Form5.Text2.Text <> "" Then
    File1.Pattern = Form5.Text2.Text
End If
Form1.Enabled = False
Form5.Hide
Form4.Show
Form4.MousePointer = 11
DoEvents
Open AppPath & "tfile" For Output As #20
back = DirExplorer(Dir1.Path)
Close #20
Open AppPath & "tfile" For Input As #20
tm = Timer
'tmpflname = Left(tm, Len(tm) - 3) & Right(tm, 2)
FileCopy CurFile, AppPath & tm
Open CurFile For Output As #1
    Print #1, "/cdi11/"
    Print #1, "<" & Form5.Text1.Text & ">"
    'For a = 0 To List1.ListCount - 1
    '    Print #1, List1.List(a)
    'Next
    Do
        Line Input #20, tfilestr
        Print #1, tfilestr
    Loop Until EOF(20)
    Open AppPath & tm For Input As #2
        Line Input #2, tmpa
        Do
            Line Input #2, tmpa
            Print #1, tmpa
        Loop Until EOF(2)
    Close #2
Close #1
Close #20
Form1.List1.AddItem Form5.Text1.Text
List1.Clear
Kill AppPath & "tfile"
Kill AppPath & tm
Form1.Enabled = True
Form4.MousePointer = 0
Form4.Hide
End Sub

