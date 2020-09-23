VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inställningar"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Sök efter ny ver. av CD Indexer"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   3840
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Enabled         =   0   'False
         Height          =   975
         Left            =   3120
         Picture         =   "Form6.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Denna function kräver att du har en internet anslutning. Om du använder modem, koppla upp innan du uppdaterar."
         ForeColor       =   &H00808080&
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1853
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Språk"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Välj ett språk ifrån listan som passar dig bäst:"
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Hide
If List1.ListIndex <> -1 Then
    Language "language\" & List1.List(List1.ListIndex) & ".lng", False
    SaveSetting "cd indexer", "settings", "lng", List1.List(List1.ListIndex)
End If
Form1.List1.ToolTipText = lngVars(14)
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Command1.Enabled = False
Text1.Text = lngVars(36) & vbCrLf
    TransferSuccess = GetInternetFile(Inet1, "http://web.tillenius.net/cdi/uver.dat", AppPath)
    If TransferSuccess = False Then
        Text1.Text = Text1.Text & lngVars(37) & vbCrLf
        Command2.Enabled = True
        Command1.Enabled = True
        Exit Sub
    Else
        
        Open AppPath & "uver.dat" For Input As #1
            Line Input #1, newVer
        Close #1
        Kill AppPath & "uver.dat"
        backarr = Split(newVer, " ", , vbTextCompare)
        If CLng(backarr(0)) > CLng(Form2.Label6.Caption) Then
            'fixa ny ver
            Text1.Text = Text1.Text & lngVars(38) & vbCrLf
            Text1.Text = Text1.Text & lngVars(39) & backarr(1) & vbCrLf
            TransferSuccess = GetInternetFile(Inet1, "http://web.tillenius.net/cdi/cdi_ud.exe", AppPath)
            If TransferSuccess = True Then
                Text1.Text = Text1.Text & lngVars(40) & vbCrLf
                ltb = Shell(AppPath & "cdi_ud.exe", vbNormalFocus)
                End
            Else
                Text1.Text = Text1.Text & lngVars(41) & vbCrLf
                Command2.Enabled = True
                Command1.Enabled = True
                Exit Sub
            End If
        Else
            Text1.Text = Text1.Text & lngVars(42) & vbCrLf
            Command2.Enabled = True
            Command1.Enabled = True
        End If
    End If
Command2.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Form_Load()
Open AppPath & "Language\alllngs.dat" For Input As #1
    Do
        Line Input #1, traad
        List1.AddItem traad
    Loop Until EOF(1)
Close #1
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
Text1.SetFocus
End Sub
