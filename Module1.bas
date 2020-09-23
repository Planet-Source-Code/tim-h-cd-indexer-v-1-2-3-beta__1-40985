Attribute VB_Name = "Module1"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public DrvOK As Boolean
Public NameOK As Boolean
Public BrakeSearch As Boolean
Public Mover As Boolean
Public lngVars(0 To 1000) As String
Public AppName As String
Public CurFile As String
Public tHead As Head
Public OX As Long
Public OY As Long
Public AppPath As String
Public Enum Head
    tContr
    tForms
    tVars
End Enum
Public Sub Language(lngFile As String, Ftime As Boolean)
'läser ur språkfilen
'öppna forms
Form2.Show
Form2.Enabled = False
Form3.Show
Form3.Enabled = False
Form4.Show
Form4.Enabled = False
Form5.Show
Form5.Enabled = False
Form6.Show
Form6.Enabled = False
Form7.Show
Form7.Enabled = False
Form8.Show
Form8.Enabled = False
Form9.Show
Form9.Enabled = False
Dim Rad As String
Open AppPath & lngFile For Input As #1
    Do
        Line Input #1, Rad
        If Rad = "[Controls]" Then
            tHead = tContr
        ElseIf Rad = "[Forms]" Then
            tHead = tForms
        ElseIf Rad = "[Vars]" Then
            tHead = tVars
        Else
            Splitter Rad, tHead
        End If
    Loop Until EOF(1)
Close #1
'stänger formsen
Form2.Enabled = True
Form2.Hide
Form3.Enabled = True
Form3.Hide
Form4.Enabled = True
Form4.Hide
Form5.Enabled = True
Form5.Hide
Form6.Enabled = True
Form6.Hide
Form7.Enabled = True
Form7.Hide
Form8.Enabled = True
Form8.Hide
Form9.Enabled = True
Form9.Hide
If Ftime = False Then
    Form1.SetFocus
End If
End Sub
Public Sub Splitter(radLng As String, tType As Head)
Dim Flag As Boolean
Dim tObj As String
Dim word As String
a = 0
rlen = Len(radLng)
tmps = Split(radLng, "=", , vbTextCompare)
word = tmps(1)
For a = 2 To UBound(tmps)
    word = word & "=" & tmps(a)
Next
If tType = tContr Then
    PasteWords CStr(tmps(0)), word
ElseIf tType = tForms Then
    fCaption CSng(Right(tmps(0), Len(tmps(0)) - 4)), word
ElseIf tType = tVars Then
    FixVar CLng(tmps(0)), word
End If
End Sub
Public Sub PasteWords(allObjs As String, word As String)
'On Error Resume Next
Dim tmpForm As Control
Dim tal As Single
TallObjs = Split(allObjs, ".", , vbTextCompare)
'MsgBox TallObjs(1)
tal = CSng(Right(TallObjs(0), Len(TallObjs(0)) - 4))
'MsgBox tal
Set tmpForm = GetControl(tal, CStr(TallObjs(1)))
tmpForm.Caption = word
End Sub
Public Function GetControl(myForm As Single, Ctrl As String) As Control
On Error GoTo errGC
Set GetControl = Forms(myForm - 1).Controls(Ctrl)
Exit Function
errGC:
Select Case Err
    Case 9
        MsgBox "Detected error while loading language: " & Error & " = " & Err & Chr(10) & "Form no. " & myForm & " not found." & Chr(10) & "Exiting program"
    Case 730
        MsgBox "Detected error while loading language: " & Error & " = " & Err & Chr(10) & "Exiting program"
    Case Else
        MsgBox "Detected error while loading language: " & Error & " = " & Err & Chr(10) & "Exiting program"
    
End Select
End
End Function
Public Sub fCaption(frm As Single, word As String)
Dim tmpForm As Form
Set tmpForm = Forms(frm - 1)
tmpForm.Caption = word
End Sub
Public Sub FixVar(place As Long, word As String)
lngVars(place) = word
End Sub
