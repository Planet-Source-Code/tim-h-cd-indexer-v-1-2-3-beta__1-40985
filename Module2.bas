Attribute VB_Name = "Module2"
Public Sub ExtrTree(GetDrive As String, ByVal FollowUp As String)
tid = Timer
Dim tback As Node
Dim tback2 As Node
Form1.Tree1.Nodes.Clear
Set tback2 = Form1.Tree1.Nodes.Add(, , GetDrive, GetDrive, 1, 1)
Form1.MousePointer = 11
Open CurFile For Input As #1
    Do
        Line Input #1, trad
        If trad = "<" & GetDrive & ">" Then

            Do
                Line Input #1, tttrad
                If tttrad = "</end of file/>" Then Exit Do
                attrad = Split(tttrad, "/", , vbTextCompare)
                ttrad = attrad(0)
                If ttrad Like "<*>" Then Exit Do
                tarray = Split(ttrad, "\", , vbTextCompare)
                If UBound(tarray) = 1 Then
                    Set tback = Form1.Tree1.Nodes.Add(GetDrive, tvwChild, mroot & "\" & tarray(1), tarray(1), 3, 3)
                Else
                    On Error Resume Next
                    utarray = UBound(tarray)
                    mkay = GetDrive
                    For a = 1 To utarray - 1
                        omkay = mkay
                        mkay = mkay & "\" & tarray(a)
                        Set tback = Form1.Tree1.Nodes.Add(omkay, tvwChild, mkay, tarray(a), 2, 2)
                    Next
                    Set tback = Form1.Tree1.Nodes.Add(mkay, tvwChild, mkay & "\" & tarray(utarray), tarray(utarray), 3, 3)
               End If
            Loop Until EOF(1)
            Exit Do
        End If
    Loop Until EOF(1)
    tback2.Child.EnsureVisible
Close #1
If FollowUp <> "" Then
    Form1.Tree1.Nodes(FollowUp).Selected = True
    Form1.Tree1.Nodes(FollowUp).EnsureVisible
    Form1.SetFocus
End If
Form1.MousePointer = 0
Form1.Text1.Text = lngVars(17) & lngVars(19) & Round(Timer - tid, 5) & " s"
End Sub

Public Sub oldExtrTree(ByVal GetDrive As String, ByVal FollowUp As String)
'Snabb men FEL!!!!!!
'/////////////////////////////////////////////////////////////////////
On Error GoTo ExtrTree_Err
  Dim bFound    As Boolean
  Dim bFoundKey As Boolean
  Dim sInput    As String
  Dim sArray()  As String
  Dim sTemp()   As String
  Dim sLast()   As String
  Dim lCount1   As Long
  Dim lCount2   As Long
  Dim BottomKey As String
  Dim ChildKey  As String
  Dim nTemp     As Node
  Dim nNodes    As Nodes
  Dim UBsArray  As Long
  Dim UBsTemp   As Long

Screen.MousePointer = vbHourglass

Set nNodes = Form1.Tree1.Nodes
nNodes.Clear

Open CurFile For Binary As #1
    sInput$ = String(LOF(1), Chr(0))
    Get #1, , sInput$
Close #1

sArray() = Split(sInput, vbCrLf)

Set nTemp = nNodes.Add(, , GetDrive, GetDrive, 1, 1)
nTemp.Expanded = True

bFound = False
'UBsArray = UBound(sArray)
For lCount1 = 0 To UBound(sArray)
    If bFound Then
        If sArray(lCount1) Like "<*>" Then
            Exit For
        End If
        sTemp = Split(Mid(sArray(lCount1), 2), "\", , vbTextCompare)
        bFoundKey = False
        For lCount2 = 0 To UBound(sTemp) - 1
            If BottomKey <> "" And Not bFoundKey Then
                If sLast(lCount2) = sTemp(lCount2) Then
                    BottomKey = BottomKey & "\" & sTemp(lCount2)
                Else
                    bFoundKey = True
                End If
            ElseIf bFoundKey Then
                BottomKey = ChildKey
            Else
                BottomKey = GetDrive
                bFoundKey = True
            End If
            If bFoundKey Then
                ChildKey = BottomKey & "\" & sTemp(lCount2)
                Set nTemp = nNodes.Add(BottomKey, tvwChild, ChildKey, sTemp(lCount2), 2, 2)
            End If
        Next lCount2
        If bFoundKey Then
            Set nTemp = nNodes.Add(ChildKey, tvwChild, ChildKey & "\" & sTemp(lCount2), sTemp(lCount2), 3, 3)
        Else
            Set nTemp = nNodes.Add(BottomKey, tvwChild, BottomKey & "\" & sTemp(lCount2), sTemp(lCount2), 3, 3)
        End If
        sLast = sTemp
        BottomKey = GetDrive
    End If
    If sArray(lCount1) = "<" & GetDrive & ">" Then
        bFound = True
    End If
Next lCount1

If Len(FollowUp) > 0 Then
    nNodes(FollowUp).EnsureVisible
    nNodes(FollowUp).Selected = True
End If

Screen.MousePointer = vbNormal
Exit Sub
ExtrTree_Err:
 Screen.MousePointer = vbNormal
 'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
