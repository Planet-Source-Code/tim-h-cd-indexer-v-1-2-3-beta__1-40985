Attribute VB_Name = "Module3"
Public Function GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String) As Boolean
    ' Written by: Blake Pell
    On Error Resume Next
    Kill AppPath & "cver.txt"
    Dim myData() As Byte
    If Inet1.StillExecuting = True Then Exit Function
    myData() = Inet1.OpenURL(myURL, icByteArray)
    
    For X = Len(myURL) To 1 Step -1
        If Left$(Right$(myURL, X), 1) = "/" Then RealFile$ = Right$(myURL, X - 1)
    Next X
    myFile$ = DestDIR + "\" + RealFile$
    
    Open myFile$ For Binary Access Write As #11
        Put #11, , myData()
    Close #11
    
    'Open myFile$ For Output As #25
    '    Print #25, myData()
    'Close #25
    
    If FileLen(myFile$) = 0 Then
        GetInternetFile = False
        Kill myFile$
    Else
        GetInternetFile = True
    End If
    Exit Function

' error handler
errha:
    GetInternetFile = False
    Resume 105
105 End Function
