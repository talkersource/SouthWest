Attribute VB_Name = "Viewer"
Option Explicit
Global lines() As String

Public Sub Plot_Viewer(atext As String, num As Integer)
'This little gem will parse the string for color codes.
'The code is extremly complicated but you shouldnt have
'to mess with it. I really should have commented it
'more heavily but oh well...
Dim count As Integer, col As String, Fhold As Long, text As String
Dim code As Boolean, codefound As Boolean, count2 As Integer
text = atext
mainForm.p(num).ToolTipText = Format$(viewer_strip(text).vDate, "mmmm d, yyyy @ h:nnam/pm")
If Len(mainForm.p(num).ToolTipText) = 1 Then
    mainForm.p(num).ToolTipText = ""
    End If
text = viewer_strip(text).vRest
mainForm.p(num).Cls
Fhold = mainForm.p(num).ForeColor
For count = 1 To Len(text)
    code = False
    col = Mid$(text, count, 1)
    If Not count + 2 > Len(text) Then
        If col = "~" Then
            If count > 1 Then
                If Mid$(text, count - 1, 1) = "/" Then
                    code = False
                    Else
                        code = True
                        End If
                Else
                    code = True
                    End If
            If code Then
                code = False
                col = Mid$(text, count + 1, 2)
                For count2 = 0 To 20
                    If col = colorShorts(count2) Then
                        If count2 >= 0 And count2 < 5 Then
                            Select Case count2
                                Case 0
                                    mainForm.p(num).ForeColor = Fhold
                                    mainForm.p(num).FontBold = False
                                    mainForm.p(num).FontUnderline = False
                                    mainForm.p(num).FontItalic = False
                                Case 1
                                    mainForm.p(num).FontBold = True
                                Case 2
                                    mainForm.p(num).FontUnderline = True
                                Case 3
                                    mainForm.p(num).FontItalic = True
                                End Select
                            ElseIf count2 < 13 Then
                                mainForm.p(num).ForeColor = colorHex(count2)
                                End If
                        text = Left$(text, count - 1) & Right$(text, Len(text) - count - 2)
                        count = count - 1
                        code = True
                        Exit For
                        End If
                    Next count2
                If Not code Then
                    col = Mid$(text, count, 1)
                    End If
                End If
            Else
                code = False
                End If
        End If
    If Not code Then
        mainForm.p(num).Print col;
        End If
    Next count
mainForm.p(num).FontBold = False: mainForm.p(num).FontItalic = False
mainForm.p(num).FontUnderline = False: mainForm.p(num).ForeColor = Fhold
End Sub

Public Sub load_text(text As String)
Dim count As Integer, CRLFcount As Integer, smaller As Integer
For count = 0 To 12
    mainForm.p(count).ToolTipText = ""
    mainForm.p(count).Cls
    Next count
text = text & CRLF
For count = 1 To Len(text) - 1
    If Mid$(text, count, 2) = CRLF Then
        CRLFcount = CRLFcount + 1
        End If
    Next count
ReDim lines(CRLFcount)
smaller = CRLFcount - 1
If smaller < 3075 Then
    smaller = 3075
    End If
If smaller > 12 Then
    smaller = 12
    End If
For count = 0 To CRLFcount - 1
    lines(count) = Left$(text, InStr(text, CRLF))
    text = Right$(text, Len(text) - (InStr(text, CRLF) + 1))
    If count <= smaller Then
        Plot_Viewer lines(count), count
        End If
    Next count
If (UBound(lines) - 13) > 0 Then
    mainForm.vscroll.Max = UBound(lines) - 12
    Else
        mainForm.vscroll.Max = 0
        End If
mainForm.vscroll.value = 0
End Sub

Public Function viewer_strip(text As String) As VI_OBJECT
Dim tospace As Integer, length As Integer, out As String
length = Len(text)
tospace = InStr(text, " ")
If tospace + 1 >= length Or tospace = 0 Then
    viewer_strip.vRest = text
    Exit Function
    End If
tospace = InStr(tospace + 1, text, " ")
If tospace = 0 Or tospace >= length Then
    viewer_strip.vRest = text
    Exit Function
    End If
viewer_strip.vDate = Left$(text, tospace - 1)
viewer_strip.vRest = Right$(text, length - tospace - 1)
End Function

Public Function treeFormat(ByVal UserName As String, shave As Integer) As String
'shaves last chars off
If Len(UserName) > shave Then
    UserName = Left$(UserName, Len(UserName) - shave)
    Else
        UserName = ""
        End If
UserName = userCap(UserName)
treeFormat = UserName
End Function

Public Function loadViewer()
mainForm.tree.Nodes.Clear
Dim Tnode As node, count As Integer, hold As String
Dim usrCount As Integer, usrArray() As String
mainForm.tree.Nodes.add , , "GRAPHER", "AutoGraph"
mainForm.tree.Nodes.add 1, tvwChild, "USER ACTIVITY", "User Activity"
mainForm.tree.Nodes.add 2, tvwNext, "USER LOGINS", "User Logins"
mainForm.tree.Nodes.add 3, tvwNext, "HTTP CONNECTIONS", "HTTP Connects"
mainForm.tree.Nodes.add 4, tvwNext, "HTTP REQUESTS", "HTTP Requests"
count = 5
count = count + 1
mainForm.tree.Nodes.add , , "USER DATA", "User Data"
count = count + 1
hold = treeFormat(Dir$(App.Path & "\Users\*.D"), 2)
ReDim usrArray(0)
Do While Not hold = ""
    usrCount = usrCount + 1
    hold = Dir$
    Loop
ReDim usrArray(usrCount)
hold = treeFormat(Dir$(App.Path & "\Users\*.D"), 2)
For usrCount = LBound(usrArray) To UBound(usrArray) - 1
    usrArray(usrCount) = hold
    hold = treeFormat(Dir$, 2)
    Next usrCount
ShellSort usrArray()
For usrCount = LBound(usrArray) + 1 To UBound(usrArray)
    If usrCount = 1 And Not usrArray(usrCount) = vbNullString Then
        mainForm.tree.Nodes.add count - 1, tvwChild, "UD " & UCase$(usrArray(usrCount)), usrArray(usrCount)
        Else
            count = count + 1
            mainForm.tree.Nodes.add count - 1, tvwNext, "UD " & UCase$(usrArray(usrCount)), usrArray(usrCount)
            End If
    Next usrCount
count = count + 1
mainForm.tree.Nodes.add , , "USER HISTORIES", "User Histories"
count = count + 1
hold = treeFormat(Dir$(App.Path & "\Users\*.His"), 2)
ReDim usrArray(0)
usrCount = 0
Do While Not hold = ""
    usrCount = usrCount + 1
    hold = Dir$
    Loop
ReDim usrArray(usrCount)
hold = treeFormat(Dir$(App.Path & "\Users\*.His"), 4)
For usrCount = LBound(usrArray) To UBound(usrArray) - 1
    usrArray(usrCount) = hold
    hold = treeFormat(Dir$, 4)
    Next usrCount
ShellSort usrArray()
For usrCount = LBound(usrArray) + 1 To UBound(usrArray)
    If usrCount = 1 And Not usrArray(usrCount) = vbNullString Then
        mainForm.tree.Nodes.add count - 1, tvwChild, "UH " & UCase$(usrArray(usrCount)), usrArray(usrCount)
        Else
            count = count + 1
            mainForm.tree.Nodes.add count - 1, tvwNext, "UH " & UCase$(usrArray(usrCount)), usrArray(usrCount)
            End If
    Next usrCount
count = count + 1
If UBound(net) > 0 Then
    mainForm.tree.Nodes.add , , "NETLINKS", "Netlinks"
    For count = LBound(net) To UBound(net)
        If Not net(count).name = "" Then
            mainForm.tree.Nodes.add "NETLINKS", tvwChild, , net(count).name
            End If
        Next count
    End If
mainForm.tree.Nodes.add , , "SERVER_INFO", "Server Information"
mainForm.tree.Nodes.add , , "SYSLOG", "System's Logbook"
mainForm.tree.Nodes(1).Selected = True
mainForm.tree.Nodes(1).Expanded = False
End Function
