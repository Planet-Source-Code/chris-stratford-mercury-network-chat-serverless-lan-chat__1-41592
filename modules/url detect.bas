Attribute VB_Name = "URLDETECT"
'HYPERLINK CODE
Public Const TVM_SETBKCOLOR = 4381&
Public Const EM_CHARFROMPOS& = &HD7


Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long


Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As _
    Any) As Long


Public Type POINTAPI
    X As Long
    Y As Long
    End Type
Public hyperlink As String

'HYPERLINK CODE

Public Function getHyperlink(X As Single, Y As Single, iBox As RichTextBox) As String
    'Some Tips that is Safe to Change/Edit f
    '     or Beginers:
    '=======================================
    '     =========================
    ' ibox.MousePointer = rtfCustom
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    ' What MousePointer to Use ? if its rtfC
    '     ustom
    ' Then you need to add a Icon to the Ric
    '     hTextBox
    '=======================================
    '     =========================
    '=======================================
    '     =========================
    'ibox.SelUnderline = True
    ' Do you want the URl to UnderLine ?
    '=======================================
    '     =========================
    '=======================================
    '     =========================
    'ibox.SelColor = purltext
    ' What Color would you like the URL to b
    '     e ?
    '=======================================
    '     =========================
    On Error Resume Next
    Dim point As POINTAPI
    Dim charpos As Long
    Dim pos_start As Long
    Dim pos_end As Long
    Dim char As String
    Dim word As String
    point.X = X \ Screen.TwipsPerPixelX
    point.Y = Y \ Screen.TwipsPerPixelY
    charpos = SendMessage(iBox.hwnd, EM_CHARFROMPOS, 0&, point)


    If charpos <= 0 Or charpos = Len(iBox.Text) Then
        iBox.MousePointer = rtfDefault
        getHyperlink = vbNullString
        Exit Function
    End If


    For pos_start = charpos To 1 Step -1


        If Mid$(iBox.Text, pos_start + 1, 1) = Chr$(13) Then
            iBox.MousePointer = rtfDefault
            getHyperlink = vbNullString
            Exit Function
        End If
        char = Mid$(iBox.Text, pos_start, 1)
        If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
    Next pos_start
    pos_start = pos_start + 1


    For pos_end = charpos To Len(iBox.Text)
        char = Mid$(iBox.Text, pos_end, 1)
        If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
    Next pos_end
    pos_end = pos_end - 1
    If pos_start <= pos_end Then word = LCase$(Mid$(iBox.Text, pos_start, _
    pos_end - pos_start + 1))
    If Left$(word, 7) = "http://" Or Left$(word, 4) = "www." Or Left$(word, 6) = _
    "ftp://" Or Left$(word, 7) = "mailto:" Then
    char = Right$(word, 1)


    Do While char = "." Or char = "," Or char = "!" Or char = "?"
        If Len(char) = 0 Then Exit Do
        word = Left$(word, Len(word) - 1)
        char = Right$(word, 1)
    Loop


    If Len(word) < 4 Then
        iBox.MousePointer = rtfDefault
        getHyperlink = vbNullString
    Else
        iBox.MousePointer = rtfCustom
        getHyperlink = word
    End If
Else
    iBox.MousePointer = rtfCustom
End If
End Function


Public Sub highlightHyperlink(iBox As RichTextBox)
    On Error Resume Next
    Dim pos As Long
    Dim posEnd As Long
    Dim char As String
    Dim link As String
    pos = InStr(1, LCase$(iBox.Text), "mailto:")


    Do While pos > 0


        For posEnd = pos To Len(iBox.Text)
            char = Mid$(iBox.Text, posEnd, 1)
            If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
            
            Next posEnd
            link = Mid$(iBox.Text, pos, posEnd - pos)
            char = Right$(link, 1)


            Do While char = "." Or char = "," Or char = "!" Or char = "?" Or _
                Len(char) <> 1
                link = Left$(link, Len(link) - 1)
                char = Right$(link, 1)
            Loop


            If Len(link) > 7 Then
                iBox.SelStart = pos - 1
                iBox.SelLength = Len(link)
                iBox.SelUnderline = True
                iBox.SelColor = pURLText
            End If
            pos = InStr(posEnd + 1, LCase$(iBox.Text), "ftp://")
        Loop
        pos = InStr(1, LCase$(iBox.Text), "ftp://")


        Do While pos > 0


            For posEnd = pos To Len(iBox.Text)
                char = Mid$(iBox.Text, posEnd, 1)
                If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
                
                Next posEnd
                link = Mid$(iBox.Text, pos, posEnd - pos)
                char = Right$(link, 1)


                Do While char = "." Or char = "," Or char = "!" Or char = "?" Or _
                    Len(char) <> 1
                    link = Left$(link, Len(link) - 1)
                    char = Right$(link, 1)
                Loop


                If Len(link) > 6 Then
                    iBox.SelStart = pos - 1
                    iBox.SelLength = Len(link)
                    iBox.SelUnderline = True
                    iBox.SelColor = pURLText
                End If
                pos = InStr(posEnd + 1, LCase$(iBox.Text), "ftp://")
            Loop
            pos = InStr(1, LCase$(iBox.Text), "http://")


            Do While pos > 0


                For posEnd = pos To Len(iBox.Text)
                    char = Mid$(iBox.Text, posEnd, 1)
                    If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
                    
                    Next posEnd
                    link = Mid$(iBox.Text, pos, posEnd - pos)
                    char = Right$(link, 1)


                    Do While char = "." Or char = "," Or char = "!" Or char = "?" Or _
                        Len(char) <> 1
                        link = Left$(link, Len(link) - 1)
                        char = Right$(link, 1)
                    Loop


                    If Len(link) > 7 Then
                        iBox.SelStart = pos - 1
                        iBox.SelLength = Len(link)
                        iBox.SelUnderline = True
                        iBox.SelColor = pURLText
                    End If
                    pos = InStr(posEnd + 1, LCase$(iBox.Text), "http://")
                Loop
                pos = InStr(1, LCase$(iBox.Text), "www.")


                Do While pos > 0


                    For posEnd = pos To Len(iBox.Text)
                        char = Mid$(iBox.Text, posEnd, 1)
                        If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
                        
                        Next posEnd
                        link = Mid$(iBox.Text, pos, posEnd - pos)
                        char = Right$(link, 1)


                        Do While char = "." Or char = "," Or char = "!" Or char = "?" Or _
                            Len(char) <> 1
                            link = Left$(link, Len(link) - 1)
                            char = Right$(link, 1)
                        Loop


                        If Len(link) > 4 Then
                            iBox.SelStart = pos - 1
                            iBox.SelLength = Len(link)
                            iBox.SelUnderline = True
                            iBox.SelColor = pURLText
                        End If
                        pos = InStr(posEnd + 1, LCase$(iBox.Text), "www.")
                    Loop
                    iBox.SelStart = Len(iBox.Text)
                End Sub


