Attribute VB_Name = "WordCountModule"
'**************************************
'Windows API/Global Declarations for :Ch
'     ange Screen Resolution (on the fly even!
'     )
'**************************************


Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean


Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    Const CCDEVICENAME = 32
    Const CCFORMNAME = 32
    Const DM_PELSWIDTH = &H80000
    Const DM_PELSHEIGHT = &H100000


Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    End Type
    Dim DevM As DEVMODE


Public Function WORDCOUNT(iSource As String) As Long
    Dim strInput As String
    Dim strWords() As String
    Dim x As Integer
    Dim y As Integer
    Dim FoundWords As Integer

FoundWords = 0
strInput = Trim(iSource) 'strip any spaces from the beginning and end to speed up the search a bit

If Len(strInput) > 0 Then 'if there's anything left after we stripped spaces...
    strWords = Split(strInput, " ") 'split the "words" into an array
    For x = 0 To UBound(strWords)   'for every one of the "words" we found...
        DoEvents
        If Len(strWords(x)) > 0 Then    'if there's actually something here then...
            'for y=97 to 122            'uncomment this block to only count "words" that contain the letters a-z
                'if instr(1,lcase(strwords(x)),chr(y))>0 then
                    FoundWords = FoundWords + 1 'update the number of words found
                    'exit for   'so we don't count the same word more than once :P
                'end if
            'next y
        End If
    Next x
End If

For i = 1 To Len(iSource)
DoEvents
    If Mid(iSource, i, 2) = vbCrLf Then
    'This Means That There Has Been An Enter!
        If Mid(iSource, i + 2, 2) <> vbCrLf Then
            'Its an enter then a word!
            If Mid(iSource, i + 2, 1) <> " " Then
            'Check Its Not Enter THen SPACE
                If Mid(iSource, i + 2, 1) <> "" Then
                'Check Its Not A Trailing Enter
                    FoundWords = FoundWords + 1
                End If
            End If
        End If
        
        'If i > 1 Then
        On Error Resume Next
            If Mid(iSource, i - 1, 1) = " " Then
            'If Its An Enter With A Space But No Word!
            'THe PRevious Code Has Counted it, Now we must Remove it!
                FoundWords = FoundWords - 1
            End If
        'End If
    End If
Next

WORDCOUNT = FoundWords

End Function



Sub ChangeRes(iWidth As Single, iHeight As Single)

    Dim a As Boolean
    Dim i&
    i = 0


    Do
        a = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (a = False)
    Dim b&
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight
    b = ChangeDisplaySettings(DevM, 0)
End Sub
