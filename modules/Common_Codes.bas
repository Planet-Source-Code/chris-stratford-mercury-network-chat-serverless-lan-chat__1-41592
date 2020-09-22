Attribute VB_Name = "Primary_Codes"
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

    Global Const GWL_EXSTYLE = (-20)
    Global Const WS_EX_LAYERED = &H80000
    Global Const WS_EX_TRANSPARENT = &H20&
    Global Const LWA_ALPHA = &H2&
Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwflags As Long) As Long
#If Win16 Then 'Conditional Compile statements


Declare Sub SetWindowPos Lib "User" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#Else


Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If


Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    End Type
    Private Const VER_PLATFORM_WIN32s = 0
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32_NT = 2


Private Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" (lpVersionInformation As _
    OSVERSIONINFOEX) As Long


Declare Function GetTickCount& Lib "kernel32" ()

Dim NormalWindowStyle As Long
Dim NORMALNUM As Long

'used for shelling out to the default web browser
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1

Private Declare Function ShellEx Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As Any, ByVal lpDirectory As Any, ByVal nShowCmd As Long) As Long

Sub ShellDef(iPath As String)
    X = ShellEx(frmMain.hwnd, "open", file_name, "", "", 10)
End Sub


Public Function WindowsRunTime() As Long
    WindowsRunTime = GetTickCount()
End Function

Public Function pReplace(strExpression As String, strFind As String, strReplace As String)
    Dim intX As Integer


    If (Len(strExpression) - Len(strFind)) >= 0 Then


        For intX = 1 To Len(strExpression)


            If Mid(strExpression, intX, Len(strFind)) = strFind Then
                strExpression = Left(strExpression, (intX - 1)) + strReplace + Mid(strExpression, intX + Len(strFind), Len(strExpression))
            End If
        Next
    End If
    pReplace = strExpression
End Function


Public Sub OpenWeb(iUrl As String)
ShellExecute hwnd, "open", iUrl, vbNullString, vbNullString, conSwNormal
End Sub

Public Function OSVersion() As String
'PSCODE.COM supplied this code to help me recognise what version of windows the user has.

    Dim udtOSVersion As OSVERSIONINFOEX
    Dim lMajorVersion As Long
    Dim lMinorVersion As Long
    Dim lPlatformID As Long
    Dim sAns As String
    udtOSVersion.dwOSVersionInfoSize = Len(udtOSVersion)
    GetVersionEx udtOSVersion
    lMajorVersion = udtOSVersion.dwMajorVersion
    lMinorVersion = udtOSVersion.dwMinorVersion
    lPlatformID = udtOSVersion.dwPlatformId


    Select Case lMajorVersion
        Case 5
        sAns = "Windows 2000"
        Case 4


        If lPlatformID = VER_PLATFORM_WIN32_NT Then
            sAns = "Windows NT 4.0/5.0"
        Else
            sAns = IIf(lMinorVersion = 0, _
            "Windows 95/ME", "Windows 98/ME")
        End If
        Case 3


        If lPlatformID = VER_PLATFORM_WIN32_NT Then
            sAns = "Windows NT 3.x"
            
            'below should only happen if person has
            '     Win32s
            'installed
        Else
            sAns = "Windows 3.x"
        End If
        Case Else
        sAns = "Unknown Windows Version"
    End Select
OSVersion = sAns
End Function



Public Sub AlwaysOnTop(Form As Form, Value As Boolean)
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

Select Case Value
    Case True: SetWindowPos Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Case False: SetWindowPos Form.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Select
End Sub

Public Function InBox(iTitle As String, iCaption As String, Optional iDefault As String, Optional iPassChar As String)
    frmMain.Enabled = False
    
    InputTemp = ""
    Load frmInputbox

    frmInputbox.tmrINPUTBOXPAUSE.Enabled = True
    frmInputbox.Caption = iTitle
    frmInputbox.Label1.Caption = iCaption


If iPassChar <> "" Then
    frmInputbox.Text1.PasswordChar = Mid(iPassChar, 1, 1)
Else
    frmInputbox.Text1.PasswordChar = ""
End If


If iDefault <> "" Then
    frmInputbox.Text1.Text = iDefault
End If

frmInputbox.Text1.SetFocus
frmInputbox.Text1.SelStart = 0
frmInputbox.Text1.SelLength = Len(frmInputbox.Text1)



Do While frmInputbox.tmrINPUTBOXPAUSE.Enabled = True
    DoEvents
    DoEvents
    DoEvents
Loop
    
    InBox = InputTemp
    frmMain.Enabled = True
    Unload frmInputbox
End Function


Public Sub Pause(Duration As Long)
'Creates A Short Break In The Code
Duration = Duration / 1000
Dim Current As Long
Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub


Public Sub PrivateChannel(NewChan As String)
If NewChan = "" Then Exit Sub


        SD "leaving"
    ADDTEXT "   |You Have Left '" & pChannel & "'", pNotifyText
        Pause 50
    pChannel = "::" & NewChan
    ADDTEXT "   |You Have Joined Private Channel - " & Mid(pChannel, 3), pNotifyText
    SD "joining"
        Pause 50
End Sub


Public Sub CHANNEL(iChannel As String)

If pChannel <> iChannel Then
    
    If FALSECHANNEL(iChannel) Then
    
        iChannel = pReplace(iChannel, " ", "_")
        
        SD "leaving"
        ADDTEXT "   |You Have Left '" & pChannel & "'", pNotifyText
            Pause 50
        pChannel = iChannel
            Pause 50
        ADDTEXT "   |You Have Joined '" & pChannel & "'", pNotifyText
            SD "joined"
       
    Else
        ADDTEXT "The Channel You Requested to Join Is Restricted!", pErrorText, True
    End If
Else
        ADDTEXT "You are Already In " & iChannel & "!", pErrorText, True
End If
SF
End Sub

Public Sub SF()
'Bad Coding - I Know
'Just Quick And Easy Way To Set The Focus Again

    On Error Resume Next
    frmMain.txtEntry.SetFocus
End Sub

Public Function FALSECHANNEL(iChannel) As Boolean
If Trim(iChannel) = "" Then
    FALSECHANNEL = False
    Exit Function
End If

If InStr(1, iChannel, "::") > 0 Then
    FALSECHANNEL = False
    Exit Function
End If

If InStr(iChannel, "%") > 0 Then FALSECHANNEL = False: Exit Function

For i = 0 To frmMain.lstBarredChannels.ListCount
    If LCase(iChannel) = LCase(frmMain.lstBarredChannels.List(i)) Then
        FALSECHANNEL = False
        Exit Function
    Else
        FALSECHANNEL = True
    End If
Next
End Function


Public Sub Load_Help()
Dim i As Integer
For i = 0 To frmMain.lstHelp.ListCount
    Select Case Even(i)
        Case True
            ADDTEXT frmMain.lstHelp.List(i), pHelpTextN, True
        Case False
            ADDTEXT frmMain.lstHelp.List(i), pHelpTextD, False, True
    End Select
Next
End Sub

Public Function Even(iNum As Integer) As Boolean

If InStr(1, Trim(Str(iNum / 2)), ".") = 0 Then
    Even = True
Else
    Even = False
End If
End Function

Public Sub ADDTEXT(iTEXT As String, Optional iColour As Long, Optional iBold As Boolean, Optional iItalic As Boolean, Optional iUnderline As Boolean)

'Shortcuts In The Text
iTEXT = pReplace(iTEXT, LCase("%time%"), Time)
iTEXT = pReplace(iTEXT, LCase("%date%"), Date)
iTEXT = pReplace(iTEXT, LCase("%now%"), Now)

iTEXT = pReplace(iTEXT, LCase("%f1%"), pBindF1)
iTEXT = pReplace(iTEXT, LCase("%f2%"), pBindF2)
iTEXT = pReplace(iTEXT, LCase("%f3%"), pBindF3)
iTEXT = pReplace(iTEXT, LCase("%f4%"), pBindF4)
iTEXT = pReplace(iTEXT, LCase("%f5%"), pBindF5)
iTEXT = pReplace(iTEXT, LCase("%f6%"), pBindF6)
iTEXT = pReplace(iTEXT, LCase("%f7%"), pBindF7)
iTEXT = pReplace(iTEXT, LCase("%f8%"), pBindF8)
iTEXT = pReplace(iTEXT, LCase("%f9%"), pBindF9)
iTEXT = pReplace(iTEXT, LCase("%f10%"), pBindF10)
iTEXT = pReplace(iTEXT, LCase("%f11%"), pBindF11)
iTEXT = pReplace(iTEXT, LCase("%f12%"), pBindF12)

iTEXT = pReplace(iTEXT, LCase("%wordcount%"), Str(WORDCOUNT(frmMain.txtChat.Text)))
iTEXT = pReplace(iTEXT, LCase("%ip%"), frmMain.WS.LocalIP)
iTEXT = pReplace(iTEXT, LCase("%title%"), frmMain.Caption)
iTEXT = pReplace(iTEXT, LCase("%windows%"), OSVersion)
iTEXT = pReplace(iTEXT, LCase("%status%"), frmMain.StatusBar1.Panels(2).Text)
iTEXT = pReplace(iTEXT, LCase("%profile%"), pProfileName)
iTEXT = pReplace(iTEXT, LCase("%bytes%"), Str(TotalBytes))
iTEXT = pReplace(iTEXT, LCase("%nick%"), pHandel)
iTEXT = pReplace(iTEXT, LCase("%font%"), pFont)
iTEXT = pReplace(iTEXT, LCase("%devmode%"), Str(pDeveloperMode))
iTEXT = pReplace(iTEXT, LCase("%id%"), Get_User_Name)
iTEXT = pReplace(iTEXT, LCase("%version%"), VersionCheck)
iTEXT = pReplace(iTEXT, vbCrLf, " :.: ")


If iTEXT = "" Then Exit Sub

With frmMain.txtChat
    .SelStart = Len(.Text)
    .SelLength = Len(.Text)
    .SelBold = iBold
    .SelItalic = iItalic
    .SelUnderline = iUnderline
    .SelColor = iColour
    .SelText = iTEXT & vbCrLf
    .SelLength = Len(.Text)
    .Font = pFont
End With

SF


If InStr(LCase(iTEXT), "http://") > 0 Then
    highlightHyperlink frmMain.txtChat
ElseIf InStr(LCase(iTEXT), "mailto:") > 0 Then
    highlightHyperlink frmMain.txtChat
ElseIf InStr(LCase(iTEXT), "ftp://") > 0 Then
    highlightHyperlink frmMain.txtChat
ElseIf InStr(LCase(iTEXT), "www.") > 0 Then
    highlightHyperlink frmMain.txtChat
End If

If Len(frmMain.txtChat.Text) > 10000 Then
    frmMain.txtChat = ""
End If

End Sub


Public Sub SD(iCom As String, Optional iEXTRA As String)
'Send Any Information, But Always Sends The Logged in User

'Diverts Error When Port In Use
On Error GoTo ErrorUDPInUse


    With frmMain.WS
        .Close
        .Protocol = sckUDPProtocol
        .LocalPort = BroadcastPORT
        .RemotePort = BroadcastPORT
        .RemoteHost = "255.255.255.255"
        'Sends Data, Totally Secure Using True Encoding With A Key - iENCODEKEY
        .SendData Encode(iCom & cSEP1 & pChannel & cSEP2 & Get_User_Name & cSEP3 & iEXTRA, iENCODEKEY)
    End With

Exit Sub:
ErrorUDPInUse:
    ADDTEXT "Message Not Sent!", pErrorText
    ADDTEXT "Port Error! Already In Use!", pErrorText

End Sub

Public Sub SDSpecial(Optional iCom As String, Optional iEXTRA As String, Optional iChannel As String, Optional iUserName As String)
'Send Any Information, But Always Sends The Logged in User

'Diverts Error When Port In Use
On Error GoTo ErrorUDPInUse


    With frmMain.WS
        .Close
        .Protocol = sckUDPProtocol
        .LocalPort = BroadcastPORT
        .RemotePort = BroadcastPORT
        .RemoteHost = "255.255.255.255"
        'Sends Data, Totally Secure Using True Encoding With A Key - iENCODEKEY
        .SendData Encode(iCom & cSEP1 & iChannel & cSEP2 & iUserName & cSEP3 & iEXTRA, iENCODEKEY)
    End With

Exit Sub:
ErrorUDPInUse:
    ADDTEXT "Message Not Sent!", pErrorText
    ADDTEXT "Port Error! Already In Use!", pErrorText

End Sub

Public Function ChatHandel() As String
'This Gets The Current HANDEL Of The User
If pHandel = "" Then
    ChatHandel = "[" & Get_User_Name & "]-" & frmMain.WS.LocalIP & pChatSep
Else
    ChatHandel = "[" & Get_User_Name & "]-" & pHandel & pChatSep
End If
End Function


Public Function SETTrans(iValue As Integer, Optional Silent As Boolean)
'All Transparency Code Adapted From
'www.planetsoureccode.com
'Alot of this has been modified by myself for more personal use, and understanding.
'Comments Have All Been Added By Myself!


'I Have Incorperated A Check that allows this code to
'Work on only Win2K. Which is what should only happen
If OSVersion <> "Windows 2000" Then
    If Silent = False Then ADDTEXT "Error: You Must Have Windows 2000 or Greater to use this function!", pErrorText
Else


'Checks That The Value is A Percentage
        If Int(iValue) < 101 And Int(iValue) > -1 Then
            
            If Silent = False Then ADDTEXT "Transparency Level Set To: " & iValue, pOtherText
        Else
            If Silent = False Then ADDTEXT "Error! Number Must Be Between 0-100", pErrorText
            Exit Function
        End If
                
'Gets The Window Style
    NormalWindowStyle = GetWindowLong(frmMain.hwnd, GWL_EXSTYLE)
        
'Sets The Default Window Style - This Is Used When Resetting To 0
    If NORMALNUM = 0 Then
        NORMALNUM = NormalWindowStyle
    End If

'This notes the current Transparency Level
    CURTrans = iValue

'If The Person Is Setting Trans, To Zero
'It Needs To Change The Window Type to Default, And Not A Transparent Frame
    If iValue = 0 And CURTrans < 1 Then
        SetLayeredWindowAttributes frmMain.hwnd, 0, 255 * (1 - (Val(iValue) / 100)), LWA_ALPHA
        SetWindowLong frmMain.hwnd, GWL_EXSTYLE, NORMALNUM
        Exit Function
    End If

'This sets the window as A Transparent frame
    SetWindowLong frmMain.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED

'This sets the Transparency Level Of The window
    SetLayeredWindowAttributes frmMain.hwnd, 0, 255 * (1 - (Val(iValue) / 100)), LWA_ALPHA


End If
End Function

Public Function DirEXISTS(iDir As String) As Boolean
On Error GoTo Nope
    frmLoading.FILELISTING_USED_TO_VALIDATA_FOLDERS.Path = iDir
    DirEXISTS = True
Exit Function
Nope:
    DirEXISTS = False
End Function




Public Function APPPATH()
If Len(App.Path) = 3 Then
    APPPATH = App.Path
Else
    APPPATH = App.Path & "\"
End If
End Function

Public Function VersionCheck() As String
    VersionCheck = App.Major & "." & App.Minor & "." & App.Revision
End Function

Function Get_User_Name()
'API Call From
'www.planetsourcecode.com
'Only This Sub & The 'Private Declare Function'
'Everything Else I Have Adapted Myself
    Dim lpBuff As String * 25
    Dim ret As Long
    ret = GetUserName(lpBuff, 25)
    Get_User_Name = UCase(Left(lpBuff, InStr(lpBuff, Chr(0)) - 1))
End Function

Public Sub Load_Ban_List(iListPath As String)
On Error GoTo NoFile
frmMain.lstBannedUsers.Clear
    pBanListPath = iListPath
Dim LineIN As String
Close #1
Open iListPath For Input As #1
    Do While Not EOF(1)
        DoEvents
            Line Input #1, LineIN
            If Trim(LineIN) <> "" Then frmMain.lstBannedUsers.AddItem LineIN
    Loop
Close #1
Exit Sub
NoFile:
'Error When No File Is Present
ADDTEXT "No Ban List Loaded", pErrorText, , True
End Sub

Public Function FileExists(iPath As String) As Boolean
If DirEXISTS(Mid(iPath, 1, InStrRev(iPath, "\") - 1)) = True Then
    frmLoading.File1.Path = Mid(iPath, 1, InStrRev(iPath, "\") - 1)
    For i = 0 To frmLoading.File1.ListCount - 1
        If frmLoading.File1.List(i) = Mid(iPath, InStrRev(iPath, "\") + 1) Then
            FileExists = True
            Exit Function
        End If
    Next
Else
    FileExists = False
    Exit Function
End If
End Function


Public Sub Save_Ban_List(iListPath)
On Error GoTo NoFile
Close #1
Open iListPath For Output As #1
    For i = 0 To frmMain.lstBannedUsers.ListCount - 1
        Print #1, frmMain.lstBannedUsers.List(i)
    Next
Close #1
Exit Sub
NoFile:
'When Error For No File Occurs
ADDTEXT "Error Saving Ban List, Profile Linked Path Is Incorrect!", pErrorText
End Sub
