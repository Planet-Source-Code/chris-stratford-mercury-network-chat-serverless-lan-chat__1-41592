Attribute VB_Name = "TxtEntry_Commands"
Type COMMD
    FirstChar As String
    iCommand As String
    iSetting As String
    iCHATTeXT As String
End Type

Public Sub ENTERCOMMAND(iCommand As String)
'Needs To Send Your:
'Nickname and all seperators!!

'This Makes Sure The Setting Are All Clear By Defining it Locally
Dim EComm As COMMD

'This next block sets up the commands involved in each entry
'Its more complicated than it looks

'Error blocking
On Error Resume Next
    'Sets The First Characted
    EComm.FirstChar = Trim(LCase(Mid(iCommand, 1, 1)))
    'If It Were A Command, What It Could Be
    EComm.iCommand = LCase(Trim(Mid(iCommand, 2)))
    'If It Has Multiple Parts Then
    If InStr(1, EComm.iCommand, " ") > 0 Then
        'Sets The Command What May Be...
        EComm.iCommand = LCase(Trim(Mid(iCommand, 2, InStr(1, iCommand, " ") - 2)))
        'Also sets the variable accompanying the Command
        EComm.iSetting = Trim(Mid(iCommand, InStr(1, iCommand, " ") + 1))
    End If
    'If Its just text... this sets the text
    EComm.iCHATTeXT = Trim(iCommand)

Do While InStr(EComm.iCHATTeXT, vbCrLf) > 0
    DoEvents
    EComm.iCHATTeXT = pReplace(EComm.iCHATTeXT, vbCrLf, "[.]")
Loop

Select Case EComm.FirstChar
    Case "\", "/"
        Select Case EComm.iCommand
        'Contains all the commands and their links to subroutines.
'
'
            Case "list"
            'Calls Up A List Of All Users
            'Very Crude, And Un Exact
            frmMain.lstUsers.Clear
            PrivateMessageRequest = False
                If EComm.iSetting <> "" Then
                    ADDTEXT "User List For '" & EComm.iSetting & "'", pOtherText
                    SD "list", EComm.iSetting
                Else
                    ADDTEXT "User List For '" & pChannel & "'", pOtherText
                    SD "list", pChannel
                    
                End If
'
'
            Case "quit"
                frmMain.mnuQuit_Click
            
'
'
            Case "bindlist"
                ENTERCOMMAND "\bind"
'
'
            Case "bigres"
                ChangeRes 1024, 768
'
'
            Case "bind"
            'This Code Binds Keys And Strings To Each Function Key
            'You can Bind Anything here!
            'Anything at all
            
                If EComm.iSetting = "" Then
                    ADDTEXT "Bind List:", pOtherText
                    ADDTEXT "  F1: " & pBindF1, pListText
                    ADDTEXT "  F2: " & pBindF2, pListText
                    ADDTEXT "  F3: " & pBindF3, pListText
                    ADDTEXT "  F4: " & pBindF4, pListText
                    ADDTEXT "  F5: " & pBindF5, pListText
                    ADDTEXT "  F6: " & pBindF6, pListText
                    ADDTEXT "  F7: " & pBindF7, pListText
                    ADDTEXT "  F8: " & pBindF8, pListText
                    ADDTEXT "  F9: " & pBindF9, pListText
                    ADDTEXT "  F10: " & pBindF10, pListText
                    ADDTEXT "  F11: " & pBindF11, pListText
                    ADDTEXT "  F12: " & pBindF12, pListText
                    ADDTEXT "End Bind List", pOtherText
                Else
                        Dim iBindKey As String
                        Dim iBindText As String
                    If InStr(1, EComm.iSetting, ":") = 0 Then
                        ADDTEXT "To Bind Enter: \bind KEY:TEXT", pOtherText
                        ADDTEXT "   eg, \bind F5:\Help", pOtherText
                    Else
                        iBindKey = Mid(EComm.iSetting, 1, InStr(1, EComm.iSetting, ":") - 1)
                        iBindText = Mid(EComm.iSetting, InStr(1, EComm.iSetting, ":") + 1)
                        Select Case LCase(iBindKey)
                            Case "f1": pBindF1 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f2": pBindF2 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f3": pBindF3 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f4": pBindF4 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f5": pBindF5 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f6": pBindF6 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f7": pBindF7 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f8": pBindF8 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f9": pBindF9 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f10": pBindF10 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f11": pBindF11 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                            Case "f12": pBindF12 = iBindText: ADDTEXT "Bind Bound!", pOtherText
                        Case Else
                            ADDTEXT "To Bind Enter: \bind KEY:TEXT", pOtherText
                            ADDTEXT "   eg, \bind F5:\Help", pOtherText
                            ADDTEXT "Bindable Keys:", pOtherText
                            ADDTEXT "F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12", pOtherText
                        End Select
                        
                    End If
                End If
'
'
            Case "wordcount"
            Dim iWords As Long
            iWords = WORDCOUNT(frmMain.txtChat.Text)
                ADDTEXT "Current Word Count = " & Str(iWords + 5), pNormalText
'
'
            Case "clear"
                frmMain.txtChat.Text = ""
'
'
            Case "banlist"
                    ADDTEXT "Ban List (" & frmMain.lstBannedUsers.ListCount & "):", pOtherText
                    For i = 0 To frmMain.lstBannedUsers.ListCount - 1
                    DoEvents
                           ADDTEXT "    " & frmMain.lstBannedUsers.List(i), pListText
                    Next
                    ADDTEXT "End Ban List", pOtherText
'
'
            Case "ban"
            'This adds a user to your ban list
            
                If EComm.iSetting = "" Then
                    ADDTEXT "Ban List (" & frmMain.lstBannedUsers.ListCount & "):", pOtherText
                    For i = 0 To frmMain.lstBannedUsers.ListCount - 1
                    DoEvents
                           ADDTEXT "    " & frmMain.lstBannedUsers.List(i), pListText
                    Next
                    ADDTEXT "End Ban List", pOtherText
                Else
                    Dim USERBANNED As Boolean
                        For i = 0 To frmMain.lstBannedUsers.ListCount
                        DoEvents
                            If frmMain.lstBannedUsers.List(i) = EComm.iSetting Then
                                'User already banned!
                                USERBANNED = True
                                ADDTEXT "User Already Banned!", pOtherText
                            Else
                                frmMain.lstBannedUsers.AddItem EComm.iSetting
                                ADDTEXT "User " & EComm.iSetting & " Has now been banned!", pOtherText
                                Exit Sub
                            End If
                        Next
                    
                    'Save Changes!
                    Save_Ban_List pBanListPath
                    
                End If
'
'
            Case "remban", "unban"
                If EComm.iSetting = "" Then
                  ADDTEXT "Error: Not A Valid User!", pOtherText
                    
                Else
                    ADDTEXT "Beginnning Unban:", pOtherText
                    For i = 0 To frmMain.lstBannedUsers.ListCount - 1
                    DoEvents
                        If frmMain.lstBannedUsers.List(i) = EComm.iSetting Then
                            frmMain.lstBannedUsers.RemoveItem (i)
                            ADDTEXT "User Has Been UnBanned (" & EComm.iSetting & ")", pListText
                        Else
                        End If
                    Next
                    ADDTEXT "Unban Complete!", pOtherText
                    'Save Changes
                    Save_Ban_List pBanListPath
                End If
'
'
            Case "rempass"
LoadAgain:
            If pDeveloperMode = False Then
                If InBox("Developer Mode", "Enter The Developer Mode Password", , "*") = DevPassword Then
                    pDeveloperMode = True
                    GoTo LoadAgain
                Else
                    ADDTEXT "Error " & EComm.iCommand, pErrorText
                    Exit Sub
                End If
            Else
                   frmMain.CD.Filter = "Profile Files(*.prof)|*.prof"
                   frmMain.CD.DialogTitle = "Remove Password From Profile"
                   frmMain.CD.ShowOpen
                If frmMain.CD.FileName <> "" Then
                  Dim Line1 As String
                  Dim ProfileName As String
                  Dim ProfileSetting As Variant
                  
                    Close #12
                    Open frmMain.CD.FileName For Input As #12
                    Close #13
                    Open frmMain.CD.FileName & ".temp" For Output As #13
                        Do While Not EOF(12)
                        DoEvents
                            Line Input #12, Line1
                                ProfileName = Mid(Line1, 1, InStr(1, Line1, "=") - 1)
                                ProfileSetting = Mid(Line1, InStr(1, Line1, "=") + 1)
                            If LCase(ProfileName) <> "ppassword" Then
                                Print #13, Line1
                            Else
                                Print #13, "ppassword=ÜÞÜÊ"
                            End If
                            
                        Loop
                    Close #12
                    Close #13
                    Kill frmMain.CD.FileName
                    DoEvents
                    Name frmMain.CD.FileName & ".temp" As frmMain.CD.FileName
                    ADDTEXT "Password Has Been Removed From Profile", pNotifyText, , , True
                End If
                
            End If
'
'
            Case "pm"
                Load frmPrivateMsg
'
'
            Case "channel", "join"
             Debug.Print pChannel
                If EComm.iSetting = "" Then
                    ADDTEXT "Current Channel: " & pChannel, pOtherText
                Else
                    CHANNEL EComm.iSetting
                End If
'
'
            Case "pop"
                Select Case EComm.iSetting
                    Case "on", "1", "true"
                        pPopUpWindowOnMessage = True
                        ADDTEXT "Window Popup Has Been Activated!", pNotifyText, , , True
                    Case "off", "0", "false"
                        pPopUpWindowOnMessage = False
                        ADDTEXT "Window Popup Has Been De-Activated!", pNotifyText, , , True
                    Case Else
                        
                End Select
'
'
            Case "aot"
                Select Case EComm.iSetting
                    Case "on", "1", "true"
                        pAlwaysOnTop = True
                        AlwaysOnTop frmMain, True
                        ADDTEXT "Always On Top Has Been Activated!", pNotifyText, , , True
                    Case "off", "0", "false"
                        pawlaysontop = False
                        AlwaysOnTop frmMain, False
                        ADDTEXT "Always On Top Has Been De-Activated!", pNotifyText, , , True
                    Case Else
                        
                End Select
'
'
            Case "devmode"
                Select Case EComm.iSetting
                    Case "on", "1", "true"
                        If InBox("Developer Mode", "Enter The Developer Mode Password", , "*") = DevPassword Then
                            pDeveloperMode = True
                            ADDTEXT "Developer Mode Has Been Activated!", pNotifyText, , , True
                            frmMain.mnuAdmin.Visible = True
                        Else
                            ADDTEXT "Error devmode." & EComm.iSetting, pErrorText
                        End If
                    Case "off", "0", "false"
                        pDeveloperMode = False
                        ADDTEXT "Developer Mode Has Been De-Activated!", pNotifyText, , , True
                        frmMain.mnuAdmin.Visible = False
                    Case Else
                        ADDTEXT "Error devmode." & EComm.iSetting, pErrorText
                End Select
'
'
            Case "fullmsg"
                Select Case EComm.iSetting
                Case "on", "1", "true"
                    If pDeveloperMode = True Then
                        ViewFullMessage = True
                        ADDTEXT "Full Messages Has Been Activated!!", pNotifyText, , , True
                    Else
                        ADDTEXT "Error " & EComm.iCommand & "." & EComm.iSetting, pErrorText
                    End If
                Case "off", "0", "false"
                    If pDeveloperMode = True Then
                        ViewFullMessage = False
                        ADDTEXT "Full Messages Has Been De-Activated!!", pNotifyText, , , True
                    Else
                        ADDTEXT "Error " & EComm.iCommand & "." & EComm.iSetting, pErrorText
                    End If
                Case Else
                    ADDTEXT "Error " & EComm.iCommand & "." & EComm.iSetting, pErrorText
                End Select
'
'
            Case "listfunctions"
                ADDTEXT "Admin Functions", pNormalText, , True
                ADDTEXT "   * Devmode", pListText, , True
                ADDTEXT "   * Fullmsg", pListText, , True
                ADDTEXT "   * RemPass", pListText, , True
                ADDTEXT "End Admin Functions", pNormalText, , True
'
'
            Case "help"
                Load_Help
'
'
            Case "[test html]"
                frmMain.txtChat.TextRTF = frmUpdates.Inet1.OpenURL("http://www.hotmail.com")
                Unload frmUpdates
                frmMain.Show
'
'
            Case "pchannel"
                PrivateChannel EComm.iSetting
'
'
            Case "bytes"
                ADDTEXT "   You Have Received: " & Str(TotalBytes), pNormalText, True
'
'
            Case "trans"
                On Error Resume Next
                SETTrans Int(EComm.iSetting)
'
'
            Case "check"
                Load frmUpdates
                frmUpdates.Show
'
'
            Case "name", "handel"
                If EComm.iSetting = "" Then
                    ADDTEXT "Current Nickname: " & pHandel, pOtherText
                Else
                    pHandel = EComm.iSetting
                    ADDTEXT "New Nickname: " & pHandel, pOtherText
                End If
'
'
            Case "info"
                ADDTEXT "Your Computer Stats:", pNormalText
                ADDTEXT "   Merc Chat Version: " & VersionCheck, pListText
                ADDTEXT "   Windows Version: " & OSVersion, pListText
                ADDTEXT "   IP: " & WS.LocalIP, pListText
                ADDTEXT "   Logged In As: " & Get_User_Name, pListText
                    Dim iNICK As String
                    If pHandel = "" Then iNICK = WS.LocalIP Else iNICK = pHandel
                ADDTEXT "   Nickname: " & iNICK, pListText
                ADDTEXT "   You Are In The Channel: " & pChannel, pListText
                ADDTEXT "   Current Profile: " & pProfileName, pListText
'
'
            Case "about"
                frmAbout.Show
'
'
            Case "profile"
                If EComm.iSetting = "" Then
                    frmMain.CD.Filter = "Profile Files(*.prof)|*.prof"
                    frmMain.CD.ShowOpen
                    If frmMain.CD.FileName <> "" Then LOAD_Profile frmMain.CD.FileName
                
                ElseIf EComm.iSetting = "testprofile" Then
                    LOAD_Profile "test"
                
                ElseIf EComm.iSetting = "default" Then
                    LOAD_Profile "default"
                
                ElseIf EComm.iSetting = "new" Then
                    frmMain.mnuNewProfile_Click
                
                Else
                    LOAD_Profile APPPATH & "skins\" & EComm.iSetting & ".prof"
                
                End If
'
'
            'If The Command Is Unknown
            Case Else
                If EComm.iSetting <> "" Then
                    ADDTEXT "Unknown Command - " & EComm.iCommand & "." & EComm.iSetting, pErrorText
                Else
                    ADDTEXT "Unknown Command - " & EComm.iCommand, pErrorText
                End If
        End Select
        
    Case "#"
        CHANNEL Mid(EComm.iCHATTeXT, 2)
    
    Case "!"
        'Send Private Message
        'PM's arent Channel Based!
        'Therefore The Channel is Irrelivant
        
        'ANTI SPAM PROTECTION!
        If frmMain.tmrAntiSpam.Enabled = True Then ADDTEXT "NO SPAM", pErrorText, , , True: Exit Sub
        
        
        If InStr(EComm.iCHATTeXT, "|") > 0 Then
        
        Dim eCHaTTexT As String
        Dim eTHeiRiD As String
        
        eTHeiRiD = Mid(EComm.iCHATTeXT, 2, InStr(EComm.iCHATTeXT, "|") - 2)
        eCHaTTexT = Mid(EComm.iCHATTeXT, InStr(EComm.iCHATTeXT, "|") + 1)
        
            ADDTEXT "[PM: " & Get_User_Name & " (" & eTHeiRiD & ")] - " & eCHaTTexT, pYourPMText, True
            SDSpecial "pmsg", "[PM: " & Get_User_Name & "] - " & eCHaTTexT, eTHeiRiD, Get_User_Name
        Else
            ADDTEXT ChatHandel & EComm.iCHATTeXT, pYourChatText
            SD "msg", ChatHandel & EComm.iCHATTeXT
        End If
        
        'ANTI SPAM PROTECTION!
        frmMain.tmrAntiSpam.Enabled = True
        
    Case "$"
        pHandel = Mid(EComm.iCHATTeXT, 2)
        ADDTEXT "New Nickname: " & pHandel, pOtherText
    
    Case Else
        'Spam Protection
        'Only Allows a message every .1 of a second (Or otherwise specified- may change)
        
        
        'I Have 2 Types Of Spam protection. This One, Erases Any Back Logged Messages
        'Its Commented Out Because If you accidentally spam, then you wont loose
        'your message
        
        If frmMain.tmrAntiSpam.Enabled = True Then ADDTEXT "NO SPAM", pErrorText, , , True: Exit Sub
        
        'This process of spam protection will log all the spammed messages, and
        'Will send them after a short .1 of a second break
        
        'Do While frmMain.tmrAntiSpam.Enabled = True
        '    DoEvents
        'Loop
        
        'If its normatl chat text
        
            ADDTEXT ChatHandel & EComm.iCHATTeXT, pYourChatText
            SD "msg", ChatHandel & EComm.iCHATTeXT
        
        'Turn Spam Protection ON!
        'This will last for .1 of a second, and will pause the program
        frmMain.tmrAntiSpam.Enabled = True
End Select
End Sub
