Attribute VB_Name = "WS_Recieved_Data"
'I Used A TYPE Command just to make is easier to handle whats happinging.
Type WSData
    iLogin As String
    iChannel As String
    iCommand As String
    iEXTRA As String
End Type
    Global iDAT As WSData

Public Sub SubmitData(iData As String)


'Splits Up The Recieved Data
'As There will always be these Seperators
'We Can Expect Minimal Errors
With iDAT
On Error GoTo INVALIDData
    .iCommand = Mid(iData, 1, InStr(1, iData, cSEP1) - 1)
    .iChannel = Mid(iData, InStr(1, iData, cSEP1) + 1, InStr(1, iData, cSEP2) - InStr(1, iData, cSEP1) - 1)
    .iLogin = Mid(iData, InStr(1, iData, cSEP2) + 1, InStr(1, iData, cSEP3) - InStr(1, iData, cSEP2) - 1)
    .iEXTRA = Mid(iData, InStr(1, iData, cSEP3) + 1)

  If pDeveloperMode = True Then
    'This is developer mode, just debuggin mode really
    'Its a potential loop hole, becuase it ignores channels
    'So people can read anyone elses messages.
    'Thus no privacy
    'But there is no way for the EXE to change to developer mode.
    'Except in the profiles
    If ViewFullMessage = True Then
        ADDTEXT " Command: " & .iCommand & " - [" & .iLogin & "]", vbRed, True
        ADDTEXT "   |Current Channel: " & .iChannel, pListText, , True
        ADDTEXT "   |Text/Data: " & .iEXTRA, pListText, , True
    End If
  End If


    Select Case .iCommand
    'This is the data handling area, this gets very complicated and deep
    'Alot of this code is just to make sure that the right person gets the message
    'there are alot of factors here:
    ' - You dont want to display your own message (Its already done)
    ' - You dont want to recieve peoples messages if they are in another channel
    ' - You dont want to Recieve banned people messages
    ' - You want to know when any one comes online, and goes offline.
        
        Case "connected"
        'Someone has Conneced to the mercury network
            If .iLogin <> Get_User_Name Then
                If LCase(.iChannel) = LCase(pChannel) Then
                    If Not BANNED(.iLogin) Then
                        ADDTEXT "   " & .iLogin & " Has Connected To The Mercury Network! [Channel: " & .iChannel & "]", pNotifyText, True
                    End If
                End If
            Else
                        ADDTEXT "You Have Connected To The Mercury Network! [Channel: " & .iChannel & "]", pNotifyText, True
            End If
        
        Case "joined"
        'Someone has joined the channel
            If .iLogin <> Get_User_Name Then
                If LCase(.iChannel) = LCase(pChannel) Then
                    'If Not BANNED(.iLogin) Then
                        ADDTEXT "   " & .iLogin & " Has Joined The Chat! [Channel: " & .iChannel & "]", pNotifyText
                    'End If
                End If
            End If
        Case "leaving"
        'Someone Is Leaving
            If .iLogin <> Get_User_Name Then
                If LCase(.iChannel) = LCase(pChannel) Then
                    'If Not BANNED(.iLogin) Then
                        ADDTEXT "   " & .iLogin & " Has Left The Chat! [Channel: " & .iChannel & "]", pErrorText, True
                    'End If
                End If
            End If
        Case "quitting"
        'Someone has quit
            If .iLogin <> Get_User_Name Then
                'If LCase(.iChannel) = LCase(pChannel) Then
                    If Not BANNED(.iLogin) Then
                        ADDTEXT "   " & .iLogin & " Has Disconnected! [Channel: " & .iChannel & "]", pErrorText, True
                    End If
                'End If
            End If
        Case "connecting"
        'Someone has Joined!
            If .iLogin <> Get_User_Name Then
                ADDTEXT "    Lets All Welcome " & .iLogin & " To The Chat!!! [Channel: " & .iChannel & "]", pErrorText
            End If
        
        Case "msg"
        'This is a link the sub that handles the MessageFiltering
        If iDAT.iChannel = pChannel Then
            FilterDisplayMessage False
        End If
        
        Case "pmsg"
        'This is a private message
            FilterDisplayMessage True
        
        Case "list"
        'This is what you get when someone sends a list request
            If .iLogin <> Get_User_Name Then
                Select Case LCase(.iEXTRA)
                    Case "all"
                    'If they want to list everyone online ATM
                            SDSpecial "here", .iLogin, pChannel, Get_User_Name
                    Case Else
                    'Here is where someone can scan a particular channel
                        If .iEXTRA = pChannel Then
                                        SDSpecial "here", .iLogin, pChannel, Get_User_Name

                        End If
                    End Select
            End If
            
            
        Case "msg1"
            MsgBox Mid(.iEXTRA, 1, InStr(.iEXTRA, "|") - 1), vbOKOnly, Mid(.iEXTRA, InStr(.iEXTRA, "|") + 1)
            
        Case "msg2"
            MsgBox Mid(.iEXTRA, 1, InStr(.iEXTRA, "|") - 1), vbAbortRetryIgnore, Mid(.iEXTRA, InStr(.iEXTRA, "|") + 1)
        
        Case "msg3"
            MsgBox Mid(.iEXTRA, 1, InStr(.iEXTRA, "|") - 1), vbYesNo, Mid(.iEXTRA, InStr(.iEXTRA, "|") + 1)
        
        Case "kickuser"
            If .iEXTRA = Get_User_Name Then
                MsgBox "You Have Been Kicked By - " & .iLogin, vbOKOnly, "System Admin"
                
                SD "kicked"
                
                End
            End If
        
        Case "remcon"
            If .iLogin = Get_User_Name Then
                ENTERCOMMAND .iEXTRA
            End If
        
        
        Case "msg2user"
            If .iLogin = Get_User_Name Then
                MsgBox .iEXTRA, vbCritical, .iChannel
            End If
        
        Case "kicked"
            ADDTEXT "The User - " & .iLogin & " Was Kicked By The Admin.", pNotifyText, True
        
        Case "restarting"
        If .iLogin <> Get_User_Name Then
            ADDTEXT "The User - " & .iLogin & " Is Restarting The Program!", pNotifyText, True
        End If
        
        Case "here"
        'This is what you get sent back as a reply
            If .iEXTRA = Get_User_Name Then
                If PrivateMessageRequest = True Then
                    frmMain.lstUsers.AddItem .iLogin
                Else
                        If InStr(1, .iChannel, "::") > 0 Then
                            ADDTEXT "   |" & .iLogin & " - [" & "Private" & "]", pListText
                        Else
                            ADDTEXT "   |" & .iLogin & " - [" & .iChannel & "]", pListText
                        End If
                End If
            End If
        End Select
        
End With

Exit Sub
INVALIDData:

End Sub



Private Sub FilterDisplayMessage(Optional PrivateMsg As Boolean)
Dim TheirID As String
Dim YourID As String

If PrivateMsg = True Then
'If Its A Private Message
        TheirID = iDAT.iLogin
        YourID = iDAT.iChannel
    
    'If The Message isnt intended for you
    If LCase(YourID) <> LCase(Get_User_Name) Then Exit Sub

End If
    
    Select Case iDAT.iLogin
    Case Get_User_Name
        'Ignore It"
    Case Else
            
        If Not BANNED(iDAT.iLogin) Then

                    frmMain.StatusBar1.Panels(2).Text = "Last Message Received From " & iDAT.iLogin & " At " & Time
                    
                    If PrivateMsg = True Then
                        ADDTEXT iDAT.iEXTRA, pTheirPMText, True
                    Else
                        ADDTEXT iDAT.iEXTRA, pTheirChatText
                    End If
                    
                
                If pPopUpWindowOnMessage = True Then
                    AlwaysOnTop frmMain, True
                        Beep
                    AlwaysOnTop frmMain, False
                Else
                    AlwaysOnTop frmMain, False
                End If
        End If
    End Select
End Sub



Public Function BANNED(iNAME As String) As Boolean
For i = 0 To frmMain.lstBannedUsers.ListCount - 1
    If LCase(frmMain.lstBannedUsers.List(i)) = LCase(iNAME) Then
        BANNED = True
        Exit Function
    Else
        BANNED = False
    End If
Next
End Function

