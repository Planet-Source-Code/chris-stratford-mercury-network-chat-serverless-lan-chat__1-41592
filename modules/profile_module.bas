Attribute VB_Name = "Profile_Module"

Public Sub LOAD_Profile(iProfile_Path As String)

'This Is The Code Which Loads Your Profile
'I Have Tried To Chop It Up As Much As I Can To Pin Point
'Errors

pPassword = ""
pProfileMessage = ""
pProfileMSGBOX = ""
pProfileURL = ""

If iProfile_Path = "default" Then
'This Loads The Defaults
    Load_Defaults
ElseIf iProfile_Path = "test" Then
    Load_TestStyle
Else
'This Loads The File Given To The Program

On Error GoTo NoFileExists
Dim LineIN As String

'Check For Password


'If No password Specified, Then Its An Edited Profile Because They Always Come With Passwords

    Close #1
        Open iProfile_Path For Input As #1
            Do While Not EOF(1)
                Line Input #1, LineIN
                    If Trim(LineIN) <> "" Then SETProfileInfo LineIN
            Loop
        Close #1

End If
'Apply The Physical Settings
If pPassword = "" Then pPassword = Encode(DevPassword, DevPassword)

If Trim(pPassword) <> "ÜÞÜÊ" Then
    If InBox("Password", "Enter Profile Password", , "*") <> Decode(pPassword, pProfileName) Then
        LOAD_Profile "default"
        DoEvents
        ADDTEXT "Incorrect Password!", vbRed, True, , True
        Exit Sub
    Else
        
    End If
End If


With frmMain
    .txtChat.Text = ""
    .txtChat.BackColor = pChatTextBack
    .txtEntry.BackColor = pEntryTextBack
    .txtEntry.ForeColor = pEntryText
    .BackColor = pWindowBack
    .Caption = pTitleTextConnected
    .txtEntry.Font = pFont


If LCase(pPicturePath) <> "(none)" Then
    If pPicturePath <> "" Then
        On Error Resume Next
        .Picture1.Picture = LoadPicture(pPicturePath)
    Else
        .Picture1.Picture = .Picture2.Picture
    End If
End If


    Load_Ban_List pBanListPath
    
    If pProfileMSGBOX <> "" Then
        MsgBox pProfileMSGBOX, vbOKOnly, pProfileName
    End If
    
    .Load_Welcome_Note
        
'This is now controlled after loading, otherwise it causes difficulty
    SETTrans pTransparency, True
    
    ADDTEXT pProfileMessage, pNotifyText, True, , True
    If pProfileURL <> "" Then ADDTEXT "Home URL - " & pProfileURL, pNotifyText, , True

End With


Exit Sub
NoFileExists:
    ADDTEXT ""
    ADDTEXT "Error Loading: " & iProfile_Path, pErrorText, True
    ADDTEXT "File Doesn't Contain Profile Information, Or Cannot Open The File", pErrorText, , True
    ADDTEXT "No Changes Have Been Made", pErrorText, , True
    ADDTEXT ""
End Sub

Private Sub SETProfileInfo(iSetting As Variant)
Dim ProfileName As String
Dim ProfileSetting As Variant

    ProfileName = Mid(iSetting, 1, InStr(1, iSetting, "=") - 1)
    ProfileSetting = Mid(iSetting, InStr(1, iSetting, "=") + 1)

Select Case LCase(ProfileName)
                Case "pprofilename"
                    pProfileName = ProfileSetting
    
                Case "pnormaltext"
                    pNormalText = ProfileSetting
                    
                Case "plisttext"
                    pListText = ProfileSetting
                    
                Case "perrortext"
                    pErrorText = ProfileSetting
                    
                Case "pbanlistpath"
                    pBanListPath = ProfileSetting
                                
                Case "pyourpmtext"
                    pYourPMText = ProfileSetting
                
                Case "ptheirpmtext"
                    pTheirPMText = ProfileSetting
                
                Case "pprofilemsgbox"
                    pProfileMSGBOX = ProfileSetting
                                
                Case "pprofileurl"
                    pProfileURL = ProfileSetting
                                             
                Case "pnotifytext"
                    pNotifyText = ProfileSetting
                    
                Case "phelptextn"
                    pHelpTextN = ProfileSetting
                    
                Case "phelptextd"
                    pHelpTextD = ProfileSetting
                    
                Case "pothertext"
                    pOtherText = ProfileSetting
                    
                Case "pentrytext"
                    pEntryText = ProfileSetting
                    
                Case "pyourchattext"
                    pYourChatText = ProfileSetting
                    
                Case "ptheirchattext"
                    pTheirChatText = ProfileSetting
                    
                Case "pwindowback"
                    pWindowBack = ProfileSetting
                    
                Case "pchattextback"
                    pChatTextBack = ProfileSetting
                    
                Case "pentrytextback"
                    pEntryTextBack = ProfileSetting
                    
                Case "ptitletextnorm"
                    pTitleTextNorm = ProfileSetting
                    
                Case "ptitletextconnected"
                    pTitleTextConnected = ProfileSetting
                    
                Case "phandel"
                    pHandel = ProfileSetting
                    
                Case "pchatsep"
                    pChatSep = ProfileSetting
                    
                Case "pchannel"
                    pChannel = ProfileSetting
                    
                Case "pprofilemessage"
                    pProfileMessage = ProfileSetting
                    
                Case "ptransparency"
                    pTransparency = ProfileSetting
                    
                Case "pfont"
                    pFont = ProfileSetting
                    
                Case "ppopupwindowonmessage"
                    pPopUpWindowOnMessage = ProfileSetting
                    
                Case "palwaysontop"
                    pAlwaysOnTop = ProfileSetting
                    
                Case "ppicturepath"
                    pPicturePath = ProfileSetting
                      
                Case "pdevelopermode"
                    pDeveloperMode = ProfileSetting
                
                Case "pbindf1"
                    pBindF1 = ProfileSetting
                
                Case "pbindf2"
                    pBindF2 = ProfileSetting
                
                Case "pbindf3"
                    pBindF3 = ProfileSetting
                
                Case "pbindf4"
                    pBindF4 = ProfileSetting
                
                Case "pbindf5"
                    pBindF5 = ProfileSetting
                
                Case "pbindf6"
                    pBindF6 = ProfileSetting
                
                Case "pbindf7"
                    pBindF7 = ProfileSetting
                
                Case "pbindf8"
                    pBindF8 = ProfileSetting
                
                Case "pbindf9"
                    pBindF9 = ProfileSetting
                
                Case "pbindf10"
                    pBindF10 = ProfileSetting
                
                Case "pbindf11"
                    pBindF11 = ProfileSetting
                
                Case "pbindf12"
                    pBindF12 = ProfileSetting
                
                Case "ppassword"
                    pPassword = ProfileSetting
                
                Case "urltext"
                    pURLText = ProfileSetting
                
                Case Else
                    ADDTEXT "Unknown Profile Item: " & ProfileName & ":" & ProfileSetting

End Select

End Sub
