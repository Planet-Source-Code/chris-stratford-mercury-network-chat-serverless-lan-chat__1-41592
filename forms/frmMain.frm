VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form2"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9615
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5220
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Tag             =   """"
   Begin VB.ListBox lstUsers 
      Height          =   450
      Left            =   1560
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox lstBarredChannels 
      Height          =   2010
      ItemData        =   "frmMain.frx":030A
      Left            =   1560
      List            =   "frmMain.frx":0323
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox lstHelp 
      Height          =   4545
      ItemData        =   "frmMain.frx":0341
      Left            =   5880
      List            =   "frmMain.frx":0495
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSComctlLib.ProgressBar SpamBar 
      Height          =   105
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Max             =   500
      Scrolling       =   1
   End
   Begin VB.ListBox lstMainCommands 
      Height          =   4545
      ItemData        =   "frmMain.frx":0D68
      Left            =   3960
      List            =   "frmMain.frx":0DED
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer tmrAntiSpam 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3480
      Top             =   480
   End
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   0
      Picture         =   "frmMain.frx":0F93
      ScaleHeight     =   675
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3480
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstBannedUsers 
      Height          =   2010
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   120
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   0
      Picture         =   "frmMain.frx":2570
      ScaleHeight     =   675
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   4815
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3351
            MinWidth        =   3351
            Picture         =   "frmMain.frx":3B4D
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8705
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Total Bytes Received"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "8:14 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3201
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":49EE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.ComboBox txtEntry 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
      VariousPropertyBits=   746604571
      MaxLength       =   1000
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2566;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSaveChat 
         Caption         =   "Save Chat..."
      End
      Begin VB.Menu mnuPrintChat 
         Caption         =   "Print Chat..."
      End
      Begin VB.Menu sepppe 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuProfiles 
      Caption         =   "Profiles"
      Begin VB.Menu mnuNewProfile 
         Caption         =   "New..."
      End
      Begin VB.Menu mnuLoadProfile 
         Caption         =   "Load..."
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDefaultProfile 
         Caption         =   "Default Profile"
      End
      Begin VB.Menu mnuDwldProfiles 
         Caption         =   "Download New Profiles"
      End
   End
   Begin VB.Menu mnuMainChannel 
      Caption         =   "Channel"
      Begin VB.Menu mnuOPENCHAT 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuChannelPrivate 
         Caption         =   "Private Room"
      End
      Begin VB.Menu SEP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Popular Channels"
         Begin VB.Menu mnuChannel 
            Caption         =   "Open#2"
            Index           =   0
         End
         Begin VB.Menu mnuChannel 
            Caption         =   "Open#3"
            Index           =   1
         End
         Begin VB.Menu mnuChannel 
            Caption         =   "DesignerChat"
            Index           =   2
         End
         Begin VB.Menu mnuChannel 
            Caption         =   "SDD"
            Index           =   3
         End
         Begin VB.Menu mnuChannel 
            Caption         =   "IPT"
            Index           =   4
         End
         Begin VB.Menu mnuChannel 
            Caption         =   "Brain Storming"
            Index           =   5
         End
         Begin VB.Menu mnuChannel 
            Caption         =   "QuickTalk"
            Index           =   6
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpList 
         Caption         =   "Help Listing"
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "View Website"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "Check For Updates"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuClearText 
         Caption         =   "Clear Text"
      End
      Begin VB.Menu sep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveText 
         Caption         =   "Save Text"
      End
      Begin VB.Menu mnuPrintText 
         Caption         =   "Print Text"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Admin"
      Begin VB.Menu mnuRemPass 
         Caption         =   "Remove Password"
      End
      Begin VB.Menu mnuFullView 
         Caption         =   "View Full Message"
      End
      Begin VB.Menu AdminRcon 
         Caption         =   "RCon"
         Begin VB.Menu mnuMessage 
            Caption         =   "Global Message Box"
         End
         Begin VB.Menu mnuIndividualMsg 
            Caption         =   "Individual Message Box"
         End
         Begin VB.Menu mnuCloseClient 
            Caption         =   "Close Client"
         End
         Begin VB.Menu mnuEnterComRemote 
            Caption         =   "Enter Remote Command"
         End
      End
      Begin VB.Menu Seppppp12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout Admin"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Picture_Inserted As Boolean
Dim SpamCOunt As Integer
Dim CtrlDwn As Boolean
Dim C As String 'to store current form's caption
Dim CO As Integer 'to store caption length
Dim FS As Long 'to store current form Width







Public Sub Load_Welcome_Note()
Picture_Inserted = False
On Error Resume Next

If LCase(pPicturePath) <> "(none)" Then
        'Adds the image to the Chat Text
        txtChat.Locked = False
        Clipboard.Clear
        Clipboard.SetData Picture1.Picture
        txtChat.SetFocus
        SendKeys "^V~"
        DoEvents
        txtChat.Locked = True
End If

    Picture_Inserted = True
ADDTEXT "Mercury Network Chat: For Use On LAN's & The Internet", pNormalText
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
ADDTEXT "For Help Type: \HELP or goto HELP in the menu!", pNormalText
ADDTEXT "", pNormalText
    Me.Show
    txtEntry.SetFocus

DoEvents


End Sub

Private Sub Form_Load()
Me.Show
DoEvents
COMMS False
DoEvents
    LOAD_Profile "default"
    'LOAD_Profile "test"
    DoEvents
    Pause 50
    SD "connected", Time
COMMS True


For i = 0 To lstMainCommands.ListCount
    txtEntry.AddItem lstMainCommands.List(i)
Next

    If pTransparency <> 0 Then
        SETTrans pTransparency
    End If
DoEvents

mnuAdmin.Visible = False
End Sub

Private Sub mnuIndividualMsg_Click()
    frmRemoteMsgBox.Show
End Sub

Private Sub txtchat_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
    On Error Resume Next
    hyperlink = getHyperlink(X, Y, txtChat)
End Sub


Private Sub txtchat_MouseUp(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
    On Error Resume Next


    If Button = vbLeftButton Then

        If Len(hyperlink) > 0 Then
            ShellExecute Me.hwnd, "Open", hyperlink, vbNullString, vbNullString, _
            vbHide
        End If
    End If
End Sub

Private Sub COMMS(iSetting As Boolean)
    Me.mnuMainChannel.Enabled = iSetting
    Me.mnuFile.Enabled = iSetting
    Me.mnuProfiles.Enabled = iSetting
    Me.mnuHelp.Enabled = iSetting
    Me.mnuAdmin.Enabled = iSetting
End Sub

Private Sub Form_Resize()
On Error Resume Next

If LOCKDOWN = True And pPassword <> "ÜÞÜÊ" Then
Me.Hide
    If InBox("Password", "Enter Profile Password", , "*") <> Decode(pPassword, pProfileName) Then
        Me.Show
        Me.WindowState = 1
        LOCKDOWN = True
    Else
        LOCKDOWN = False
        Me.WindowState = 0
        Me.Show
        
        
    End If
End If

'Restrict The Resize
If Me.WindowState <> 1 Then
    If Me.Width < 5000 Then Me.Width = 5000
    If Me.Height < 4000 Then Me.Height = 4000
    
'Resize the objects on the form when the window is
'Resized
    txtChat.Left = 50
    txtEntry.Left = 50
    txtChat.Top = 50
    
    txtChat.Width = Me.Width - 200
    txtChat.Height = Me.Height - txtEntry.Height - 1250

    txtEntry.Width = Me.Width - 200
    txtEntry.Top = Me.Height - txtEntry.Height - 1220
   
    txtEntry.SetFocus
    
    SpamBar.Top = txtEntry.Top + txtEntry.Height + 10
    SpamBar.Width = txtEntry.Width
    SpamBar.Left = txtEntry.Left

Else
    If CtrlDwn = True Then
        'This Means Its Been Locked!
        LOCKDOWN = True
    Else
        LOCKDOWN = False
    End If
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    SD "quitting"
    End
End Sub

Private Sub mnuAbout_Click()
ENTERCOMMAND "\about"
End Sub

Private Sub mnuChannel_Click(Index As Integer)
Select Case Index
    Case 0: CHANNEL mnuChannel.Item(Index).Caption
    Case 1: CHANNEL mnuChannel.Item(Index).Caption
    Case 2: CHANNEL mnuChannel.Item(Index).Caption
    Case 3: CHANNEL mnuChannel.Item(Index).Caption
    Case 4: CHANNEL mnuChannel.Item(Index).Caption
    Case 5: CHANNEL mnuChannel.Item(Index).Caption
    Case 6: CHANNEL mnuChannel.Item(Index).Caption
    Case 7: CHANNEL mnuChannel.Item(Index).Caption
    Case 8: CHANNEL mnuChannel.Item(Index).Caption
    Case 9: CHANNEL mnuChannel.Item(Index).Caption

End Select
End Sub

Private Sub mnuChannelPrivate_Click()
Dim NewChan As String
NewChan = InBox("Private Room", "Enter the Private Room Key")
If NewChan = "" Then Exit Sub


        SD "leaving"
    ADDTEXT "   |You Have Left '" & pChannel & "'", pNotifyText
        Pause 50
    pChannel = "::" & NewChan
    ADDTEXT "   |You Have Joined Private Channel - " & Mid(pChannel, 3), pNotifyText
    SD "joining"
        Pause 50
End Sub

Private Sub mnuCheck_Click()
ENTERCOMMAND "\check"
End Sub

Private Sub mnuClearText_Click()
    txtChat.Text = ""
End Sub

Private Sub mnuCloseClient_Click()
    frmAdminKick.Show
End Sub

Private Sub mnuDefaultProfile_Click()
    LOAD_Profile "default"
End Sub

Private Sub mnuDwldProfiles_Click()
    Load frmDownloadProfiles
    frmDownloadProfiles.Show
End Sub

Private Sub mnuEnterComRemote_Click()
frmRemCon.Show

End Sub

Private Sub mnuFullView_Click()

Select Case ViewFullMessage
    Case True
        ENTERCOMMAND "\fullmsg 0"
        mnuFullView.Checked = False
    Case False
        ENTERCOMMAND "\fullmsg 1"
        mnuFullView.Checked = True
End Select

End Sub

Private Sub mnuHelpList_Click()
ENTERCOMMAND "\help"
End Sub

Private Sub mnuLoadProfile_Click()
                    frmNewProfile.CD.Filter = "Profile Files(*.prof)|*.prof"
                    frmNewProfile.CD.ShowOpen
                    If frmNewProfile.CD.FileName <> "" Then LOAD_Profile frmNewProfile.CD.FileName
End Sub

Private Sub mnuLogout_Click()
ENTERCOMMAND "\devmode 0"

End Sub

Private Sub mnuMessage_Click()
Dim iTitle As String
Dim iMessage As String
Dim iType As String

iTitle = InBox("Message Box", "Enter The Message Box Title")
iMessage = InBox("Message Box", "Enter The Message Box Message")
iType = InBox("Message Box", "Type: 1, 2 or 3 (OK, Abort Retry Cancel, Yes No)")


Select Case iType
    Case "1"
        SD "msg1", iMessage & "|" & iTitle
    Case "2"
        SD "msg2", iMessage & "|" & iTitle
    Case "3"
        SD "msg3", iMessage & "|" & iTitle
    Case Else
        SD "msg1", iMessage & "|" & iTitle
End Select
End Sub

Public Sub mnuNewProfile_Click()
    Load frmNewProfile
    frmNewProfile.Show
    Me.Hide
End Sub

Private Sub mnuOPENCHAT_Click()
ENTERCOMMAND "\channel open"
End Sub


Private Sub mnuPB_Click()
    Load frmPluginManager
    frmPluginManager.Show
End Sub


Private Sub mnuPrintChat_Click()
CD.ShowPrinter
VB.Printer.Copies = CD.Copies
If MsgBox("Are You Sure You Want To Print This Chat Session " & CD.Copies & " time(s)?", vbYesNo + vbQuestion, "Print") = vbYes Then
    VB.Printer.Print txtChat.Text
        DoEvents
    VB.Printer.EndDoc
Else

End If
End Sub

Private Sub mnuPrintText_Click()
    mnuPrintChat_Click
End Sub

Private Sub mnuProfileBrowser_Click()
frmBrowser.Show
End Sub

Public Sub mnuQuit_Click()
SD "quitting"
DoEvents
End
End Sub

Private Sub SpamBar2_Click()

End Sub

Private Sub mnuRemPass_Click()
ENTERCOMMAND "\rempass"

End Sub

Private Sub mnuSaveChat_Click()
CD.Filter = "Word Document (*.doc)|*.doc|Text File (*.txt)|*.txt|All Files (*.*)|*.*"
CD.ShowSave
If CD.FileName <> "" Then
Close #9
    Open CD.FileName For Output As #9
        If Right(CD.FileName, 3) = "doc" Then
            Print #9, txtChat.TextRTF
        Else
            Print #9, txtChat.Text
        End If
     Close #9
End If
End Sub

Private Sub mnuSaveText_Click()
    mnuSaveChat_Click
End Sub




Private Sub tmrAntiSpam_Timer()
SpamCOunt = SpamCOunt + 25
SpamBar.Value = 500 - SpamCOunt
If SpamCOunt = 500 Then tmrAntiSpam.Enabled = False: SpamCOunt = 0
End Sub



Private Sub txtChat_KeyPress(KeyAscii As Integer)
        txtChat.SelStart = Len(txtChat)
If Picture_Inserted = True Then
    txtEntry.SetFocus
    SendKeys Chr(KeyAscii)
    KeyAscii = 0
End If

End Sub


Private Sub txtChat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuPopup, , , , mnuClearText
End If
End Sub

Private Sub txtEntry_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

Dim TEMPTEXT As String
Select Case KeyCode
    Case 80:
        If CtrlDwn = True Then
            Load frmPrivateMsg
            frmPrivateMsg.Show
            KeyCode = 0
            CtrlDwn = False
        End If
    Case 77:
        If CtrlDwn = True Then
            Me.WindowState = 1
            KeyCode = 0
        End If
    Case 112:
        If pBindF1 <> "" Then
        TEMPTEXT = pBindF1
        End If
        
    Case 113:
        If pBindF2 <> "" Then
        TEMPTEXT = pBindF2
        End If
        
    Case 114:
        If pBindF3 <> "" Then
        TEMPTEXT = pBindF3
        End If
        
    Case 115:
        If pBindF4 <> "" Then
        TEMPTEXT = pBindF4
        End If
        
    Case 116:
        If pBindF5 <> "" Then
        TEMPTEXT = pBindF5
        End If
        
    Case 117:
        If pBindF6 <> "" Then
        TEMPTEXT = pBindF6
        End If
        
    Case 118:
        If pBindF7 <> "" Then
        TEMPTEXT = pBindF7
        End If
        
    Case 119:
        If pBindF8 <> "" Then
        TEMPTEXT = pBindF8
        End If
        
    Case 120:
        If pBindF9 <> "" Then
        TEMPTEXT = pBindF9
        End If
        
    Case 121:
        If pBindF10 <> "" Then
        TEMPTEXT = pBindF10
        End If
        
    Case 122:
        If pBindF11 <> "" Then
        TEMPTEXT = pBindF11
        End If
        
    Case 123:
        If pBindF12 <> "" Then
        TEMPTEXT = pBindF12
        End If
            
    Case 17: CtrlDwn = True
    Case 9: txtEntry.SelStart = Len(txtEntry): KeyCode = 0
    Case 13:
    'The Commenetd Out code is what i coul dhave to allow multi line
    'messages
    'But i dont like the look with them.
    'I left the code to show what it could look like.
    
        'If CtrlDwn = False Then
            If txtEntry <> "" Then ENTERCOMMAND Trim(txtEntry.Text): txtEntry.Text = "": KeyCode = 0
        'Else
        '    txtEntry = txtEntry & vbCrLf
        '    txtEntry.SelStart = Len(txtEntry)
        'End If
    Case Else
        
End Select

Select Case KeyCode
    Case 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123
    KeyCode = 0
        If TEMPTEXT <> "" Then
            ENTERCOMMAND TEMPTEXT
        End If
End Select
End Sub

Private Sub txtEntry_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 17 Then CtrlDwn = False
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
Dim iData As String
WS.GetData iData
'Decodes Data
'The EncodeKey will change when major version changes have been made
'This will eliminate old versions And force updates
SubmitData Decode(iData, iENCODEKEY)

On Error GoTo Clear_Bytes
TotalBytes = TotalBytes + bytesTotal

Me.StatusBar1.Panels(3).Text = "B: " & TotalBytes

Exit Sub

Clear_Bytes:
'This error is probably becuase the limit to a LONG has been reached
    ADDTEXT "You Have Recieved: " & Str(TotalBytes), vbNormal
    ADDTEXT "This Number Has Been Reset!", vbNormal
    TotalBytes = 0
End Sub

