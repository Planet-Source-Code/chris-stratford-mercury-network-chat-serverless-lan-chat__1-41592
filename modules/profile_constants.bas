Attribute VB_Name = "Profile_Constants"
'Name
    Global pProfileName As String   'The Name

'Defines The Global Colours
    Global pNormalText As Long      'Normal Window Dialog
    Global pListText As Long        'When A List Is Presented
    Global pErrorText As Long       'When An Error Occurs
    Global pNotifyText As Long      'When A User Comes/Goes
    Global pHelpTextN As Long       'The "HELP" Text For The Command
    Global pHelpTextD As Long       'The "HELP" Text For The Definition
    Global pOtherText As Long       'Any Other Text
    Global pEntryText As Long       'The Entry Textbox Forecolour
    Global pYourChatText As Long    'Your Chat Text Colour
    Global pTheirChatText As Long   'Their Chat Text Colour
    Global pYourPMText As Long      'Your Private Message Colour
    Global pTheirPMText As Long     'Their Private Message Colour
    Global pURLText As Long         'URL Text Colour

'Defines The Text & Form Backgrounds
    Global pWindowBack As Long      'Window Backcolour
    Global pChatTextBack As Long    'The Main Text Backcolour
    Global pEntryTextBack As Long   'The Entry Text Backcolour

'Defines Text Variables
    Global pTitleTextNorm As String 'The Window Title Bar Text
    Global pTitleTextConnected As String    'The Window Title Bar Text (When Connected)
    Global pHandel As String        'Your Nickname
    Global pChatSep As String       'The Seperator On The Chat
    Global pProfileMessage As String        'The Profiles 1Line Message

'Defines Interger Values
    Global pTransparency As Integer 'The Transparency Level

'Defines Fonts
    Global pFont As String          'The Font For The Text
    
'Defines Other Settings
    Global pPopUpWindowOnMessage As Boolean 'If The Window Pops Up When Msg Recieved
    Global pAlwaysOnTop As Boolean  'If The Window Is Always Ontop
    Global pChannel As String       'The Current Channel
    Global pDeveloperMode As Boolean
    Global pBanListPath As String   'The Path To THe Ban List
    Global pPassword As String      'The Password
    Global pProfileMSGBOX           'If The Profile Has A Message Box
    Global pProfileURL              'Profile URL Home Page

'Defines Images To use
    Global pPicturePath As String   'If The User Wants His/Her Own Image

'Defines TEXT Binds To Function Keys
    Global pBindF1 As String
    Global pBindF2 As String
    Global pBindF3 As String
    Global pBindF4 As String
    Global pBindF5 As String
    Global pBindF6 As String
    Global pBindF7 As String
    Global pBindF8 As String
    Global pBindF9 As String
    Global pBindF10 As String
    Global pBindF11 As String
    Global pBindF12 As String


Public Sub Load_TestStyle()
'This is a test style, Its just a test that all the commands are working
'And they all can change
            pProfileName = "Test Style"
             pNormalText = vbWhite
               pListText = vbRed
              pErrorText = &H8000&
             pNotifyText = vbRed
              pHelpTextN = vbBlue
              pHelpTextD = vbMagenta
              pOtherText = vbYellow
              pEntryText = vbWhite
           pYourChatText = &HFF8080
          pTheirChatText = &H80FF&
             pYourPMText = &HFF8080
            pTheirPMText = &H80FF&
             pWindowBack = vbBlack
           pChatTextBack = vbBlack
          pEntryTextBack = vbBlack
          pTitleTextNorm = "Mercury Chat - Test Styllllle!"
     pTitleTextConnected = "Mercury Chat - Test Styllllle!"
                 pHandel = "Test StyLE!"
                pChatSep = ":-:"
                pChannel = "open - Test"
         pProfileMessage = "Welcome To Mercury Netwrr.k C.at ahssdoasdjds"
           pTransparency = 15
                   pFont = "arial"
          pProfileMSGBOX = "TEST STYLE"
             pProfileURL = "http://www.teststylerules.com"
   pPopUpWindowOnMessage = False
            pAlwaysOnTop = True
            pPicturePath = "C:\Documents and Settings\chris\My Documents\My Pictures\sample.jpg"
          pDeveloperMode = True
                 pBindF1 = "\help"
                 pBindF2 = "\quit"
                 pBindF3 = "\trans 99"
                 pBindF4 = "\trans 0"
                 pBindF5 = "\noooo"
                 pBindF6 = "\channel 69er"
                 pBindF7 = "\channel H4KaZ"
                 pBindF8 = "HEY MAN"
                 pBindF9 = "GTG EVERY 1"
                pBindF10 = "=)"
                pBindF11 = "=("
                pBindF12 = "=|"
            pBanListPath = "C:\Program Files\Microsoft Visual Studio\VB98\Projects\SDD Year 11 Major\22-6-02\testprof.ban"
               pPassword = Encode("dude", "Test Style")

               

'No Banned Users By Default
frmMain.lstBannedUsers.Clear
          
End Sub


Public Sub Load_Defaults()
'This Is Where The Default Settings Are Set
'These Are Typically The Same That They Are Now
'This Is Here So That You Can Switch Between Profiles
'And Still Come Back To Default
'Defines The Global Colours

            pProfileName = "Default"
             pNormalText = vbBlack
               pListText = vbBlue
              pErrorText = vbRed
             pNotifyText = &H8000&
              pHelpTextN = &HC000C0
              pHelpTextD = &H404040
              pOtherText = &H404080
              pEntryText = vbBlack
           pYourChatText = &HFF8080
          pTheirChatText = &H80FF&
             pWindowBack = &H8000000F
           pChatTextBack = vbWhite
          pEntryTextBack = vbWhite
          pTitleTextNorm = "Mercury Chat"
     pTitleTextConnected = "Mercury Chat - Connected!"
                 pHandel = ""
                pChatSep = ": "
             pYourPMText = &H80FF&
          pProfileMSGBOX = ""
             pProfileURL = ""
            pTheirPMText = &H80FF&
                pChannel = "open"
         pProfileMessage = "Welcome To Mercury Network Chat!"
           pTransparency = 0
                   pFont = "verdana"
   pPopUpWindowOnMessage = False
            pAlwaysOnTop = False
            pPicturePath = ""
          pDeveloperMode = False
                 pBindF1 = "\help"
                 pBindF2 = ""
                 pBindF3 = ""
                 pBindF4 = "\quit"
                 pBindF5 = "\trans 0"
                 pBindF6 = "\trans 25"
                 pBindF7 = "\trans 50"
                 pBindF8 = "\trans 75"
                 pBindF9 = "\pm"
                pBindF10 = "\channel 9ine"
                pBindF11 = "\Hey %id% How Are You Man?"
                pBindF12 = "\channel backchat"
            pBanListPath = APPPATH & "default.ban"
               pPassword = Encode("none", "none")

           
       
'Make Sure Local Default Ban List is There
    On Error GoTo NoFile
    Close #1
        Open APPPATH & "default.ban" For Input As #1
    Close #1
    
Exit Sub

NoFile:
'Error When No File Is Present
Close #1
    Open APPPATH & "default.ban" For Output As #1
        Print #1, ""
    Close #1
End Sub

