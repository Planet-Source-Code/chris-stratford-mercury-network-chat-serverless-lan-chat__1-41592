VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Checking For Updates"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet DL 
      Left            =   3720
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Goto Profile Explorer"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3960
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Close Info"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Print Info"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Save Info"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3600
      Width           =   5535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "View Release Notes"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Goto Website"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Download Now"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1785
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Now"
      Default         =   -1  'True
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   2480
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   5760
      X2              =   0
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   5760
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   5760
      X2              =   0
      Y1              =   3050
      Y2              =   3050
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   5760
      Y1              =   3050
      Y2              =   3050
   End
End
Attribute VB_Name = "frmUpdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim eMESSAGE As String
Dim eURL As String
Dim eURL2 As String

Private Sub Command1_Click()
PB 0
    List1.Clear
    Text1.Text = ""
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command9.Enabled = False
PB 5
Command2.Caption = "Close"
Command1.Enabled = False
List1.AddItem "Connecting To Webarchive"
PB 15
Close #6
Open "c:\tempdata.data" For Output As #6
    On Error GoTo ERRORWeb
    List1.AddItem "Downloading Information File"
    Print #6, Inet1.OpenURL("http://www.neester.com/sif/mercurychat/sifmain.sif")
Close #6
PB 25
List1.AddItem "Analysing Retrieved Data"
CheckFile

Command1.Caption = "Refresh"
Command1.Enabled = True

Exit Sub
ERRORWeb:
    Close #6
    List1.AddItem "Error: Couldn't Retreive The Host!"
End Sub

Private Sub CheckFile()
eMESSAGE = ""

Dim LineIN As String
Close #4
Open "C:\tempdata.data" For Input As #4
PB 50
Do While Not EOF(4)
        Line Input #4, LineIN
        PB 70
        CheckNow LineIN
    Loop
Close #4
PB 90
Kill "C:\tempdata.data"

Text1.Text = Text1.Text & "Downloaded At: " & Now & vbCrLf & "Downloaded From: Mercury Network Server" & vbCrLf & "******************************************"
PB 100
End Sub

Private Sub CheckNow(iData As String)
Dim iCommand As String
Dim iEXTRA As String

If InStr(1, iData, "©") < 1 Then
    iData = "msg©" & iData
End If
iCommand = LCase(Mid(iData, 1, InStr(1, iData, "©") - 1))
iEXTRA = Mid(iData, InStr(1, iData, "©") + 1)
Select Case iCommand
Case "version"
        List1.AddItem "Your Version: " & VersionCheck
        List1.AddItem "Newest Version: " & iEXTRA
        Command9.Enabled = True
    If iEXTRA <> VersionCheck Then
        List1.AddItem "Click To Download The New Version!"
        Command3.Enabled = True
    Else
        List1.AddItem "You Have The Current Version!"
    End If
Case "downloadurl"
    eURL = iEXTRA

Case "websiteurl"
    eURL2 = iEXTRA
    Command4.Enabled = True
Case "msg"
    Text1 = Text1 & iEXTRA & vbCrLf
    Command5.Enabled = True

End Select
End Sub


Private Sub Command2_Click()
frmMain.Show
Unload Me
End Sub

Private Sub Command3_Click()
'OpenWeb eURL
Dim byteData() As Byte
Dim intFile As Integer
 

intFile = FreeFile()

byteData() = DL.OpenURL(eURL, icByteArray)

Open "C:\temp.exe" For Binary Access Write As #intFile
    Put #intFile, , byteData()
Close #intFile

MsgBox "Download Complete!", vbInformation, "Update!"

Name "C:\temp.exe" As APPPATH & Mid(eURL, InStrRev(eURL, "/") + 1)

MsgBox "Restaring Mercury Chat! Please Wait!", vbInformation, "New Update Installed!"
SD "restarting"
Shell APPPATH & Mid(eURL, InStrRev(eURL, "/") + 1)
End

End Sub

Private Sub Command4_Click()
OpenWeb eURL2
End Sub

Private Sub Command5_Click()
Me.Height = 7440
Me.Show
End Sub

Private Sub Command6_Click()
frmMain.CD.Filter = "Text Files (*.txt)|*.txt"
frmMain.CD.ShowSave

If frmMain.CD.FileName <> "" Then
Close #2
    Open frmMain.CD.FileName For Output As #2
        Print #2, Text1.Text
    Close #2
    MsgBox "Info Has Been Saved!", vbInformation, "Saved"
End If

End Sub
Private Sub PB(iVal As Integer)
ProgressBar1.Value = iVal
End Sub

Private Sub Command7_Click()
frmMain.CD.ShowPrinter
VB.Printer.Copies = frmMain.CD.Copies
    VB.Printer.Print Text1.Text
    DoEvents
    VB.Printer.EndDoc
    MsgBox "Info Has Been Printed!", vbInformation, "Printed"
End Sub

Private Sub Command8_Click()
Me.Height = 3450
Me.Show
End Sub

Private Sub Command9_Click()
frmDownloadProfiles.Show
Unload Me
End Sub

Private Sub Form_Load()
    Me.BackColor = pWindowBack
    List1.BackColor = pEntryTextBack
    List1.ForeColor = pEntryText
    Text1.BackColor = pEntryTextBack
    Text1.ForeColor = pEntryText

frmMain.Hide
List1.Clear

Me.Height = 3450
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
Unload Me
End Sub
