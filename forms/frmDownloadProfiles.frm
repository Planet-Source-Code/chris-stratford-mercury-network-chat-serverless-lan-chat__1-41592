VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmDownloadProfiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download Profiles"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDownloadProfiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblInfo 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2400
      Width           =   6255
   End
   Begin VB.ListBox List3 
      Height          =   645
      Left            =   5160
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Download Now"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   5760
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Retrieve Profile Archive"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   3360
      Width           =   2295
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDownloadProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.Clear
List2.Clear
List3.Clear
Command3.Enabled = False

Dim iFF As Integer
iFF = FreeFile

Close iFF
Open "C:\tempprofiledata.data" For Output As iFF
    Command1.Enabled = False
    On Error Resume Next
    Print #iFF, Inet1.OpenURL("http://www.neester.com/sif/mercurychat/sifmain.sif")
    
Close iFF

Dim LineIN As String

iFF = FreeFile

Close iFF

Open "C:\tempprofiledata.data" For Input As iFF
    Do While Not EOF(iFF)
        Line Input #iFF, LineIN
        CheckData LineIN
    Loop
Close iFF

On Error Resume Next
Kill "C:\tempprofiledata.data"
Command1.Enabled = True
End Sub

Private Sub CheckData(iData As String)

Dim iCom As String
Dim iSet As String
Dim TempURL As String
On Error Resume Next

iCom = LCase(Mid(iData, 1, InStr(1, iData, "©") - 1))
iSet = Mid(iData, InStr(1, iData, "©") + 1)

Select Case iCom
    Case "profile"
        List1.AddItem " ·  " & Mid(iSet, 1, InStr(iSet, "|") - 1)
        TempURL = Mid(iSet, InStr(iSet, "|") + 1)
        List2.AddItem Mid(TempURL, 1, InStr(TempURL, "[") - 1)
        List3.AddItem Mid(TempURL, InStr(TempURL, "[") + 1)
    Case Else
End Select
End Sub

Private Sub Command2_Click()
frmMain.Show
Unload Me
End Sub

Private Sub Command3_Click()
List1_DblClick
End Sub

Private Sub Form_Load()
Me.BackColor = pWindowBack
List1.BackColor = pEntryTextBack
List1.ForeColor = pEntryText
lblInfo.BackColor = pEntryTextBack
lblInfo.ForeColor = pEntryText
    frmMain.Hide

End Sub

Private Sub Form_Terminate()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
    Unload Me
End Sub

Private Sub List1_Click()
If List1.ListIndex > -1 Then
    Command3.Enabled = True
    lblInfo.Text = List3.List(List1.ListIndex)
    lblInfo.Text = pReplace(lblInfo.Text, "¦", vbCrLf)
    
Else
    Command3.Enabled = False
    lblInfo.Text = ""
End If
End Sub

Private Sub List1_DblClick()
frmMain.CD.Filter = "Profile Files (*.prof)|*.prof"
frmMain.CD.DialogTitle = "Save Profile - " & Mid(List1.List(List1.ListIndex), 4)
frmMain.CD.FileName = Mid(List1.List(List1.ListIndex), 4)
frmMain.CD.InitDir = APPPATH & "skins\"
frmMain.CD.ShowSave

If frmMain.CD.FileName <> "" Then
List2.Visible = False
On Error GoTo ENDNOW
Close #4
    Open frmMain.CD.FileName For Output As #4
        Print #4, Inet1.OpenURL(List2.List(List1.ListIndex))
    Close #4

If MsgBox("Your Profile Has Been Downloaded Sucessfully!" & vbCrLf & "Would You Like To Load It Now?", vbYesNo + vbInformation) = vbYes Then
    LOAD_Profile frmMain.CD.FileName
    frmMain.Show
    Unload Me
End If
End If


Exit Sub
ENDNOW:
    ADDTEXT "Error Loading Profile", pErrorText
End Sub
