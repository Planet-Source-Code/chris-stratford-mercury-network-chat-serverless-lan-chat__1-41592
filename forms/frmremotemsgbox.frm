VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRemoteMsgBox 
   Caption         =   "Remote Message Box"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5953;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   165
      Width           =   855
   End
End
Attribute VB_Name = "frmRemoteMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
If ComboBox1.Text <> "" Then
    Command1.Item(0).Enabled = True
Else
    Command1.Item(0).Enabled = False
End If

End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        If txtMsg(0).Text <> "" And txtMsg(1).Text <> "" Then
            SDSpecial "msg2user", txtMsg(1).Text, txtMsg(0).Text, ComboBox1.Text
            Command1.Item(2).Caption = "Close"
        End If
        txtMsg(0).SetFocus
    Case 1
        ComboBox1.ListIndex = -1
        txtMsg(0).Text = ""
        txtMsg(1).Text = ""
        txtMsg(0).SetFocus
    Case 2
        Unload Me
End Select
End Sub

Private Sub Form_Load()
Me.Icon = frmMain.Icon
ADDTEXT "Please Hold... Loading User List...", pNormalText, , True
AlwaysOnTop Me, True
Me.ComboBox1.Clear
frmMain.lstUsers.Clear
'TEST DATA
    'ComboBox1.AddItem "USER11"
    'ComboBox1.AddItem "USEasdR"
    'ComboBox1.AddItem "USdsfER!"
    'ComboBox1.AddItem "12312USER!"

PrivateMessageRequest = True
    SD "list", "all"
Pause 1000
    For i = 0 To frmMain.lstUsers.ListCount
        ComboBox1.AddItem frmMain.lstUsers.List(i)
    Next
    
Me.BackColor = pChatTextBack
Label1.ForeColor = pNormalText
Label2.ForeColor = pNormalText
Label3.ForeColor = pNormalText
txtMsg(0).BackColor = pEntryTextBack
txtMsg(0).ForeColor = pEntryText
txtMsg(1).BackColor = pEntryTextBack
txtMsg(1).ForeColor = pEntryText
ComboBox1.BackColor = pEntryTextBack
ComboBox1.ForeColor = pEntryText
Me.Show


End Sub

