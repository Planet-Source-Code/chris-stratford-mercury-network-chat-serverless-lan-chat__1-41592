VERSION 5.00
Begin VB.Form frmInputbox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInputbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4620
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Timer tmrINPUTBOXPAUSE 
      Interval        =   1
      Left            =   4680
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   310
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmInputbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txtInput As Boolean


Private Sub Command1_Click()
    If Text1.Text <> "" Then
        InputTemp = Text1.Text
        txtInput = True
    End If
    

End Sub

Private Sub Command2_Click()
        InputTemp = ""
        txtInput = True
End Sub

Private Sub Form_Load()
Me.BackColor = frmMain.BackColor
Label1.ForeColor = pNormalText
Label1.BackColor = frmMain.BackColor
Text1.BackColor = frmMain.txtEntry.BackColor
Text1.ForeColor = frmMain.txtEntry.ForeColor

If Label1.ForeColor = Me.BackColor Then Label1.BackStyle = 1

Me.Show
    AlwaysOnTop Me, True
    txtInput = False

    AlwaysOnTop Me, True
End Sub

Private Sub Form_Resize()

If Me.WindowState <> 1 Then
    Me.Width = 4740
    Me.Height = 2070
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not txtInput Then Cancel = True
End Sub

Private Sub tmrINPUTBOXPAUSE_Timer()
If txtInput <> True Then
    tmrINPUTBOXPAUSE.Enabled = True
Else
    tmrINPUTBOXPAUSE.Enabled = False
End If
End Sub
