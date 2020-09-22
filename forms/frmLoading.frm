VERSION 5.00
Begin VB.Form frmLoading 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2760
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoading.frx":0000
   ScaleHeight     =   2760
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.FileListBox FILELISTING_USED_TO_VALIDATA_FOLDERS 
      Height          =   1455
      Left            =   5040
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
Me.Show
DoEvents
Pause 1000
DoEvents

If DirEXISTS(APPPATH & "skins") = False Then
    MkDir APPPATH & "skins"
End If

Select Case OSVersion
    Case "Windows 95/ME"
        
    Case "Windows 98/ME"

    Case "Windows NT 4.0/5.0"

End Select

AlwaysOnTop Me, True
   Load frmMain
   frmMain.Show
   DoEvents
Me.Hide

End Sub

Private Sub tmrPAUSE_Timer()
    tmrPAUSE.Enabled = False
End Sub
