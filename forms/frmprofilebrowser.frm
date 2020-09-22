VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile Browser"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProfileBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstPath 
      Height          =   1035
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Use This Profile"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3720
      Width           =   1815
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5655
      BackColor       =   14737632
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "9975;5741"
      MatchEntry      =   0
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Profile From List:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
LOAD_Profile lstPath.List(ListBox1.ListIndex)
    frmMain.Show
    Unload Me
End Sub

Private Sub Command2_Click()
If MsgBox("Load Default Profile?", vbQuestion + vbYesNo, "Profile Browser") = vbYes Then
    LOAD_Profile "default"
    frmMain.Show
    Unload Me
Else
    frmMain.Show
    Unload Me
End If

End Sub

Private Sub Form_Load()
    
Me.BackColor = pWindowBack
Label1.ForeColor = pNormalText
ListBox1.BackColor = pEntryTextBack
ListBox1.ForeColor = pEntryText

    
    Me.ListBox1.Clear
    lstPath.Clear
    frmMain.Hide
    Load_Profile_List
End Sub

Private Sub Load_Profile_List()
Dim iPATH As String
Dim LineIN As String
Dim iCom As String
Dim iSet As String

'Set a huge loop so you can check all these directories if they exist.
'Easy To Add Directories, Just Add It To The List Below

For j = 0 To 3
    
    Select Case j
        Case 0:  iPATH = APPPATH
        Case 1:  iPATH = APPPATH & "profiles\"
        Case 2:  iPATH = APPPATH & "profile\"
        Case 3:  iPATH = APPPATH & "skins\"
    End Select
    
    If DirEXISTS(iPATH) = True Then
        File1.Path = iPATH
        File1.Refresh
        
            For i = 0 To File1.ListCount
                If Right(File1.List(i), 4) = "prof" Then
                Close #1
                    Open iPATH & File1.List(i) For Input As #1
                        Do While Not EOF(1)
                            Line Input #1, LineIN
                            If LineIN <> "" Then
                                iCom = Mid(LineIN, 1, InStr(LineIN, "=") - 1)
                                iSet = Mid(LineIN, InStr(LineIN, "=") + 1)
                                    Select Case LCase(iCom)
                                        Case "pprofilename"
                                            ListBox1.AddItem iSet
                                            Me.lstPath.AddItem iPATH & File1.List(i)
                                    End Select
                            End If
                        Loop
                    Close #1
                    
                End If
             Next
    
    Else
    
    End If
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub ListBox1_Click()
If ListBox1.ListIndex >= 0 Then
    Command1.Enabled = True
End If
End Sub

Private Sub ListBox1_DblClick(Cancel As MSForms.ReturnBoolean)
LOAD_Profile lstPath.List(ListBox1.ListIndex)
    frmMain.Show
    Unload Me
End Sub
