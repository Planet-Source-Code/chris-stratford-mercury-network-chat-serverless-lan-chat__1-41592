VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewProfile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Profile Wizard"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame STEP0 
      Caption         =   "Profile Wizard"
      Height          =   4695
      Left            =   5520
      TabIndex        =   12
      Top             =   5040
      Width           =   6975
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   105
         Top             =   2880
         Width           =   3135
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   240
         Picture         =   "frmNewProfile.frx":030A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   59
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox txtPProfileName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Top             =   2400
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   285
         Left            =   6480
         TabIndex        =   19
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtPFilePath 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Profile Password:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   106
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Profile Name:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   20
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Filename:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   17
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"frmNewProfile.frx":0614
         Height          =   975
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Timer tmrDemo 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3360
      Top             =   3240
   End
   Begin VB.Frame STEP1 
      Caption         =   "Step 1 of 6"
      Height          =   4695
      Left            =   4680
      TabIndex        =   22
      Top             =   5040
      Visible         =   0   'False
      Width           =   6975
      Begin VB.Frame framCOlourSettings 
         Caption         =   "Colours And Descriptions"
         Enabled         =   0   'False
         Height          =   3855
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   6735
         Begin VB.CommandButton Command3 
            Caption         =   "..."
            Height          =   255
            Left            =   4320
            TabIndex        =   41
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Example of Default:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   42
            Top             =   3120
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "New Colour:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   40
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblColourSetting 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   5
            Left            =   2400
            TabIndex        =   39
            Top             =   3120
            Width           =   4095
         End
         Begin VB.Label lblColourSetting 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   1215
            Index           =   4
            Left            =   2400
            TabIndex        =   38
            Top             =   1800
            Width           =   4095
         End
         Begin VB.Label lblColourSetting 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   37
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblColourSetting 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   36
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblColourSetting 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   35
            Top             =   720
            Width           =   4095
         End
         Begin VB.Label lblColourSetting 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   34
            Top             =   360
            Width           =   4095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Default Colour:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Colour Use:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Colour Name:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmNewProfile.frx":06EA
         Left            =   1800
         List            =   "frmNewProfile.frx":0715
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Text Colours"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1560
      End
   End
   Begin VB.Frame STEP2 
      Caption         =   "Step 2 of 6"
      Height          =   4695
      Left            =   4200
      TabIndex        =   23
      Top             =   5040
      Visible         =   0   'False
      Width           =   6975
      Begin VB.Frame frmBackColourSettings 
         Caption         =   "Colours And Descriptions"
         Enabled         =   0   'False
         Height          =   3855
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   6735
         Begin VB.CommandButton Command4 
            Caption         =   "..."
            Height          =   255
            Left            =   4320
            TabIndex        =   46
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Colour Name:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Colour Use:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   57
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Default Colour:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   56
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   55
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblColourSetting 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   2400
            TabIndex        =   54
            Top             =   360
            Width           =   4095
         End
         Begin VB.Label lblColourSetting 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   2400
            TabIndex        =   53
            Top             =   720
            Width           =   4095
         End
         Begin VB.Label lblColourSetting 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   2400
            TabIndex        =   52
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblColourSetting 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   30
            Left            =   2400
            TabIndex        =   51
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblColourSetting 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   1215
            Index           =   40
            Left            =   2400
            TabIndex        =   50
            Top             =   1800
            Width           =   4095
         End
         Begin VB.Label lblColourSetting 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   50
            Left            =   2400
            TabIndex        =   49
            Top             =   3120
            Width           =   4095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "New Colour:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   48
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Example of Default:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   47
            Top             =   3120
            Width           =   2175
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmNewProfile.frx":07E9
         Left            =   1800
         List            =   "frmNewProfile.frx":07F6
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Back Colours"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   1560
      End
   End
   Begin VB.Frame STEP3 
      Caption         =   "Step 3 of 6"
      Height          =   4695
      Left            =   3720
      TabIndex        =   24
      Top             =   5040
      Visible         =   0   'False
      Width           =   6975
      Begin VB.Frame Frame1 
         Caption         =   "Colour Preview"
         Height          =   4335
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   6735
         Begin VB.TextBox txtTESTENTRY 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   3840
            Width           =   6495
         End
         Begin RichTextLib.RichTextBox txtTEXTCHAT 
            Height          =   2415
            Left            =   120
            TabIndex        =   61
            Top             =   1320
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   4260
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmNewProfile.frx":083D
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Long Caption Set On_Load"
            Height          =   855
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   6495
         End
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6480
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<< Back"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   14
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next >>"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   13
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame STEP4 
      Caption         =   "Step 4 of 6"
      Height          =   4695
      Left            =   3000
      TabIndex        =   25
      Top             =   5040
      Visible         =   0   'False
      Width           =   6975
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2385
         ScaleWidth      =   4305
         TabIndex        =   119
         Top             =   240
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Message Font"
         Height          =   855
         Left            =   120
         TabIndex        =   111
         Top             =   2880
         Width           =   6735
         Begin VB.TextBox txtFont 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   120
            TabIndex        =   113
            Text            =   "Arial"
            Top             =   360
            Width           =   6015
         End
         Begin VB.CommandButton Command7 
            Caption         =   "..."
            Height          =   375
            Left            =   6240
            TabIndex        =   112
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Always On Top"
         Height          =   855
         Index           =   1
         Left            =   3000
         TabIndex        =   73
         Top             =   840
         Width           =   2775
         Begin VB.CheckBox chkAOT 
            Caption         =   "Have Window Default Set To Always On Top?"
            Height          =   495
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   2600
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Image Path"
         Height          =   855
         Left            =   120
         TabIndex        =   70
         Top             =   3720
         Width           =   6735
         Begin VB.CommandButton Command8 
            Caption         =   "P"
            Height          =   375
            Left            =   6240
            TabIndex        =   118
            ToolTipText     =   "Preview The Selected Image"
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton Command5 
            Caption         =   "..."
            Height          =   375
            Left            =   5760
            TabIndex        =   72
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtPicPath 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Text            =   "(Default)"
            ToolTipText     =   "HIT 'DEL' TO REMOVE THE IMAGE - HIT 'DEL' AGAIN TO HAVE NO IMAGE"
            Top             =   360
            Width           =   5535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Incomming Messages"
         Height          =   855
         Index           =   0
         Left            =   3000
         TabIndex        =   68
         Top             =   1800
         Width           =   2775
         Begin VB.CheckBox chkPOP 
            Caption         =   "Make Window Pop Up On Top Of Others Briefly"
            Height          =   495
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame frmTrans 
         Caption         =   "Default Transparency"
         Height          =   855
         Left            =   120
         TabIndex        =   65
         Top             =   840
         Width           =   2775
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   66
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "0 %"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Transparency Settings && Other Settings"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   2520
      End
   End
   Begin VB.Frame STEP5 
      Caption         =   "Step 5 of 6"
      Height          =   4695
      Left            =   2640
      TabIndex        =   26
      Top             =   5040
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtPURL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   116
         Top             =   3480
         Width           =   4695
      End
      Begin VB.TextBox txtPPopup 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   114
         Top             =   3120
         Width           =   4695
      End
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   109
         Text            =   "Welcome To The Mercury Chat Network"
         Top             =   2760
         Width           =   4695
      End
      Begin VB.TextBox txtBanFilePath 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   2160
         TabIndex        =   86
         Text            =   "(Default)"
         Top             =   2400
         Width           =   4215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   285
         Left            =   6480
         TabIndex        =   85
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtChannel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   83
         Text            =   "open"
         Top             =   2040
         Width           =   4695
      End
      Begin VB.TextBox txtHeading 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   82
         Text            =   "Mercury Chat"
         Top             =   1680
         Width           =   4695
      End
      Begin VB.TextBox txtChatSep 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   81
         Text            =   "..:."
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtNickname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   80
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Profile URL:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   117
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Profile Popup Msg:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   115
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Profile Message:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   110
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ban List File:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   84
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Default Channel:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   79
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Window Heading:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   78
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Chat Seperator:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   77
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nickname:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   76
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Other Chat Options && More General Settings"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   2520
      End
   End
   Begin VB.Frame STEP6 
      Caption         =   "Step 6 of 6"
      Height          =   4695
      Left            =   2280
      TabIndex        =   87
      Top             =   5040
      Visible         =   0   'False
      Width           =   6975
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3900
         Left            =   240
         ScaleHeight     =   3900
         ScaleWidth      =   6255
         TabIndex        =   91
         Top             =   650
         Width           =   6255
         Begin VB.PictureBox frmBinDS 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5895
            Left            =   0
            ScaleHeight     =   5895
            ScaleWidth      =   6135
            TabIndex        =   92
            Top             =   0
            Width           =   6135
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F1:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   104
               Top             =   120
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F2:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   103
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F3:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   102
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F4:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   101
               Top             =   1560
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F5:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   100
               Top             =   2040
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F6:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   99
               Top             =   2520
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F7:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   98
               Top             =   3000
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F8:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   97
               Top             =   3480
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F9:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   120
               TabIndex        =   96
               Top             =   3960
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F10:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   95
               Top             =   4440
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F11:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   120
               TabIndex        =   94
               Top             =   4920
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "F12:"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   120
               TabIndex        =   93
               Top             =   5400
               Width           =   735
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   335
               Index           =   0
               Left            =   960
               TabIndex        =   0
               Top             =   120
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   1
               Left            =   960
               TabIndex        =   1
               Top             =   600
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   2
               Left            =   960
               TabIndex        =   2
               Top             =   1080
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   3
               Left            =   960
               TabIndex        =   3
               Top             =   1560
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   4
               Left            =   960
               TabIndex        =   4
               Top             =   2040
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   5
               Left            =   960
               TabIndex        =   5
               Top             =   2520
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   6
               Left            =   960
               TabIndex        =   6
               Top             =   3000
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   7
               Left            =   960
               TabIndex        =   7
               Top             =   3480
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   8
               Left            =   960
               TabIndex        =   8
               Top             =   3960
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   9
               Left            =   960
               TabIndex        =   9
               Top             =   4440
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   10
               Left            =   960
               TabIndex        =   10
               Top             =   4920
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbBind 
               Height          =   330
               Index           =   11
               Left            =   960
               TabIndex        =   11
               Top             =   5400
               Width           =   5175
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9128;591"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Verdana"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   6705
         TabIndex        =   89
         Top             =   240
         Width           =   6735
         Begin VB.Line Line4 
            X1              =   6360
            X2              =   6480
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Bind List"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   90
            Top             =   0
            Width           =   1320
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3975
         LargeChange     =   200
         Left            =   6600
         Max             =   2000
         SmallChange     =   50
         TabIndex        =   88
         Top             =   600
         Width           =   255
      End
      Begin VB.Line Line5 
         X1              =   6555
         X2              =   6555
         Y1              =   600
         Y2              =   4680
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   120
         Y1              =   600
         Y2              =   4680
      End
   End
   Begin VB.Frame frmSAVENOW 
      Caption         =   "Saving Profile Now"
      Height          =   4695
      Left            =   1800
      TabIndex        =   107
      Top             =   5040
      Width           =   6975
      Begin VB.TextBox txtLOG 
         Appearance      =   0  'Flat
         Height          =   4215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   108
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7200
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   7200
      Y1              =   4920
      Y2              =   4920
   End
End
Attribute VB_Name = "frmNewProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Steps As Integer
Dim IGNOREErrors As Boolean

Private Sub SETBACKCOLOURS(iIndex As Integer)
    'pWindowBack
    'pChatTextBack
    'pEntryTextBack
Select Case iIndex
    Case 0
        With Me.lblColourSetting
            .Item(11).Caption = "Window Back Colour"   'Name
            .Item(10).Caption = "N/A"   'Frequently Used?
            .Item(40).Caption = "The Window Background Colour" 'Description
            .Item(20).BackColor = -2147483633 'Default Bakcolour

            .Item(50).ForeColor = -2147483633 'Forecolour
            .Item(50).Caption = "N/A" 'Demonstration Text\
            
            .Item(30).BackColor = Combo2.ItemData(iIndex)
        End With
    Case 1
        With Me.lblColourSetting
            .Item(11).Caption = "Text Chat Background"   'Name
            .Item(10).Caption = "N/A"   'Frequently Used?
            .Item(40).Caption = "The Text Chats Backcolour. . Remember It Should Contrast All The Text Colours" 'Description
            .Item(20).BackColor = vbWhite 'Default Bakcolour
            
            .Item(50).ForeColor = vbWhite 'Forecolour
            .Item(50).Caption = "N/A"   'Demonstration Text
            
            .Item(30).BackColor = Combo2.ItemData(iIndex)
        End With
    Case 2
        With Me.lblColourSetting
            .Item(11).Caption = "Text Entry Background"   'Name
            .Item(10).Caption = "N/A"   'Frequently Used?
            .Item(40).Caption = "This Is The Text Entry Background Colour. Remember It Should Contrast The Text Colour"   'Description
            .Item(20).BackColor = vbWhite 'Default Bakcolour
            
            .Item(50).ForeColor = vbWhite 'Forecolour
            .Item(50).Caption = "N/A"   'Demonstration Text
            
            .Item(30).BackColor = Combo2.ItemData(iIndex)
        End With

End Select
    
End Sub


Private Sub SETCOLOURS(iIndex As Integer)

'Normal Text
'List Text
'Error Text
'Notification Text
'Your Chat Text
'Their Chat Text
'Help (Definition)
'Help (Command)
'Entry Text
'Other Text
'Combo1.ItemData(iINDEX)

Select Case iIndex
    Case 0
        With Me.lblColourSetting
            .Item(0).Caption = "Normal Text"   'Name
            .Item(1).Caption = "All The Time"   'Frequently Used?
            .Item(4).Caption = "This Text Is Used in the Welcome Notes And Other Text Displayed" & vbrlf & _
            "It is used alot, and is recommended to be opposite the Back Colour" 'Description
            .Item(2).BackColor = vbBlack 'Default Bakcolour
            
            .Item(5).ForeColor = vbBlack 'Forecolour
            .Item(5).Caption = "Mercury Network Chat: For Use On LAN's && The Internet" 'Demonstration Text\
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
    Case 1
        With Me.lblColourSetting
            .Item(0).Caption = "List Text"   'Name
            .Item(1).Caption = "A Lot"   'Frequently Used?
            .Item(4).Caption = "This is the text used when listing Anything!" & vbCrLf & _
            "Its used when displaying: Bind List, Ban List and Online Lists" 'Description
            .Item(2).BackColor = vbBlue 'Default Bakcolour
            
            .Item(5).ForeColor = vbBlue 'Forecolour
            .Item(5).Caption = "    |40380 [BackCHAt]" & vbCrLf & "     |1881 [LoserChat]"   'Demonstration Text
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
    Case 2
        With Me.lblColourSetting
            .Item(0).Caption = "Error Text"   'Name
            .Item(1).Caption = "Hopefully Not Much"   'Frequently Used?
            .Item(4).Caption = "This Text Is Used When An Error Occurs"   'Description
            .Item(2).BackColor = vbRed 'Default Bakcolour
            
            .Item(5).ForeColor = vbRed 'Forecolour
            .Item(5).Caption = "Unknown Command - help.now"   'Demonstration Text
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
    Case 3
        With Me.lblColourSetting
            .Item(0).Caption = "Notification Text"   'Name
            .Item(1).Caption = "Not So Often"   'Frequently Used?
            .Item(4).Caption = "Used When Notification Needed" & vbCrLf & _
            "Eg- When You Change Channels, Someone joins the channel" 'Description
            .Item(2).BackColor = &H8000& 'Default Bakcolour
            
            .Item(5).ForeColor = &H8000& 'Forecolour
            .Item(5).Caption = "   |You Have Left 'open'" & vbCrLf & "    |You Have Joined '69eRz!'" 'Demonstration Text
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
    Case 4
        With Me.lblColourSetting
            .Item(0).Caption = "Local Chat Colour"   'Name
            .Item(1).Caption = "A Great Lot"   'Frequently Used?
            .Item(4).Caption = "The Colour of Your Chat Text"   'Description
            .Item(2).BackColor = &HFF8080 'Default Bakcolour
            
            .Item(5).ForeColor = &HFF8080 'Forecolour
            .Item(5).Caption = "[CHRIS]-192.168.0.2: Hey Ya!"   'Demonstration Text
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
    Case 5
        With Me.lblColourSetting
            .Item(0).Caption = "Any Remote Text"   'Name
            .Item(1).Caption = "Another Great Lot"   'Frequently Used?
            .Item(4).Caption = "The Colour Of The Remote Chat Text"   'Description
            .Item(2).BackColor = &H80FF& 'Default Bakcolour
            
            .Item(5).ForeColor = &H80FF& 'Forecolour
            .Item(5).Caption = "[EDDIE]-Ed4Bread: Wassup?"   'Demonstration Text
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
    Case 6
        With Me.lblColourSetting
            .Item(0).Caption = "Help Definition"   'Name
            .Item(1).Caption = "Whenever you call \help"   'Frequently Used?
            .Item(4).Caption = "The Colour Of The HELP Contexts DESCRIPTION"   'Description
            .Item(2).BackColor = &H404040 'Default Bakcolour
            
            .Item(5).ForeColor = &H404040 'Forecolour
            .Item(5).Caption = "The \channel <CHANNEL NAME> Command Changes The Current Channel, Enter \Channel alone to view your current channel."   'Demonstration Text
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With

    Case 7
        With Me.lblColourSetting
            .Item(0).Caption = "Help Command Text"   'Name
            .Item(1).Caption = "Whenever you call \help"   'Frequently Used?
            .Item(4).Caption = "The Colour Of The HELP Contexts COMMAND"   'Description
            .Item(2).BackColor = &HC000C0 'Default Bakcolour
            
            .Item(5).ForeColor = &HC000C0 'Forecolour
            .Item(5).Caption = "\Channel <CHANNEL or Clear>"   'Demonstration Text
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
        
    Case 8
        With Me.lblColourSetting
            .Item(0).Caption = "Entry Text"   'Name
            .Item(1).Caption = "Everytime You Enter Text"   'Frequently Used?
            .Item(4).Caption = "This Is The Text Colour Of The TEXTBOX Where You Enter Text && Commands"   'Description
            .Item(2).BackColor = vbBlack 'Default Bakcolour
            
            .Item(5).ForeColor = vbBlack 'Forecolour
            .Item(5).Caption = "\channel 69eRz!"   'Demonstration Text
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
        
    Case 9
        With Me.lblColourSetting
            .Item(0).Caption = "All Other Text"   'Name
            .Item(1).Caption = "Often Used"   'Frequently Used?
            .Item(4).Caption = "This Covers All Other Text Uses"   'Description
            .Item(2).BackColor = &H404080 'Default Bakcolour
            
            .Item(5).ForeColor = &H404080 'Forecolour
            .Item(5).Caption = "Begin Ban List(0)" & vbCrLf & "End Ban List"   'Demonstration Text
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
        
    Case 10
        With Me.lblColourSetting
            .Item(0).Caption = "Your Private Message Colour"   'Name
            .Item(1).Caption = "Often Used"   'Frequently Used?
            .Item(4).Caption = "This Covers Your Display Of Private Messages"   'Description
            .Item(2).BackColor = &H80FF& 'Default Bakcolour
            
            .Item(5).ForeColor = &H80FF& 'Forecolour
            .Item(5).Caption = ":.PM.:[CHRIS]: Hey Man!"
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
        
    Case 11
        With Me.lblColourSetting
            .Item(0).Caption = "All Other Text"   'Name
            .Item(1).Caption = "Often Used"   'Frequently Used?
            .Item(4).Caption = "This Covers All Other Text Uses"   'Description
            .Item(2).BackColor = &H80FF& 'Default Bakcolour
            
            .Item(5).ForeColor = &H80FF& 'Forecolour
            .Item(5).Caption = ":.PM.:[EDDIE]: Wassup?"
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
    Case 12
        With Me.lblColourSetting
            .Item(0).Caption = "URL Text"   'Name
            .Item(1).Caption = "Whenever A URL is Displayed"   'Frequently Used?
            .Item(4).Caption = "Color Of A URL"   'Description
            .Item(2).BackColor = vbBlue 'Default Bakcolour
            
            .Item(5).ForeColor = vbBlue 'Forecolour
            .Item(5).Caption = "http://www.hotmail.com"
            
            .Item(3).BackColor = Combo1.ItemData(iIndex)
        End With
End Select

End Sub

Private Sub Combo1_Click()
    framCOlourSettings.Enabled = True
    SETCOLOURS Combo1.ListIndex
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    framCOlourSettings.Enabled = True
    SETCOLOURS Combo1.ListIndex
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    frmBackColourSettings.Enabled = True
    SETBACKCOLOURS Combo2.ListIndex
End Sub

Private Sub Combo2_Click()
    frmBackColourSettings.Enabled = True
    SETBACKCOLOURS Combo2.ListIndex
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0: IGNOREErrors = False: STEPNOW Steps + 1
    Case 1: IGNOREErrors = True: STEPNOW Steps - 1
    Case 2: frmMain.Show: Unload Me
End Select
End Sub

Private Sub Command2_Click()
    CD.Filter = "Profile Packages (*.prof)|*.prof"
    CD.DialogTitle = "Save Your Profile"
    CD.ShowSave
    Me.txtPFilePath = CD.FileName
    txtPFilePath.SelStart = Len(txtPFilePath)
    If txtPProfileName.Text = "" Then
        txtPProfileName.Text = Mid(CD.FileName, InStrRev(CD.FileName, "\") + 1)
    Else
        If MsgBox("Do You Want To Set The Profile Name As The Filename?", vbYesNo + vbQuestion, "New Path") = vbYes Then
            txtPProfileName.Text = Mid(CD.FileName, InStrRev(CD.FileName, "\") + 1)
        End If
    End If
End Sub

Private Sub Command3_Click()
    CD.ShowColor
    Me.lblColourSetting.Item(3).BackColor = CD.Color
    Combo1.ItemData(Combo1.ListIndex) = Int(CD.Color)
    Combo1.SetFocus
End Sub

Private Sub Command4_Click()
    CD.ShowColor
    Me.lblColourSetting.Item(30).BackColor = CD.Color
    Combo2.ItemData(Combo2.ListIndex) = Int(CD.Color)
    Combo2.SetFocus
End Sub

Private Sub Command5_Click()
    CD.Filter = "Jpg Encoded Image (*.jpg)|*.jpg|Bitmap Image (*.bmp)|*.bmp|GiF Encoded Image (*.gif)|*.gif|All Files (*.*)|*.*|"
    CD.ShowOpen
    
    txtPicPath.Text = CD.FileName
End Sub

Private Sub Command6_Click()
CD.Filter = "Ban List Files (*.ban)|*.ban"
CD.ShowSave
txtBanFilePath.Text = CD.FileName

End Sub



Private Sub Command7_Click()
    CD.Flags = &H3
    CD.ShowFont
    txtFont.Text = CD.FontName
End Sub

Private Sub Command8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Picture4.Picture = LoadPicture(txtPicPath.Text)
    Picture4.Visible = True
End Sub

Private Sub Command8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture4.Visible = False
End Sub

Private Sub Form_Load()

STEP0.Left = 120
STEP0.Top = 120

STEP1.Left = 120
STEP1.Top = 120

STEP2.Left = 120
STEP2.Top = 120

STEP3.Left = 120
STEP3.Top = 120

STEP4.Left = 120
STEP4.Top = 120

STEP5.Left = 120
STEP5.Top = 120

STEP6.Left = 120
STEP6.Top = 120

frmSAVENOW.Left = 120
frmSAVENOW.Top = 120

If OSVersion <> "Windows 2000" Then frmTrans.Enabled = False

Steps = 0
For i = 0 To Combo1.ListCount - 1
Dim iDefCol As Long
Dim iDefCol2 As Long
    Select Case i
        Case 0:
            iDefCol = vbBlack
            iDefCol2 = -2147483633
        Case 1:
            iDefCol = vbBlue
            iDefCol2 = vbWhite
        Case 2:
            iDefCol = vbRed
            iDefCol2 = vbWhite
        Case 3: iDefCol = 32768
        Case 4: iDefCol = 16744576
        Case 5: iDefCol = 33023
        Case 6: iDefCol = 4210752
        Case 7: iDefCol = 12583104
        Case 8: iDefCol = vbBlack
        Case 9: iDefCol = 4210816
        Case 10: iDefCol = &H80FF&
        Case 11: iDefCol = &H80FF&
        Case 12: iDefCol = vbBlue
    End Select
    Combo1.ItemData(i) = iDefCol
    On Error Resume Next
    Combo2.ItemData(i) = iDefCol2
Next

For i = 0 To 11
    cmbBind.Item(i).Clear
Next

For i = 0 To 11
    For j = 0 To frmMain.txtEntry.ListCount
        cmbBind.Item(i).AddItem frmMain.txtEntry.List(j)
    Next
Next

HScroll1.Value = CURTrans
chkPOP.Value = pPopUpWindowOnMessage
chkAOT.Value = pAlwaysOnTop
txtNickname.Text = pHandel
txtChannel = pChannel
txtChatSep = pChatSep
txtFont = pFont
If pPicturePath <> "" Then txtPicPath = pPicturePath
txtHeading = pTitleTextConnected
txtMessage = pProfileMessage
cmbBind(0).Text = pBindF1
cmbBind(1).Text = pBindF2
cmbBind(2).Text = pBindF3
cmbBind(3).Text = pBindF4
cmbBind(4).Text = pBindF5
cmbBind(5).Text = pBindF6
cmbBind(6).Text = pBindF7
cmbBind(7).Text = pBindF8
cmbBind(8).Text = pBindF9
cmbBind(9).Text = pBindF10
cmbBind(10).Text = pBindF11
cmbBind(11).Text = pBindF12

txtPassword.Text = ""

Label5.Caption = "This is a preview of your chat window, It contains some Text, and thats how it will look." & vbCrLf & _
                 "NOTE: There is some text on EVERY Line, so if something is blank, you will need to change it!"

End Sub


Private Sub HideALL()
STEP0.Visible = False
STEP1.Visible = False
STEP2.Visible = False
STEP3.Visible = False
STEP4.Visible = False
STEP5.Visible = False
STEP6.Visible = False

End Sub
Private Sub STEPNOW(iStep As Integer)
    Dim ErrorText As String
Select Case iStep
Case 0:
        HideALL
        STEP0.Visible = True
        Command1.Item(1).Enabled = False
        Steps = 0
        Me.Caption = "New Profile Wizard"
Case 1



    If Me.txtPFilePath.Text <> "" And Me.txtPProfileName.Text <> "" Then
        If txtPassword.Text <> "" Then
            If IGNOREErrors = False Then
        If InBox("Password Check", "Please Re-Enter The Password!", , "*") = txtPassword Then
            
        Else
            MsgBox "Passwords Do Not Match!", vbCritical, "Password Check"
            Exit Sub
            End If
        End If
            End If
            Command1.Item(1).Enabled = True
            HideALL
            STEP0.Visible = False
            STEP1.Visible = True
            Steps = 1
            Me.Caption = "New Profile Wizard - " & Me.txtPProfileName & "*"
        Else
            ErrorText = "You Must Enter The Filename AND The Profile Name!" & vbCrLf & "One Or More Of Them Are Missing!'"
    End If
        
        
        
Case 2
    If Combo1.ListIndex = -1 Then
        If IGNOREErrors = False Then MsgBox "You Haven't Viewed Any Colour Options!" & vbCrLf & "Continuing Will Leave Them As Defaults, You Can Always Go Back And Change Them Later", vbInformation, "Profile Wizard"
    End If
       HideALL
       STEP1.Visible = False
       STEP2.Visible = True
            Steps = 2
        
            
Case 3
    If Combo2.ListIndex = -1 Then
        If IGNOREErrors = False Then MsgBox "You Haven't Viewed Any Colour Options!" & vbCrLf & "Continuing Will Leave Them As Defaults, You Can Always Go Back And Change Them Later", vbInformation
    End If
    
       STEP2.Visible = False
            HideALL
            Steps = 3
'***************************************************************************
'ALL THIS CODE CREATES A DEMONTRATION CHAT SAMPLE!                         *
'It May Be Crude, but If Followed Correctly You Can See Whats Happening.   *
'                                                                          *
'0 Normal Text .                                                           *
'1 List Text .                                                             *
'2 Error Text .                                                            *
'3 Notification Text .                                                     *
'4 Your Chat Text                                                          *
'5 Their Chat Text                                                         *
'6 Help (Definition) .                                                     *
'7 Help (Command) .                                                        *
'8 Entry Text .                                                            *
'9 Other Text .                                                            *
'Combo1.ItemData(iINDEX)                                                   *
'                                                                          *
            txtTEXTCHAT.Text = "" '                                        *
            Frame1.BackColor = Combo2.ItemData(0) '                        *
            txtTEXTCHAT.BackColor = Combo2.ItemData(1) '                   *
            txtTESTENTRY.BackColor = Combo2.ItemData(2) '                  *
            txtTESTENTRY.ForeColor = Combo1.ItemData(8) '                  *
            txtTESTENTRY.Text = "\Channel 9ine" '                          *
            TestText "Mercury Network Chat: For Use On LAN's & The Internet", Combo1.ItemData(0)
            TestText "Your Computer Stats:", Combo1.ItemData(0) '          *
            TestText "   Merc Chat Version: " & VersionCheck, Combo1.ItemData(1)
            TestText "   IP: " & frmMain.WS.LocalIP, Combo1.ItemData(1) '  *
            TestText "   Logged In As: " & Get_User_Name, Combo1.ItemData(1)
                Dim iNICK As String '                                      *
                If pHandel = "" Then iNICK = frmMain.WS.LocalIP Else iNICK = pHandel
            TestText "   Nickname: " & iNICK, Combo1.ItemData(1) '         *
            TestText "   You Are In The Channel: " & pChannel, Combo1.ItemData(1)
            TestText "   Current Profile: " & pProfileName, Combo1.ItemData(1)
            
            TestText "Ban List(3)", Combo1.ItemData(9) '                   *
            TestText "   |1881: " & iNICK, Combo1.ItemData(1) '            *
            TestText "   |Bradley: " & iNICK, Combo1.ItemData(1) '         *
            TestText "   |15126: " & iNICK, Combo1.ItemData(1) '           *
            TestText "End Ban List", Combo1.ItemData(9) '                  *
            
            TestText "Error! Number Must Be Between 0-100", Combo1.ItemData(2)
            TestText "Transparency Level Set To: 99", Combo1.ItemData(9) ' *
            
            TestText "   1881" & " Has Joined The Chat! [Channel: Se7en]", Combo1.ItemData(3)
            
            TestText "[" & Get_User_Name & "]-" & frmMain.WS.LocalIP & pChatSep & "Hey Ed! Whats Up?", Combo1.ItemData(4)
            TestText "[" & 1881 & "]-" & frmMain.WS.LocalIP & ".122" & pChatSep & "Not Much Man!, How Bout You??", Combo1.ItemData(5)
            TestText "[" & 1881 & "]-" & frmMain.WS.LocalIP & ".122" & pChatSep & "http://www.urbanterror.net", Combo1.ItemData(12)
             
            TestText "[PM: " & Get_User_Name & "(1881)]-" & pChatSep & "Hey Ed! Whats Up?", Combo1.ItemData(10)
            TestText "[PM: 1881 (" & Get_User_Name & ")]-" & pChatSep & "Not Much Man!, How Bout You??", Combo1.ItemData(11)
                        
            
            TestText "\Channel <CHANNEL NAME or Clear>", Combo1.ItemData(7)
            TestText "This Command Changes The Current Channel (If A Channel Name Is Entered), If its clear It Returns Your Current Channel", Combo1.ItemData(6)
'***************************************************************************
    
    
    
       STEP3.Visible = True
            Steps = 3
        
Case 4
        HideALL
        STEP3.Visible = False
        STEP4.Visible = True

            Steps = 4
Case 5
        txtBanFilePath = Mid(txtPFilePath, 1, InStrRev(txtPFilePath, ".")) & "ban"
        If txtNickname.Text = "" Then
            txtNickname.Text = Mid(txtPFilePath, InStrRev(txtPFilePath, "\") + 1, InStrRev(txtPFilePath, ".") - InStrRev(txtPFilePath, "\") - 1)
        End If
        txtBanFilePath.SelStart = Len(txtBanFilePath)
        HideALL
        STEP4.Visible = False
        STEP5.Visible = True
        Command1.Item(0).Caption = "Next >>"
        
            Steps = 5
            
Case 6
        HideALL
        STEP5.Visible = False
        STEP6.Visible = True
        Command1.Item(0).Caption = "Finished"
        
            Steps = 6
            
Case 7
        HideALL
        frmSAVENOW.ShowWhatsThis
        Command1.Item(1).Enabled = False
        Command1.Item(0).Caption = "Save Profile"
        Steps = 7
        
        SAVE_PROFiLE_NOW
            
End Select
If ErrorText <> "" Then
    If IGNOREErrors = False Then MsgBox ErrorText, vbOKOnly + vbCritical, "Error"
End If
End Sub

Public Sub TestText(iTEXT As String, Optional iColour As Long, Optional iBold As Boolean, Optional iItalic As Boolean, Optional iUnderline As Boolean)

With frmNewProfile.txtTEXTCHAT
    .SelStart = Len(.Text)
    .SelLength = Len(.Text)
    .SelBold = iBold
    .SelItalic = iItalic
    .SelUnderline = iUnderline
    .SelColor = iColour
    .SelText = iTEXT & vbCrLf
    .SelLength = Len(.Text)
End With


End Sub

Private Sub TransSetVal()
    Label6.Caption = Str(HScroll1.Value) & " %"
End Sub

Private Sub HScroll1_Change()
    TransSetVal
End Sub

Private Sub HScroll1_KeyPress(KeyAscii As Integer)
    TransSetVal
End Sub

Private Sub HScroll1_Scroll()
TransSetVal
End Sub

Private Sub tmrDemo_Timer()
    Me.Caption = Me.Tag
    tmrDemo.Enabled = False
End Sub

Private Sub txtBanFilePath_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    KeyCode = 0
    txtBanFilePath = "(Default)"
End If

End Sub

Private Sub txtBanFilePath_KeyPress(KeyAscii As Integer)
KeyAscii = 0


End Sub

Private Sub txtFont_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    KeyCode = 0
    txtFont.Text = "Arial"
End If
End Sub

Private Sub txtFont_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtPicPath_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    KeyCode = 0
    If txtPicPath.Text <> "(Default)" Then
        txtPicPath.Text = "(Default)"
    Else
        txtPicPath.Text = "(None)"
    End If
End If
End Sub

Private Sub txtPicPath_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub VScroll1_Change()
    frmBinDS.Top = 0 - VScroll1.Value
End Sub

Private Sub VScroll1_KeyPress(KeyAscii As Integer)
    frmBinDS.Top = 0 - VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    frmBinDS.Top = 0 - VScroll1.Value
End Sub

Private Sub AL(iTEXT As String)
    txtLOG.Text = txtLOG.Text & iTEXT & vbCrLf
    txtLOG.SelStart = Len(txtLOG)
End Sub

Private Sub SAVE_PROFiLE_NOW()
On Error GoTo FiLEACCESSERROR
Close #1
Open Me.txtPFilePath For Output As #1

'******************************************************************
'Saves the profile NAM, PASS and the DEveloper Mode (Always False)
'******************************************************************
    Print #1, "pProfileName=" & Me.txtPProfileName
        AL "Saving Profile Name: " & Me.txtPProfileName
    If Me.txtPassword <> "" Then
        Print #1, "pPassword=" & Encode(Me.txtPassword, txtPProfileName)
            AL "Saving Profile Password: *************"
    Else
        Print #1, "pPassword=ÜÞÜÊ"
            AL "Saving Profile Password: *************"
    End If
    
    Print #1, "pDeveloperMode=" & "false"
        AL "Saving DevMode Status"
        
'******************************************************************
'Saves the text Colours
'******************************************************************
'0 Normal Text
'1 List Text
'2 Error Text
'3 Notification Text
'4 Your Chat Text
'5 Their Chat Text
'6 Help (Definition)
'7 Help (Command)
'8 Entry Text
'9 Other Text
'10 Your Private Msg Colour
'11 Their Private Msg Colour
'12 URL Text
'Combo1.ItemData(iINDEX)
        AL "Saving Normal Text Colour"
    Print #1, "pNormalText=" & Combo1.ItemData(0)
        AL "Saving List Text Colour"
    Print #1, "pListText=" & Combo1.ItemData(1)
        AL "Saving Error Text Colour"
    Print #1, "pErrorText=" & Combo1.ItemData(2)
        AL "Saving Notification Text Colour"
    Print #1, "pNotifyText=" & Combo1.ItemData(3)
        AL "Saving Your Chat Text Colour"
    Print #1, "pYourChatText=" & Combo1.ItemData(4)
        AL "Saving Remote Chat Text Colour"
    Print #1, "pTheirChatText=" & Combo1.ItemData(5)
        AL "Saving Help Definition Text Colour"
    Print #1, "pHelpTextD=" & Combo1.ItemData(6)
        AL "Saving Help Command Text Colour"
    Print #1, "pHelpTextN=" & Combo1.ItemData(7)
        AL "Saving Entry Text Colour"
    Print #1, "pEntryText=" & Combo1.ItemData(8)
        AL "Saving Other Text Colours"
    Print #1, "pOtherText=" & Combo1.ItemData(9)
        AL "Saving PM (Yours) Text Colours"
    Print #1, "pyourpmtext=" & Combo1.ItemData(10)
        AL "Saving PM (Their) Text Colours"
    Print #1, "ptheirpmtext=" & Combo1.ItemData(11)
        AL "Saving URL Text Colour"
    Print #1, "urltext=" & Combo1.ItemData(12)
'******************************************************************
'Saves the back colours
'******************************************************************
'pWindowBack
'pChatTextBack
'pEntryTextBack
    
        AL "Saving The Windows Backcolour"
    Print #1, "pWindowBack=" & Combo2.ItemData(0)

        AL "Saving Chat Back Colour"
    Print #1, "pChatTextBack=" & Combo2.ItemData(1)
    
            AL "Saving Entry Text Back Colour"
    Print #1, "pEntryTextBack=" & Combo2.ItemData(2)
    
    
    
'******************************************************************
'Save all other settings
'******************************************************************

            AL "Saving Profile Popup Message"
    Print #1, "pprofilemsgbox=" & txtPPopup.Text

            AL "Saving Title Text"
    Print #1, "pprofileurl=" & txtPURL.Text


            AL "Saving Title Text"
    Print #1, "pTitleTextConnected=" & txtHeading.Text
            
            AL "Saving Handel Name"
    Print #1, "pHandel=" & txtNickname.Text
            
            AL "Saving Chat Seperator"
    Print #1, "pChatSep=" & txtChatSep.Text
    
            AL "Saving Profile Message"
    Print #1, "pProfileMessage=" & txtMessage.Text
    
            AL "Saving Font"
            If txtFont.Text = "" Then
                txtFont.Text = "arial"
            End If
    Print #1, "pFont=" & txtFont.Text
    
            AL "Saving Popup Status"
    Print #1, "pPopUpWindowOnMessage=" & chkPOP.Value
    
            AL "Saving Always On Top Status"
    Print #1, "pAlwaysOnTop=" & chkAOT.Value
        
            AL "Saving Transparency Level Status"
    Print #1, "pTransparency=" & HScroll1.Value

            AL "Saving Picture Path"
            If txtPicPath.Text = "(Default)" Then
                txtPicPath.Text = ""
            End If
    Print #1, "pPicturePath=" & txtPicPath.Text


            AL "Saving Default Channel"
    Print #1, "pChannel=" & txtChannel.Text

            AL "Saving Ban List File"
    Print #1, "pBanListPath=" & txtBanFilePath.Text


'******************************************************************
'Save Binds
'******************************************************************

            AL "Begin Bind Save..."
            For i = 0 To 11
                AL "   Binding: F" & Trim(Str(i + 1))
                Print #1, "pBindF" & Trim(Str(i + 1)) & "=" & cmbBind(i).Text
            Next
            
            AL "End Binding Commands"
            
            'Global pProfileName As String   'The Name
            'Global pNormalText As Long      'Normal Window Dialog
            'Global pListText As Long        'When A List Is Presented
            'Global pErrorText As Long       'When An Error Occurs
            'Global pNotifyText As Long      'When A User Comes/Goes
            'Global pHelpTextN As Long       'The "HELP" Text For The Command
            'Global pHelpTextD As Long       'The "HELP" Text For The Definition
            'Global pOtherText As Long       'Any Other Text
            'Global pEntryText As Long       'The Entry Textbox Forecolour
            'Global pYourChatText As Long    'Your Chat Text Colour
            'Global pTheirChatText As Long   'Their Chat Text Colour
            'Global pWindowBack As Long      'Window Backcolour
            'Global pChatTextBack As Long    'The Main Text Backcolour
            'Global pEntryTextBack As Long   'The Entry Text Backcolour
            'Global pTitleTextNorm As String 'The Window Title Bar Text
            'Global pTitleTextConnected As String    'The Window Title Bar Text (When Connected)
            'Global pHandel As String        'Your Nickname
            'Global pChatSep As String       'The Seperator On The Chat
            'Global pProfileMessage As String        'The Profiles 1Line Message
            'Global pTransparency As Integer 'The Transparency Level
            'Global pFont As String          'The Font For The Text
            'Global pPopUpWindowOnMessage As Boolean 'If The Window Pops Up When Msg Recieved
            'Global pAlwaysOnTop As Boolean  'If The Window Is Always Ontop
            'Global pChannel As String
            'Global pDeveloperMode As Boolean
            'Global pBanListPath As String
            'Global pPassword As String
            'Global pPicturePath As String   'If The User Wants His/Her Own Image
            'Global pBindF1 As String
            'Global pBindF2 As String
            'Global pBindF3 As String
            'Global pBindF4 As String
            'Global pBindF5 As String
            'Global pBindF6 As String
            'Global pBindF7 As String
            'Global pBindF8 As String
            'Global pBindF9 As String
            'Global pBindF10 As String
            'Global pBindF11 As String
            'Global pBindF12 As String
Close #1


If MsgBox("Save Complete!" & vbCrLf & "Would You Like To Load This Profile Now?", vbQuestion + vbYesNo, "Save Complete!") = vbYes Then
    frmMain.Show
    LOAD_Profile txtPFilePath.Text
    Unload Me
Else
    frmMain.Show
    Unload Me
End If

Exit Sub

FiLEACCESSERROR:
If MsgBox("There Has Been A Fatal Error In Compiling Your Profile, Do You Want To Save The Error Log?", vbQuestion + vbYesNo, "Error") = vbYes Then
Close #2
    Open "C:\ProfileErrorLog.txt" For Output As #2
        Print #2, txtLOG.Text
    Close #2
    MsgBox txtLOG.Text
    'Shell "C:\ProfileErrorLog.txt"
End If
End Sub

