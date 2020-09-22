VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   120
      Top             =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1572
      Left            =   4440
      TabIndex        =   10
      Top             =   1560
      Width           =   1572
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   252
         Left            =   420
         TabIndex        =   12
         Top             =   480
         Width           =   852
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "5"
         Top             =   960
         Width           =   288
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   288
         Left            =   768
         TabIndex        =   13
         Top             =   960
         Width           =   252
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "Frame1"
         BuddyDispid     =   196619
         OrigLeft        =   1800
         OrigTop         =   840
         OrigRight       =   2052
         OrigBottom      =   1092
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   252
      Left            =   3000
      TabIndex        =   9
      Top             =   2400
      Value           =   -1  'True
      Width           =   852
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   252
      Left            =   3000
      TabIndex        =   8
      Top             =   2760
      Width           =   1104
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   252
      Left            =   3000
      TabIndex        =   7
      Top             =   3120
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   612
      Left            =   2880
      TabIndex        =   6
      Top             =   1560
      Width           =   1212
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   252
      LargeChange     =   33
      Left            =   480
      Max             =   99
      TabIndex        =   5
      Top             =   5520
      Value           =   49
      Width           =   2772
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1812
      LargeChange     =   25
      Left            =   3600
      Max             =   100
      TabIndex        =   4
      Top             =   3960
      Value           =   50
      Width           =   252
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   4200
      List            =   "Form1.frx":000D
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3480
      Width           =   1812
   End
   Begin VB.ListBox List1 
      Height          =   510
      ItemData        =   "Form1.frx":003B
      Left            =   4200
      List            =   "Form1.frx":004E
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   4080
      Width           =   1812
   End
   Begin VB.ListBox List2 
      Height          =   645
      ItemData        =   "Form1.frx":00B1
      Left            =   4200
      List            =   "Form1.frx":00C4
      TabIndex        =   1
      Top             =   4920
      Width           =   1812
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   3960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      _Version        =   327682
      LargeChange     =   10
      Max             =   100
      TickFrequency   =   10
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   4920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2055
      Left            =   480
      TabIndex        =   16
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3625
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tab 1"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tab 2"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tab 3"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please Note: You need to compile this code into a .exe and run the .exe for this code to work !!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    '#####################################################################

    'USE THIS CODE IF YOU WANT ONLY ONE INSTANCE OF YOUR PROGRAM TO RUN
    If GetSetting(App.EXEName, "Settings", "CanRun", "NO") <> "YES" Then
        If App.PrevInstance Then
            MsgBox "Another instance of this program is already running!" & Chr(13) & "This program will now exit", vbCritical, "WinXP"
            End
        End If
    End If
        
    SaveSetting App.EXEName, "Settings", "CanRun", "NO"
    
    '######################################################################
    
    'CHECK IF WINDOWS XP (DEMO PURPOSES)
    If Val(Win32Ver) < 6 Then
        MsgBox "ERROR: This code only works if you have Windows XP", vbCritical, "WinXP"
    End If
        
End Sub

Private Sub Timer1_Timer() 'PROGRESS BAR MOVEMENT (DEMO PURPOSES)
    ProgressBar1.Value = ProgressBar1.Value + 0.5
    If ProgressBar1.Value >= 100 Then ProgressBar1.Value = 0
End Sub
