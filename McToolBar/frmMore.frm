VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "McToolBar - Styles!"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Win 98"
      TabPicture(0)   =   "frmMore.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Shape4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "McToolBar5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "McToolBar4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "McToolBar3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "McToolBar2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "McToolBar1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "XP Style"
      TabPicture(1)   =   "frmMore.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "McToolBar6"
      Tab(1).Control(1)=   "McToolBar9"
      Tab(1).Control(2)=   "McToolBar10"
      Tab(1).Control(3)=   "McToolBar13"
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(8)=   "Shape1"
      Tab(1).Control(9)=   "Label7"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "frmMore.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label15"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label16"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Shape2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label10"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "McToolBar8"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "McToolBar12"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "McToolBar11"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdTest"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin VB.CommandButton cmdTest 
         Caption         =   "Goto test form >>>"
         Height          =   495
         Left            =   -72720
         TabIndex        =   26
         Top             =   5400
         Width           =   3015
      End
      Begin ToolBar.McToolBar McToolBar1 
         Height          =   1200
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   2117
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   6
         ButtonsWidth    =   60
         TooTipStyle     =   0
         ButtonsStyle    =   1
         ButtonCaption0  =   "Button 0"
         ButtonCaption1  =   "Button 1"
         ButtonCaption2  =   "Button 2"
         ButtonCaption3  =   "Button 3"
         ButtonCaption4  =   "Button 4"
         ButtonCaption5  =   "Button 5"
         ButtonPressed5  =   -1  'True
      End
      Begin ToolBar.McToolBar McToolBar2 
         Height          =   1200
         Left            =   360
         TabIndex        =   3
         Top             =   2520
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   2117
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   6
         ButtonsWidth    =   60
         TooTipStyle     =   0
         ButtonCaption0  =   "Button 0"
         ButtonCaption1  =   "Button 1"
         ButtonCaption2  =   "Button 2"
         ButtonCaption3  =   "Button 3"
         ButtonCaption4  =   "Button 4"
         ButtonCaption5  =   "Button 5"
      End
      Begin ToolBar.McToolBar McToolBar3 
         Height          =   600
         Left            =   360
         TabIndex        =   4
         Top             =   3960
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1058
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   3
         ButtonsWidth    =   60
         TooTipStyle     =   0
         ButtonsStyle    =   1
         ButtonCaption0  =   ""
         ButtonPicture0  =   "frmMore.frx":0054
         ButtonCaption1  =   ""
         ButtonPicture1  =   "frmMore.frx":07CE
         ButtonCaption2  =   ""
         ButtonPicture2  =   "frmMore.frx":0B68
         ButtonPressed2  =   -1  'True
      End
      Begin ToolBar.McToolBar McToolBar4 
         Height          =   600
         Left            =   360
         TabIndex        =   7
         Top             =   5040
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   1058
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   3
         Button_Index    =   1
         ButtonsWidth    =   110
         AutoSize        =   0   'False
         TooTipStyle     =   0
         ButtonsStyle    =   1
         ButtonCaption0  =   "Addresses"
         ButtonPicture0  =   "frmMore.frx":12E2
         ButtonToolTipText0=   "Click here to view the addresses"
         ButtonToolTipIcon0=   2
         ButtonIconAllignment0=   2
         ButtonCaption1  =   "Dialup NetWork"
         ButtonPicture1  =   "frmMore.frx":1A5C
         ButtonToolTipText1=   "Click here to get connected!"
         ButtonToolTipIcon1=   1
         ButtonIconAllignment1=   2
         ButtonCaption2  =   "MS Info"
         ButtonPicture2  =   "frmMore.frx":2336
         ButtonToolTipText2=   "MS Info!"
         ButtonToolTipIcon2=   3
         ButtonIconAllignment2=   2
      End
      Begin ToolBar.McToolBar McToolBar5 
         Height          =   2700
         Left            =   4320
         TabIndex        =   10
         Top             =   1200
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   4763
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   3
         ButtonsWidth    =   60
         AutoSize        =   0   'False
         ButtonsHeight   =   60
         ButtonsPerRow   =   1
         ButtonsStyle    =   1
         ButtonCaption0  =   "DirectX"
         ButtonPicture0  =   "frmMore.frx":2C10
         ButtonToolTipText0=   "Direct X Driver!"
         ButtonToolTipIcon0=   2
         ButtonCaption1  =   "Desktop"
         ButtonPicture1  =   "frmMore.frx":34EA
         ButtonToolTipText1=   "Desktop!"
         ButtonToolTipIcon1=   1
         ButtonEnabled1  =   0   'False
         ButtonCaption2  =   "IE"
         ButtonPicture2  =   "frmMore.frx":3DC4
         ButtonToolTipText2=   "IE"
         ButtonToolTipIcon2=   3
      End
      Begin ToolBar.McToolBar McToolBar6 
         Height          =   1200
         Left            =   -74640
         TabIndex        =   12
         Top             =   2760
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   2117
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   6
         ButtonsWidth    =   60
         HoverColor      =   16422674
         TooTipStyle     =   0
         ButtonsStyle    =   2
         ButtonCaption0  =   "Button 0"
         ButtonCaption1  =   "Button 1"
         ButtonCaption2  =   "Button 2"
         ButtonCaption3  =   "Button 3"
         ButtonCaption4  =   "Button 4"
         ButtonCaption5  =   "Button 5"
         ButtonPressed5  =   -1  'True
      End
      Begin ToolBar.McToolBar McToolBar9 
         Height          =   600
         Left            =   -74640
         TabIndex        =   13
         Top             =   5040
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   1058
         BackColor       =   16422674
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   3
         ButtonsWidth    =   110
         AutoSize        =   0   'False
         HoverColor      =   16699301
         TooTipStyle     =   0
         BackGradient    =   3
         BackGradientCol =   16777215
         ButtonsStyle    =   2
         ButtonCaption0  =   "Addresses"
         ButtonPicture0  =   "frmMore.frx":469E
         ButtonToolTipText0=   "Click here to view the addresses"
         ButtonToolTipIcon0=   2
         ButtonIconAllignment0=   2
         ButtonCaption1  =   "Dialup Network"
         ButtonPicture1  =   "frmMore.frx":4E18
         ButtonToolTipText1=   "Click here to get connected!"
         ButtonToolTipIcon1=   1
         ButtonIconAllignment1=   2
         ButtonCaption2  =   "MS Info"
         ButtonPicture2  =   "frmMore.frx":56F2
         ButtonToolTipText2=   "MS Info!"
         ButtonToolTipIcon2=   3
         ButtonIconAllignment2=   2
      End
      Begin ToolBar.McToolBar McToolBar10 
         Height          =   3600
         Left            =   -70680
         TabIndex        =   14
         Top             =   1080
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   6350
         BackColor       =   16777215
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   4
         ButtonsWidth    =   60
         AutoSize        =   0   'False
         ButtonsHeight   =   60
         ButtonsPerRow   =   1
         HoverColor      =   12632319
         BackGradient    =   1
         ButtonsStyle    =   2
         ButtonCaption0  =   "DirectX"
         ButtonPicture0  =   "frmMore.frx":5FCC
         ButtonToolTipText0=   "Direct X Driver!"
         ButtonToolTipIcon0=   2
         ButtonCaption1  =   "Desktop"
         ButtonPicture1  =   "frmMore.frx":68A6
         ButtonToolTipText1=   "Desktop!"
         ButtonToolTipIcon1=   1
         ButtonEnabled1  =   0   'False
         ButtonCaption2  =   "IE"
         ButtonPicture2  =   "frmMore.frx":7180
         ButtonCaption3  =   "Home"
         ButtonPicture3  =   "frmMore.frx":7A5A
      End
      Begin ToolBar.McToolBar McToolBar11 
         Height          =   600
         Left            =   -74760
         TabIndex        =   19
         Top             =   960
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1058
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   3
         ButtonsWidth    =   60
         TooTipStyle     =   0
         ButtonCaption0  =   "Button 0"
         ButtonCaption1  =   "Button 1"
         ButtonCaption2  =   "Button 2"
      End
      Begin ToolBar.McToolBar McToolBar12 
         Height          =   600
         Left            =   -74760
         TabIndex        =   20
         Top             =   1920
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1058
         Appearance      =   1
         BorderStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   3
         ButtonsWidth    =   60
         TooTipStyle     =   0
         ButtonCaption0  =   "Button 0"
         ButtonCaption1  =   "Button 1"
         ButtonCaption2  =   "Button 2"
      End
      Begin ToolBar.McToolBar McToolBar13 
         Height          =   780
         Left            =   -74640
         TabIndex        =   25
         Top             =   1080
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1376
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   10
         ButtonsWidth    =   30
         ButtonsHeight   =   26
         ButtonsPerRow   =   6
         HoverColor      =   16422674
         TooTipStyle     =   0
         ToolTipBackCol  =   -2147483624
         ButtonsStyle    =   2
         BorderColor     =   16761024
         ButtonCaption0  =   ""
         ButtonPicture0  =   "frmMore.frx":8334
         ButtonToolTipText0=   "Open"
         ButtonCaption1  =   ""
         ButtonPicture1  =   "frmMore.frx":86CE
         ButtonToolTipText1=   "Save"
         ButtonCaption2  =   ""
         ButtonPicture2  =   "frmMore.frx":8A68
         ButtonToolTipText2=   "Copy"
         ButtonCaption3  =   ""
         ButtonPicture3  =   "frmMore.frx":8E02
         ButtonToolTipText3=   "Cut"
         ButtonCaption4  =   ""
         ButtonPicture4  =   "frmMore.frx":919C
         ButtonToolTipText4=   "Delete"
         ButtonCaption5  =   ""
         ButtonPicture5  =   "frmMore.frx":9536
         ButtonToolTipText5=   "Find"
         ButtonEnabled5  =   0   'False
         ButtonCaption6  =   ""
         ButtonPicture6  =   "frmMore.frx":98D0
         ButtonToolTipText6=   "Home"
         ButtonPressed6  =   -1  'True
         ButtonCaption7  =   ""
         ButtonPicture7  =   "frmMore.frx":9C6A
         ButtonToolTipText7=   "Mail"
         ButtonCaption8  =   ""
         ButtonPicture8  =   "frmMore.frx":A004
         ButtonToolTipText8=   "MSN"
         ButtonCaption9  =   ""
         ButtonPicture9  =   "frmMore.frx":A39E
         ButtonToolTipText9=   "Options"
      End
      Begin ToolBar.McToolBar McToolBar8 
         Height          =   600
         Left            =   -74760
         TabIndex        =   27
         Top             =   2880
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   1058
         BackColor       =   16422674
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   3
         ButtonsWidth    =   110
         AutoSize        =   0   'False
         HoverColor      =   16699301
         TooTipStyle     =   0
         BackGradient    =   3
         BackGradientCol =   16777215
         ButtonsStyle    =   1
         ButtonCaption0  =   "Addresses"
         ButtonPicture0  =   "frmMore.frx":A738
         ButtonToolTipText0=   "Click here to view the addresses"
         ButtonToolTipIcon0=   2
         ButtonIconAllignment0=   2
         ButtonCaption1  =   "Dialup Network"
         ButtonPicture1  =   "frmMore.frx":AEB2
         ButtonToolTipText1=   "Click here to get connected!"
         ButtonToolTipIcon1=   1
         ButtonIconAllignment1=   2
         ButtonCaption2  =   "MS Info"
         ButtonPicture2  =   "frmMore.frx":B78C
         ButtonToolTipText2=   "MS Info!"
         ButtonToolTipIcon2=   3
         ButtonIconAllignment2=   2
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "No Caption (Only Icon)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74640
         TabIndex        =   29
         Top             =   840
         Width           =   1890
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Style 98, But with gradients"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   28
         Top             =   2640
         Width           =   2325
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000010&
         Height          =   5535
         Left            =   -74880
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label16 
         Caption         =   "Note : Icon shadow can be enabled for both XP and non-XP styles!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -74760
         TabIndex        =   24
         Top             =   4680
         Width           =   5025
      End
      Begin VB.Label Label15 
         Caption         =   $"frmMore.frx":C066
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -74760
         TabIndex        =   23
         Top             =   3960
         Width           =   5025
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Downed border"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   22
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Raised Border"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   21
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Without icon (Raised Buttons)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74640
         TabIndex        =   18
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "With Icon, caption ( Icon allined to left)  (Test ToolTip)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74640
         TabIndex        =   17
         Top             =   4800
         Width           =   4575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Note: Caption having more length is split down to two lines!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74640
         TabIndex        =   16
         Top             =   5640
         Width           =   4965
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000010&
         Height          =   5535
         Left            =   -74880
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label7 
         Caption         =   "Vertically aligned    ( index1- disabled!)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -71400
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Vertically aligned    ( index1- disabled!)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3600
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000010&
         Height          =   5535
         Left            =   120
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Note: Caption having more length is split down to two lines!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   5640
         Width           =   4965
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "With Icon, caption ( Icon allined to left)  (Test ToolTip)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   4800
         Width           =   4575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "With Icon (No caption, allined to center)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   3720
         Width           =   3360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Flat Buttons ( Will raise on Hover)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Without icon (Raised Buttons)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
   End
   Begin ToolBar.McToolBar McToolBar7 
      Align           =   4  'Align Right
      Height          =   6705
      Left            =   5820
      TabIndex        =   30
      Top             =   0
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   11827
      Appearance      =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      Button_Count    =   12
      ButtonsHeight   =   33
      ButtonsPerRow   =   1
      HoverColor      =   16699301
      TooTipStyle     =   1
      ToolTipBackCol  =   16777215
      BackGradientCol =   -2147483633
      ButtonsStyle    =   1
      ButtonCaption0  =   "Next"
      ButtonPicture0  =   "frmMore.frx":C0EE
      ButtonToolTipText0=   "Move to next page!"
      ButtonToolTipIcon0=   2
      ButtonIconAllignment0=   2
      ButtonCaption1  =   "Home"
      ButtonPicture1  =   "frmMore.frx":C868
      ButtonToolTipText1=   "Click to open Home!"
      ButtonToolTipIcon1=   1
      ButtonPressed1  =   -1  'True
      ButtonIconAllignment1=   3
      ButtonCaption2  =   "Sync"
      ButtonPicture2  =   "frmMore.frx":CFE2
      ButtonToolTipText2=   "Syncronize!"
      ButtonToolTipIcon2=   3
      ButtonIconAllignment2=   2
      ButtonCaption3  =   "Address"
      ButtonPicture3  =   "frmMore.frx":D75C
      ButtonToolTipText3=   "Click here to view the Addresses!"
      ButtonToolTipIcon3=   2
      ButtonIconAllignment3=   3
      ButtonCaption4  =   "Attach"
      ButtonPicture4  =   "frmMore.frx":DED6
      ButtonToolTipText4=   "Attach files!"
      ButtonToolTipIcon4=   1
      ButtonIconAllignment4=   2
      ButtonCaption5  =   "Disabled"
      ButtonPicture5  =   "frmMore.frx":E650
      ButtonToolTipText5=   "Disabled button!"
      ButtonToolTipIcon5=   2
      ButtonEnabled5  =   0   'False
      ButtonIconAllignment5=   3
      ButtonCaption6  =   "Music"
      ButtonPicture6  =   "frmMore.frx":EDCA
      ButtonToolTipText6=   "Get mad with Music!"
      ButtonIconAllignment6=   2
      ButtonCaption7  =   "Tools"
      ButtonPicture7  =   "frmMore.frx":F544
      ButtonToolTipText7=   "Advanced diagonizing tools!"
      ButtonIconAllignment7=   3
      ButtonCaption8  =   "History"
      ButtonPicture8  =   "frmMore.frx":FCBE
      ButtonToolTipText8=   "View History"
      ButtonEnabled8  =   0   'False
      ButtonIconAllignment8=   2
      ButtonCaption9  =   "MSN"
      ButtonPicture9  =   "frmMore.frx":10438
      ButtonToolTipText9=   "Enter to msn!"
      ButtonIconAllignment9=   3
      ButtonCaption10 =   "Windows"
      ButtonPicture10 =   "frmMore.frx":10D12
      ButtonToolTipText10=   "Windows update!"
      ButtonIconAllignment10=   2
      ButtonCaption11 =   "Paste"
      ButtonPicture11 =   "frmMore.frx":115EC
      ButtonToolTipText11=   "Paste Image!"
      ButtonIconAllignment11=   3
   End
End
Attribute VB_Name = "frmMore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdTest_Click()
    frmTest.Show
End Sub
