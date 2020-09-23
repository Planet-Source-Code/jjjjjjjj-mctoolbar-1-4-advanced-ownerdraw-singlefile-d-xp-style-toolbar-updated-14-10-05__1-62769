VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "McToolBar - Test Form !"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8850
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   590
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   8850
      TabIndex        =   1
      Top             =   0
      Width           =   8850
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "     An advanced XP style Toolbar (Single file'd) with ""Hover effect"" and ""Custom ToolTip""..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Width           =   7830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "McToolBar 1.4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   1740
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmTest.frx":000C
         Top             =   120
         Width           =   480
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   529
      TabCaption(0)   =   "About"
      TabPicture(0)   =   "frmTest.frx":08D6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label12"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label13"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label14"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Shape3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Properties"
      TabPicture(1)   =   "frmTest.frx":08F2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label17"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label19"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label20"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Shape2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label18"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "sss"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Shape5"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label21"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label27"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtindex"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtCaption"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtRow"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtTooltip"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmbIcon"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cmbStyle"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkAuto"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtHeight"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtWidth"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmdApply"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "chkEnabled"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "chklWarp"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "chkEnablectl"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "chkshadow"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "chkPress"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cmbCapAln"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "Appearance"
      TabPicture(2)   =   "frmTest.frx":090E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label24"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label23"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Shape4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Image2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "picCol"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "optBoder"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "optBack"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "optHover"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "optFore"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "optBackGrd"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "optTipBack"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "optTipFore"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "chkBorder"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "chkAppearance"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "cmbGradient"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "cmbHover"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "cmdTile"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).ControlCount=   18
      TabCaption(3)   =   "Operations"
      TabPicture(3)   =   "frmTest.frx":092A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label22"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Shape6"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdRemove"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdMove"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.CommandButton cmdMove 
         Caption         =   "Move Button "
         Height          =   375
         Left            =   -74520
         TabIndex        =   58
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove Button"
         Height          =   375
         Left            =   -72000
         TabIndex        =   57
         Top             =   2640
         Width           =   1935
      End
      Begin VB.ComboBox cmbCapAln 
         Height          =   360
         ItemData        =   "frmTest.frx":0946
         Left            =   3360
         List            =   "frmTest.frx":0956
         TabIndex        =   56
         Text            =   "[ALN_Top] = 0"
         Top             =   3840
         Width           =   2175
      End
      Begin VB.CheckBox chkPress 
         Caption         =   "Pressed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   52
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox chkshadow 
         Caption         =   "Hover Icon Shadow"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2040
      End
      Begin VB.CheckBox chkEnablectl 
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1320
      End
      Begin VB.CheckBox chklWarp 
         Caption         =   "Warp Size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   49
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   48
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   47
         Top             =   4320
         Width           =   2175
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1080
         TabIndex        =   44
         Text            =   "60"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2640
         TabIndex        =   43
         Text            =   "55"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdTile 
         Caption         =   "Insert Tile..."
         Height          =   375
         Left            =   -71640
         TabIndex        =   42
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "AutoSize"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   41
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1320
      End
      Begin VB.ComboBox cmbHover 
         Height          =   360
         ItemData        =   "frmTest.frx":099C
         Left            =   -74640
         List            =   "frmTest.frx":09A9
         TabIndex        =   39
         Text            =   "[WinXP_Hover] = 2"
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox cmbGradient 
         Height          =   360
         ItemData        =   "frmTest.frx":09E6
         Left            =   -71640
         List            =   "frmTest.frx":09FF
         TabIndex        =   37
         Text            =   "[Fill_None] = 0"
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox chkAppearance 
         Caption         =   "Appearance [3d] "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   36
         Top             =   4200
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.CheckBox chkBorder 
         Caption         =   "BorderStyle =1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   35
         Top             =   4560
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.OptionButton optTipFore 
         Caption         =   "ToolTip BackCol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74280
         TabIndex        =   34
         Top             =   3480
         Width           =   2055
      End
      Begin VB.OptionButton optTipBack 
         Caption         =   "ToolTip backCol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74280
         TabIndex        =   33
         Top             =   3240
         Width           =   1935
      End
      Begin VB.OptionButton optBackGrd 
         Caption         =   "Back Gradient Col"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74280
         TabIndex        =   32
         Top             =   2280
         Width           =   1695
      End
      Begin VB.OptionButton optFore 
         Caption         =   "Fore Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74280
         TabIndex        =   31
         Top             =   2880
         Width           =   1575
      End
      Begin VB.OptionButton optHover 
         Caption         =   "Hover Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74280
         TabIndex        =   30
         Top             =   2640
         Width           =   1575
      End
      Begin VB.OptionButton optBack 
         Caption         =   "BackColor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74280
         TabIndex        =   29
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton optBoder 
         Caption         =   "BorderColor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74280
         TabIndex        =   28
         Top             =   1680
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.PictureBox picCol 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   -71640
         Picture         =   "frmTest.frx":0AB6
         ScaleHeight     =   2100
         ScaleWidth      =   2100
         TabIndex        =   27
         Top             =   1440
         Width           =   2130
      End
      Begin VB.ComboBox cmbStyle 
         Height          =   360
         ItemData        =   "frmTest.frx":F0A8
         Left            =   1080
         List            =   "frmTest.frx":F0B2
         TabIndex        =   26
         Text            =   "[Tip_Normal] = 1"
         Top             =   4320
         Width           =   2175
      End
      Begin VB.ComboBox cmbIcon 
         Height          =   360
         ItemData        =   "frmTest.frx":F0DB
         Left            =   1080
         List            =   "frmTest.frx":F0EB
         TabIndex        =   25
         Text            =   "[Icon_None] = 0"
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox txtTooltip 
         Height          =   375
         Left            =   1080
         TabIndex        =   22
         Text            =   "Move to next page!"
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtRow 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   4320
         TabIndex        =   20
         Text            =   "3"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtCaption 
         Height          =   375
         Left            =   1080
         TabIndex        =   19
         Text            =   "Next"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txtindex 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Text            =   "0"
         Top             =   2280
         Width           =   615
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H80000010&
         Height          =   4575
         Left            =   -74880
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Icon Allignment!"
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
         Left            =   3360
         TabIndex        =   61
         Top             =   3480
         Width           =   1380
      End
      Begin VB.Label Label21 
         Caption         =   "Tool Style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   60
         Top             =   4320
         Width           =   540
      End
      Begin VB.Label Label22 
         Caption         =   "Try tooltip on the Gradient toolbar>>>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73080
         TabIndex        =   59
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H80000010&
         Height          =   3015
         Left            =   120
         Top             =   2040
         Width           =   5535
      End
      Begin VB.Label sss 
         Caption         =   "Buttons Width"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   46
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label18 
         Caption         =   "Buttons Height"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   45
         Top             =   720
         Width           =   540
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -71640
         Picture         =   "frmTest.frx":F137
         Top             =   4080
         Width           =   615
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000010&
         Height          =   4575
         Left            =   -74880
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Hover Style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74640
         TabIndex        =   40
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "BackGradient"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71640
         TabIndex        =   38
         Top             =   720
         Width           =   945
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000010&
         Height          =   4575
         Left            =   -74880
         Top             =   480
         Width           =   5535
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000010&
         Height          =   1455
         Left            =   120
         Top             =   480
         Width           =   5535
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000010&
         Height          =   2535
         Left            =   -74760
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label20 
         Caption         =   "Tool Icon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   24
         Top             =   3840
         Width           =   540
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "ToolTip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   3480
         Width           =   510
      End
      Begin VB.Label Label17 
         Caption         =   "Buttons per Row"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3600
         TabIndex        =   21
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   3000
         Width           =   555
      End
      Begin VB.Label Label15 
         Caption         =   "Button Index"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   360
         TabIndex        =   16
         Top             =   2280
         Width           =   465
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   -74760
         X2              =   -69720
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   -74760
         X2              =   -69720
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label14 
         Caption         =   "Give the new index to the property ""ButtonMove"". Selected Button will move to the given index"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74520
         TabIndex        =   15
         Top             =   4560
         Width           =   5055
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Move Button :"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74760
         TabIndex        =   14
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "In property window set ""ButtonRemove"" to ""Yes!"". The selected button will be removed!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74520
         TabIndex        =   13
         Top             =   3840
         Width           =   5055
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Remove Button :"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74760
         TabIndex        =   12
         Top             =   3600
         Width           =   1440
      End
      Begin VB.Label Label10 
         Caption         =   "In property window set ""Button_Count"". This much buttons will be created instantly with default values"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74520
         TabIndex        =   11
         Top             =   1440
         Width           =   5055
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Create Button :"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74760
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   $"frmTest.frx":FAD9
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74520
         TabIndex        =   9
         Top             =   2880
         Width           =   5055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Assign Properties :"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74760
         TabIndex        =   8
         Top             =   2640
         Width           =   1680
      End
      Begin VB.Label Label5 
         Caption         =   "In property window set the ""Button_Index"" (is shown on the control, underlined in design time )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74520
         TabIndex        =   7
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Label Label6 
         Caption         =   "          Asign all the properties at design time without using property window! [ Read the following carefully! ]"
         Height          =   495
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select Button :"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74760
         TabIndex        =   5
         Top             =   1920
         Width           =   1290
      End
   End
   Begin ToolBar.McToolBar McToolBar2 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   53
      Top             =   975
      Width           =   8850
      _ExtentX        =   15081
      _ExtentY        =   688
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
      BackGround      =   "frmTest.frx":FB8B
      ButtonsWidth    =   30
      ButtonsHeight   =   26
      ButtonsPerRow   =   19
      HoverColor      =   16422674
      TooTipStyle     =   0
      ToolTipBackCol  =   -2147483624
      ButtonsStyle    =   2
      BorderColor     =   16761024
      ButtonCaption0  =   ""
      ButtonPicture0  =   "frmTest.frx":1053D
      ButtonToolTipText0=   "Open"
      ButtonCaption1  =   ""
      ButtonPicture1  =   "frmTest.frx":108D7
      ButtonToolTipText1=   "Save"
      ButtonCaption2  =   ""
      ButtonPicture2  =   "frmTest.frx":10C71
      ButtonToolTipText2=   "Copy"
      ButtonCaption3  =   ""
      ButtonPicture3  =   "frmTest.frx":1100B
      ButtonToolTipText3=   "Cut"
      ButtonCaption4  =   ""
      ButtonPicture4  =   "frmTest.frx":113A5
      ButtonToolTipText4=   "Delete"
      ButtonCaption5  =   ""
      ButtonPicture5  =   "frmTest.frx":1173F
      ButtonToolTipText5=   "Find"
      ButtonEnabled5  =   0   'False
      ButtonCaption6  =   ""
      ButtonPicture6  =   "frmTest.frx":11AD9
      ButtonToolTipText6=   "Home"
      ButtonPressed6  =   -1  'True
      ButtonCaption7  =   ""
      ButtonPicture7  =   "frmTest.frx":11E73
      ButtonToolTipText7=   "Mail"
      ButtonCaption8  =   ""
      ButtonPicture8  =   "frmTest.frx":1220D
      ButtonToolTipText8=   "MSN"
      ButtonCaption9  =   ""
      ButtonPicture9  =   "frmTest.frx":125A7
      ButtonToolTipText9=   "Options"
   End
   Begin ToolBar.McToolBar McToolBar4 
      Align           =   4  'Align Right
      Height          =   5640
      Left            =   7950
      TabIndex        =   62
      Top             =   1365
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   10186
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   7
      ButtonsWidth    =   60
      ButtonsHeight   =   55
      ButtonsPerRow   =   1
      HoverColor      =   8421631
      TooTipStyle     =   1
      ToolTipBackCol  =   16777215
      BackGradient    =   1
      ButtonsStyle    =   2
      BorderColor     =   16711680
      ButtonCaption0  =   "Next"
      ButtonPicture0  =   "frmTest.frx":12941
      ButtonToolTipText0=   "Move to next page!"
      ButtonToolTipIcon0=   2
      ButtonCaption1  =   "Home"
      ButtonPicture1  =   "frmTest.frx":130BB
      ButtonToolTipText1=   "Click to open Home!"
      ButtonToolTipIcon1=   1
      ButtonPressed1  =   -1  'True
      ButtonCaption2  =   "Sync"
      ButtonPicture2  =   "frmTest.frx":13835
      ButtonToolTipText2=   "Syncronize!"
      ButtonToolTipIcon2=   3
      ButtonCaption3  =   "Address"
      ButtonPicture3  =   "frmTest.frx":13FAF
      ButtonToolTipText3=   "Click here to view the Addresses!"
      ButtonToolTipIcon3=   2
      ButtonCaption4  =   "Attach"
      ButtonPicture4  =   "frmTest.frx":14729
      ButtonToolTipText4=   "Attach files!"
      ButtonToolTipIcon4=   1
      ButtonCaption5  =   "Disabled"
      ButtonPicture5  =   "frmTest.frx":14EA3
      ButtonToolTipText5=   "Disabled button!"
      ButtonToolTipIcon5=   2
      ButtonEnabled5  =   0   'False
      ButtonCaption6  =   "Music"
      ButtonPicture6  =   "frmTest.frx":1561D
      ButtonToolTipText6=   "Get mad with Music!"
   End
   Begin ToolBar.McToolBar McToolBar1 
      Height          =   4950
      Left            =   120
      TabIndex        =   63
      Top             =   1680
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   8731
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
      Button_Count    =   12
      ButtonsWidth    =   60
      ButtonsHeight   =   55
      ButtonsPerRow   =   2
      HoverColor      =   8421631
      TooTipStyle     =   1
      ButtonsStyle    =   2
      ButtonCaption0  =   "Next"
      ButtonPicture0  =   "frmTest.frx":15D97
      ButtonToolTipText0=   "Move to next page!"
      ButtonToolTipIcon0=   2
      ButtonCaption1  =   "Home"
      ButtonPicture1  =   "frmTest.frx":16511
      ButtonToolTipText1=   "Click to open Home!"
      ButtonToolTipIcon1=   1
      ButtonCaption2  =   "Sync"
      ButtonPicture2  =   "frmTest.frx":16C8B
      ButtonToolTipText2=   "Syncronize!"
      ButtonToolTipIcon2=   3
      ButtonCaption3  =   "Address"
      ButtonPicture3  =   "frmTest.frx":17405
      ButtonToolTipText3=   "Click here to view the Addresses!"
      ButtonToolTipIcon3=   2
      ButtonCaption4  =   "Attach"
      ButtonPicture4  =   "frmTest.frx":17B7F
      ButtonToolTipText4=   "Attach files!"
      ButtonToolTipIcon4=   1
      ButtonCaption5  =   "Button 5"
      ButtonPicture5  =   "frmTest.frx":182F9
      ButtonToolTipText5=   "Move to last page!"
      ButtonToolTipIcon5=   2
      ButtonCaption6  =   "Music"
      ButtonPicture6  =   "frmTest.frx":18A73
      ButtonToolTipText6=   "Get mad with Music!"
      ButtonCaption7  =   "Tools"
      ButtonPicture7  =   "frmTest.frx":191ED
      ButtonToolTipText7=   "Advanced diagonizing tools!"
      ButtonCaption8  =   "History"
      ButtonPicture8  =   "frmTest.frx":19967
      ButtonToolTipText8=   "View History"
      ButtonCaption9  =   "MSN"
      ButtonPicture9  =   "frmTest.frx":1A0E1
      ButtonToolTipText9=   "Enter to msn!"
      ButtonCaption10 =   "Windows"
      ButtonPicture10 =   "frmTest.frx":1A9BB
      ButtonToolTipText10=   "Windows update!"
      ButtonCaption11 =   "Paste"
      ButtonPicture11 =   "frmTest.frx":1B295
      ButtonToolTipText11=   "Paste Image!"
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "Style XP >>"
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
      Left            =   6840
      TabIndex        =   55
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Style Normal ^^"
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
      Left            =   240
      TabIndex        =   54
      Top             =   6720
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Style Raised ^^"
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
      TabIndex        =   4
      Top             =   1440
      Width           =   1365
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub chkAppearance_Click()
    McToolBar1.Appearance = chkAppearance
End Sub

Private Sub chkAuto_Click()
    McToolBar1.AutoSize = chkAuto
End Sub

Private Sub chkBorder_Click()
    McToolBar1.BorderStyle = chkBorder
End Sub

Private Sub chkEnablectl_Click()
    McToolBar1.Enabled = chkEnablectl
End Sub

Private Sub chklWarp_Click()
    McToolBar1.WarpSize = chklWarp
End Sub

Private Sub chlshadow_Click()

End Sub

Private Sub chkshadow_Click()
    McToolBar1.HoverIconShadow = chkshadow
End Sub

Private Sub cmbGradient_Click()
    McToolBar1.BackGradient = cmbGradient.ListIndex
End Sub

Private Sub cmbHover_Click()
    McToolBar1.ButtonsStyle = cmbHover.ListIndex
End Sub



Private Sub cmdApply_Click()
    McToolBar1.Button_Index = Val(txtindex)
    McToolBar1.ButtonCaption = txtCaption
    McToolBar1.ToolTipText = txtTooltip
    McToolBar1.ButtonToolTipIcon = cmbIcon.ListIndex
    McToolBar1.TooTipStyle = cmbStyle.ListIndex
    McToolBar1.ButtonEnabled = chkEnabled
    McToolBar1.ButtonPressed = chkPress
    McToolBar1.ButtonIconAllignment = cmbCapAln.ListIndex
End Sub

Private Sub cmdMove_Click()
    McToolBar1.Button_Index = Val(txtindex)
    McToolBar1.ButtonMoveTo = Val(InputBox("Enter new index!", "NewIndex!"))
End Sub

Private Sub cmdRemove_Click()
    McToolBar1.Button_Index = Val(txtindex)
    McToolBar1.ButtonRemove = 1
End Sub

Private Sub cmdTile_Click()
    Set McToolBar1.BackGround = Image2.Picture
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    cmbCapAln.ListIndex = 0
    cmbHover.ListIndex = 2
    cmbIcon.ListIndex = 0
    cmbStyle.ListIndex = 0
End Sub

Private Sub McToolBar1_Click(ByVal vButton_Index As Long)
    Debug.Print vbCrLf & "Clicked on " & vButton_Index & vbCrLf
End Sub

Private Sub picCol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If optBack Then McToolBar1.BackColor = picCol.Point(X, Y)
    If optBackGrd Then McToolBar1.BackGradientCol = picCol.Point(X, Y)
    If optBoder Then McToolBar1.BorderColor = picCol.Point(X, Y)
    If optFore Then McToolBar1.ForeColor = picCol.Point(X, Y)
    If optHover Then McToolBar1.HoverColor = picCol.Point(X, Y)
    If optTipBack Then McToolBar1.ToolTipBackCol = picCol.Point(X, Y)
    If optTipFore Then McToolBar1.ToolTipForeCol = picCol.Point(X, Y)
End Sub

Private Sub txtHeight_Change()
    McToolBar1.ButtonsHeight = Val(txtHeight)
End Sub

Private Sub txtRow_Change()
    McToolBar1.ButtonsPerRow = Val(txtRow)
End Sub


Private Sub txtWidth_Change()
    McToolBar1.ButtonsWidth = Val(txtWidth)
End Sub
