VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "xScanner"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wSw 
      Left            =   6930
      Top             =   495
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar PbarS 
      Height          =   240
      Left            =   4545
      TabIndex        =   34
      Top             =   4095
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog cD 
      Left            =   6525
      Top             =   3690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      DialogTitle     =   "Open/Save file"
      Filter          =   "Plain Text Files (*.txt, *.log)|*.txt; *.log"
   End
   Begin MSComctlLib.ImageList imBar 
      Left            =   6750
      Top             =   3690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1E72
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":212E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":228A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":23E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2702
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":285E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":29BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2B16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   635
      ButtonWidth     =   1455
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Log  "
            Object.ToolTipText     =   "Write log to file"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View "
            Object.ToolTipText     =   "View Log"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " "
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ping  "
            Object.ToolTipText     =   "Ping a host"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DNS "
            Object.ToolTipText     =   "Convert DNS names"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Scan "
            Object.ToolTipText     =   "Scan ports"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Who "
            Object.ToolTipText     =   "Run a WhoIs query"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " "
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help  "
            Object.ToolTipText     =   "Get Help"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Info  "
            Object.ToolTipText     =   "About this program"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imMenu 
      Left            =   6660
      Top             =   3465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2E32
            Key             =   ""
            Object.Tag             =   "&Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2F8E
            Key             =   ""
            Object.Tag             =   "&View"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":30EA
            Key             =   ""
            Object.Tag             =   "E&xit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3246
            Key             =   ""
            Object.Tag             =   "&Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":33A2
            Key             =   ""
            Object.Tag             =   "C&opy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":34FE
            Key             =   ""
            Object.Tag             =   "&Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":365A
            Key             =   ""
            Object.Tag             =   "&Select"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":37B6
            Key             =   ""
            Object.Tag             =   "&Options..."
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3912
            Key             =   ""
            Object.Tag             =   "&Contents"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3A6E
            Key             =   ""
            Object.Tag             =   "&Search..."
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BCA
            Key             =   ""
            Object.Tag             =   "&Website"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3D26
            Key             =   ""
            Object.Tag             =   "&About"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4042
            Key             =   ""
            Object.Tag             =   "&Undo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":419E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":44BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":47D6
            Key             =   ""
            Object.Tag             =   "&Browser"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4932
            Key             =   ""
            Object.Tag             =   "&Ping"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4A8E
            Key             =   ""
            Object.Tag             =   "Activity &Log"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4BF2
            Key             =   ""
            Object.Tag             =   "W&hoIs"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4D4E
            Key             =   ""
            Object.Tag             =   "Port S&can"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tTab 
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   405
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      MouseIcon       =   "main.frx":506A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Activity &Log"
      TabPicture(0)   =   "main.frx":5086
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "wsP(0)"
      Tab(0).Control(1)=   "tmrSave"
      Tab(0).Control(2)=   "txLog"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Ping/DNS"
      TabPicture(1)   =   "main.frx":51E0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txTimeout"
      Tab(1).Control(1)=   "txDatasize"
      Tab(1).Control(2)=   "cmPing"
      Tab(1).Control(3)=   "cmIP"
      Tab(1).Control(4)=   "cmHostname"
      Tab(1).Control(5)=   "cbPing"
      Tab(1).Control(6)=   "Label1"
      Tab(1).Control(7)=   "Label2"
      Tab(1).Control(8)=   "Label3"
      Tab(1).Control(9)=   "Label4"
      Tab(1).Control(10)=   "Label5"
      Tab(1).Control(11)=   "Label6"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "W&hoIs"
      TabPicture(2)   =   "main.frx":533A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txWhois"
      Tab(2).Control(1)=   "txWhoisH"
      Tab(2).Control(2)=   "cmWhois"
      Tab(2).Control(3)=   "cbWhois"
      Tab(2).Control(4)=   "Label9"
      Tab(2).Control(5)=   "Label7"
      Tab(2).Control(6)=   "Label8"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Port S&can"
      TabPicture(3)   =   "main.frx":5494
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label10"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label11"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label12"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label13"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "opCommon"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "opCustom"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmConfigurePorts"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "lsPorts"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "lvAPorts"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txAPorts"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "tScan"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "imlScan"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Frame1"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "txScanHost"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "txHowMany"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "wsPc(0)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "tmPort"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "PBarM"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).ControlCount=   18
      Begin MSComctlLib.ProgressBar PBarM 
         Height          =   330
         Left            =   4635
         TabIndex        =   38
         Top             =   2700
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Timer tmPort 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2385
         Top             =   3105
      End
      Begin MSWinsockLib.Winsock wsPc 
         Index           =   0
         Left            =   1935
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txHowMany 
         Height          =   315
         Left            =   1575
         TabIndex        =   37
         Text            =   "15"
         Top             =   3060
         Width           =   510
      End
      Begin MSWinsockLib.Winsock wsP 
         Index           =   0
         Left            =   -73110
         Top             =   450
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageCombo txScanHost 
         Height          =   330
         Left            =   4635
         TabIndex        =   35
         Top             =   630
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "Host to Scan"
         ImageList       =   "imBar"
      End
      Begin VB.Frame Frame1 
         Height          =   2985
         Left            =   2205
         TabIndex        =   31
         Top             =   450
         Width           =   25
      End
      Begin MSComctlLib.ImageList imlScan 
         Left            =   6435
         Top             =   3150
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":55EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":574E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":5A72
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":5BCE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tScan 
         Height          =   330
         Left            =   2925
         TabIndex        =   30
         Top             =   3105
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   582
         ButtonWidth     =   1746
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imlScan"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Star&t   "
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Stop   "
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "S&ave  "
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cle&ar   "
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txAPorts 
         Height          =   1680
         Left            =   4635
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   29
         Top             =   990
         Width           =   2400
      End
      Begin MSComctlLib.ListView lvAPorts 
         Height          =   2400
         Left            =   2340
         TabIndex        =   28
         Top             =   630
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   4233
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imMenu"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Scan Results"
            Object.Width           =   3881
         EndProperty
      End
      Begin MSComctlLib.ListView lsPorts 
         Height          =   1455
         Left            =   225
         TabIndex        =   27
         Top             =   900
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Select ports"
            Object.Width           =   2822
         EndProperty
      End
      Begin VB.CommandButton cmConfigurePorts 
         Caption         =   "Configure..."
         Height          =   375
         Left            =   270
         TabIndex        =   26
         Top             =   2610
         Width           =   1365
      End
      Begin VB.OptionButton opCustom 
         Caption         =   "Cus&tom specification"
         Height          =   195
         Left            =   225
         TabIndex        =   25
         Top             =   2385
         Width           =   1770
      End
      Begin VB.OptionButton opCommon 
         Caption         =   "&Common ports only"
         Height          =   240
         Left            =   225
         TabIndex        =   24
         Top             =   675
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.Timer tmrSave 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   -74820
         Top             =   360
      End
      Begin VB.TextBox txTimeout 
         Height          =   315
         Left            =   -72465
         TabIndex        =   14
         Text            =   "500"
         Top             =   1035
         Width           =   510
      End
      Begin VB.TextBox txDatasize 
         Height          =   315
         Left            =   -72465
         TabIndex        =   13
         Text            =   "32"
         Top             =   1380
         Width           =   510
      End
      Begin VB.CommandButton cmPing 
         Caption         =   "&Start Ping"
         Default         =   -1  'True
         Height          =   600
         Left            =   -73335
         Picture         =   "main.frx":5D2A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1755
         Width           =   1050
      End
      Begin VB.CommandButton cmIP 
         Caption         =   "Get &IP"
         Height          =   600
         Left            =   -72255
         Picture         =   "main.frx":5E74
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1755
         Width           =   960
      End
      Begin VB.CommandButton cmHostname 
         Caption         =   "Get H&ostname"
         Height          =   600
         Left            =   -71265
         Picture         =   "main.frx":5FBE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1755
         Width           =   1230
      End
      Begin VB.TextBox txWhois 
         BackColor       =   &H80000004&
         Height          =   2085
         Left            =   -74550
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   990
         Width           =   6180
      End
      Begin VB.TextBox txWhoisH 
         Height          =   315
         Left            =   -71130
         TabIndex        =   5
         Top             =   630
         Width           =   1950
      End
      Begin VB.CommandButton cmWhois 
         Caption         =   "S&tart"
         Height          =   315
         Left            =   -69150
         TabIndex        =   4
         Top             =   630
         Width           =   735
      End
      Begin VB.TextBox txLog 
         Height          =   3120
         Left            =   -74910
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   405
         Width           =   6900
      End
      Begin MSComctlLib.ImageCombo cbWhois 
         Height          =   330
         Left            =   -74100
         TabIndex        =   7
         Top             =   630
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "WhoIs Hostname"
         ImageList       =   "imBar"
      End
      Begin MSComctlLib.ImageCombo cbPing 
         Height          =   330
         Left            =   -72465
         TabIndex        =   15
         Top             =   675
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "imBar"
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Simultaneou&s scan:"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   3105
         Width           =   1380
      End
      Begin VB.Label Label12 
         Caption         =   "Host&name:"
         Height          =   195
         Left            =   4635
         TabIndex        =   33
         Top             =   420
         Width           =   1230
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Scan Res&ults"
         Height          =   195
         Left            =   2385
         TabIndex        =   32
         Top             =   425
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Scan m&ethod:"
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   450
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Pick a server and type in your query, click the 'Start' button when you're ready."
         Height          =   195
         Left            =   -74325
         TabIndex        =   22
         Top             =   3195
         Width           =   5730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Host Name:"
         Height          =   195
         Left            =   -73320
         TabIndex        =   21
         Top             =   765
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ti&meout:"
         Height          =   195
         Left            =   -73320
         TabIndex        =   20
         Top             =   1110
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(Milliseconds)"
         Height          =   195
         Left            =   -71880
         TabIndex        =   19
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data &size:"
         Height          =   195
         Left            =   -73320
         TabIndex        =   18
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label5 
         Caption         =   "(Bytes)"
         Height          =   240
         Left            =   -71880
         TabIndex        =   17
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label6 
         Caption         =   $"main.frx":6108
         Height          =   420
         Left            =   -74190
         TabIndex        =   16
         Top             =   2745
         Width           =   5505
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Host:"
         Height          =   195
         Left            =   -74505
         TabIndex        =   9
         Top             =   705
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Query for:"
         Height          =   195
         Left            =   -71940
         TabIndex        =   8
         Top             =   705
         Width           =   765
      End
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   4035
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7964
            Picture         =   "main.frx":61AC
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1561
            MinWidth        =   1569
            TextSave        =   "8:36 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileBrowser 
         Caption         =   "&Browser"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveLog 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileOpenLog 
         Caption         =   "&Open..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "E&dit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelect 
         Caption         =   "&Select"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "Vie&w"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utility"
      Begin VB.Menu mnuUtil 
         Caption         =   "Activity &Log"
         Index           =   0
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuUtil 
         Caption         =   "&Ping/DNS"
         Index           =   1
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuUtil 
         Caption         =   "W&hoIs Info"
         Index           =   2
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuUtil 
         Caption         =   "Port S&canner"
         Index           =   3
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuHelpWebsite 
         Caption         =   "&Website"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuHelpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
         Shortcut        =   {F8}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================
'   xScanner : Port Scanner, Ping & WhoIs Utility
'=====================================
' By Sushant Pandurangi [sushant@phreaker.net]
'=====================================
'This utility allows you to go through hosts & find
'out which ports are active on them. You can also
'run a WhoIs query or a ping and there is also the
'added functionality of converting domains to IPs
'and also the other way round.
'=====================================
'For more software, including the fast, vast & free
'VB6LIB, click to http://sushantshome.tripod.com.
'=====================================
'In this project I have used a Menus OCX that is, to
'be true, not one made by me. My menus module is
'not used because I wanted to experiment with this
'other OCX that I had found elsewhere on the net.
'=====================================
'Thanks to Scott Pierce (webmaster@calclinks.net)
'=====================================
Option Explicit
Public bCompleted As Boolean

Private Sub cbPing_Change()
'Accordingly enable button
cmPing.Enabled = (cbPing.Text <> "")
End Sub

Private Sub cmConfigurePorts_Click()
'Show options form
mnuToolsOptions_Click
End Sub

Private Sub cmIP_Click()
Screen.MousePointer = 11
'Get IP address of given host
cbPing.Text = Trim(cbPing.Text)
Dim Res As String, sErr As String
Res = GetIPAddress(cbPing.Text, sErr)
If (ParseInt(sErr) = 0 And Res = "") Then MsgBox "Unexpected error.", vbCritical, "DNS": WriteLog "Unexpected error.": Exit Sub
If Res = "" Then
MsgBox "'" & cbPing.Text & "' could not be resolved." & vbNewLine & "(" & GetStatusCode(ParseInt(sErr)) & ")", vbExclamation, "DNS"
WriteLog "'" & cbPing.Text & "' could not be resolved." & vbNewLine & "(" & GetStatusCode(ParseInt(sErr)) & ")"
Else
MsgBox "IP Address for '" & cbPing.Text & "' is " & Res & ".", vbInformation, "DNS"
WriteLog "IP Address for '" & cbPing.Text & "' is " & Res & "."
If Trim(cbPing.Text) <> "" Then cbPing.ComboItems.Add , , cbPing.Text, 3
End If
Screen.MousePointer = 0
End Sub

Private Sub cmHostname_Click()
'Get Host name of given IP
Screen.MousePointer = 11
cbPing.Text = Trim(cbPing.Text)
Dim Res As String, sErr As String
Res = GetHostFromIP(cbPing.Text, sErr)
If cbPing.Text = "" Then MsgBox "Your Hostname is " & GetIPHostName(), vbInformation, "DNS": WriteLog "Your Hostname is " & GetIPHostName(): Screen.MousePointer = 0: Exit Sub
If (ParseInt(sErr) = 0 And Res = "") Then MsgBox "Unexpected error.", vbCritical, "DNS": WriteLog "Unexpected Error.": Exit Sub
If Res = "" Then
MsgBox "'" & cbPing.Text & "' could not be resolved." & vbNewLine & "(" & GetStatusCode(ParseInt(sErr)) & ")", vbExclamation, "DNS"
WriteLog "'" & cbPing.Text & "' could not be resolved." & vbNewLine & "(" & GetStatusCode(ParseInt(sErr)) & ")"
Else
MsgBox "Hostname for '" & cbPing.Text & "' is " & Res & ".", vbInformation, "DNS"
WriteLog "Hostname for '" & cbPing.Text & "' is " & Res & "."
If Trim(cbPing.Text) <> "" Then cbPing.ComboItems.Add , , cbPing.Text, 3
End If
Screen.MousePointer = 0
End Sub

Private Sub cmPing_Click()
'This proc will do the ping
cbPing.Text = Trim(cbPing.Text)
If cbPing.Text = "" Then Exit Sub
Screen.MousePointer = 11
Dim strRTT As String, fDM As Boolean, result As Long
Screen.MousePointer = vbHourglass
result = Ping(cbPing.Text, strRTT, fDM, CLng(txDatasize.Text), CLng(txTimeout.Text))
If fDM = True Then
MsgBox ("Successfully pinged " & cbPing.Text & " in " & strRTT & "."), vbInformation, "Ping"
If Trim(cbPing.Text) <> "" Then cbPing.ComboItems.Add , , cbPing.Text, 3
WriteLog "Successfully pinged " & cbPing.Text & " in " & strRTT & "."
Else
MsgBox "Could not ping " & cbPing.Text & "." & vbNewLine & "(" & GetStatusCode(result) & ")", vbExclamation, "Ping"
WriteLog "Could not ping " & cbPing.Text & "." & vbNewLine & "(" & GetStatusCode(result) & ")"
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub cmWhois_Click()
'Initialize whois
Screen.MousePointer = 11
On Error Resume Next
wSw.Close
wSw.Connect Trim(cbWhois.Text), 43
End Sub

Private Sub Form_Load()
bCompleted = False 'whether or not scanning completed
App.HelpFile = App.Path & "\scanner.hlp"
cD.Flags = cdlOFNFileMustExist And cdlOFNOverwritePrompt
GetPrefs
SetFonts Me, "MS Sans Serif", txLog
AddCommonPorts
Progress False, Me, PbarS, sBar, 1 'Merge progress bar
opCommon.Value = ReadValue("Settings", "ScanCommon", True)
opCustom.Value = Not opCommon.Value
End Sub

Private Sub Form_Resize()
On Error Resume Next
Progress True, Me, PbarS, sBar, 1
tTab.Width = ScaleWidth
tTab.Height = ScaleHeight - sBar.Height - Tbar.Height - 45
tTab.Top = Tbar.Height + 45
txLog.Width = tTab.Width - 190
txLog.Height = tTab.Height - 490
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetPrefs
End Sub

Private Sub lsPorts_ItemClick(ByVal Item As MSComctlLib.ListItem)
sBar.Panels(1).Text = "Specify if or not to scan " & Item.Text
End Sub

Private Sub mnuEditCopy_Click()
On Error Resume Next
Clipboard.SetText ActiveControl.SelText, vbCFText
End Sub

Private Sub mnuEditCut_Click()
On Error Resume Next
Clipboard.SetText ActiveControl.SelText, vbCFText
ActiveControl.SelText = ""
End Sub

Private Sub mnuEditPaste_Click()
On Error Resume Next
ActiveControl.SelText = Clipboard.GetText()
End Sub

Private Sub mnuEditSelect_Click()
On Error Resume Next
ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub mnuEditUndo_Click()
On Error Resume Next
SendMessage ActiveControl.hWnd, EM_UNDO, 0, 0&
End Sub

Private Sub mnuFileBrowser_Click()
ShellExecute Me.hWnd, "open", "http://www.yahoo.com", "", App.Path, 10
End Sub

Private Sub mnuFileExit_Click()
'Long procedure; but anyways...
Dim pf As Form
For Each pf In Forms
Unload pf
Set pf = Nothing
Next pf
Unload Me
Set fMain = Nothing
End Sub

Private Sub mnuFileOpenLog_Click()
On Error GoTo hell
cD.ShowOpen
Open cD.FileName For Input As #1
txLog.Text = Input(LOF(1), 1)
Close #1
sBar.Panels(1).Text = "Viewing " & cD.FileTitle
tTab.Tab = 0
Exit Sub
hell:
Close #1
End Sub

Private Sub mnuFileSaveLog_Click()
On Error GoTo hell
cD.ShowSave
Open cD.FileName For Output As #1
WriteLog "Saved log file as " & cD.FileName
Print #1, txLog.Text
Close #1
sBar.Panels(1).Text = "Saved " & cD.FileTitle
tTab.Tab = 0
Exit Sub
hell:
Close #1
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuHelpContents_Click()
ShellExecute Me.hWnd, "open", App.HelpFile, "", App.Path, 10
End Sub

Private Sub mnuHelpWebsite_Click()
ShellExecute Me.hWnd, "open", "http://sushantshome.tripod.com/xs/index.html", "", "", 10
End Sub

Private Sub mnuToolsOptions_Click()
frmOpts.Show vbModal
End Sub

Private Sub mnuUtil_Click(Index As Integer)
'accordingly show tab
Dim p As Integer
For p = 0 To mnuUtil.Count - 1
mnuUtil(p).Checked = False
Next p
tTab.Tab = Index
mnuUtil(Index).Checked = True
End Sub

Private Sub mnuUtility_Click()
'which is selected
Dim p As Integer
For p = 0 To mnuUtil.Count - 1
mnuUtil(p).Checked = False
Next p
mnuUtil(tTab.Tab).Checked = True
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1: mnuFileSaveLog_Click
Case 2: mnuFileOpenLog_Click
Case 4: tTab.Tab = 1
Case 5: tTab.Tab = 1
Case 6: tTab.Tab = 3
Case 7: tTab.Tab = 2
Case 9: mnuHelpContents_Click
Case 10: mnuHelpAbout_Click
End Select
End Sub

Sub WriteLog(Contents As String)
On Error Resume Next
txLog.Text = txLog.Text & vbNewLine & Time() & vbNewLine & Contents & vbNewLine & String(50, "-")
End Sub

Sub GetPrefs()
On Error Resume Next
Dim p As Integer, pt As Integer
pt = ReadValue("PingHosts", "Count", 0)
With Me
.cbPing.ComboItems.Clear
.cbWhois.ComboItems.Clear
.txScanHost.Text = ReadValue("Settings", "DefaultPortHost", "")
.cbPing.Text = ReadValue("Settings", "DefaultPingHost", "")
.txDatasize.Text = ReadValue("Settings", "PingDataSize", "32")
.txTimeout.Text = ReadValue("Settings", "PingTimeout", "250")
For p = 1 To pt
.cbPing.ComboItems.Add , , ReadValue("PingHosts", "Host" & p, ""), 3
Next p
pt = ReadValue("PortHosts", "Count", 0)
For p = 1 To pt
.txScanHost.ComboItems.Add , , ReadValue("PortHosts", "Host" & p, ""), 3
Next p
End With
tmrSave.Enabled = ReadValue("Settings", "AutoSave", True)
Caption = ReadValue("Settings", "Caption", "xScanner - Default User")
txLog.Text = "Session started " & Date & ", " & Time() & vbNewLine & String(50, "-") & vbNewLine
pt = ReadValue("WServers", "Count", 0)
For p = 1 To pt
cbWhois.ComboItems.Add , , ReadValue("WServers", "Server" & p, ""), 3
Next p
cbWhois.SelectedItem = 1
End Sub

Sub SetPrefs()
Dim p As Integer, pt As Integer
pt = cbPing.ComboItems.Count - 1
SaveValue "PingHosts", "Count", CStr(pt)
For p = 1 To pt
SaveValue "PingHosts", "Host" & p, cbPing.ComboItems.Item(p).Text
Next p
pt = txScanHost.ComboItems.Count - 1
SaveValue "PortHosts", "Count", CStr(pt)
For p = 1 To pt
SaveValue "PortHosts", "Host" & p, txScanHost.ComboItems.Item(p).Text
Next p
SaveValue "Settings", "DefaultPingHost", cbPing.Text
SaveValue "Settings", "DefaultPortHost", txScanHost.Text
SaveValue "Settings", "PingTimeout", txTimeout.Text
SaveValue "Settings", "PingDataSize", txDatasize.Text
SaveValue "Settings", "Caption", Caption
SaveValue "Settings", "AutoSave", tmrSave.Enabled
End Sub

Private Sub tmPort_Timer()
sBar.Panels(1).Text = "Scanning ports, scanned " & lngNextPort - ParseInt(ReadValue("Settings", "ScanFrom", 0)) & "; ports remaining " & ParseInt(ReadValue("Settings", "ScanTo", CStr(lngNextPort))) - lngNextPort
End Sub

Private Sub tmrSave_Timer()
'AutoSave proc
On Error Resume Next
Open App.Path & "\logfile.log" For Output As #1
WriteLog "AutoSave completed; saved logfile.log"
sBar.Panels(1).Text = "AutoSave completed; saved logfile.log"
Print #1, txLog.Text
Close #1
End Sub

Private Sub tScan_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
DoScan
Case 2
StopScan
Case 3
SaveScanFile
Case 4
lvAPorts.ListItems.Clear
txAPorts.Text = ""
PbarS.Value = 0
PbarM.Value = 0
End Select
End Sub

Private Sub wsP_Connect(Index As Integer)
'success
txAPorts.Text = txAPorts.Text & vbNewLine & "Connected on port " & Index & vbNewLine
lvAPorts.ListItems.Add , , Index & " ACTIVE", , 17
WriteLog "Scanning " & txScanHost.Text & vbNewLine & "Port #" & Index & " is ACTIVE"
PbarS.Value = PbarS.Value + 1
PbarM.Value = PbarS.Value
wsP(Index).Close
Unload wsP(Index)
If PbarM.Value = PbarM.Max Then bCompleted = True Else bCompleted = False
If bCompleted = True Then sBar.Panels(1).Text = "Completed scanning.": WriteLog "Completed scanning."
If bCompleted = True Then txScanHost.ComboItems.Add , , txScanHost.Text, 3
End Sub

Private Sub wsP_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'inactive port
txAPorts.Text = txAPorts.Text & vbNewLine & "Can't connect on " & Index & vbNewLine & "(" & Description & ")" & vbNewLine
lvAPorts.ListItems.Add , , Index & " Inactive", , 3
WriteLog "Scanning " & txScanHost.Text & vbNewLine & "Port #" & Index & " is inactive"
PbarS.Value = PbarS.Value + 1
PbarM.Value = PbarS.Value
wsP(Index).Close
Unload wsP(Index)
If PbarM.Value = PbarM.Max Then bCompleted = True Else bCompleted = False
If bCompleted = True Then sBar.Panels(1).Text = "Completed scanning.": WriteLog "Completed scanning."
If bCompleted = True Then txScanHost.ComboItems.Add , , txScanHost.Text, 3
End Sub

Private Sub wsPc_Connect(Index As Integer)
lvAPorts.ListItems.Add , , lngNextPort & " ACTIVE", , 3
txAPorts.Text = txAPorts.Text & vbNewLine & "Connected to port " & lngNextPort
WriteLog "Connected to port " & lngNextPort
PbarM.Value = PbarM.Value + 1
PbarS.Value = PbarM.Value
TryNextPort Index
End Sub

Private Sub wsPc_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lvAPorts.ListItems.Add , , lngNextPort & " Inactive", , 3
txAPorts.Text = txAPorts.Text & vbNewLine & "Can't Connect to port " & lngNextPort
WriteLog "Can't Connect to port #" & lngNextPort & vbNewLine & "(" & Description & ")"
PbarM.Value = PbarM.Value + 1
PbarS.Value = PbarM.Value
TryNextPort Index
End Sub

Private Sub wSw_Connect()
'Whois connected
sBar.Panels(1).Text = "Sending Data to server"
wSw.SendData txWhoisH.Text & vbCrLf
Screen.MousePointer = 0
End Sub

Private Sub wSw_DataArrival(ByVal bytesTotal As Long)
'after data is in
sBar.Panels(1).Text = "Recieving Data from server"
Dim strData As String
txWhois.Text = ""
wSw.GetData strData
txWhois.Text = txWhois.Text & strData
sBar.Panels(1).Text = "Recieved " & bytesTotal & " bytes, transfer complete"
End Sub

Private Sub wSw_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'uh oh...
MsgBox Description, vbExclamation + vbMsgBoxHelpButton, "Error", HelpFile, HelpContext
Screen.MousePointer = 0
End Sub

Private Sub wSw_SendComplete()
'now wait for response
sBar.Panels(1).Text = "Waiting for Data to arrive"
End Sub

Sub AddCommonPorts()
'for portscan
With lsPorts.ListItems
.Add , , "SyStat (11)"
.Add , , "Echo (7)"
.Add , , "Discard (9)"
.Add , , "IMAP (143)"
.Add , , "Daytime (13)"
.Add , , "CharGen (19)"
.Add , , "NetStat (15)"
.Add , , "FTP (21)"
.Add , , "Time (37)"
.Add , , "SSH (22)"
.Add , , "Telnet (23)"
.Add , , "SMTP (25)"
.Add , , "WhoIs (43)"
.Add , , "DNS (53)"
.Add , , "NNTP (119)"
.Add , , "Finger (79)"
.Add , , "Gopher (70)"
.Add , , "HTTP (80)"
.Add , , "POP3 (110)"
.Add , , "iDent (113)"
.Add , , "SHTTP (443)"
.Add , , "B1FF (512)"
.Add , , "rLogin (513)"
.Add , , "Shell (514)"
.Add , , "Route (520)"
.Add , , "WinGate (1080)"
.Add , , "IRC (6667)"
.Add , , "SubSeven (1243)"
.Add , , "DipStix (2002)"
.Add , , "NetBus (12345)"
.Add , , "SubSeven (27374)"
.Add , , "BackOrifice (31337)"
End With
Dim p As Integer
For p = 1 To lsPorts.ListItems.Count
lsPorts.ListItems.Item(p).Checked = True
Next p
End Sub

Sub DoScan()
If txScanHost.Text = "" Then Exit Sub
'Either of two options
If opCommon.Value = True Then DoCommonScan Else DoCustomScan ReadValue("Settings", "ScanFrom", 0), ReadValue("Settings", "ScanTo", 0)
End Sub

Sub StopScan()
'Terminate; abort
On Error Resume Next
bCompleted = False
Dim p As Integer
For p = 0 To wsP.Count - 1
wsP(p).Close
Unload wsP(p)
Next p
For p = 0 To wsPc.Count - 1
wsPc(p).Close
Unload wsPc(p)
Next p
PbarS.Value = 0
PbarM.Value = 0
Screen.MousePointer = 0
End Sub

Sub SaveScanFile()
'Log results
On Error GoTo hell
cD.Filter = "Plain Text Files (*.txt, *.log)|*.txt; *.log"
cD.ShowSave
Open cD.FileName For Output As #1
Print #1, txAPorts.Text
Close #1
hell:
End Sub

Sub DoCommonScan()
'common ports
On Error Resume Next
StopScan
bCompleted = False
PbarS.Value = 0
PbarM.Value = 0
txAPorts.Text = ""
lvAPorts.ListItems.Clear
Dim p As Integer, port As Integer
For p = 1 To lsPorts.ListItems.Count
port = ParseInt(lsPorts.ListItems.Item(p).Text)
Load wsP(port)
wsP(port).Connect txScanHost.Text, port
Next p
PbarS.Max = lsPorts.ListItems.Count
PbarM.Max = PbarS.Max
End Sub

Sub DoCustomScan(vStartFrom As Integer, vGoUpTo As Integer)
'specified ports
On Error Resume Next
StopScan
tmPort.Enabled = True
bCompleted = False
txAPorts.Text = ""
lvAPorts.ListItems.Clear
PbarS.Max = vGoUpTo - vStartFrom
PbarM.Max = PbarS.Max
lngNextPort = vStartFrom
Screen.MousePointer = 11
Dim p As Integer
For p = 0 To (ParseInt(txHowMany.Text))
Load wsPc(p)
wsPc(p).Connect txScanHost.Text, lngNextPort
Err.Raise 13
lngNextPort = lngNextPort + 1
Next p
End Sub

Private Sub TryNextPort(Index As Integer)
On Error Resume Next
'This function will close the current winsock connection, check to see if
'upper port bound has been reached, then try a connection to the next
'available port or unload the current winsock control depending on if there
'are more ports to scan or not.
   Me.wsPc(Index).Close
   'If lngNextPort >= Val(ReadValue("Settings", "ScanTo")) Then StopScan: Exit Sub
   If lngNextPort < Val(ReadValue("Settings", "ScanTo")) Then
          Me.wsPc(Index).Connect txScanHost.Text, lngNextPort
          lngNextPort = lngNextPort + 1
   Else
          Unload Me.wsPc(Index)
          Screen.MousePointer = 0
          bCompleted = True
          tmPort.Enabled = False
          sBar.Panels(1).Text = "Completed Scanning."
          WriteLog "Completed Scanning."
          StopScan
   End If
End Sub

Sub SetFonts(pForm As Object, pFontname As String, Optional Exclude1 As Object, Optional Exclude2 As Object, Optional Exclude3 As Object)
On Error Resume Next
Dim p As Control
Dim OldFont1, OldFont2, OldFont3
OldFont1 = Exclude1.Font.Name
OldFont2 = Exclude2.Font.Name
OldFont3 = Exclude3.Font.Name
For Each p In pForm.Controls
p.Font.Name = pFontname
Next p
Exclude1.Font.Name = OldFont1
Exclude2.Font.Name = OldFont2
Exclude3.Font.Name = OldFont3
End Sub
