VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "xScanner"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -405
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "about.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   3555
      TabIndex        =   1
      Top             =   1710
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   582
      ButtonWidth     =   1482
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close "
            Object.ToolTipText     =   "Close this window"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   125
      Left            =   135
      TabIndex        =   0
      Top             =   1800
      Width           =   3210
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "about.frx":0626
      Top             =   225
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"about.frx":2498
      Height          =   870
      Left            =   180
      TabIndex        =   4
      Top             =   855
      Width           =   4290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Sushant Pandurangi (sushant@phreaker.net)"
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
      Left            =   840
      TabIndex        =   3
      Top             =   585
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xScanner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1440
      TabIndex        =   2
      Top             =   90
      Width           =   2040
   End
End
Attribute VB_Name = "frmAbout"
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
Option Explicit

Private Sub Form_Load()
Label1.Caption = "xScanner " & App.Major & "." & App.Minor
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Unload Me
End Sub
