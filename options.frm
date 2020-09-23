VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Options"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   Icon            =   "options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmUser 
      Caption         =   "U&ser..."
      Height          =   330
      Left            =   3150
      TabIndex        =   14
      Top             =   45
      Width           =   735
   End
   Begin VB.CommandButton cmHistory 
      Caption         =   "Hi&story..."
      Height          =   375
      Left            =   2790
      TabIndex        =   13
      Top             =   1080
      Width           =   1005
   End
   Begin MSComctlLib.ListView lvWSer 
      Height          =   1140
      Left            =   1035
      TabIndex        =   11
      Top             =   1530
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   2011
      View            =   3
      Arrange         =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "iml"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "WhoIs servers (click to add)"
         Object.Width           =   4057
      EndProperty
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   3555
      Top             =   540
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
            Picture         =   "options.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "options.frx":0328
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "options.frx":0644
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "options.frx":07A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txpTo 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   1755
      TabIndex        =   9
      Text            =   "32767"
      Top             =   1170
      Width           =   690
   End
   Begin VB.TextBox txPFrom 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   585
      TabIndex        =   7
      Text            =   "1"
      Top             =   1170
      Width           =   645
   End
   Begin VB.CheckBox chScanDefault 
      Caption         =   "&Scan common ports by default"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   675
      Value           =   1  'Checked
      Width           =   2760
   End
   Begin VB.CheckBox chAutoSave 
      Caption         =   "Au&tomatically save log file to disk"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   450
      Value           =   1  'Checked
      Width           =   2760
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Preferences"
      Height          =   125
      Left            =   225
      TabIndex        =   1
      Top             =   135
      Width           =   2850
   End
   Begin MSComctlLib.Toolbar Tbar 
      Height          =   330
      Left            =   1395
      TabIndex        =   0
      Top             =   2745
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   582
      ButtonWidth     =   1482
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "iml"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save "
            Object.ToolTipText     =   "Saves the settings"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close "
            Object.ToolTipText     =   "Closes this window"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit  "
            Object.ToolTipText     =   "Edit the settings file"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Build your server list in the list to the right."
      Height          =   1005
      Left            =   135
      TabIndex        =   12
      Top             =   1620
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3285
      Picture         =   "options.frx":08FC
      Top             =   450
      Width           =   480
   End
   Begin VB.Label lStat 
      Caption         =   "xScanner "
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   2835
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Up to:"
      Height          =   195
      Left            =   1305
      TabIndex        =   8
      Top             =   1230
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "Fr&om:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1235
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "Port Scan&ning (custom)"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   945
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User &Preferences"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   135
      Width           =   1230
   End
End
Attribute VB_Name = "frmOpts"
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

Private Sub cmHistory_Click()
If MsgBox("History is a list of Internet Hosts that you have Pinged, Used DNS conversions on, or scanned the ports of. If you clear this history, the WhoIs server list is not cleared. Do you want to clear history?", vbYesNo + vbQuestion, "History") = vbYes Then
fMain.txScanHost.ComboItems.Clear
fMain.cbPing.ComboItems.Clear
End If
End Sub

Private Sub cmUser_Click()
Dim ps As String
ps = InputBox("Please enter your name.", "User", Right(fMain.Caption, Len(fMain.Caption) - 11))
If ps = "" Then Exit Sub
fMain.Caption = "xScanner - " & ps
End Sub

Private Sub Form_Load()
GetPrefs
lStat.Caption = "xScanner " & App.Major & "." & App.Minor
lvWSer.ColumnHeaders(1).Width = lvWSer.Width - 285
End Sub

Private Sub lvWSer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim ps As String
ps = InputBox("Enter the name of the WhoIs server you want to add to the list. Do not include an http:// in case there is one.", "Add Server")
If ps = "" Then Exit Sub
lvWSer.ListItems.Add , , ps, , 4
End Sub

Private Sub lvWSer_DblClick()
On Error Resume Next
If lvWSer.SelectedItem Is Nothing Then Exit Sub
If MsgBox("Do you want to remove:" & vbNewLine & lvWSer.SelectedItem.Text & vbNewLine & "from your list of servers?", vbQuestion + vbYesNo, "Remove") = vbYes Then lvWSer.ListItems.Remove lvWSer.SelectedItem.Index
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 2 Then
Unload Me
ElseIf Button.Index = 1 Then
SetPrefs
ElseIf Button.Index = 3 Then
ManualEdit
End If
End Sub

Sub SetPrefs()
fMain.tmrSave.Enabled = CBool(chAutoSave.Value)
SaveValue "Settings", "ScanFrom", txPFrom.Text
SaveValue "Settings", "ScanTo", txpTo.Text
SaveValue "Settings", "ScanCommon", CBool(Me.chScanDefault.Value)
fMain.SetPrefs
Dim p As Integer
For p = 1 To lvWSer.ListItems.Count
SaveValue "WServers", "Server" & p, lvWSer.ListItems.Item(p).Text
Next p
SaveValue "WServers", "Count", lvWSer.ListItems.Count
fMain.GetPrefs
Unload Me
End Sub

Sub ManualEdit()
ShellExecute Me.hWnd, "open", WindowsDir & "\notepad.exe", App.Path & "\settings.ini", App.Path, 10
End Sub

Sub GetPrefs()
chAutoSave.Value = CBinary(ReadValue("Settings", "AutoSave", True))
chScanDefault.Value = CBinary(ReadValue("Settings", "ScanCommon", True))
txPFrom.Text = ReadValue("Settings", "ScanFrom", 7)
txpTo.Text = ReadValue("Settings", "ScanTo", 520)
Dim p As Integer
For p = 1 To ReadValue("WServers", "Count", 0)
lvWSer.ListItems.Add , , ReadValue("WServers", "Server" & p, ""), , 4
Next p
End Sub
