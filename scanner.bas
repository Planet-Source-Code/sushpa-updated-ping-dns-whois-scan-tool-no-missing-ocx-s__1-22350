Attribute VB_Name = "Declarations"
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
'======================
Public lngNextPort As Integer
'======================
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_UNDO = &HC7

Public fMain As frmMain
Public Enum IP_STATUS
    
    IP_STATUS_BASE = 11000
    IP_SUCCESS = 0
    IP_BUF_TOO_SMALL = (11000 + 1)
    IP_DEST_NET_UNREACHABLE = (11000 + 2)
    IP_DEST_HOST_UNREACHABLE = (11000 + 3)
    IP_DEST_PROT_UNREACHABLE = (11000 + 4)
    IP_DEST_PORT_UNREACHABLE = (11000 + 5)
    IP_NO_RESOURCES = (11000 + 6)
    IP_BAD_OPTION = (11000 + 7)
    IP_HW_ERROR = (11000 + 8)
    IP_PACKET_TOO_BIG = (11000 + 9)
    IP_REQ_TIMED_OUT = (11000 + 10)
    IP_BAD_REQ = (11000 + 11)
    IP_BAD_ROUTE = (11000 + 12)
    IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
    IP_TTL_EXPIRED_REASSEM = (11000 + 14)
    IP_PARAM_PROBLEM = (11000 + 15)
    IP_SOURCE_QUENCH = (11000 + 16)
    IP_OPTION_TOO_BIG = (11000 + 17)
    IP_BAD_DESTINATION = (11000 + 18)
    IP_ADDR_DELETED = (11000 + 19)
    IP_SPEC_MTU_CHANGE = (11000 + 20)
    IP_MTU_CHANGE = (11000 + 21)
    IP_UNLOAD = (11000 + 22)
    IP_ADDR_ADDED = (11000 + 23)
    IP_GENERAL_FAILURE = (11000 + 50)
    MAX_IP_STATUS = 11000 + 50
    IP_PENDING = (11000 + 255)
    PING_TIMEOUT = 200
End Enum
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const ERROR_SUCCESS       As Long = 0
Private Const WS_VERSION_REQD     As Long = &H101
Private Const WS_VERSION_MAJOR    As Long = _
    WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR    As Long = _
   WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD    As Long = 1
Private Const SOCKET_ERROR        As Long = -1

Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Dim ICMPOPT As ICMP_OPTIONS

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long  'formerly integer
  '  Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Private Type HOSTENT
    hName      As Long
    hAliases   As Long
    hAddrType  As Integer
    hLen       As Integer
    hAddrList  As Long
End Type

Private Type WSADATA
    wVersion      As Integer
    wHighVersion  As Integer
    szDescription(0 To MAX_WSADescription)   As Byte
    szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
    wMaxSockets   As Integer
    wMaxUDPDG     As Integer
    dwVendorInfo  As Long
End Type

Private Declare Function gethostbyaddr Lib "wsock32.dll" (ByRef dwHost As Long, ByVal hLen As Integer, ByVal aType As Integer) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal szHost As String) As Long
Private Declare Function lstrlen Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" _
        (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function gethostname Lib "wsock32.dll" _
        (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)


Private Function HiByte(ByVal wParam As Integer)
HiByte = wParam \ &H1 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
LoByte = wParam And &HFF&
End Function

Private Sub SocketsCleanup()
If WSACleanup() <> ERROR_SUCCESS Then
    App.LogEvent "Socket error occurred in Cleanup.", vbLogEventTypeError
End If
End Sub

Private Function SocketsInitialize(Optional sErr As String) As Boolean
Dim WSAD As WSADATA, sLoByte As String, sHiByte As String
If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
    sErr = "The 32-bit Windows Socket is not responding."
    SocketsInitialize = False
    Exit Function
End If
If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
    sErr = "This application requires a minimum of " & _
            CStr(MIN_SOCKETS_REQD) & " supported sockets."
    SocketsInitialize = False
    Exit Function
End If

If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
        (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
        HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
    sHiByte = CStr(HiByte(WSAD.wVersion))
    sLoByte = CStr(LoByte(WSAD.wVersion))
    sErr = "Sockets version " & sLoByte & "." & sHiByte & _
            " is not supported by 32-bit Windows Sockets."
    SocketsInitialize = False
    Exit Function
End If
SocketsInitialize = True
End Function

Private Function DoPing(szAddress As String, sDataToSend As String, ECHO As ICMP_ECHO_REPLY, Optional Timeout As Long = PING_TIMEOUT) As Long
Dim hPort As Long, dwAddress As Long, iOpt As Long
    dwAddress = AddressStringToLong(szAddress)
hPort = IcmpCreateFile()
If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, _
    ECHO, Len(ECHO), Timeout) Then
    'the ping succeeded,
    '.Status will be 0
    '.RoundTripTime is the time in ms for
    '               the ping to complete,
    '.Data is the data returned (NULL terminated)
    '.Address is the Ip address that actually replied
    '.DataSize is the size of the string in .Data
    DoPing = IP_SUCCESS
Else
    If ECHO.status = 0 Then
        DoPing = -1
    Else
        DoPing = ECHO.status * -1
    End If
End If
IcmpCloseHandle (hPort)
End Function
   
Public Function AddressStringToLong(ByVal tmp As String) As Long
Dim i As Integer, parts(1 To 4) As String
i = 0
'we have to extract each part of the
'123.456.789.123 string, delimited by
'a period
While InStr(tmp, ".") > 0
    i = i + 1
    parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
    tmp = Mid(tmp, InStr(tmp, ".") + 1)
Wend
i = i + 1
parts(i) = tmp
If i <> 4 Then
    AddressStringToLong = 0
    Exit Function
End If
'build the long value out of the
'hex of the extracted strings
AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
                     Right("00" & Hex(parts(3)), 2) & _
                     Right("00" & Hex(parts(2)), 2) & _
                     Right("00" & Hex(parts(1)), 2))
   
End Function

Public Function GetStatusCode(status As IP_STATUS) As String
Dim msg As String
   Select Case status
      Case IP_SUCCESS:               msg = "IP Success"
      Case IP_BUF_TOO_SMALL:         msg = "IP Buffer too small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "IP Destination Net Unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "IP Destination Host Unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "IP Destination Protocol Unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "IP Destination Port Unreachable"
      Case IP_NO_RESOURCES:          msg = "IP No Resources"
      Case IP_BAD_OPTION:            msg = "IP Bad Option"
      Case IP_HW_ERROR:              msg = "IP Hardware Error"
      Case IP_PACKET_TOO_BIG:        msg = "IP Packet too big"
      Case IP_REQ_TIMED_OUT:         msg = "IP Reqquest timed out"
      Case IP_BAD_REQ:               msg = "IP Bad Request"
      Case IP_BAD_ROUTE:             msg = "IP Bad Route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "IP TTL Expired Transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "IP TTL Expired Reassem"
      Case IP_PARAM_PROBLEM:         msg = "IP Parameter Problem"
      Case IP_SOURCE_QUENCH:         msg = "IP Source Quench"
      Case IP_OPTION_TOO_BIG:        msg = "IP Option too big"
      Case IP_BAD_DESTINATION:       msg = "IP Bad Destination"
      Case IP_ADDR_DELETED:          msg = "IP Address Deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "IP Spec MTU Change"
      Case IP_MTU_CHANGE:            msg = "IP MTU Change"
      Case IP_UNLOAD:                msg = "IP Unload"
      Case IP_ADDR_ADDED:            msg = "IP Address Added"
      Case IP_GENERAL_FAILURE:       msg = "IP General Failure"
      Case IP_PENDING:               msg = "IP Pending"
      Case PING_TIMEOUT:             msg = "Ping timeout"
      Case -1:                       msg = "Destination host unreachable."
      Case Else:                     msg = "Unknown message returned"
   End Select
Debug.Print status
   GetStatusCode = msg
End Function

Public Function GetIPAddress(Optional sHost As String, _
Optional serrmsg As String) As String
'Resolves the host-name (or current machine if balnk) to an IP address
Dim sHostName   As String * 256
Dim lpHost      As Long
Dim HOST        As HOSTENT
Dim dwIPAddr    As Long
Dim tmpIPAddr() As Byte
Dim i           As Integer
Dim sIPAddr     As String
Dim werr        As Long
    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    If sHost = "" Then
        If gethostname(sHostName, 256) = SOCKET_ERROR Then
            werr = WSAGetLastError()
            GetIPAddress = ""
            serrmsg = "Windows Sockets error " & Str$(werr) & _
                " has occurred. Unable to successfully get Host Name." & vbCrLf
            GetIPAddress = ""
            SocketsCleanup
            Exit Function
        End If
        sHostName = Trim$(sHostName)
    Else
        sHostName = Trim$(sHost) & Chr$(0)
    End If
    lpHost = gethostbyname(sHostName)
    If lpHost = 0 Then
        werr = WSAGetLastError()
        GetIPAddress = ""
        serrmsg = "Windows Sockets error " & Str$(werr) & _
                " has occurred. Unable to successfully get Host Name." & vbCrLf
        GetIPAddress = ""
        SocketsCleanup
        Exit Function
    End If
    CopyMemory HOST, lpHost, Len(HOST)
    CopyMemory dwIPAddr, HOST.hAddrList, 4
    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
    For i = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
End Function

Public Function GetIPHostName() As String
'Returns the current machine's name
Dim sHostName As String * 256
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup
End Function

Public Function Ping(Address As String, RoundTripTime As String, DataMatch As Boolean, Optional DataSize As Long = 32, Optional Timeout As Long = PING_TIMEOUT) As Long
Dim ECHO As ICMP_ECHO_REPLY, pos As Integer, Dt As String, sAddress As String
On Error GoTo DPErr
   If AddressStringToLong(Address) = 0 Then
      sAddress = GetIPAddress(Address)
 Else
   sAddress = Address
End If
   If SocketsInitialize() Then
       If DataSize <= 0 Then DataSize = 32
      For pos = 1 To DataSize
      Dt = Dt & Chr$(Rnd() * 254 + 1)
    Next pos
        'ping an ip address, passing the
        'address and the ECHO structure
   Ping = DoPing(sAddress, Dt, ECHO, Timeout)
       'display the results from the ECHO structure
   RoundTripTime = ECHO.RoundTripTime & " ms"
   'DataSize = ECHO.DataSize & " bytes"
     If Left$(ECHO.Data, 1) <> Chr$(0) Then
       pos = InStr(ECHO.Data, Chr$(0))
     DataMatch = (Left$(ECHO.Data, pos - 1) = Dt)
    End If
     SocketsCleanup
    Else
      Ping = IP_GENERAL_FAILURE
    End If
     Exit Function
DPErr:
    Ping = IP_GENERAL_FAILURE
End Function

Private Function PointerToString(lpString As Long) As String
' The PointerToString function is used to convert a
' pointer to a string into a string variable:
Dim Buffer() As Byte
Dim nLen As Long
             If lpString Then
              nLen = lstrlen(lpString)
              If nLen Then
                 ReDim Buffer(0 To (nLen - 1)) As Byte
                 CopyMemory Buffer(0), ByVal lpString, nLen
                 PointerToString = StrConv(Buffer, vbUnicode)
              End If
           End If
End Function

Public Function GetHostFromIP(sIPAddr As String, Optional serrmsg As String) As String
' Finally, the GetHostFromIP function returns the host name
' from an IP address string:
'Resolves the IP address to a host name
        Dim dwIPAddr    As Long
        Dim lpHost      As Long
        Dim HOST        As HOSTENT
        Dim werr        As Long
            If Not SocketsInitialize() Then
                GetHostFromIP = ""
                Exit Function
            End If
            dwIPAddr = inet_addr(sIPAddr)
            lpHost = gethostbyaddr(dwIPAddr, Len(dwIPAddr), 2)
            If lpHost = 0 Then
                werr = WSAGetLastError()
                serrmsg = "Windows Sockets error " & Str$(werr) & _
                  " has occurred. Unable to successfully get Host Name." & vbCrLf
                GetHostFromIP = ""
                       SocketsCleanup
                Exit Function
            End If
            CopyMemory HOST, lpHost, Len(HOST)
            GetHostFromIP = PointerToString(HOST.hName)
                   SocketsCleanup
End Function

Public Function ReadValue(Section As String, Key As String, Optional Default As String)
    Dim sReturn As String
    sReturn = String(255, Chr(0))
    ReadValue = Left(sReturn, GetPrivateProfileString(Section, Key, Default, sReturn, Len(sReturn), App.Path & "\settings.ini"))
End Function

Public Sub SaveValue(Section As String, Key As String, Value As String)
    ' Write to INI file
    WritePrivateProfileString Section, Key, Value, App.Path & "\settings.ini"
End Sub

Public Function ParseInt(Expression As Variant) As Integer
If Expression = "" Then ParseInt = 0: Exit Function
'This will return only the Integer portion
Dim pos As Integer, TEMP As String
For pos = 1 To Len(Expression)
If IsNumeric(Mid(Expression, pos, 1)) = True Then TEMP = TEMP & CStr(Mid(Expression, pos, 1))
Next pos
ParseInt = CInt(TEMP)
End Function

Sub Main()
Set fMain = New frmMain
fMain.Show
End Sub

Function CBinary(Expression As Boolean) As Integer
If Expression = True Then CBinary = 1 Else CBinary = 0
End Function

Function WindowsDir() As String
    Dim strBuf As String
    strBuf = Space$(255)
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    If GetWindowsDirectory(strBuf, 255) > 0 Then
        WindowsDir = Left(strBuf, GetWindowsDirectory(strBuf, 255))
    Else
        WindowsDir = vbNullString
    End If
End Function

Public Function GetPort(PortName As String) As Integer
If PortName = "SyStat" Then GetPort = 11
If PortName = "NetStat" Then GetPort = 15
If PortName = "FTP" Then GetPort = 21
If PortName = "SSH" Then GetPort = 22
If PortName = "TelNet" Then GetPort = 23
If PortName = "STMP" Then GetPort = 25
If PortName = "Whois" Then GetPort = 43
If PortName = "Name Server" Then GetPort = 53
If PortName = "Finger" Then GetPort = 79
If PortName = "HTTP" Then GetPort = 80
If PortName = "Pop3" Then GetPort = 110
If PortName = "Ident" Then GetPort = 113
If PortName = "Secure HTTP" Then GetPort = 400
If PortName = "Wingate" Then GetPort = 1080
If PortName = "SubSeven (1243)" Then GetPort = 1243
If PortName = "DipStix" Then GetPort = 2002
If PortName = "NetBus" Then GetPort = 12345
If PortName = "SubSeven (27374)" Then GetPort = 27374
If PortName = "Back Orifice" Then GetPort = 31337
End Function

Public Sub Progress(ByVal bShowProgressBar As Boolean, pForm As Object, Progressbar As Object, Statusbar As Object, Panel As Integer)
    Dim tRC As RECT
    If bShowProgressBar Then
        SendMessageAny Statusbar.hWnd, SB_GETRECT, Panel, tRC
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
        With Progressbar
            SetParent .hWnd, Statusbar.hWnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 0
        End With
    Else
        SetParent Progressbar.hWnd, pForm.hWnd
        Progressbar.Visible = False
    End If
End Sub
