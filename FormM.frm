VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormM 
   Caption         =   "D-MSN"
   ClientHeight    =   4335
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TypingTimeout 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   20000
      Left            =   60
      Top             =   3600
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4080
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Chatlist 
      Height          =   4035
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7117
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   1
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   795
   End
   Begin MSWinsockLib.Winsock Switchboard 
      Index           =   0
      Left            =   2280
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   360
      TabIndex        =   1
      Top             =   7140
      Width           =   4215
   End
   Begin VB.Timer cPing 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1680
      Top             =   2700
   End
   Begin MSWinsockLib.Winsock SSL 
      Left            =   1080
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock http 
      Index           =   0
      Left            =   540
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   1035
   End
   Begin MSWinsockLib.Winsock Client 
      Index           =   0
      Left            =   0
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   21
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormM.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormM.frx":03CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormM.frx":0719
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormM.frx":0ACB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormM.frx":169D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormM.frx":1A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormM.frx":1E17
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormM.frx":21C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mainopt 
      Caption         =   "Options"
      Begin VB.Menu testune 
         Caption         =   "test unescape"
      End
      Begin VB.Menu mLog 
         Caption         =   "Login"
      End
   End
   Begin VB.Menu chatmenu 
      Caption         =   "chatmenu"
      Visible         =   0   'False
      Begin VB.Menu mConv 
         Caption         =   "Start Conversation"
      End
      Begin VB.Menu infView 
         Caption         =   "View Info"
      End
   End
End
Attribute VB_Name = "FormM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Buffer(10) As String, hBuffer(10) As String, HTTP_Header As String, Auth_Challenge As String
Dim Auth_Login As String, Ticket As String, curIndex As Integer
 Dim C As Long

Private Sub Chatlist_DblClick()
LoadIM Chatlist.SelectedItem.Text
End Sub

Private Sub Chatlist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu chatmenu
End Sub

Private Sub Chatlist_NodeClick(ByVal Node As MSComctlLib.Node)
'Chatlist.Nodes.Remove Node.index
End Sub

Private Sub Client_Connect(index As Integer)
MsnSend index, "VER 1 MSNP8 CVR0"
AddStatus Me, "Connected to Server"
End Sub

Private Sub Client_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim Packet As String, tmp() As String, n As Long
Client(index).GetData Packet
Buffer(index) = Buffer(index) & Packet
tmp = Split(Buffer(index), vbCrLf)
For n = 0 To UBound(tmp) - 1
    Handle index, tmp(n)
    Buffer(index) = Replace$(Buffer(index), tmp(n) & vbCrLf, "")
Next n
End Sub

Private Sub Client_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Debug.Print "Error: "; index; Description
Client(index).Close
End Sub

Private Sub Command1_Click()
Client(0).Connect "messenger.hotmail.com", 1863
End Sub

Private Sub Command2_Click()
'Dim S As String
'S = "MIME-Version: 1.0" & _
'"Content-Type: text/x-msmsgscontrol" & _
'"TypingUser: email@email.net" & _
'vbCrLf & vbCrLf
'MsgBox Len(S)
MsgBox Len("MIME-Version: 1.0\rContent-Type: text/x-msmsgscontrol\rTypingUser: \r\r\n")
End Sub

Private Sub cPing_Timer()
MsnSend 0, "PNG"
End Sub

Private Sub Form_Load()
Dim n As Long
For n = 1 To 1024
    Load Switchboard(n)
    Load TypingTimeout(n)
Next n
For n = 1 To 10
    Load http(n)
    Load Client(n)
Next n
ReDim Contacts(0)


End Sub

Sub MsnSend(index As Integer, ByVal Packet As String)
Client(index).SendData Packet & vbCrLf
End Sub

Sub Handle(index As Integer, ByVal Packet As String)
Debug.Print "hnd: "; index, Packet
Dim pType As String, tmp() As String, tmp1() As String
pType = Mid(Packet, 1, 3)
Select Case pType
    Case "VER"
        AddStatus Me, "Setting Version"
        MsnSend index, "CVR 2 0x0409 winnt 5.1 i386 MSNMSGR 7.0.0816 MSMSGS " & Username
    Case "CVR"
        MsnSend index, "USR 3 TWN I " & Username
    Case "XFR"
        tmp = Split(Packet, " ")
        Client(index).Close
        tmp1 = Split(tmp(3), ":")
        Client(index + 1).Connect tmp1(0), tmp1(1)
    Case "USR"
        tmp = Split(Packet, " ")
        Select Case tmp(2)
        Case "TWN"
            AddStatus Me, "Authorising"
            Auth_Challenge = tmp(4)
            Debug.Print Auth_Challenge
            HTTP_Header = "GET https://nexus.passport.com/rdr/pprdr.asp" & vbCrLf
            SSL.Connect "nexus.passport.com", 443
            curIndex = index
        Case "OK"
            Me.Caption = tmp(3) & " [" & tmp(4) & "]"
            MsnSend index, "SYN 8 6"
        End Select
    Case "SYN"
        tmp = Split(Packet, " ")
        ContactCount = tmp(3)
    Case "LST"
        tmp = Split(Packet, " ")
        ReDim Preserve Contacts(C + 1)
        C = UBound(Contacts)
        Contacts(C).Email = tmp(1)
        Contacts(C).Friendly_Name = tmp(2)
        AssignIndex Contacts(C).Email
         List1.AddItem Unescape(Contacts(C).Friendly_Name)
        Chatlist.Nodes.Add , , , Unescape(Contacts(C).Friendly_Name), 1
        If UBound(Contacts) = ContactCount Then MsnSend index, "CHG 5 NLN 1073856564"
        AddStatus Me, "Recieving Contacts"



    Case "MSG"
    Case "RNG"
        tmp = Split(Packet, " ")
        'RNG 14422 207.46.4.198:1863 CKI 1128549075.11374 tel@xxzcxc.net tel
        SB_Connect tmp(2), tmp(1), tmp(5), tmp(4)
    Case "CHL"
        tmp = Split(Packet, " ")
        Client(index).SendData "QRY 1049 msmsgs@msnmsgr.com 32" & vbCrLf & CalculateMD5(tmp(2) & "Q1P7W2E4J9R8U3S5")
    Case Else
    
End Select
End Sub

Sub AssignIndex(ByVal Email As String)
Dim n As Long
For n = 1 To UBound(Contacts) - 1
    If Contacts(n).Email = Email Then
        
    End If
Next n
End Sub

Sub LoadIM(ByVal Email As String)
Dim IMWin As Form, n As Long
For Each IMWin In Forms
If IMWin.Tag = Email Then
    FlashWin IMWin.hWnd, FLASHW_TRAY
    Exit Sub
End If
Next
Set IMWin = New IM
IMWin.Tag = Email
IMWin.Caption = "Conversation - " & Email
IMWin.Show
End Sub

'--------------------------------------- Switchboard Sockets ---------------------------------'

Sub SB_Connect(ByVal Address As String, ByVal SessionID As String, ByVal Caller As String, ByVal Challenge As String)
Dim tmp() As String, n As Long
tmp = Split(Address, ":")
For n = 0 To Switchboard.UBound - 1
    If Switchboard(n).State = sckClosed Then
        Switchboard(n).Connect tmp(0), tmp(1)
        RNG(n).Caller = Caller
        RNG(n).Challenge = Challenge
        RNG(n).SessionID = SessionID
        Exit For
    End If
Next n
End Sub

Private Sub Form_Resize()
On Error Resume Next
Chatlist.Width = ScaleWidth
Chatlist.Height = ScaleHeight - (StatusBar1.Height)
StatusBar1.Width = ScaleWidth
StatusBar1.Panels(1).Width = ScaleWidth
End Sub

Private Sub infView_Click()
MsgBox Chatlist.SelectedItem.Text
End Sub

Private Sub List1_DblClick()
LoadIM List1.Text
End Sub

Private Sub mConv_Click()
LoadIM Chatlist.SelectedItem.Text
End Sub

Private Sub mLog_Click()
FormL.Show
End Sub

Private Sub Switchboard_Connect(index As Integer)
SB_Send index, "ANS 1 " & Username & " " & RNG(index).Challenge & " " & RNG(index).SessionID
End Sub

Sub SB_Send(index As Integer, ByVal Packet As String)
Switchboard(index).SendData Packet & vbCrLf
End Sub

Private Sub Switchboard_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim Packet As String
Switchboard(index).GetData Packet
Debug.Print "SB: "; index; Packet
Parse index, Packet
End Sub

Sub Parse(index As Integer, ByVal Packet As String)
Dim Args() As String, n As Long, Params() As String, Data() As String, Typing As Boolean
Args = Split(Packet, vbCrLf)
Data = Split(Args(0), " ")
For n = 1 To UBound(Args) - 1
    If Args(n) <> "" Then
        Params = Split(Args(n), ":")
        Select Case Params(0)
            Case "MIME-Version"
            Case "TypingUser"
                Typing = True
                'If Typing = True Then AddStatus Params(1) & " is typing a message"
                'Me.Caption = Params(1) & " is typing a message"
                'Message index, 2, Params(1)
                Message index, 2, Data(1), Data(2), Params(1) & " is typing a message"
            Case "Content-Type"
                Select Case Params(1)
                    Case " text/x-msnmsgr-datacast"
                        'AddStatus Data(2) & " sent you a nudge!!"
                        Message index, 3, Data(1), Data(2), Data(2) & " sent you a nudge!!"
                    Case Else
                        ''text/plain; charset=UTF-8
                End Select
            Case "X-MMS-IM-Format"
                'FontFormat Params(1)
        End Select
    End If
Next n
If Args(n) <> "" Then
    'LoadIM Data(1)
    Message index, 1, Data(1), Data(2), Args(n)
End If
End Sub

Sub Message(index As Integer, ByVal mType As Long, Optional ByVal Email As String, Optional ByVal FriendlyName As String, Optional ByVal Msg As String)
'0
'1 - text
'2 - typing
'3 - Nudge
Dim IMWin As Form, n As Long, FoundWin As Boolean, Win As Long
Select Case mType
Case 0
Case 1

For Each IMWin In Forms
If IMWin.Tag = Email Then
    FoundWin = True
    IMWin.RichTextBox1.Text = IMWin.RichTextBox1.Text & vbCrLf & FriendlyName & " says: " & vbCrLf & Space(3) & Msg
    IMWin.StatusBar1.Panels(1).Text = "Last Message Recieved @ " & Now
    Exit Sub
End If
Next
If FoundWin = False Then
    LoadIM Email
    For Each IMWin In Forms
        If IMWin.Tag = Email Then
            IMWin.RichTextBox1.Text = IMWin.RichTextBox1.Text & vbCrLf & FriendlyName & " says: " & vbCrLf & Space(3) & Msg
            IMWin.StatusBar1.Panels(1).Text = "Last Message Recieved @ " & Now
            Exit Sub
        End If
    Next
End If
Case 2
For Each IMWin In Forms
If IMWin.Tag = Email Then
    IMWin.StatusBar1.Panels(1).Text = Msg
    Exit Sub
End If
Next
Case 3
For Each IMWin In Forms
    If IMWin.Tag = Email Then
        FoundWin = True
        IMWin.RichTextBox1.Text = IMWin.RichTextBox1.Text & vbCrLf & Msg
        IMWin.StatusBar1.Panels(1).Text = "Last Message Recieved @ " & Now
        Exit Sub
    End If
Next
If FoundWin = False Then
    LoadIM Email
    For Each IMWin In Forms
        If IMWin.Tag = Email Then
             IMWin.RichTextBox1.Text = IMWin.RichTextBox1.Text & vbCrLf & Msg
             IMWin.StatusBar1.Panels(1).Text = "Last Message Recieved @ " & Now
            Exit Sub
        End If
    Next
End If
Case Else
End Select
End Sub

Sub AddStatus(frm As Form, ByVal Msg As String)
frm.StatusBar1.Panels(1).Text = Msg
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim n As Long
For n = 0 To 10
    If http(n).State <> sckClosed Then http(n).Close
    If Client(n).State <> sckClosed Then Client(n).Close
Next n
End
End Sub

'--------------------------------- HTTP Socks --------------------------------'
Private Sub http_Connect(index As Integer)
httpSend index, HTTP_Header
End Sub

Private Sub http_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim Packet As String, tmp() As String, n As Long
http(index).GetData Packet
Debug.Print "headers: "; Packet
hBuffer(index) = hBuffer(index) & Packet
If Right(hBuffer(index), 4) = vbCrLf & vbCrLf Then
tmp = Split(Buffer(index), vbCrLf)
For n = 0 To UBound(tmp) - 1
    Headers index, tmp(n)
    hBuffer(index) = ""
Next n
End If
End Sub

Sub httpSend(index As Integer, ByVal Packet As String)
http(index).SendData Packet
End Sub

Sub Headers(index As Integer, ByVal Packet As String)
Debug.Print index; Packet
End Sub

'------------------------------------- SSL Sockets ---------------------------------'
'SSLv2 for VB, coded by Jason K. Resch & Seth Taylor

Private Sub SSL_Close()
    Me.Caption = "Closed."
    SSL.Close
    If Layer = 3 Then
        Layer = 4
        Call SSL_DataArrival(0)
    End If
    Layer = 0
    Set SecureSession = Nothing
End Sub

Private Sub SSL_Connect()
    Processing = False
    Set SecureSession = New CryptoCls
    Call SendClientHello(SSL)
End Sub

Private Sub SSL_DataArrival(ByVal bytesTotal As Long)
 Dim TheData As String
    Dim Response As String
    Response = ""
    
    ' Buffer incoming data while connection is open or being opened
    If Layer < 4 Then
        Call SSL.GetData(TheData, vbString, bytesTotal)
        DataBuffer = DataBuffer & TheData
    End If
    
    If Layer = 3 Then
        ' Download complete response before processing
        Exit Sub
    End If
    
    'Parse each SSL Record
    Do
    
        If SeekLen = 0 Then
            If Len(DataBuffer) >= 2 Then
                TheData = GetBufferDataPart(2)
                SeekLen = BytesToLen(TheData)
            Else
                Exit Sub
            End If
        End If
        
        If Len(DataBuffer) >= SeekLen Then
            TheData = GetBufferDataPart(SeekLen)
        Else
            Exit Sub
        End If
        
        
        Select Case Layer
            Case 0:
                ENCODED_CERT = Mid(TheData, 12, BytesToLen(Mid(TheData, 6, 2)))
                CONNECTION_ID = Right(TheData, BytesToLen(Mid(TheData, 10, 2)))
                Call IncrementRecv
                Call SendMasterKey(SSL)
            Case 1:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If Right(TheData, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
                    If VerifyMAC(TheData) Then
                        Call SendClientFinish(SSL)
                    Else
                        ' SSL Error -- send SSL error to server
                        MsgBox ("SSL Error: Invalid MAC data ... aborting connection.")
                        SSL.Close
                    End If
                Else
                    ' SSL Error -- send SSL error to server
                    MsgBox ("SSL Error: Invalid Challenge data ... aborting connection.")
                    SSL.Close
                End If
             Case 2:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If VerifyMAC(TheData) = False Then
                    ' SSL Error -- send SSL error to server
                    MsgBox ("SSL Error: Invalid MAC data ... aborting connection.")
                    SSL.Close
                End If
                Layer = 3
                DoEvents
                SSLSend SSL, HTTP_Header & vbCrLf
             Case 3:
                ' Do nothing while buffer is filled ... wait for connection to close
             Case 4:
                'SSLSend SSL, HTTP_Header & vbCrLf
                TheData = SecureSession.RC4_Decrypt(TheData)
                If VerifyMAC(TheData) Then
                    Response = Response & Mid(TheData, 17)
                Else
                    ' SSL Error -- data is corrupt and must be discarded
                    MsgBox ("SSL Error: Invalid MAC data ... Data discarded.")
                    Layer = 0
                    DataBuffer = ""
                    Response = ""
                    Exit Sub
                End If
        End Select
        SeekLen = 0
    Loop Until Len(DataBuffer) = 0
    
    If Layer = 4 Then
        Layer = 0
        Handle_SSL Response
    End If
   ' SSLSend SSL, HTTP_Header & vbCrLf
End Sub

Sub Handle_SSL(ByVal Packet As String)
Dim Headers() As String, Params() As String, Args() As String, n As Long, l As Long
Debug.Print Packet
Headers = Split(Packet, vbCrLf)
For n = 0 To UBound(Headers) - 2
    If Headers(n) <> "" Then
    Params = Split(Headers(n), ":")
    Select Case Params(0)
    Case "PassportURLs"
        Args = Split(Params(1), ",")
        Auth_Login = Mid(Args(1), 9)
        Args = Split(Auth_Login, "/")
        Debug.Print Auth_Login
        HTTP_Header = "GET /" & Args(1) & " HTTP/1.1" & vbCrLf & _
            "Authorization: Passport1.4 OrgVerb=GET,OrgURL=http%3A%2F%2Fmessenger%2Emsn%2Ecom,sign-in=" & Username & ",pwd=" & Password & "," & Auth_Challenge & vbCrLf & _
            "Host: " & Args(0) & vbCrLf
        SSL.Connect Args(0), 443
    Case "Authentication-Info"
        Args = Split(Params(1), ",")
        For l = 0 To UBound(Args) - 1
        If Mid(Args(l), 1, 9) = "from-PP='" Then
            Ticket = Mid(Args(l), 10, Len(Mid(Args(l), 10)) - 1)
            Exit For
        End If
        Next l
        Debug.Print Ticket
        MsnSend curIndex, "USR 4 TWN S " & Ticket
    End Select
    End If
Next n
End Sub


Function GetBufferDataPart(ByVal Length As Long) As String
    Dim l As Long
    l = Len(DataBuffer)
    If Length > l Then
        Length = l
        GetBufferDataPart = Left(DataBuffer, l)
    Else
        GetBufferDataPart = Left(DataBuffer, Length)
    End If
    If Length = l Then
        DataBuffer = ""
    Else
        DataBuffer = Mid(DataBuffer, Length + 1)
    End If
End Function

Private Sub testune_Click()
MsgBox Unescape("hello%20%world%25")
End Sub
