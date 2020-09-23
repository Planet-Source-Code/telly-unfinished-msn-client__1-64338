Attribute VB_Name = "Msnmod"
Option Explicit

Public Type MSN_Contacts
    Email               As String
    Friendly_Name       As String
    Group               As String
    Active              As Boolean
    index               As Long
End Type

Public Type Message
    Caller              As String
    Challenge           As String
    SessionID           As String
End Type
Global ContactCount As Long, Contacts() As MSN_Contacts, RNG(1024) As Message, tID As Long, Username As String, Password As String, Status As String

Function Unescape(ByVal Enc As String) As String
Dim i As Long
For i = Len(Enc) To 1 Step -1
    If Mid$(Enc, i, 1) = "%" Then Enc = Replace$(Enc, Mid$(Enc, i, 3), Chr$(Asc(Chr$("&H" & Mid$(Enc, i + 1, 2)))))
Next i
Unescape = Enc
End Function

Function Escape(ByVal Enc As String) As String
Dim i As Long, tmp As String
Do
    i = i + 1
    tmp = Mid$(Enc, i, 1): If tmp = "" Then Exit Do
    If Asc(tmp) < 48 Then Enc = Replace$(Enc, tmp, "%" & Hex(Asc(Mid$(Enc, i))))
Loop
Escape = Enc
End Function

Function Typing(ByVal User As String) As String
'type trID ack len packet
Typing = "MSG " & TRID & " U " & CStr(Len(User) + 73) & vbCrLf & _
    "MIME-Version: 1.0" & vbCrLf & _
    "Content-Type: text/x-msmsgscontrol" & vbCrLf & _
"TypingUser: " & User & vbCrLf & vbCrLf
End Function

Function TRID() As String
If tID < 32767 Then tID = tID + 1 Else tID = 5
TRID = CStr(tID)
End Function
