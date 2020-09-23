VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form IM 
   Caption         =   "Form2"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9315
   LinkTopic       =   "Form2"
   ScaleHeight     =   7710
   ScaleWidth      =   9315
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   4260
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   582
      ButtonWidth     =   1852
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Font"
            Key             =   "font"
            Object.ToolTipText     =   "Change Font"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Emote"
            Key             =   "smiley"
            Object.ToolTipText     =   "Emote"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nudge"
            Key             =   "nudge"
            Object.ToolTipText     =   "Send Nudge!"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "File Send"
            Key             =   "files"
            Object.ToolTipText     =   "Send File"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Send!"
            Key             =   "send"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7740
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483638
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IM.frx":0000
            Key             =   "a"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IM.frx":0E52
            Key             =   "s"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IM.frx":0EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IM.frx":1488
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   7470
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7329
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"IM.frx":1A22
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   4620
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   2566
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"IM.frx":1AA4
   End
   Begin VB.Menu mIMOpt 
      Caption         =   "Options"
   End
End
Attribute VB_Name = "IM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rX As Double, rY As Double, Moving As Boolean

Private Sub Form_Activate()
FlashWin Me.hWnd, FLASHW_STOP
End Sub

Private Sub Form_GotFocus()
FlashWin Me.hWnd, FLASHW_STOP
End Sub

Private Sub Form_Initialize()
FlashWin Me.hWnd, FLASHW_STOP
End Sub

Private Sub Form_Resize()
On Error Resume Next
RichTextBox1.Width = ScaleWidth - 1
RichTextBox1.Height = ScaleHeight / 1.4
Toolbar1.Top = (RichTextBox1.Top + RichTextBox1.Height) + 25
Toolbar1.Width = RichTextBox1.Width
RichTextBox2.Top = (Toolbar1.Top + Toolbar1.Height) + 25
RichTextBox2.Width = (ScaleWidth) - 25
RichTextBox2.Height = (ScaleHeight - RichTextBox2.Top) - (StatusBar1.Height)
StatusBar1.Panels(1).Width = ScaleWidth
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
rX = X
rY = Y
Moving = True
Me.MousePointer = vbCrosshair
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Moving Then RichTextBox1.Height = X
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False: Screen.MousePointer = 0
End Sub

Private Sub RichTextBox2_GotFocus()
FlashWin Me.hWnd, FLASHW_STOP
End Sub

