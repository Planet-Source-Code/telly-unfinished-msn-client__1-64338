VERSION 5.00
Begin VB.Form FormL 
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   2955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   2955
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   3
         Text            =   "NLN"
         Top             =   1140
         Width           =   2115
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Top             =   780
         Width           =   2595
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   2595
      End
   End
End
Attribute VB_Name = "FormL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Username = Text1
Password = Text2
FormM.Client(0).Connect "messenger.hotmail.com", 1863
FormM.Chatlist.Nodes.Clear
Me.Hide
End Sub
