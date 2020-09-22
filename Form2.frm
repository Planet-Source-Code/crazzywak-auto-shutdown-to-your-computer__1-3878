VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "About - Read me"
   ClientHeight    =   2715
   ClientLeft      =   4710
   ClientTop       =   3090
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2715
   ScaleWidth      =   3270
   Begin VB.CommandButton Command2 
      Caption         =   "Back To The Program"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "                      Enter My Site!                    http://crazzywak.cjb.net"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form2.frx":0000
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell ("explorer http://crazzywak.cjb.net")
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
