VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autoshut"
   ClientHeight    =   3195
   ClientLeft      =   4695
   ClientTop       =   2970
   ClientWidth     =   2130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2130
   Begin VB.CommandButton Command2 
      Caption         =   "Deactivate"
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Activate"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time To Shutdown:"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2055
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type:"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2055
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   840
         Top             =   960
      End
      Begin VB.OptionButton optExitOption 
         Caption         =   "Normal Shutdown"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optExitOption 
         Caption         =   "Forced Shutdown"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton optExitOption 
         Caption         =   "Log Off"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton optExitOption 
         Caption         =   "Reboot"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Menu mnuAboutItem 
      Caption         =   "About..."
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Win32 Then
    Private Declare Function Shutdown Lib "user32" Alias "ExitWindowsEx" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
    Private Const EWX_LOGOFF = 0
    Private Const EWX_SHUTDOWN = 1
    Private Const EWX_REBOOT = 2
    Private Const EWX_FORCE = 4
#Else
    Private Declare Function Shutdown Lib "User" Alias "ExitWindows" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer
    Private Const EW_REBOOTSYSTEM = &H43
    Private Const EW_RESTARTWINDOWS = &H42
#End If

Private SelectedOption As Integer

Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = True
Timer1.Enabled = True
Text1.Enabled = False
End Sub

Private Sub Command2_Click()
Command1.Enabled = True
Command2.Enabled = False
Timer1.Enabled = False
Text1.Enabled = True
End Sub

Private Sub Form_Load()
    #If Win32 Then
        ' Prepare for 32 bit ExitWindowsEx.
        optExitOption(0).Caption = "Normal Shutdown"
        optExitOption(0).Tag = EWX_SHUTDOWN

        optExitOption(1).Caption = "Reboot"
        optExitOption(1).Tag = EWX_REBOOT
        
        optExitOption(2).Caption = "Log Off"
        optExitOption(2).Tag = EWX_LOGOFF
        
        optExitOption(3).Caption = "Forced Shutdown"
        optExitOption(3).Tag = EWX_FORCE
    #Else
        ' Prepare for 16 bit ExitWindows.
        optExitOption(0).Caption = "Normal Shutdown"
        optExitOption(0).Tag = 0

        optExitOption(1).Caption = "Reboot"
        optExitOption(1).Tag = EW_REBOOTSYSTEM

        optExitOption(2).Caption = "Restart Windows"
        optExitOption(2).Tag = EW_RESTARTWINDOWS
        
        optExitOption(3).Visible = False
    #End If
    Text1.Text = Time
End Sub

Private Sub mnuAboutItem_Click()
Load Form2
Form2.Show
End Sub

Private Sub Timer1_Timer()
Dim CurrentTime
CurrentTime = Format(Time, "hh:mm")
If CurrentTime = Text1.Text Then
Dim exit_option As Long
exit_option = CLng(optExitOption(SelectedOption).Tag)
Shutdown exit_option, 0
Timer1.Enabled = False
End
End If
End Sub
