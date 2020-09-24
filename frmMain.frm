VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API Examples By WhiteKnight    witenite87@excite.com http://camalot.virtualave.net"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command18 
      Caption         =   "Draw Edges 2"
      Height          =   255
      Left            =   3840
      TabIndex        =   29
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Restart Windows"
      Height          =   255
      Left            =   5400
      TabIndex        =   28
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Shut Down Windows"
      Height          =   255
      Left            =   5400
      TabIndex        =   27
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Command15 
      Caption         =   "LogOff Windows"
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Normal"
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Minimize"
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Maximize"
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Launch Web Browser"
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Download File"
      Height          =   255
      Left            =   5400
      TabIndex        =   21
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   19
      Text            =   "4"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   5400
      TabIndex        =   17
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Draw Edges 1"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cause Fatal Error"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Percent Example"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Flash On"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create Shortcut"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "C"
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Free Drive Space"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Drive Type"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   765
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encrypt / Decrypt"
      Height          =   1815
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   2535
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Encrypt"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Decrypt"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Encrypt Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4800
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4800
      Top             =   1800
   End
   Begin VB.Label Label5 
      Caption         =   "Thick:"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Key Watcher:"
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Drive:"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Yes I Used Text1 Text2 Command1 ect. Only because This is an example I usually dont _
do that and I don't sugest that you do it either.  I Call them What they Do (If a button exits the _
app then I'd Call it bttn_exit). This Is just an example of how to do Some things with API.  If you _
Like My example then please vote for me.  Why is it all in a class well I plan on making a dll with
'these functions and More too.  If you have any comments of sugestions please e-mail them _
me at witenite87@excite.com.  Feel Free to use this in any way you want.  You can Modify it any way _
you like as long as the comments remain and you give credit where credit is do.  These are not that hard _
of API Function I myself am just learning API.  Please Visit my site http://camalot.virtualave.net.
Dim bINOUT As Boolean
Dim thick As Integer
Private ddd As New clsAPIstuff

Private Sub Command1_Click()
Text1.Text = ddd.Get_DriveType(Text2.Text)
End Sub

Private Sub Command10_Click()
ddd.DownLoadFile "http://camalot.virtualave.net/vb/", "apiex.zip"
End Sub

Private Sub Command11_Click()
ddd.LaunchWebBrowser Me.hwnd, "http://camalot.virtualave.net"
End Sub

Private Sub Command12_Click()
ddd.MaxWindow (Me.hwnd)
End Sub

Private Sub Command13_Click()
ddd.MinWindow (Me.hwnd)
End Sub

Private Sub Command14_Click()
ddd.NormWindow (Me.hwnd)
End Sub

Private Sub Command15_Click()
ddd.LogOffWindows (True)
End Sub

Private Sub Command16_Click()
ddd.ShutDownWindows (True)
End Sub

Private Sub Command18_Click()
bINOUT = False
thick% = Text8.Text
Call Form_Paint
End Sub

Private Sub Command2_Click()
Text1.Text = ddd.FreeDiscSpace(Text2.Text)
End Sub

Private Sub Command3_Click()
Call ddd.DesktopShortcut("Char Map ShortCut Created By API Example By WhiteKnight", "C:\WINDOWS\CHARMAP.EXE", True)
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Flash On" Then Timer1.Enabled = True: Timer2.Enabled = False: Command4.Caption = "Flash Off"
    If Command4.Caption = "Flash Off" Then Timer1.Enabled = False: Timer2.Enabled = True: Command4.Caption = "Flash On"
End Sub

Private Sub Command5_Click()
Text1.Text = "This is step number 50 of 321.  Percent Complete is: " & ddd.GetPercent(50, 321, 100) & "%"
End Sub

Private Sub Command6_Click()
'You could use this if a user enters the wrong password or something
ddd.FatalErrorExit "You gone and Messed Up This Program!!!!!!!!!!!!!!!!!"
End Sub

Private Sub Command7_Click()
Text4.Text = ddd.Encrypt(Text3.Text, Text5.Text)
End Sub

Private Sub Command8_Click()
Text4.Text = ddd.Decrypt(Text4.Text, Text5.Text)
End Sub

Private Sub Command9_Click()
bINOUT = True
thick% = Text8.Text
Call Form_Paint
End Sub

Private Sub Form_Load()
thick = 3
bINOUT = True
End Sub

Private Sub Form_Paint()
ddd.DrawFormEdges Me, thick%, bINOUT
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
Text1.Text = ddd.GetPressedKey
End Sub

Private Sub Timer1_Timer()
Dim x
'You can do this a lot quicker with a 'Pause' sub but I didnt feel like adding one  hehehe
x = ddd.FlashWin(Me.hwnd, True)
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Dim x
'You can do this a lot quicker with a 'Pause' sub but I didnt feel like adding one  hehehe
x = ddd.FlashWin(Me.hwnd, False)
Timer1.Enabled = True
Timer2.Enabled = False
End Sub
