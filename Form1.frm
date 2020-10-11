VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menampilkan Utilitas Control Panel"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command20 
      Caption         =   "Command5"
      Height          =   495
      Left            =   6480
      TabIndex        =   19
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Command4"
      Height          =   495
      Left            =   6480
      TabIndex        =   18
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Command3"
      Height          =   495
      Left            =   6480
      TabIndex        =   17
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6480
      TabIndex        =   16
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6480
      TabIndex        =   15
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Command5"
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command4"
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command5"
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command4"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  'Menampilkan Control Panel
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL ", vbNormalFocus)
End Sub

Private Sub Command2_Click()
  'Menampilkan Accessibility Properties
  Call Shell("rundll32.exe shell32.dll,Control_RunDLL access.cpl", vbNormalFocus)
End Sub

Private Sub Command3_Click()
  'Menampilkan Add/Remove Programs
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL appwiz.cpl", vbNormalFocus)
End Sub

Private Sub Command4_Click()
  'Menampilkan Display Settings (Background)
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL desk.cpl,,0", vbNormalFocus)
End Sub

Private Sub Command5_Click()
  'Menampilkan Display Settings (Screensaver)
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL desk.cpl,,1", vbNormalFocus)
End Sub

Private Sub Command6_Click()
  'Menampilkan the Display Settings (Appearance)
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL desk.cpl,,2", vbNormalFocus)
End Sub

Private Sub Command7_Click()
  'Menampilkan the Display Settings (Settings)
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL desk.cpl,,3", vbNormalFocus)
End Sub

Private Sub Command8_Click()
  'Menampilkan Internet Properties
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL inetcpl.cpl", vbNormalFocus)
End Sub

Private Sub Command9_Click()
  'Menampilkan Regional Settings
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL intl.cpl", vbNormalFocus)
End Sub

Private Sub Command10_Click()
  'Menampilkan Joystick Settings
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL joy.cpl", vbNormalFocus)
End Sub

Private Sub Command11_Click()
  'Menampilkan Mouse Settings
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL main.cpl @0", vbNormalFocus)
End Sub

Private Sub Command12_Click()
  'Menampilkan Keyboard Settings
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL main.cpl @1", vbNormalFocus)
End Sub

Private Sub Command13_Click()
  'Menampilkan Printers
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL main.cpl @2", vbNormalFocus)
End Sub

Private Sub Command14_Click()
  'Menampilkan Fonts
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL main.cpl @3", vbNormalFocus)
End Sub

Private Sub Command15_Click()
  'Menampilkan Multimedia Settings
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL mmsys.cpl", vbNormalFocus)
End Sub

Private Sub Command16_Click()
  'Menampilkan Modem Settings
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL modem.cpl", vbNormalFocus)
End Sub

Private Sub Command17_Click()
  'Menampilkan Dial-Up Networking Wizard (pada Win9x)
  Call Shell("rundll32.exe rnaui.dll, RnaWizard", vbNormalFocus)
End Sub

Private Sub Command18_Click()
  'Menampilkan System Properties
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL sysdm.cpl", vbNormalFocus)
End Sub

Private Sub Command19_Click()
  'Menjalankan 'Add New Hardware' Wizard (pada Win9x)
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL sysdm.cpl @1", vbNormalFocus)
End Sub

Private Sub Command20_Click()
  'Menampilkan 'Add New Printer' Wizard (pada Win9x)
  Call Shell("rundll32.exe shell32.dll, SHHelpShortcuts_RunDLL AddPrinter", vbNormalFocus)
End Sub

Private Sub Command21_Click()
  'Menampilkan Themes Settings
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL themes.cpl", vbNormalFocus)
End Sub

Private Sub Command22_Click()
  'Menampilkan Time/Date Settings
  Call Shell("rundll32.exe shell32.dll, Control_RunDLL timedate.cpl", vbNormalFocus)
End Sub


