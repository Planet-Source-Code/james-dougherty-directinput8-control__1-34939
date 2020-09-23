VERSION 5.00
Object = "{7EF5B47A-424E-4C23-B3BB-489AF85F7F47}#1.0#0"; "XInput.ocx"
Begin VB.Form frmTest 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UltimaX XInput.ocx Demo"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8205
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   547
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin UltimaX_XInput.XInput XInput1 
      Height          =   2745
      Left            =   120
      TabIndex        =   58
      Top             =   240
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   4842
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Keyboard Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdConfig 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Configure"
         Height          =   255
         Index           =   0
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Esc"
         Height          =   375
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F1"
         Height          =   375
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F2"
         Height          =   375
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F3"
         Height          =   375
         Index           =   3
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F4"
         Height          =   375
         Index           =   4
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F5"
         Height          =   375
         Index           =   5
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F6"
         Height          =   375
         Index           =   6
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F7"
         Height          =   375
         Index           =   7
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F8"
         Height          =   375
         Index           =   8
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F9"
         Height          =   375
         Index           =   9
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F10"
         Height          =   375
         Index           =   10
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F11"
         Height          =   375
         Index           =   11
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F12"
         Height          =   375
         Index           =   12
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblKey 
         BackStyle       =   0  'Transparent
         Caption         =   "Press One Of The following keys On Your Keyboard."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   585
         Left            =   480
         TabIndex        =   27
         Top             =   360
         Width           =   2355
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Mouse Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton cmdConfig 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Configure"
         Height          =   255
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse X -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   840
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Y -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   825
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   2
         Left            =   990
         TabIndex        =   10
         Top             =   360
         Width           =   105
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   3
         Left            =   990
         TabIndex        =   9
         Top             =   720
         Width           =   105
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Z -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   5
         Left            =   990
         TabIndex        =   7
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left Click -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right Click -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scroll Click -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   9
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   105
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   10
         Left            =   1200
         TabIndex        =   2
         Top             =   1800
         Width           =   105
      End
      Begin VB.Label MousePos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   11
         Left            =   1230
         TabIndex        =   1
         Top             =   2160
         Width           =   105
      End
   End
   Begin VB.Timer InTimer 
      Interval        =   1
      Left            =   120
      Top             =   2760
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   3120
      Width           =   5055
      Begin VB.ComboBox cmbAvialJoy 
         Height          =   315
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avialible Joysticks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   80
         TabIndex        =   30
         Top             =   120
         Width           =   1530
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   2880
      TabIndex        =   45
      Top             =   4035
      Width           =   2295
      Begin VB.ListBox lstFX 
         Height          =   1815
         ItemData        =   "frmTest.frx":030A
         Left            =   120
         List            =   "frmTest.frx":030C
         TabIndex        =   46
         Top             =   400
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avialible Joystick Effects"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   75
         TabIndex        =   47
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   120
      TabIndex        =   31
      Top             =   4035
      Width           =   2730
      Begin VB.CommandButton cmdConfig 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Configure"
         Height          =   255
         Index           =   2
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joystick Informtion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   75
         TabIndex        =   44
         Top             =   120
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Version - "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label lblNFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DriveVer"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   0
         Left            =   1485
         TabIndex        =   42
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Firmware Revision -"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   240
         TabIndex        =   41
         Top             =   720
         Width           =   1530
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hardware Revision -"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   240
         TabIndex        =   40
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label lblNFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Firm"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   1
         Left            =   1800
         TabIndex        =   39
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblNFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hard"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   2
         Left            =   1800
         TabIndex        =   38
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avialible Buttons -"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   240
         TabIndex        =   37
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avialible Axis -"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avialible POV's -"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   240
         TabIndex        =   35
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label lblNFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Butt"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   3
         Left            =   1680
         TabIndex        =   34
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblNFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Axis"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   4
         Left            =   1440
         TabIndex        =   33
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lblNFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POV"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   5
         Left            =   1560
         TabIndex        =   32
         Top             =   1680
         Width           =   360
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Height          =   3330
      Left            =   5250
      TabIndex        =   51
      Top             =   3120
      Width           =   2820
      Begin VB.CommandButton cmdFX 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Shoot 44 Magnum"
         Height          =   375
         Index           =   4
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdFX 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Shoot Gatling Gun"
         Height          =   375
         Index           =   3
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdFX 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Shoot 9MM"
         Height          =   375
         Index           =   2
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdFX 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Shoot BB Gun"
         Height          =   375
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdFX 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Shoot Arrow"
         Height          =   375
         Index           =   0
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Play Force Feedback Effect"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   75
         TabIndex        =   52
         Top             =   135
         Width           =   2190
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IF YOU JUST GOT AN ERROR
'MAKE SURE TO RENAME XInput.oc to XInput.ocx
'(If you don't like pre-compiled .ocx's compile your own and just
'place it in this folder)

'In a real engine just make the XInput.ocx invisible when your
'ready to make the .exe. Until then, if you wish, you can leave
'it visible so you can see if any errors occour.

'Also note, in any game engine(whether you use the .ocx or pull
'the .cls out of the control) you must call "Joystick_EnumJoysticks"
'and "Joystick_EnumEffects" for it to work correctly. In and app.
'just put a combo box or list box on the form and set the visibility
'to false.

'Enjoy it,
'   James
Option Explicit

'This opens the devices configuration screen
Private Sub cmdConfig_Click(Index As Integer)
 Select Case Index
  Case 0: XInput1.Keyboard_RunControlPanel Me.hWnd
  Case 1: XInput1.Mouse_RunControlPanel Me.hWnd
  Case 2: XInput1.Joystick_RunControlPanel Me.hWnd
 End Select
End Sub

'Play some sample effects I made
Private Sub cmdFX_Click(Index As Integer)
 On Local Error Resume Next
 Select Case Index
  Case 0
   XInput1.Joystick_PlayPredefinedFX ShootArrow
  Case 1
   XInput1.Joystick_PlayPredefinedFX BB_Gun
  Case 2
   XInput1.Joystick_PlayPredefinedFX Gun_9MM
  Case 3
   XInput1.Joystick_PlayPredefinedFX Gatling_Gun
  Case 4
   XInput1.Joystick_PlayPredefinedFX Gun_44_Magnum
 End Select
End Sub

'Initialize the input engine
Private Sub Form_Load()
 XInput1.Initialize_InputEngine frmTest.hWnd
 ShowInput XInput1.Joystick_HasJoystick
End Sub

'In a real world game you would never use a timer but this
'fits the needs of the demo.

'Shows you how to detect what keys are presses and the mouse data
Private Sub InTimer_Timer()
 
 XInput1.Mouse_Update
 MousePos(2) = XInput1.Mouse_InputX
 MousePos(3) = XInput1.Mouse_InputY
 MousePos(5) = XInput1.Mouse_InputZ
 MousePos(9) = XInput1.Mouse_LeftClick
 MousePos(10) = XInput1.Mouse_RightClick
 MousePos(11) = XInput1.Mouse_WheelClick
 
 If XInput1.Keyboard_KeyState(X_Escape) <> 0 Then cmdKey_Click (0)
 If XInput1.Keyboard_KeyState(X_F1) <> 0 Then cmdKey_Click (1)
 If XInput1.Keyboard_KeyState(X_F2) <> 0 Then cmdKey_Click (2)
 If XInput1.Keyboard_KeyState(X_F3) <> 0 Then cmdKey_Click (3)
 If XInput1.Keyboard_KeyState(X_F4) <> 0 Then cmdKey_Click (4)
 If XInput1.Keyboard_KeyState(X_F5) <> 0 Then cmdKey_Click (5)
 If XInput1.Keyboard_KeyState(X_F6) <> 0 Then cmdKey_Click (6)
 If XInput1.Keyboard_KeyState(X_F7) <> 0 Then cmdKey_Click (7)
 If XInput1.Keyboard_KeyState(X_F8) <> 0 Then cmdKey_Click (8)
 If XInput1.Keyboard_KeyState(X_F9) <> 0 Then cmdKey_Click (9)
 If XInput1.Keyboard_KeyState(X_F10) <> 0 Then cmdKey_Click (10)
 If XInput1.Keyboard_KeyState(X_F11) <> 0 Then cmdKey_Click (11)
 If XInput1.Keyboard_KeyState(X_F12) <> 0 Then cmdKey_Click (12)

End Sub

'Highlight selected key pressed
Private Sub cmdKey_Click(Index As Integer)
 Dim i As Long
 
 For i = 0 To 12
  cmdKey(i).BackColor = &HE0E0E0
 Next
 cmdKey(Index).BackColor = vbRed
 
End Sub

' If a joystick is found then gather information and enable he buttons
Private Sub ShowInput(IsAvialible As Boolean)
 Dim i As Long

 'If we have a joystick then
 If IsAvialible = True Then
  cmdConfig(2).Enabled = True
  XInput1.Joystick_EnumJoysticks cmbAvialJoy
  cmbAvialJoy.ListIndex = 0
  
  'If it has force feedback then
  If XInput1.Joystick_HasForceFeedback Then
   XInput1.Joystick_EnumEffects lstFX
  Else
   lstFX.AddItem "No Force Feedback"
  End If
  
  For i = 0 To 4
   cmdFX(i).Enabled = True
  Next
  
  DisplayNFO
 Else
  cmdConfig(2).Enabled = False
  cmbAvialJoy.AddItem "No avialable devices attached"
  cmbAvialJoy.ListIndex = 0
  lstFX.AddItem "No Force Feedback"
  
  For i = 0 To 5
   lblNFO(i).Caption = "N/A"
  Next
  
  For i = 0 To 4
   cmdFX(i).Enabled = False
  Next
  
 End If

End Sub

'Called in ShowInput()
'Just gathers information about the avialable joystick
Private Sub DisplayNFO()
 On Local Error Resume Next

 lblNFO(0).Caption = XInput1.Joystick_GetDriverVersion
 lblNFO(1).Caption = XInput1.Joystick_GetFirmwareRevision
 lblNFO(2).Caption = XInput1.Joystick_GetHardwareRevision
 lblNFO(3).Caption = XInput1.Joystick_NumberOfButtons
 lblNFO(4).Caption = XInput1.Joystick_NumberOfAxis
 lblNFO(5).Caption = XInput1.Joystick_NumberOfPOVs
 
End Sub
