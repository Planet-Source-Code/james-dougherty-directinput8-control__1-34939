VERSION 5.00
Begin VB.UserControl XInput 
   BackColor       =   &H00000000&
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   ScaleHeight     =   183
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   188
   ToolboxBitmap   =   "XInput.ctx":0000
   Begin VB.Label lblLogo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UltimaX"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   945
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   2340
   End
   Begin VB.Label lblDisable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enabled"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label OCXType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input Engine"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   683
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   870
   End
   Begin VB.Label lblPreError 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Error Detected -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   1305
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "None"
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
      Height          =   690
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   2265
   End
End
Attribute VB_Name = "XInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
'|¶¶             © 2001-2002 Ariel Productions          ¶¶|'
'|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
'|¶¶             Programmer - James Dougherty           ¶¶|'
'|¶¶             Source - XInput.ctl                    ¶¶|'
'|¶¶             Object - UltimaX.dll                   ¶¶|'
'|¶¶             Version - 2.1                          ¶¶|'
'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

'This control will display an error(if there happens to be one)
'directly on the control its self, I hate message boxes popping
'up.

'For explaination of the Input Function look at the UXInput.cls

Option Explicit

Public Enum XInputKey
 X_1 = 2
 X_2 = 3
 X_3 = 4
 X_4 = 5
 X_5 = 6
 X_6 = 7
 X_7 = 8
 X_8 = 9
 X_9 = 10
 X_0 = 11
 X_A = &H1E
 X_B = &H30
 X_C = &H2E
 X_D = &H20
 X_E = &H12
 X_F = &H21
 X_G = &H22
 X_H = &H23
 X_I = &H17
 X_J = &H24
 X_K = &H25
 X_L = &H26
 X_M = &H32
 X_N = &H31
 X_O = &H18
 X_P = &H19
 X_Q = &H10
 X_R = &H13
 X_S = &H1F
 X_T = &H14
 X_U = &H16
 X_V = &H2F
 X_W = &H11
 X_X = &H2D
 X_Y = &H15
 X_Z = &H2C
 X_F1 = &H3B
 X_F2 = &H3C
 X_F3 = &H3D
 X_F4 = &H3E
 X_F5 = &H3F
 X_F6 = &H40
 X_F7 = &H41
 X_F8 = &H42
 X_F9 = &H43
 X_F10 = &H44
 X_F11 = &H57
 X_F12 = &H58
 X_Num1 = &H4F
 X_Num2 = &H50
 X_Num3 = &H51
 X_Num4 = &H4B
 X_Num5 = &H4C
 X_Num6 = &H4D
 X_Num7 = &H47
 X_Num8 = &H48
 X_Num9 = &H49
 X_Num0 = &H52
 X_NumEnter = &H9C
 X_UP = &HC8
 X_Down = &HD0
 X_Left = &HCB
 X_Right = &HCD
 X_Escape = 1
 X_Enter = &H1C
 X_LShift = &H2A
 X_RShift = &H36
 X_LControl = &H1D
 X_RControl = &H9D
 X_Space = &H39
 X_Insert = &HD2
 X_Delete = &HD3
 X_Home = &HC7
 X_End = &HCF
 X_PageUp = &HC9
 X_PageDown = &HD1
 X_BackSpace = 14
 X_Add = &H4E
 X_Subtract = &H4A
 X_Period = &H34
 X_Tab = 15
End Enum

Public Enum FXDirection
 North = 0
 North_East = 1
 East = 2
 South_East = 3
 South = 4
 South_West = 5
 West = 6
 North_West = 7
End Enum

Public Enum FXType
 ShootArrow = 0
 Gun_44_Magnum = 1
 Gun_9MM = 2
 BB_Gun = 3
 Gatling_Gun = 4
End Enum

Public Enum UXBorderStyle
 None = 0
 Fixed_Single = 1
End Enum

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Private XIN As New UXInput
Private InitOK As Boolean

Public Property Get BackColor() As OLE_COLOR
 On Local Error Resume Next
 BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
 On Local Error Resume Next
 UserControl.BackColor() = New_BackColor
 PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As UXBorderStyle
 On Local Error Resume Next
 BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As UXBorderStyle)
 On Local Error Resume Next
 UserControl.BorderStyle() = New_BorderStyle
 PropertyChanged "BorderStyle"
End Property

Public Property Get Enabled() As Boolean
 On Local Error Resume Next
 Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 On Local Error Resume Next
 If Not New_Enabled Then lblDisable.Caption = "Disabled": lblDisable.ForeColor = vbRed Else lblDisable.Caption = "Enabled": lblDisable.ForeColor = vbGreen
 UserControl.Enabled() = New_Enabled
 PropertyChanged "Enabled"
End Property

Private Sub UserControl_Click()
 On Local Error Resume Next
 RaiseEvent Click
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 On Local Error Resume Next
 RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
 On Local Error Resume Next
 RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
 On Local Error Resume Next
 RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 On Local Error Resume Next
 UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
 UserControl.BackColor = PropBag.ReadProperty("BackColor", vbBlack)
 UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
 If Not UserControl.Enabled Then lblDisable.Caption = "Disabled": lblDisable.ForeColor = vbRed Else lblDisable.Caption = "Enabled": lblDisable.ForeColor = vbGreen
End Sub

Private Sub UserControl_Resize()
 On Local Error Resume Next
 UserControl.Width = 2820
 UserControl.Height = 2745
End Sub

Private Sub UserControl_Terminate()
 On Local Error Resume Next
 Set XIN = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 On Local Error Resume Next
 Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
 Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbBlack)
 Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
 If Not UserControl.Enabled Then lblDisable.Caption = "Disabled": lblDisable.ForeColor = vbRed Else lblDisable.Caption = "Enabled": lblDisable.ForeColor = vbGreen
End Sub

'|¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤|
'|¤¤                    Main Functions                        ¤¤|
'|¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤|

Public Function Initialize_InputEngine(hWnd As Long) As Boolean
 On Local Error GoTo errOut
 
 Initialize_InputEngine = XIN.Initialize_Input_Engine(hWnd)
 If Not Initialize_InputEngine Then GoTo errOut
 DoEvents
 InitOK = True
 Exit Function
 
errOut:
 lblError = "Unable To Initialize Engine"
End Function

'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|
'|ºº                     Keyboard                           ºº|
'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|

Public Function Keyboard_KeyState(Key As XInputKey) As Boolean
 If Not UserControl.Enabled Then Exit Function
 Keyboard_KeyState = XIN.Keyboard_KeyState(Key)
End Function

Public Sub Keyboard_RunControlPanel(hWnd As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Keyboard_RunControlPanel hWnd
End Sub

'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|
'|ºº                       Mouse                            ºº|
'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|

Public Sub Mouse_RunControlPanel(hWnd As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Mouse_RunControlPanel hWnd
End Sub

Public Sub Mouse_Update()
 If Not UserControl.Enabled Then Exit Sub
 XIN.Mouse_UpdateState
End Sub

Public Function Mouse_InputX() As Long
 If Not UserControl.Enabled Then Exit Function
 Mouse_InputX = XIN.Mouse_InputX
End Function

Public Function Mouse_InputY() As Long
 If Not UserControl.Enabled Then Exit Function
 Mouse_InputY = XIN.Mouse_InputY
End Function

Public Function Mouse_InputZ() As Long
 If Not UserControl.Enabled Then Exit Function
 Mouse_InputZ = XIN.Mouse_InputZ
End Function

Public Function Mouse_LeftClick() As Long
 If Not UserControl.Enabled Then Exit Function
 Mouse_LeftClick = XIN.Mouse_LeftClick
End Function

Public Function Mouse_RightClick() As Long
 If Not UserControl.Enabled Then Exit Function
 Mouse_RightClick = XIN.Mouse_RightClick
End Function

Public Function Mouse_WheelClick() As Long
 If Not UserControl.Enabled Then Exit Function
 Mouse_WheelClick = XIN.Mouse_WheelClick
End Function

'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|
'|ºº                      JoyStick                          ºº|
'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|

Public Sub Joystick_RunControlPanel(hWnd As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_RunControlPanel hWnd
End Sub

Public Function Joystick_GetDriverVersion() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_GetDriverVersion = XIN.Joystick_GetDriverVersion
End Function

Public Function Joystick_GetFirmwareRevision() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_GetFirmwareRevision = XIN.Joystick_GetFirmwareRevision
End Function

Public Function Joystick_GetHardwareRevision() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_GetHardwareRevision = XIN.Joystick_GetHardwareRevision
End Function

Public Function Joystick_HasJoystick() As Boolean
 If Not UserControl.Enabled Then Exit Function
 Joystick_HasJoystick = XIN.Joystick_HasJoystick
End Function

Public Function Joystick_HasForceFeedback() As Boolean
 If Not UserControl.Enabled Then Exit Function
 Joystick_HasForceFeedback = XIN.Joystick_HasForceFeedback
End Function

Public Function Joystick_EnumJoysticks(ListBox As Object) As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_EnumJoysticks = XIN.Joystick_EnumJoysticks(ListBox)
End Function

Public Sub Joystick_EnumEffects(ListBox As Object)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_EnumEffects ListBox
End Sub

Public Sub Joystick_Update()
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_Update
End Sub

Public Function Joystick_NumberOfButtons() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_NumberOfButtons = XIN.Joystick_NumberOfButtons
End Function

Public Function Joystick_NumberOfAxis() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_NumberOfAxis = XIN.Joystick_NumberOfAxis
End Function

Public Function Joystick_NumberOfPOVs() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_NumberOfPOVs = XIN.Joystick_NumberOfPOVs
End Function

Public Function Joystick_Button(Button As Long) As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_Button = XIN.Joystick_Button(Button)
End Function

Public Function Joystick_POV() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_POV = XIN.Joystick_POV
End Function

Public Function Joystick_XAxis() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_XAxis = XIN.Joystick_XAxis
End Function

Public Function Joystick_YAxis() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_YAxis = XIN.Joystick_YAxis
End Function

Public Function Joystick_ZAxis() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_ZAxis = XIN.Joystick_ZAxis
End Function

Public Function Joystick_RotXAxis() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_RotXAxis = XIN.Joystick_RotXAxis
End Function

Public Function Joystick_RotYAxis() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_RotYAxis = XIN.Joystick_RotYAxis
End Function

Public Function Joystick_RotZAxis() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_RotZAxis = XIN.Joystick_RotZAxis
End Function

Public Function Joystick_Slider0() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_Slider0 = XIN.Joystick_Slider0
End Function

Public Function Joystick_Slider1() As Long
 If Not UserControl.Enabled Then Exit Function
 Joystick_Slider1 = XIN.Joystick_Slider1
End Function

Public Sub Joystick_TurnOffAutocenter()
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_TurnOffAutocenter
End Sub

Public Sub Joystick_SetFXStart(FXIndex As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXStart FXIndex
End Sub

Public Sub Joystick_SetFXUnload(FXIndex As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXUnload FXIndex
End Sub

Public Sub Joystick_SetFXStop(FXIndex As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXStop FXIndex
End Sub

Public Sub Joystick_SetFXEnvelopeEffect(FXIndex As Long, AttackLevel As Long, AttackTime As Long, _
                                        FadeLevel As Long, FadeTime As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXEnvelopeEffect FXIndex, AttackLevel, AttackTime, FadeLevel, FadeTime
End Sub

Public Sub Joystick_SetFXDuration(FXIndex As Long, Duration As Long, Optional Infinite As Boolean = False)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXDuration FXIndex, Duration, Infinite
End Sub

Public Sub Joystick_SetFXGain(FXIndex As Long, Gain As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXGain FXIndex, Gain
End Sub

Public Sub Joystick_SetFXSampleRate(FXIndex As Long, Rate As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXSampleRate FXIndex, Rate
End Sub

Public Sub Joystick_SetFXConstantForce(FXIndex As Long, Force As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXConstantForce FXIndex, Force
End Sub

Public Sub Joystick_SetFXDirection(FXIndex As Long, Direction As FXDirection)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXDirection FXIndex, Direction
End Sub

Public Sub Joystick_SetFXConditionX(FXIndex As Long, DeadBand As Long, _
                                    NegCoeff As Long, NegSat As Long, _
                                    PosCoeff As Long, PosSat As Long, Offset As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXConditionX FXIndex, DeadBand, NegCoeff, NegSat, PosCoeff, PosSat, Offset
End Sub

Public Sub Joystick_SetFXConditionY(FXIndex As Long, DeadBand As Long, _
                                    NegCoeff As Long, NegSat As Long, _
                                    PosCoeff As Long, PosSat As Long, Offset As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXConditionY FXIndex, DeadBand, NegCoeff, NegSat, PosCoeff, PosSat, Offset
End Sub

Public Sub Joystick_SetFXRampForce(FXIndex As Long, StartRange As Long, EndRange As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXRampForce FXIndex, StartRange, EndRange
End Sub

Public Sub Joystick_SetFXPeriodicForce(FXIndex As Long, Magnitude As Long, Offset As Long, Period As Long, Phase As Long)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_SetFXPeriodicForce FXIndex, Magnitude, Offset, Period, Phase
End Sub

Public Sub Joystick_PlayPredefinedFX(Effect As FXType)
 If Not UserControl.Enabled Then Exit Sub
 XIN.Joystick_PlayPredefinedFX Effect
End Sub

