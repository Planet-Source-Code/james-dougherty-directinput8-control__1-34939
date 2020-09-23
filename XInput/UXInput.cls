VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UXInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
'|¶¶             © 2001-2002 Ariel Productions          ¶¶|'
'|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
'|¶¶             Programmer - James Dougherty           ¶¶|'
'|¶¶             Source - UXInput.cls                   ¶¶|'
'|¶¶             Object - UltimaX.dll                   ¶¶|'
'|¶¶             Version - 2.1                          ¶¶|'
'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

'This was pulled out of my updated engine.

Option Explicit

Private DX As New DirectX8
Private DI As DirectInput8
Private DIEnum As DirectInputEnumDevices8
Private DIDev As DirectInputDevice8
Private DIState As DIKEYBOARDSTATE
Private DIMouse As DirectInputDevice8
Private DIMState As DIMOUSESTATE
Private DIJoy As DirectInputDevice8
Private DIJOYSTATE As DIJOYSTATE
Private DIEffEnum As DirectInputEnumEffects
Private DIFFEffect() As DirectInputEffect
Private DIEffectType As Long
Private DIFFStaticParams As Long
Private DIJoyPropRange As DIPROPRANGE
Private DIJoyPropDead As DIPROPLONG
Private DIJoyPropSaturation As DIPROPLONG
Private EffectParams() As Long
Private AxisPresent(1 To 8) As Boolean
Private LastEffect As Long
Private FFEffect As DIEFFECT
Private IsJoystick As Boolean
Private KeyboardCaps As DIDEVCAPS
Private MouseCaps As DIDEVCAPS
Private JoyCaps As DIDEVCAPS
Private WasEnumerated As Boolean
Private mhWnd As Long

'Programmer access to the DirectInput8 object
Public Function DXInput8() As DirectInput8
 Set DXInput8 = DI
End Function

'Initialize the keyboard, mouse, and joystick(If one is avialable)
Public Function Initialize_Input_Engine(hWnd As Long) As Boolean
 On Local Error GoTo errOut
 
 Set DI = DX.DirectInputCreate()
 
 Set DIDev = DI.CreateDevice("GUID_SysKeyboard")
 DIDev.SetCommonDataFormat DIFORMAT_KEYBOARD
 DIDev.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
 DIDev.GetCapabilities KeyboardCaps
 DIDev.Acquire
 
 Set DIMouse = DI.CreateDevice("GUID_SysMouse")
 DIMouse.SetCommonDataFormat DIFORMAT_MOUSE
 DIMouse.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
 DIMouse.GetCapabilities MouseCaps
 DIMouse.Acquire
 
 Set DIEnum = DI.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)
 If DIEnum.GetCount > 0 Then
  Set DIJoy = DI.CreateDevice(DIEnum.GetItem(1).GetGuidInstance)
  DIJoy.SetCommonDataFormat DIFORMAT_JOYSTICK
  DIJoy.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  SetDefaultJoystickProperties
  IdentifyAxes
  DIJoy.GetCapabilities JoyCaps
  DIJoy.Acquire
  IsJoystick = True
 End If
 
 mhWnd = hWnd
 Initialize_Input_Engine = True
 Exit Function
 
errOut:
 Initialize_Input_Engine = False
End Function

'Notice this is private and its called under Class_Terminate
Private Sub Cleanup_Input_Engine()
 On Local Error Resume Next
 DIDev.Unacquire
 DIMouse.Unacquire
 If IsJoystick Then DIJoy.Unacquire: Set DIJoy = Nothing
 Set DIDev = Nothing
 Set DIMouse = Nothing
 Set DI = Nothing
 Set DX = Nothing
End Sub

'Cleanup when the class terminates, let it do the work
Private Sub Class_Terminate()
 Cleanup_Input_Engine
End Sub

'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|
'|ºº                     Keyboard                           ºº|
'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|

'Check what keys are being pressed
Public Function Keyboard_KeyState(Key As XInputKey) As Boolean
 Keyboard_KeyState = DIState.Key(Key)
 DIDev.GetDeviceStateKeyboard DIState
End Function

'Shows the devices configuration
Public Sub Keyboard_RunControlPanel(hWnd As Long)
 DIDev.RunControlPanel hWnd
End Sub

'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|
'|ºº                       Mouse                            ºº|
'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|

'Shows the devices configuration
Public Sub Mouse_RunControlPanel(hWnd As Long)
 DIMouse.RunControlPanel hWnd
End Sub

'Call this in your game loop so it always checks the mouse state
Public Sub Mouse_UpdateState()
 On Local Error Resume Next
 DIMouse.GetDeviceStateMouse DIMState
End Sub

'X Coord
Public Function Mouse_InputX() As Long
 Mouse_InputX = DIMState.lX
End Function

'Y Coord
Public Function Mouse_InputY() As Long
 Mouse_InputY = DIMState.lY
End Function

'Z Coord (When the scroll is clicked then slide forward or backward)
Public Function Mouse_InputZ() As Long
 Mouse_InputZ = DIMState.lZ
End Function

'Left click
Public Function Mouse_LeftClick() As Long
 Mouse_LeftClick = DIMState.Buttons(0)
End Function

'Right click
Public Function Mouse_RightClick() As Long
 Mouse_RightClick = DIMState.Buttons(1)
End Function

'Scroll Click
Public Function Mouse_WheelClick() As Long
 Mouse_WheelClick = DIMState.Buttons(2)
End Function

'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|
'|ºº                      JoyStick                          ºº|
'|ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº|

'Joysticks and Force Feedback are a devious task, very hard to
'explain. For complete explaination please see Microsofts Demo

'The function names pretty much explain what the purpose is
'The main thing to be concerned with is making sure the users
'system has a joystick and if it has force feedback. Thats why
'theres a lot of error handling. Also the error handling makes
'sure we do not loose the device and the feedback effects values
'are in the range of the joystick.

Public Sub Joystick_RunControlPanel(hWnd As Long)
 If Not IsJoystick Then Exit Sub
 DIJoy.RunControlPanel hWnd
End Sub

Public Function Joystick_GetDriverVersion() As Long
 If Not IsJoystick Then Exit Function
 Joystick_GetDriverVersion = JoyCaps.lDriverVersion
End Function

Public Function Joystick_GetFirmwareRevision() As Long
 If Not IsJoystick Then Exit Function
 Joystick_GetFirmwareRevision = JoyCaps.lFirmwareRevision
End Function

Public Function Joystick_GetHardwareRevision() As Long
 If Not IsJoystick Then Exit Function
 Joystick_GetHardwareRevision = JoyCaps.lHardwareRevision
End Function

Public Function Joystick_HasJoystick() As Boolean
 Joystick_HasJoystick = IsJoystick
End Function

Public Function Joystick_HasForceFeedback() As Boolean
 Set DIEnum = Nothing
 Set DIEnum = DI.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY Or DIEDFL_FORCEFEEDBACK)
 If DIEnum.GetCount = 0 Then
  Joystick_HasForceFeedback = False
 Else
  Joystick_HasForceFeedback = True
 End If
End Function

Private Sub SetDefaultJoystickProperties()

 With DIJoyPropDead
  .lData = 1000
  .lHow = DIPH_BYOFFSET
  .lObj = DIJOFS_X
  DIJoy.SetProperty "DIPROP_DEADZONE", DIJoyPropDead
  .lObj = DIJOFS_Y
  DIJoy.SetProperty "DIPROP_DEADZONE", DIJoyPropDead
 End With
    
 With DIJoyPropSaturation
  .lData = 9500
  .lHow = DIPH_BYOFFSET
  .lObj = DIJOFS_X
  DIJoy.SetProperty "DIPROP_SATURATION", DIJoyPropSaturation
  .lObj = DIJOFS_Y
  DIJoy.SetProperty "DIPROP_SATURATION", DIJoyPropSaturation
 End With
    
 With DIJoyPropRange
  .lHow = DIPH_DEVICE
  .lMin = 0
  .lMax = 10000
  DIJoy.SetProperty "DIPROP_RANGE", DIJoyPropRange
 End With

End Sub

Public Function Joystick_EnumJoysticks(ListBox As Object) As Long
 If Not IsJoystick Then Exit Function
 Dim i As Integer

 Set DIEnum = Nothing
 Set DIEnum = DI.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)
 If DIEnum.GetCount = 0 Then
  ListBox.AddItem "No avialable devices attached"
  Joystick_EnumJoysticks = 0
  Exit Function
 Else
  For i = 1 To DIEnum.GetCount
   Call ListBox.AddItem(DIEnum.GetItem(i).GetInstanceName)
  Next
  Joystick_EnumJoysticks = 1
 End If

End Function

Public Sub Joystick_EnumEffects(ListBox As Object)
 If Not IsJoystick Then Exit Sub
 Dim i As Integer
 Set DIEffEnum = DIJoy.GetEffectsEnum(DIEFT_ALL)

 For i = 1 To DIEffEnum.GetCount
  DIEffectType = DIEffEnum.GetType(i) And &HFF
  DIFFStaticParams = DIEffEnum.GetStaticParams(i)
                                                   
  If (DIEffectType = DIEFT_HARDWARE) And (DIFFStaticParams And DIEP_TYPESPECIFICPARAMS) <> 0 Then
   GoTo Ignore
  ElseIf DIEffectType = DIEFT_CUSTOMFORCE Then
   GoTo Ignore
  End If

  ListBox.AddItem DIEffEnum.GetName(i)
  ReDim Preserve EffectParams(ListBox.ListCount - 1)
  EffectParams(ListBox.ListCount - 1) = DIEffectType
  ReDim Preserve DIFFEffect(ListBox.ListCount - 1)
  On Local Error GoTo ErrorHandler
  Set DIFFEffect(UBound(DIFFEffect)) = DIJoy.CreateEffect(DIEffEnum.GetEffectGuid(i), CreateDefaultFX(i))
  DIFFEffect(UBound(DIFFEffect)).Unload
Ignore:
 Next
 WasEnumerated = True
 Exit Sub

ErrorHandler:
    
 If Err.Number = 5 Then
  ListBox.RemoveItem ListBox.ListCount - 1
  ReDim Preserve DIFFEffect(ListBox.ListCount - 1)
  Resume Next
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If

End Sub

Private Function CreateDefaultFX(Index As Integer) As DIEFFECT
    
 With CreateDefaultFX
  .lDuration = &HFFFFFFFF
  .lGain = 10000
  .lSamplePeriod = 0
  .lTriggerButton = DIEB_NOTRIGGER
  .lTriggerRepeatInterval = -1
  .constantForce.lMagnitude = 10000
  .rampForce.lRangeStart = 0
  .rampForce.lRangeEnd = 0
  .conditionFlags = DICONDITION_USE_BOTH_AXES
  With .conditionX
   .lDeadBand = 0
   .lNegativeSaturation = 10000
   .lOffset = 0
   .lPositiveSaturation = 10000
  End With
  With .conditionY
   .lDeadBand = 0
   .lNegativeSaturation = 10000
   .lOffset = 0
   .lPositiveSaturation = 10000
  End With
  With .periodicForce
   .lMagnitude = 10000
   .lOffset = 0
   .lPeriod = 1
   .lPhase = 0
  End With
 End With
    
End Function

Private Sub IdentifyAxes()
 Dim didoEnum As DirectInputEnumDeviceObjects
 Dim dido As DirectInputDeviceObjectInstance
 Dim sGuid As String
 Dim i As Integer
   
 For i = 1 To 8
  AxisPresent(i) = False
 Next
   
 Set didoEnum = DIJoy.GetDeviceObjectsEnum(DIDFT_AXIS)
 For i = 1 To didoEnum.GetCount
  Set dido = didoEnum.GetItem(i)
  sGuid = dido.GetGuidType
  Select Case sGuid
   Case "GUID_XAxis"
    AxisPresent(1) = True
   Case "GUID_YAxis"
    AxisPresent(2) = True
   Case "GUID_ZAxis"
    AxisPresent(3) = True
   Case "GUID_RxAxis"
    AxisPresent(4) = True
   Case "GUID_RyAxis"
    AxisPresent(5) = True
   Case "GUID_RzAxis"
    AxisPresent(6) = True
   Case "GUID_Slider"
    AxisPresent(7) = True
    AxisPresent(8) = True
  End Select
 Next
End Sub

Private Sub CheckLoss()
 If Err.Number = DIERR_NOTACQUIRED Or Err.Number = DIERR_INPUTLOST Then
  DIJoy.Acquire
  Exit Sub
 End If
End Sub

Public Sub Joystick_Update()
 If Not IsJoystick Then Exit Sub
 DIJoy.Poll
 CheckLoss
End Sub

Public Function Joystick_NumberOfButtons() As Long
 If Not IsJoystick Then Exit Function
 Joystick_NumberOfButtons = JoyCaps.lButtons
End Function

Public Function Joystick_NumberOfAxis() As Long
 If Not IsJoystick Then Exit Function
 Joystick_NumberOfAxis = JoyCaps.lAxes
End Function

Public Function Joystick_NumberOfPOVs() As Long
 If Not IsJoystick Then Exit Function
 Joystick_NumberOfPOVs = JoyCaps.lPOVs
End Function

Public Function Joystick_Button(Button As Long) As Long
 If Not IsJoystick Then Exit Function
 If Button > Joystick_NumberOfButtons Then Exit Function
 If Button <= 0 Then Exit Function
 Select Case DIJOYSTATE.Buttons(Button - 1)
  Case 0
   Joystick_Button = 0
  Case Else
   Joystick_Button = 1
 End Select
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
End Function

Public Function Joystick_POV() As Long
 If Not IsJoystick Then Exit Function
 Dim i As Long
 For i = 0 To JoyCaps.lPOVs - 1
  Joystick_POV = DIJOYSTATE.POV(i)
 Next
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_XAxis() As Long
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(1) Then Exit Function
 Joystick_XAxis = DIJOYSTATE.x
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_YAxis() As Long
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(2) Then Exit Function
 Joystick_YAxis = DIJOYSTATE.y
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_ZAxis() As Long
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(3) Then Exit Function
 Joystick_ZAxis = DIJOYSTATE.z
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_RotXAxis() As Long
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(4) Then Exit Function
 Joystick_RotXAxis = DIJOYSTATE.rx
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_RotYAxis() As Long
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(5) Then Exit Function
 Joystick_RotYAxis = DIJOYSTATE.ry
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_RotZAxis() As Long
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(6) Then Exit Function
 Joystick_RotZAxis = DIJOYSTATE.rz
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_Slider0() As Long
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(7) Then Exit Function
 Joystick_Slider0 = DIJOYSTATE.slider(0)
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Function Joystick_Slider1() As Long
 If Not IsJoystick Then Exit Function
 If Not AxisPresent(8) Then Exit Function
 Joystick_Slider1 = DIJOYSTATE.slider(1)
 DIJoy.GetDeviceStateJoystick DIJOYSTATE
 CheckLoss
End Function

Public Sub Joystick_TurnOffAutocenter()
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 Dim Prop As DIPROPLONG

 Prop.lData = 0
 Prop.lHow = DIPH_DEVICE
 Prop.lObj = 0
 DIJoy.SetProperty "DIPROP_AUTOCENTER", Prop
End Sub

Public Sub Joystick_SetFXStart(FXIndex As Long)
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub
 DIFFEffect(FXIndex).Start 1, 0
End Sub

Public Sub Joystick_SetFXUnload(FXIndex As Long)
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub
 DIFFEffect(FXIndex).Unload
End Sub

Public Sub Joystick_SetFXStop(FXIndex As Long)
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub
 DIFFEffect(FXIndex).Stop
End Sub

Public Sub Joystick_SetFXEnvelopeEffect(FXIndex As Long, AttackLevel As Long, AttackTime As Long, _
                                      FadeLevel As Long, FadeTime As Long)
 On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub

 FFEffect.bUseEnvelope = True
 With FFEffect.envelope
  .lAttackLevel = AttackLevel
  .lAttackTime = AttackTime
  .lFadeLevel = FadeLevel
  .lFadeTime = FadeTime
 End With
 
 DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_ENVELOPE
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXDuration(FXIndex As Long, Duration As Long, Optional Infinite As Boolean = False)
 On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub

 If Infinite Then FFEffect.lDuration = -1 Else FFEffect.lDuration = Duration * 100
 DIFFEffect(FXIndex).Stop
 DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_DURATION
 DIFFEffect(FXIndex).Start 1, 0
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXGain(FXIndex As Long, Gain As Long)
 On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub

 FFEffect.lGain = Gain
 DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_GAIN
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXSampleRate(FXIndex As Long, Rate As Long)
 On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub

 FFEffect.lSamplePeriod = Rate
 DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_SAMPLEPERIOD
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXConstantForce(FXIndex As Long, Force As Long)
 On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub

 FFEffect.constantForce.lMagnitude = Force
 DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_TYPESPECIFICPARAMS
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXDirection(FXIndex As Long, Direction As FXDirection)
 On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub

 With FFEffect
  If Direction = North Then
   .x = 0
  ElseIf Direction = North_East Then
   .x = 4500
  ElseIf Direction = East Then
   .x = 9000
  ElseIf Direction = South_East Then
   .x = 13500
  ElseIf Direction = South Then
   .x = 18000
  ElseIf Direction = South_West Then
   .x = 22500
  ElseIf Direction = West Then
   .x = 27000
  ElseIf Direction = North_West Then
   .x = 31500
  End If
 End With

 DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_DIRECTION
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXConditionX(FXIndex As Long, DeadBand As Long, _
                                    NegCoeff As Long, NegSat As Long, _
                                    PosCoeff As Long, PosSat As Long, Offset As Long)
 On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub

 With FFEffect.conditionX
  .lDeadBand = DeadBand
  .lNegativeCoefficient = NegCoeff
  .lNegativeSaturation = NegSat
  .lOffset = Offset
  .lPositiveCoefficient = PosCoeff
  .lPositiveSaturation = PosSat
 End With
 
 DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_TYPESPECIFICPARAMS
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXConditionY(FXIndex As Long, DeadBand As Long, _
                                    NegCoeff As Long, NegSat As Long, _
                                    PosCoeff As Long, PosSat As Long, Offset As Long)
On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub

 With FFEffect.conditionY
  .lDeadBand = DeadBand
  .lNegativeCoefficient = NegCoeff
  .lNegativeSaturation = NegSat
  .lOffset = Offset
  .lPositiveCoefficient = PosCoeff
  .lPositiveSaturation = PosSat
 End With
 
 DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_TYPESPECIFICPARAMS
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXRampForce(FXIndex As Long, StartRange As Long, EndRange As Long)
 On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub

 With FFEffect.rampForce
  .lRangeStart = StartRange
  .lRangeEnd = EndRange
 End With
 
 DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_TYPESPECIFICPARAMS
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_SetFXPeriodicForce(FXIndex As Long, Magnitude As Long, Offset As Long, Period As Long, Phase As Long)
 On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub

 With FFEffect.periodicForce
  .lMagnitude = Magnitude
  .lOffset = Offset
  .lPeriod = Period
  .lPhase = Phase
 End With
 
 DIFFEffect(FXIndex).SetParameters FFEffect, DIEP_TYPESPECIFICPARAMS
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub

Public Sub Joystick_PlayPredefinedFX(Effect As FXType)
 On Local Error GoTo ErrorHandler
 If Not IsJoystick Then Exit Sub
 If Not Joystick_HasForceFeedback Then Exit Sub
 If Not WasEnumerated Then Exit Sub
 Dim HIndex As Long

 If Effect = ShootArrow Then
  Joystick_SetFXConstantForce 0, 7000
  Joystick_SetFXDirection 0, North
  Joystick_SetFXDuration 0, 1000
  HIndex = 0
 ElseIf Effect = Gun_44_Magnum Then
  Joystick_SetFXDirection 0, South
  Joystick_SetFXConstantForce 0, 10000
  Joystick_SetFXDuration 0, 1500
  Joystick_SetFXDirection 0, North
  Joystick_SetFXConstantForce 0, 8000
  Joystick_SetFXDuration 0, 1000
  HIndex = 0
 ElseIf Effect = Gun_9MM Then
  Joystick_SetFXDirection 0, South
  Joystick_SetFXConstantForce 0, 5000
  Joystick_SetFXDuration 0, 1000
  Joystick_SetFXDirection 0, North
  Joystick_SetFXConstantForce 0, 5000
  Joystick_SetFXDuration 0, 500
  HIndex = 0
 ElseIf Effect = BB_Gun Then
  Joystick_SetFXDirection 0, South
  Joystick_SetFXConstantForce 0, 2000
  Joystick_SetFXDuration 0, 500
  Joystick_SetFXDirection 0, North
  Joystick_SetFXConstantForce 0, 2000
  Joystick_SetFXDuration 0, 500
  HIndex = 0
 ElseIf Effect = Gatling_Gun Then
  Joystick_SetFXDirection 2, North
  Joystick_SetFXPeriodicForce 2, 10000, 0, 40000, 0
  Joystick_SetFXDuration 2, 5000
  HIndex = 2
 End If
 Joystick_SetFXUnload HIndex
 Joystick_SetFXStart HIndex
 DoEvents
 Exit Sub

ErrorHandler:
 If Err.Number = 445 Then
  Exit Sub
 ElseIf Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then
  DIJoy.Unacquire
  DIJoy.SetCooperativeLevel mhWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
  DIJoy.Acquire
  Resume
 End If
End Sub
