VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwGlobe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cwGlobe
' Author    : beededea
' Date      : 28/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private ringxo As Integer
Private ringyo As Integer
Private surroundxo As Integer
Private surroundyo As Integer
Private switchFacesButtonxo As Integer
Private switchFacesButtonyo As Integer
Private startButtonxo As Integer
Private startButtonyo As Integer
Private stopButtonxo As Integer
Private stopButtonyo As Integer
Private bigReflectionxo As Integer
Private bigReflectionyo As Integer
Private windowReflectionxo As Integer
Private windowReflectionyo As Integer
Private pinxo As Integer
Private pinyo As Integer
Private prefsxo As Integer
Private prefsyo As Integer
Private helpButtonxo As Integer
Private helpButtonyo As Integer
Private tickSwitchxo As Integer
Private tickSwitchyo As Integer
Private earthGlobexo As Integer
Private earthGlobeyo  As Integer
Private earthCounter As Integer

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrRotate As cTimer
Attribute tmrRotate.VB_VarHelpID = -1

Private mRotationSpeed As Double
Private mZoom As Double
Private mZoomDirection As String
Private mRotating As Boolean
Private mLocked As Boolean
Private mHidden As Boolean
Private mOpacity As Double
Private mWasRotating As Boolean

' property mRotationSpeed
' property mZoom
' property mZoomDirection
' property mRotating
' property mLocked
' property mHidden
' property mOpacity
' property mWasRotating
 
'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    :
' Date      : 03/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

   On Error GoTo Class_Initialize_Error

    Call initialiseVars
    
    mRotationSpeed = 0: mZoom = 0
    surroundxo = 0: surroundyo = 10
    ringxo = 38: ringyo = 51
    switchFacesButtonxo = 585: switchFacesButtonyo = 267
    startButtonxo = 585: startButtonyo = 135
    stopButtonxo = 585: stopButtonyo = 395
    bigReflectionxo = 44: bigReflectionyo = 69
    windowReflectionxo = 386: windowReflectionyo = 210
    pinxo = 37: pinyo = 60
    prefsxo = 36: prefsyo = 516
    helpButtonxo = 500: helpButtonyo = 516
    tickSwitchxo = 500: tickSwitchyo = 59
    earthGlobexo = 15: earthGlobeyo = 27
    
    earthCounter = 1
      
    Set W = Cairo.WidgetBase '<- this is required in each cwImplementation...
        W.BorderColor = &H444444
        W.ImplementsWheelMessages = True ' enables Private Sub W_MouseWheel
             
        mZoom = 0.1 ' the globe size initially very small
        ZoomDirection = mZoomDirection
    Set tmrRotate = New_c.Timer(mRotationSpeed, False)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwGlobe"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : initialiseVars
' Author    : beededea
' Date      : 07/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub initialiseVars() ' set all vars to zero
    On Error GoTo initialiseVars_Error

    surroundxo = 0: surroundyo = 0
    switchFacesButtonxo = 0: switchFacesButtonyo = 0
    startButtonxo = 0: startButtonyo = 0
    stopButtonxo = 0: stopButtonyo = 0
    bigReflectionxo = 0: bigReflectionyo = 0
    windowReflectionxo = 0: windowReflectionyo = 0
    pinxo = 0: pinyo = 0
    prefsxo = 0: prefsyo = 0
    helpButtonxo = 0: helpButtonyo = 0
    tickSwitchxo = 0: tickSwitchyo = 0
    earthGlobexo = 0: earthGlobeyo = 0
    
    earthCounter = 0

    On Error GoTo 0
    Exit Sub

initialiseVars_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseVars of Class Module cwGlobe"
            Resume Next
          End If
    End With

End Sub

Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets:  Set Widgets = W.Widgets: End Property ' does this really need to be exposed?
'---------------------------------------------------------------------------------------
' Procedure : Zoom
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Zoom() As Double
   On Error GoTo ZoomGet_Error

    Zoom = mZoom

   On Error GoTo 0
   Exit Property

ZoomGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Zoom of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Zoom
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Zoom(ByVal newValue As Double)
   On Error GoTo ZoomLet_Error

  If mZoom <> newValue Then mZoom = newValue Else Exit Property
  If mZoom > 1.2 Then mZoom = 1.2 Else If mZoom < 0.05 Then mZoom = 0.05
  W.Refresh

   On Error GoTo 0
   Exit Property

ZoomLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Zoom of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : ZoomDirection
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ZoomDirection() As String
   On Error GoTo ZoomDirectionGet_Error

  ZoomDirection = mZoomDirection

   On Error GoTo 0
   Exit Property

ZoomDirectionGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ZoomDirection of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : ZoomDirection
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ZoomDirection(ByVal newValue As String)
   On Error GoTo ZoomDirectionLet_Error

  If mZoomDirection <> newValue Then mZoomDirection = newValue Else Exit Property
  W.Refresh

   On Error GoTo 0
   Exit Property

ZoomDirectionLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ZoomDirection of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let opacity(ByVal newValue As Double)
   On Error GoTo opacityLet_Error

    If mOpacity <> newValue Then mOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

opacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure opacity of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get opacity() As Double
   On Error GoTo opacityGet_Error

    opacity = mOpacity

   On Error GoTo 0
   Exit Property

opacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure opacity of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : RotationSpeed
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get RotationSpeed() As Double
   On Error GoTo RotationSpeedGet_Error

  RotationSpeed = mRotationSpeed

   On Error GoTo 0
   Exit Property

RotationSpeedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RotationSpeed of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : RotationSpeed
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let RotationSpeed(ByVal newValue As Double)
   On Error GoTo RotationSpeedLet_Error

  If mRotationSpeed <> newValue Then mRotationSpeed = newValue Else Exit Property
  If mRotationSpeed Then Set tmrRotate = New_c.Timer(mRotationSpeed, True) Else Set tmrRotate = Nothing

   On Error GoTo 0
   Exit Property

RotationSpeedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RotationSpeed of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Locked() As Boolean
   On Error GoTo LockedGet_Error

  Locked = mLocked

   On Error GoTo 0
   Exit Property

LockedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Locked of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Boolean)
   On Error GoTo LockedLet_Error

  If mLocked <> newValue Then mLocked = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Locked of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Hidden
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Hidden() As Boolean
   On Error GoTo HiddenGet_Error

  Hidden = mHidden

   On Error GoTo 0
   Exit Property

HiddenGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Hidden of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Hidden
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Hidden(ByVal newValue As Boolean)
   On Error GoTo HiddenLet_Error

    If mHidden <> newValue Then mHidden = newValue Else Exit Property
    If mHidden = True Then
        If Rotating = True Then
            WasRotating = True
        Else
            WasRotating = False
        End If
        Rotating = False
        opacity = 0
        W.Refresh
    Else
        If WasRotating = True Then Rotating = True
        opacity = Val(PzEOpacity) / 100
        W.Refresh
    End If

   On Error GoTo 0
   Exit Property

HiddenLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Hidden of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : WasRotating
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get WasRotating() As Boolean
   On Error GoTo WasRotatingGet_Error

  WasRotating = mWasRotating

   On Error GoTo 0
   Exit Property

WasRotatingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WasRotating of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : WasRotating
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let WasRotating(ByVal newValue As Boolean)
   On Error GoTo WasRotatingLet_Error

    If mWasRotating <> newValue Then mWasRotating = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

WasRotatingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WasRotating of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Rotating
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Rotating() As Boolean
   On Error GoTo RotatingGet_Error

  Rotating = mRotating

   On Error GoTo 0
   Exit Property

RotatingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Rotating of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Rotating
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Rotating(ByVal newValue As Boolean)
   On Error GoTo RotatingLet_Error

  If mRotating <> newValue Then mRotating = newValue Else Exit Property
    If mRotating Then
        Set tmrRotate = New_c.Timer(RotationSpeed, True)
    Else
        Set tmrRotate = Nothing
    End If

   On Error GoTo 0
   Exit Property

RotatingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Rotating of Class Module cwGlobe"
End Property ' the earth rotating flag
'---------------------------------------------------------------------------------------
' Procedure : tmrRotate_Timer
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrRotate_Timer()
   On Error GoTo tmrRotate_Timer_Error

    If Rotating = False Then
        Set tmrRotate = Nothing
        Exit Sub
    End If
    earthCounter = earthCounter + 1
    If earthCounter >= 36 Then
      earthCounter = 1
    End If
    W.Refresh

   On Error GoTo 0
   Exit Sub

tmrRotate_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrRotate_Timer of Class Module cwGlobe"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_MouseWheel
' Author    : beededea
' Date      : 09/05/2023
' Purpose   : uses ScaleDrawings mZoom, mZoom
'---------------------------------------------------------------------------------------
'
Private Sub W_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
   On Error GoTo W_MouseWheel_Error

    If PzEIgnoreMouse = "1" Then Exit Sub

    Select Case MouseKeys
    Case 8 ' ctrl
        
        If ZoomDirection = "up" Then
            Zoom = Zoom - 0.02 * LineOffset
        Else
            Zoom = Zoom + 0.02 * LineOffset
        End If
        PzEGaugeSize = LTrim$(Str$(globeWidget.Zoom * 100)) ' store the value dynamically
        sPutINISetting "Software\PzEarth", "gaugeSize", PzEGaugeSize, PzESettingsFile
        
        If panzerEarthPrefs.IsVisible = True Then panzerEarthPrefs.sliGaugeSize = Val(PzEGaugeSize)

    End Select

   On Error GoTo 0
   Exit Sub

W_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseWheel of Class Module cwGlobe"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
   On Error GoTo W_MouseDown_Error

    If Button = vbRightButton Then
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        If PzEIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If


   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwGlobe"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : beededea
' Date      : 20/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
   On Error GoTo W_Paint_Error

    Draw CC, dx_Aligned, dy_Aligned, mOpacity

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwGlobe"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : Draw
' Author    :
' Date      : 03/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Draw(ByVal CC As cCairoContext, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, ByVal mOpacity As Double)
    
   On Error GoTo Draw_Error
          
    CC.TranslateDrawings 0, 0  ' dx / 2, dy / 2
    CC.ScaleDrawings mZoom, mZoom

    CC.Save
   
    CC.RenderSurfaceContent "Ring", ringxo, ringyo, , , , mOpacity
    CC.RenderSurfaceContent "surround", surroundxo, surroundyo, , , , mOpacity
    CC.RenderSurfaceContent "switchFacesButton", switchFacesButtonxo, switchFacesButtonyo, , , , mOpacity
    CC.RenderSurfaceContent "startButton", startButtonxo, startButtonyo, , , , mOpacity
    CC.RenderSurfaceContent "stopButton", stopButtonxo, stopButtonyo, , , , mOpacity
    CC.RenderSurfaceContent "pin", pinxo, pinyo, , , , mOpacity
    CC.RenderSurfaceContent "prefs", prefsxo, prefsyo, , , , mOpacity
    CC.RenderSurfaceContent "helpButton", helpButtonxo, helpButtonyo, , , , mOpacity
    CC.RenderSurfaceContent "tickSwitch", tickSwitchxo, tickSwitchyo, , , , mOpacity
    CC.RenderSurfaceContent "EarthGlobe" & earthCounter, earthGlobexo, earthGlobeyo, 570, 570, , mOpacity
    
    CC.Restore
    CC.RenderSurfaceContent "bigReflection", bigReflectionxo, bigReflectionyo, , , , mOpacity
    CC.RenderSurfaceContent "windowReflection", windowReflectionxo, windowReflectionyo, , , , mOpacity
    CC.RenderSurfaceContent "Glow", 25, 33, 555, 555, , mOpacity


   On Error GoTo 0
   Exit Sub

Draw_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Draw of Class Module cwGlobe"

End Sub

 


