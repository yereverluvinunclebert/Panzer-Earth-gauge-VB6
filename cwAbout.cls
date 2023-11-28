VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module      : cwAbout
' Author            : beededea
' Date              : 17/05/2023
' Purpose           :
'---------------------------------------------------------------------------------------

Option Explicit
 
Private WithEvents W As cWidgetBase '<- this is required in each cwImplementation...
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrFadeIn As cTimer
Attribute tmrFadeIn.VB_VarHelpID = -1
Private WithEvents tmrFadeOut As cTimer
Attribute tmrFadeOut.VB_VarHelpID = -1
Private mMyOpacity As Double
Private mshowMe As Boolean
Private aboutxo As Integer
Private aboutyo As Integer

'property MyOpacity
'property showMe


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

   On Error GoTo Class_Initialize_Error

    aboutxo = 0
    aboutyo = 0
    MyOpacity = 0
    showMe = False

    Set W = Cairo.WidgetBase '<- this is required in each cwImplementation...
   
    Set tmrFadeIn = New_c.Timer(20, False)
    Set tmrFadeOut = New_c.Timer(20, False)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwAbout"
End Sub

Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets:  Set Widgets = W.Widgets: End Property ' does this really need to be exposed?
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
   On Error GoTo W_MouseDown_Error

    tmrFadeOut.Enabled = True

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwAbout"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : MyOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let MyOpacity(ByVal newValue As Double)
   On Error GoTo MyOpacityLet_Error

  If mMyOpacity <> newValue Then mMyOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

MyOpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyOpacity of Class Module cwAbout"
End Property
'---------------------------------------------------------------------------------------
' Procedure : MyOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get MyOpacity() As Double
   On Error GoTo MyOpacityGet_Error

  MyOpacity = mMyOpacity

   On Error GoTo 0
   Exit Property

MyOpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyOpacity of Class Module cwAbout"
End Property

'---------------------------------------------------------------------------------------
' Procedure : showMe
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let showMe(ByVal newValue As Boolean)
    On Error GoTo showMeLet_Error

    mshowMe = newValue
    If mshowMe = True Then tmrFadeIn.Enabled = True
    
    On Error GoTo 0
    Exit Property

showMeLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showMe of Class Module cwAbout"
End Property
'---------------------------------------------------------------------------------------
' Procedure : showMe
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get showMe() As Boolean
   On Error GoTo showMeGet_Error

   showMe = mshowMe

   On Error GoTo 0
   Exit Property

showMeGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showMe of Class Module cwAbout"
End Property

'
'---------------------------------------------------------------------------------------
' Procedure : tmrFadeIn_Timer
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : timer to make the about form or about image fade in opacity
'---------------------------------------------------------------------------------------
'
Private Sub tmrFadeIn_Timer()
      
   On Error GoTo tmrFadeIn_Timer_Error
        fMain.aboutForm.Show
        MyOpacity = MyOpacity + 0.05
        If MyOpacity >= 1 Then
            MyOpacity = 1

            tmrFadeIn.Enabled = False

            Exit Sub
        End If
    W.Refresh

   On Error GoTo 0
   Exit Sub

tmrFadeIn_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrFadeIn_Timer of Class Module cwAbout"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : tmrFadeOut_Timer
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : timer to make the about form or about image fade out in opacity
'---------------------------------------------------------------------------------------
'
Private Sub tmrFadeOut_Timer()
      
   On Error GoTo tmrFadeOut_Timer_Error

        MyOpacity = MyOpacity - 0.05
        If MyOpacity < 0 Then
            MyOpacity = 0

            tmrFadeOut.Enabled = False
            fMain.aboutForm.Hide

            Exit Sub
        End If
    W.Refresh

   On Error GoTo 0
   Exit Sub

tmrFadeOut_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrFadeOut_Timer of Class Module cwAbout"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
   On Error GoTo W_Paint_Error

    Draw CC, dx_Aligned, dy_Aligned, mMyOpacity

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwAbout"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Draw
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Draw(ByRef CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal mMyOpacity As Double)
   On Error GoTo Draw_Error

    CC.RenderSurfaceContent "about", aboutxo, aboutyo, , , , mMyOpacity
        
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontItalic = True
    W.FontName = PzEPrefsFont
    W.ForeColor = vbBlack
    W.SelectFontSettingsInto CC
    
    CC.DrawText 0, 0, dx + 350, dy - 390, "                                                                                                                            version " & App.Major & "." & App.Minor & "." & App.Revision & vbNullString, , , 2, True, , mMyOpacity

   On Error GoTo 0
   Exit Sub

Draw_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Draw of Class Module cwAbout"

End Sub

