Attribute VB_Name = "mKeyboard"
Option Explicit

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub ProcessKeyboardInput(CurrentObject As mdr2DObject)

    On Error GoTo errTrap
        
    Dim lngKeyCombinations As Long
    Dim lngKeyState As Long
    Dim sngSpeedIncrement As Single
    Dim sngSpeedMagnitude As Single
    
    sngSpeedIncrement = 10 ' <<< Change this value to change speed of rotation.
    
    
    lngKeyState = GetKeyState(vbKeyLeft)
    If (lngKeyState And &H8000) Then lngKeyCombinations = lngKeyCombinations Or 1
    
    lngKeyState = GetKeyState(vbKeyRight)
    If (lngKeyState And &H8000) Then lngKeyCombinations = lngKeyCombinations Or 2
        
    lngKeyState = GetKeyState(vbKeyUp)
    If (lngKeyState And &H8000) Then lngKeyCombinations = lngKeyCombinations Or 4
    
    lngKeyState = GetKeyState(vbKeyDown)
    If (lngKeyState And &H8000) Then lngKeyCombinations = lngKeyCombinations Or 8
    
    
    ' Rotate CurrentObject (SpaceShip) using Classic Mode (easy and fun)
    ' ==================================================================
    CurrentObject.SpinMagnitude = 0  ' Reset the SpinVector
    If (lngKeyCombinations And 1) = 1 Then CurrentObject.SpinMagnitude = sngSpeedIncrement
    If (lngKeyCombinations And 2) = 2 Then CurrentObject.SpinMagnitude = -sngSpeedIncrement
    
    
    If (lngKeyCombinations And 8) = 8 Then
        ' Reset Position
        CurrentObject.WorldPosition.X = 0
        CurrentObject.WorldPosition.Y = 0
        
        CurrentObject.Vector.X = 0
        CurrentObject.Vector.Y = 0
        
    End If
    
     
    ' ===============
    ' Thrust Forwards
    ' ===============
    If (lngKeyCombinations And 4) = 4 Then
        
        Dim tempVector As mdrVector3
        Dim sngRadians As Single
        
        sngRadians = ConvertDeg2Rad(CurrentObject.RotationAboutZ + 90)
        tempVector.X = Cos(sngRadians) * 0.1 ' <<< Change this 0.1 value for speed.
        tempVector.Y = Sin(sngRadians) * 0.1
        tempVector.w = 1
    
        CurrentObject.Vector = Vect3Addition(tempVector, CurrentObject.Vector)
        sngSpeedMagnitude = Vec3Length(CurrentObject.Vector)
        
    End If

    Exit Sub
errTrap:
    
End Sub


