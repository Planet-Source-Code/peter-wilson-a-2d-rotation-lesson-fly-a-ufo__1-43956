Attribute VB_Name = "mDataStructures"
Option Explicit

Public Type mdrVector3
    X As Single
    Y As Single
    w As Single
End Type

Public Type mdrMATRIX3x3
    rc11 As Single: rc12 As Single: rc13 As Single
    rc21 As Single: rc22 As Single: rc23 As Single
    rc31 As Single: rc32 As Single: rc33 As Single
End Type


Public Type mdr2DObject
    Caption As String
    
    ' 2D-Geometery to define the Object's Shape
    Vertex() As mdrVector3  ' Original Vertices (these never change once defined)
    TVertex() As mdrVector3 ' Transformed Vertices (these change all the time)
    Face() As Variant       ' Connect the dots [Vertices] together to form shapes.
    
    ' 2D World Coordinates (ie. The object's position in the game/world.)
    WorldPosition As mdrVector3
    
    ' Direction/Speed Vector
    ' (Typically changes when the user presses the arrow keys to move something)
    Vector As mdrVector3
    
    SpinMagnitude As Single
    RotationAboutZ As Single
 
End Type
