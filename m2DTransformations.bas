Attribute VB_Name = "m2DTransformations"
Option Explicit

Public Function MatrixIdentity() As mdrMATRIX3x3
    
    ' Creates a new Identity matrix
    With MatrixIdentity
        .rc11 = 1: .rc12 = 0: .rc13 = 0
        .rc21 = 0: .rc22 = 1: .rc23 = 0
        .rc31 = 0: .rc32 = 0: .rc33 = 1
    End With
    
End Function

Public Function MatrixMultiplyVector(MatrixIn As mdrMATRIX3x3, VectorIn As mdrVector3) As mdrVector3
    
    ' Multiplies a Vector with a Matrix.
    With MatrixMultiplyVector
    
        .X = (MatrixIn.rc11 * VectorIn.X) + (MatrixIn.rc12 * VectorIn.Y) + (MatrixIn.rc13 * VectorIn.w)
        .Y = (MatrixIn.rc21 * VectorIn.X) + (MatrixIn.rc22 * VectorIn.Y) + (MatrixIn.rc23 * VectorIn.w)
        .w = 1
        
    End With
    
End Function

Public Function MatrixMultiply(m1 As mdrMATRIX3x3, m2 As mdrMATRIX3x3) As mdrMATRIX3x3
        
    Dim m1b As mdrMATRIX3x3
    Dim m2b As mdrMATRIX3x3
    m1b = m1
    m2b = m2
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixMultiply = MatrixIdentity
    
    ' Multiply the two matrices together.
    With MatrixMultiply
    
        .rc11 = (m1b.rc11 * m2b.rc11) + (m1b.rc21 * m2b.rc12) + (m1b.rc31 * m2b.rc13)
        .rc12 = (m1b.rc12 * m2b.rc11) + (m1b.rc22 * m2b.rc12) + (m1b.rc32 * m2b.rc13)
        .rc13 = (m1b.rc13 * m2b.rc11) + (m1b.rc23 * m2b.rc12) + (m1b.rc33 * m2b.rc13)
        
        .rc21 = (m1b.rc11 * m2b.rc21) + (m1b.rc21 * m2b.rc22) + (m1b.rc31 * m2b.rc23)
        .rc22 = (m1b.rc12 * m2b.rc21) + (m1b.rc22 * m2b.rc22) + (m1b.rc32 * m2b.rc23)
        .rc23 = (m1b.rc13 * m2b.rc21) + (m1b.rc23 * m2b.rc22) + (m1b.rc33 * m2b.rc23)
        
        .rc31 = (m1b.rc11 * m2b.rc31) + (m1b.rc21 * m2b.rc32) + (m1b.rc31 * m2b.rc33)
        .rc32 = (m1b.rc12 * m2b.rc31) + (m1b.rc22 * m2b.rc32) + (m1b.rc32 * m2b.rc33)
        .rc33 = (m1b.rc13 * m2b.rc31) + (m1b.rc23 * m2b.rc32) + (m1b.rc33 * m2b.rc33)
    
    End With
    
End Function

Public Function Vect3Addition(V1 As mdrVector3, V2 As mdrVector3) As mdrVector3

    ' Adds V1 and V2 together. Nothing special here, just simple addition.
    
    Vect3Addition.X = V1.X + V2.X
    Vect3Addition.Y = V1.Y + V2.Y
    
    ' We can safely ignore the W component.
    Vect3Addition.w = 1
    
End Function
Public Function MatrixTranslation(OffsetX As Single, OffsetY As Single) As mdrMATRIX3x3
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixTranslation = MatrixIdentity
    
    ' This is a Translation Matrix (pretty simple hey?)
    MatrixTranslation.rc13 = OffsetX
    MatrixTranslation.rc23 = OffsetY
    
End Function


Public Function Vec3Length(V1 As mdrVector3) As Single

    ' Returns the length of a 3-D vector.
    ' The length of a vector is from the origin (0,0) to x,y
    ' We work this out using Pythagoras theorem:  c^2 = a^2 + b^2
    
    Vec3Length = Sqr((V1.X ^ 2) + (V1.Y ^ 2))
    
    ' We can safely ignore the W component.
    
End Function
Public Function MatrixRotationZ(Radians As Single) As mdrMATRIX3x3

    ' In this VB application:
    '   * The positive X axis points towards the right.
    '   * The positive Y axis points upwards to the top of the screen.
    '   * The positive Z axis points *into* the monitor.
    
    Dim sngCosine As Single
    Dim sngSine As Single
    
    sngCosine = Cos(Radians)
    sngSine = Sin(Radians)
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixRotationZ = MatrixIdentity
    
    
    ' Define a Rotation Matrix for the Z Axis.
    With MatrixRotationZ
        .rc11 = sngCosine
        .rc12 = -sngSine
        .rc21 = sngSine
        .rc22 = sngCosine
    End With
    
End Function

