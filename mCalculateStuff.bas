Attribute VB_Name = "mCalculateStuff"
Option Explicit

Public Sub Calculate(CurrentObject As mdr2DObject)

    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    
    With CurrentObject
            
        ' Translate (ie. Move) the object (spaceship) to the correct location
        ' ===================================================================
        .WorldPosition.X = .WorldPosition.X + .Vector.X
        .WorldPosition.Y = .WorldPosition.Y + .Vector.Y
        
        ' Build the Translation Matrix
        ' =============================
        matTranslate = MatrixTranslation(.WorldPosition.X, .WorldPosition.Y)
        
        ' Apply the spin value
        ' ====================
        .RotationAboutZ = .RotationAboutZ + .SpinMagnitude
        
        ' Build the Rotation Matrix
        ' =========================
        matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
        
        
        ' Multiply the Matrices together (in the correct order)
        ' =====================================================
        matResult = MatrixIdentity
        matResult = MatrixMultiply(matResult, matRotationAboutZ)
        matResult = MatrixMultiply(matResult, matTranslate)
        
        ' Multiply the Vertex points with the result Matrix.
        ' ==================================================
        For intJ = LBound(.Vertex) To UBound(.Vertex)
            .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
        Next intJ
        
    End With

End Sub


