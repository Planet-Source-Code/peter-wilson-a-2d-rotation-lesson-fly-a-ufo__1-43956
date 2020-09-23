Attribute VB_Name = "mDrawStuff"
Option Explicit

Public Sub DrawCrossHairs(CurrentForm As Form)

    CurrentForm.DrawMode = vbCopyPen
    CurrentForm.DrawStyle = vbSolid
    CurrentForm.DrawWidth = 1
    CurrentForm.ForeColor = RGB(64, 64, 64)
    
    ' Draw vertical line.
    CurrentForm.Line (0, CurrentForm.ScaleTop)-(0, CurrentForm.ScaleHeight)

    ' Draw horizontal line.
    CurrentForm.Line (CurrentForm.ScaleLeft, 0)-(CurrentForm.ScaleWidth, 0)
    
End Sub

Public Sub DrawFaces(CurrentObject As mdr2DObject, CurrentForm As Form)

    Dim intN As Integer, intK As Integer
    Dim intFaceIndex As Integer
    Dim intVertexIndex As Integer
    Dim xPos As Single
    Dim yPos As Single
    
        
    CurrentForm.DrawStyle = vbSolid
    CurrentForm.DrawMode = vbCopyPen
    CurrentForm.DrawWidth = 1
    CurrentForm.ForeColor = RGB(0, 255, 255)
    
    With CurrentObject
        For intFaceIndex = LBound(.Face) To UBound(.Face)
            For intK = LBound(.Face(intFaceIndex)) To UBound(.Face(intFaceIndex))
                
                intVertexIndex = .Face(intFaceIndex)(intK)
                xPos = .TVertex(intVertexIndex).X
                yPos = -.TVertex(intVertexIndex).Y
                
                ' Normal Face; move to first point, then draw to the others.
                ' ==========================================================
                If intK = LBound(.Face(intFaceIndex)) Then
                    ' Move to first point
                    CurrentForm.Line (xPos, yPos)-(xPos, yPos)
                Else
                    ' Draw to point
                    CurrentForm.Line -(xPos, yPos)
                End If
                
            Next intK
        Next intFaceIndex
    End With

End Sub

