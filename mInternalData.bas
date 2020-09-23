Attribute VB_Name = "mInternalData"
Option Explicit

Public Function CreateSpaceShip() As mdr2DObject

    ' There's not too much to learn here.
    ' Just create your X, Y dots, then connect them up using faces.
    
    With CreateSpaceShip
    
        ReDim .Vertex(7)
        ReDim .TVertex(7)
                
        ' Define the X and Y coordinates of the Space Ship
        .Vertex(0).X = 1: .Vertex(0).Y = 2
        .Vertex(1).X = 2: .Vertex(1).Y = 1
        .Vertex(2).X = 4: .Vertex(2).Y = 0
        .Vertex(3).X = 2: .Vertex(3).Y = -1
        .Vertex(4).X = -2: .Vertex(4).Y = -1
        .Vertex(5).X = -4: .Vertex(5).Y = 0
        .Vertex(6).X = -2: .Vertex(6).Y = 1
        .Vertex(7).X = -1: .Vertex(7).Y = 2
        
        ' Reset all the W's (ok... I know this will confuse many, I'll explain in a later tutorial)
        .Vertex(0).w = 1
        .Vertex(1).w = 1
        .Vertex(2).w = 1
        .Vertex(3).w = 1
        .Vertex(4).w = 1
        .Vertex(5).w = 1
        .Vertex(6).w = 1
        .Vertex(7).w = 1
        
        ' Connect the Dots
        ReDim .Face(1)
        .Face(0) = Array(1, 2, 3, 4, 5, 6, 1, 0, 7, 6)
        .Face(1) = Array(5, 2)
        
    End With
    
End Function

