Attribute VB_Name = "mGameLoop"
Option Explicit

Public m_strCurrentState As String

Public m_SpaceShip As mdr2DObject

Public Sub DoMainGameLoop()

    Select Case m_strCurrentState
        Case ""
            m_SpaceShip = CreateSpaceShip
            m_strCurrentState = "RunGame"
            
        Case Else
            
            Call DrawCrossHairs(Form1)
            
            Call ProcessKeyboardInput(m_SpaceShip)
            Call Calculate(m_SpaceShip)
            Call DrawFaces(m_SpaceShip, Form1)
            
    End Select
    

End Sub


Public Sub MouseMoved(X As Single, Y As Single)

    m_SpaceShip.WorldPosition.X = X
    m_SpaceShip.WorldPosition.Y = -Y
    
    m_SpaceShip.Vector.X = 0
    m_SpaceShip.Vector.Y = 0
    
End Sub


