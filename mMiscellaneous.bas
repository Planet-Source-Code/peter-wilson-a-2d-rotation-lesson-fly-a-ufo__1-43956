Attribute VB_Name = "mMiscellaneous"
Option Explicit

Private Const m_sngPIDivideBy180 = 0.0174533!
Public Function ConvertDeg2Rad(Degress As Single) As Single

    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (m_sngPIDivideBy180)
    
End Function

