Attribute VB_Name = "modGLOBAL"
Public Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nwidth As Long, ByVal nheight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long

