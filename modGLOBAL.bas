Attribute VB_Name = "modGLOBAL"
'hDestDC - destination object, X - destination X axis, Y - destination Y axis
'nwidth - width to copy, nheight - height to copy,
'hSrcDC - source object, xSrc - start at xSrc on X axis, ySrc - start at ySrc on Y axis,
'dwRop - way to copy

Public Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nwidth As Long, ByVal nheight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long

Public Function max(intNUM1 As Integer, intNUM2 As Integer) As Integer
    If intNUM1 < intNUM2 Then
        max = intNUM2
    Else
        max = intNUM1
    End If
End Function

Public Function min(intNUM1 As Integer, intNUM2 As Integer) As Integer
    If intNUM1 > intNUM2 Then
        min = intNUM2
    Else
        min = intNUM1
    End If
End Function
