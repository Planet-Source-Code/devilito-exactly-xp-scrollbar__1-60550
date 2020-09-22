Attribute VB_Name = "Mod_API"
Option Explicit
'===============================================================================================================================
' Created By: Osen Kusnadi <support@osenxpsuite.net>
' Created Date: 2005-05-16
'==============================================================================================================================

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Type POINTAPI
    X                 As Long
    Y                 As Long
End Type

Public Function IsMouseOver(m_hWnd) As Boolean

    Dim pt As POINTAPI

    GetCursorPos pt
    IsMouseOver = (WindowFromPoint(pt.X, pt.Y) = m_hWnd)

End Function

Public Sub InitBarColor(m_ColorScheme, MyBarColor() As Long)
    If m_ColorScheme = 0 Then
        MyBarColor(1, 0) = 16777215
        MyBarColor(1, 1) = 16763817
        MyBarColor(1, 2) = 16771030
        MyBarColor(1, 3) = 16771288
        MyBarColor(1, 4) = 16771288
        MyBarColor(1, 5) = 16771030
        MyBarColor(1, 6) = 16771030
        MyBarColor(1, 7) = 16771030
        MyBarColor(1, 8) = 16771030
        MyBarColor(1, 9) = 16771030
        MyBarColor(1, 10) = 16770000
        MyBarColor(1, 11) = 16769226
        MyBarColor(1, 12) = 16769226
        MyBarColor(1, 13) = 16763302
        MyBarColor(1, 14) = 16777215
        MyBarColor(1, 15) = 13809055
        MyBarColor(2, 0) = 16777215
        MyBarColor(2, 1) = 13933443
        MyBarColor(2, 2) = 15513248
        MyBarColor(2, 3) = 16105128
        MyBarColor(2, 4) = 16368042
        MyBarColor(2, 5) = 16433064
        MyBarColor(2, 6) = 16433064
        MyBarColor(2, 7) = 16498085
        MyBarColor(2, 8) = 16432032
        MyBarColor(2, 9) = 16432032
        MyBarColor(2, 10) = 16431257
        MyBarColor(2, 11) = 16364434
        MyBarColor(2, 12) = 16364438
        MyBarColor(2, 13) = 15642776
        MyBarColor(2, 14) = 16777215
        MyBarColor(2, 15) = 13809055
        MyBarColor(0, 0) = 16777215
        MyBarColor(0, 1) = 16108215
        MyBarColor(0, 2) = 16504520
        MyBarColor(0, 3) = 16505032
        MyBarColor(0, 4) = 16504520
        MyBarColor(0, 5) = 16635590
        MyBarColor(0, 6) = 16635590
        MyBarColor(0, 7) = 16635331
        MyBarColor(0, 8) = 16569282
        MyBarColor(0, 9) = 16569282
        MyBarColor(0, 10) = 16568762
        MyBarColor(0, 11) = 16502454
        MyBarColor(0, 12) = 16502198
        MyBarColor(0, 13) = 15977401
        MyBarColor(0, 14) = 16777215
        MyBarColor(0, 15) = 13809055
    ElseIf m_ColorScheme = 1 Then
        MyBarColor(1, 0) = 16777215
        MyBarColor(1, 1) = 9882557
        MyBarColor(1, 2) = 11195848
        MyBarColor(1, 3) = 11195849
        MyBarColor(1, 4) = 11195848
        MyBarColor(1, 5) = 11131339
        MyBarColor(1, 6) = 11131339
        MyBarColor(1, 7) = 10999754
        MyBarColor(1, 8) = 10802375
        MyBarColor(1, 9) = 10802375
        MyBarColor(1, 10) = 10212294
        MyBarColor(1, 11) = 9883843
        MyBarColor(1, 12) = 9883843
        MyBarColor(1, 13) = 9947068
        MyBarColor(1, 14) = 16777215
        MyBarColor(1, 15) = 7836555
        MyBarColor(2, 0) = 16777215
        MyBarColor(2, 1) = 6590846
        MyBarColor(2, 2) = 7906193
        MyBarColor(2, 3) = 8432280
        MyBarColor(2, 4) = 8564123
        MyBarColor(2, 5) = 8432794
        MyBarColor(2, 6) = 8432794
        MyBarColor(2, 7) = 8367514
        MyBarColor(2, 8) = 8170649
        MyBarColor(2, 9) = 8170649
        MyBarColor(2, 10) = 7842712
        MyBarColor(2, 11) = 7514774
        MyBarColor(2, 12) = 7514516
        MyBarColor(2, 13) = 7447182
        MyBarColor(2, 14) = 16777215
        MyBarColor(2, 15) = 7836555
        MyBarColor(0, 0) = 16777215
        MyBarColor(0, 1) = 7447181
        MyBarColor(0, 2) = 9155748
        MyBarColor(0, 3) = 9745830
        MyBarColor(0, 4) = 9745830
        MyBarColor(0, 5) = 9353125
        MyBarColor(0, 6) = 9353125
        MyBarColor(0, 7) = 9155748
        MyBarColor(0, 8) = 8827040
        MyBarColor(0, 9) = 8827040
        MyBarColor(0, 10) = 8170908
        MyBarColor(0, 11) = 7448726
        MyBarColor(0, 12) = 7710613
        MyBarColor(0, 13) = 7314570
        MyBarColor(0, 14) = 16777215
        MyBarColor(0, 15) = 7836555
    Else
        MyBarColor(1, 0) = 6645339
        MyBarColor(1, 1) = 16777215
        MyBarColor(1, 2) = 16776958
        MyBarColor(1, 3) = 16381172
        MyBarColor(1, 4) = 16183536
        MyBarColor(1, 5) = 15985901
        MyBarColor(1, 6) = 15985901
        MyBarColor(1, 7) = 15656165
        MyBarColor(1, 8) = 15261148
        MyBarColor(1, 9) = 15261148
        MyBarColor(1, 10) = 14931155
        MyBarColor(1, 11) = 14536140
        MyBarColor(1, 12) = 14470089
        MyBarColor(1, 13) = 14535883
        MyBarColor(1, 14) = 16777215
        MyBarColor(1, 15) = 6645339
        MyBarColor(2, 0) = 4737091
        MyBarColor(2, 1) = 16777215
        MyBarColor(2, 2) = 14206917
        MyBarColor(2, 3) = 14535882
        MyBarColor(2, 4) = 14535883
        MyBarColor(2, 5) = 14931156
        MyBarColor(2, 6) = 14931156
        MyBarColor(2, 7) = 15392991
        MyBarColor(2, 8) = 15919850
        MyBarColor(2, 9) = 15919850
        MyBarColor(2, 10) = 16183536
        MyBarColor(2, 11) = 16446966
        MyBarColor(2, 12) = 16710909
        MyBarColor(2, 13) = 16777215
        MyBarColor(2, 14) = 16777215
        MyBarColor(2, 15) = 4737091
        MyBarColor(0, 0) = 10655124
        MyBarColor(0, 1) = 16777215
        MyBarColor(0, 2) = 16381943
        MyBarColor(0, 3) = 15854828
        MyBarColor(0, 4) = 15854828
        MyBarColor(0, 5) = 15657448
        MyBarColor(0, 6) = 15657448
        MyBarColor(0, 7) = 15393506
        MyBarColor(0, 8) = 15064283
        MyBarColor(0, 9) = 15064283
        MyBarColor(0, 10) = 14800340
        MyBarColor(0, 11) = 14536655
        MyBarColor(0, 12) = 14405068
        MyBarColor(0, 13) = 14470861
        MyBarColor(0, 14) = 16777215
        MyBarColor(0, 15) = 10655124
    End If
End Sub

Public Sub DrawXLine(DestDC As Long, _
                     X As Long, _
                     Y As Long, _
                     X1 As Long, _
                     Y1 As Long, _
                     oColor As OLE_COLOR, _
                     Optional IWidth As Long = 1)

    Dim pt    As POINTAPI
    Dim iPen  As Long
    Dim iPen1 As Long

    iPen = CreatePen(0, IWidth, oColor)
    iPen1 = SelectObject(DestDC, iPen)
    MoveToEx DestDC, X, Y, pt
    LineTo DestDC, X1, Y1
    SelectObject DestDC, iPen1
    DeleteObject iPen
    On Error GoTo 0

End Sub

