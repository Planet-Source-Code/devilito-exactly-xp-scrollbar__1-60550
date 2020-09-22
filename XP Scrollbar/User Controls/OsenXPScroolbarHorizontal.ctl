VERSION 5.00
Begin VB.UserControl OsenXPVScrollBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   645
   FontTransparent =   0   'False
   ForwardFocus    =   -1  'True
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   43
   ToolboxBitmap   =   "OsenXPScroolbarHorizontal.ctx":0000
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   2010
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   420
      Top             =   990
   End
End
Attribute VB_Name = "OsenXPVScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 8224
Attribute VB_Description = "osenxpsuite2005.OsenXPVScrollBar"
Option Explicit
Public Enum XPTheme
    XP_Blue
    XP_OliveGreen
    XP_Silver
End Enum

Private bSlideBar           As Boolean
Private bLoaded             As Boolean

Private PrevYPosition       As Integer

Private m_Min               As Long
Private m_Max               As Long
Private m_Value             As Long
Private m_Small             As Long
Private m_Large             As Long
Private m_State             As Byte
Private m_Last_State        As Byte
Private m_Last_Down         As Byte

Private m_ColorScheme       As XPTheme
Private m_RealSpace         As Long
Private m_IsDone            As Boolean

Private m_Last_Value            As Long
Private ResID                   As Long
Private YPos                    As Single

Private Type RectBar
    lTop    As Long
    LHeight As Long
End Type

Private RcBar                   As RectBar

Public Event Change()
Attribute Change.VB_HelpID = 8225
Public Event Scroll()
Attribute Scroll.VB_HelpID = 8226

Private m_TMR_Scroll            As Boolean
Private m_TMR_Over              As Boolean
Private m_POS                   As Byte

Private MyBarColor(2, 15)       As Long

Private Function CheckValidPosition() As Boolean

    CheckValidPosition = (YPos >= RcBar.lTop And YPos <= RcBar.LHeight + RcBar.lTop)

End Function


Private Sub PaintTransMyBlt(DstX As Long, DstY As Long, DstW As Long, DstH As Long, SrcPic As StdPicture)
    On Error Resume Next

    Dim OriW As Long, OriH As Long
    OriW = UserControl.ScaleX(SrcPic.Width, vbHimetric, vbPixels)
    OriH = UserControl.ScaleY(SrcPic.Height, vbHimetric, vbPixels)

    UserControl.PaintPicture SrcPic, DstX, DstY, DstW, DstH, 0, 0, OriW, OriH

End Sub


Public Property Get Enabled() As Boolean
Attribute Enabled.VB_HelpID = 8229

    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)
    On Error Resume Next

    UserControl.Enabled() = bEnabled
    PropertyChanged "Enabled"
    DrawBasicScrollBar

End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_HelpID = 8230

    hwnd = UserControl.hwnd

End Property

Public Property Let LargeChange(lVal As Long)
Attribute LargeChange.VB_HelpID = 8232
    On Error Resume Next

    If lVal >= 0 And lVal <= 10000 Then
        m_Large = lVal
    Else
        MsgBox "Invalid property value", vbCritical
        m_Large = 1
    End If
    UserControl_Resize
    PropertyChanged "LargeChange"

End Property

Public Property Get LargeChange() As Long

    LargeChange = m_Large

End Property

Public Property Get Max() As Long
Attribute Max.VB_HelpID = 8233

    Max = m_Max

End Property

Public Property Let Max(mVal As Long)
    On Error Resume Next

    If mVal < 0 Then
        mVal = 32767
    End If
    m_Max = mVal
    Value = IIf(m_Value > m_Max, m_Max, m_Value)
    UserControl_Resize
    PropertyChanged "Max"

End Property

Public Property Get Min() As Long
Attribute Min.VB_HelpID = 8234

    Min = m_Min

End Property

Public Property Let Min(mVal As Long)
    On Error Resume Next

    If mVal < 0 Then
        mVal = 0
    End If
    If mVal > 32767 Then
        mVal = 0
    End If
    m_Min = mVal
    Value = IIf(m_Value < m_Min, m_Min, m_Value)
    UserControl_Resize
    PropertyChanged "Min"

End Property

Public Property Get ColorScheme() As XPTheme
Attribute ColorScheme.VB_HelpID = 8235

    ColorScheme = m_ColorScheme

End Property

Public Property Let ColorScheme(mVal As XPTheme)
    On Error Resume Next

    If mVal = -1 Then mVal = 0
    If m_ColorScheme <> mVal Then
        m_ColorScheme = mVal
        InitBarColor m_ColorScheme, MyBarColor
        DrawBasicScrollBar
        PropertyChanged "ColorScheme"
    End If

End Property

Public Property Let SmallChange(sVal As Long)
Attribute SmallChange.VB_HelpID = 8237
    On Error Resume Next

    If sVal >= 0 And sVal <= 32767 Then
        m_Small = sVal
    Else
        MsgBox "Invalid property value", vbCritical
        m_Small = 1
    End If
    PropertyChanged "SmallChange"

End Property

Public Property Get SmallChange() As Long

    SmallChange = m_Small

End Property
Private Sub Timer1_Timer()

    If Not IsMouseOver(UserControl.hwnd) And Not bSlideBar Then
        UserControl.AutoRedraw = True
        UpdateLastPosition
        m_Last_State = 10
        UserControl.Refresh
        UserControl.AutoRedraw = False
        m_TMR_Over = False
        Timer1.Enabled = False
    End If

End Sub

Private Sub Timer2_Timer()
    If Not m_IsDone Then
        If bSlideBar Then
            Select Case m_Last_Down
              Case 1
                If Value - 1 >= Min Then
                    Value = Value - m_Small
                  Else
                    Value = 0
                End If
              Case 2

              Case 3
                If Value + 1 <= Max Then
                    Value = Value + m_Small
                  Else
                    Value = Max
                End If
              Case 4
                If Value - m_Large >= Min Then
                    Value = Value - m_Large
                  Else
                    Value = Min
                End If
              Case 5
                If Value + m_Large <= Max Then
                    Value = Value + m_Large
                  Else
                    Value = Max
                End If
            End Select
            RaiseEvent Change
        End If
    End If

    If (Not bSlideBar Or CheckValidPosition) And m_TMR_Scroll Then
        UserControl.AutoRedraw = True
        PaintBar
        UserControl.Refresh
        UserControl.AutoRedraw = False
        m_TMR_Scroll = False
        Timer2.Enabled = False
    End If

End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next

    m_Large = 10
    m_Max = 255
    m_Small = 1
    UserControl.Width = 255
    InitBarColor m_ColorScheme, MyBarColor
    UserControl_Resize

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    m_Last_Down = m_State

    If m_State + 10 <> m_POS Then
        UserControl.AutoRedraw = True

        Select Case m_Last_Down
            Case 4
                If Value - m_Large >= Min Then
                    Value = Value - m_Large
                Else
                    Value = Min
                End If
            Case 5
                If Value + m_Large <= Max Then
                    Value = Value + m_Large
                Else
                    Value = Max
                End If
        End Select

        If Button = 1 Then
            Select Case m_State
                Case 1
                    DrawUPButton 2
                Case 2
                    PaintBar 2
                    m_IsDone = True

                Case 3
                    DrawBotButton 2

            End Select

            bSlideBar = True
            YPos = Y
            m_Last_Value = Value

            If m_State <> 2 Then

                If Not m_TMR_Scroll And Not m_IsDone Then
                    If m_State > 3 Then
                        Timer2.Interval = 200
                    Else
                        Timer2.Interval = 100
                    End If
                    Timer2.Enabled = True
                    m_TMR_Scroll = True
                End If

            End If

        Else

        End If
        UserControl.Refresh
        UserControl.AutoRedraw = False
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If bSlideBar And Button = 1 Then
        If m_Last_Down = 2 Then
            If m_Last_Value + CLng(((Y - YPos) / m_RealSpace) * m_Max) > m_Max Then
                Value = m_Max
            ElseIf m_Last_Value + CLng(((Y - YPos) / m_RealSpace) * m_Max) < m_Min Then
                Value = m_Min
            Else
                Value = m_Last_Value + CLng(((Y - YPos) / m_RealSpace) * m_Max)
            End If
        ElseIf m_Last_Down > 3 Then
            YPos = Y
        End If
        RaiseEvent Scroll
    Else
        If Y <= 17 Then
            m_State = 1
        ElseIf Y >= RcBar.lTop And Y <= RcBar.lTop + RcBar.LHeight Then
            m_State = 2
        ElseIf Y >= ScaleHeight - 17 Then
            m_State = 3
        ElseIf Y < RcBar.lTop Then
            m_State = 4
        ElseIf Y >= RcBar.LHeight + RcBar.lTop Then
            m_State = 5
        End If

        If m_State <> m_Last_State Then
            If m_State < 4 Then
                UserControl.AutoRedraw = True
                UpdateLastPosition

                If Button = 1 Then
                    Select Case m_State
                        Case 1
                            DrawUPButton 2
                        Case 2
                            PaintBar 2
                        Case 3
                            DrawBotButton 2
                    End Select

                Else
                    Select Case m_State
                        Case 1
                            DrawUPButton 1
                        Case 2
                            PaintBar 1
                        Case 3
                            DrawBotButton 1
                    End Select
                End If

                UserControl.Refresh
                UserControl.AutoRedraw = False
            Else
                If m_Last_State < 4 Then
                    UserControl.AutoRedraw = True
                    Select Case m_Last_State
                        Case 1
                            DrawUPButton
                        Case 2
                            PaintBar
                        Case 3
                            DrawBotButton
                    End Select
                    UserControl.Refresh
                    UserControl.AutoRedraw = False
                End If
            End If
        End If

    End If

    m_Last_State = m_State
    If Not m_TMR_Over Then
        Timer1.Enabled = True
        m_TMR_Over = True
    End If

End Sub

Private Sub UpdateLastPosition()
    On Error Resume Next

    If m_Last_State < 4 Then
        Select Case m_Last_State
            Case 1
                DrawUPButton
            Case 2
                PaintBar
            Case 3
                DrawBotButton
        End Select
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    bSlideBar = False
    DrawBasicScrollBar
    m_IsDone = False

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next

    With PropBag
        UserControl.Enabled = .ReadProperty("Enabled", True)
        LargeChange = .ReadProperty("LargeChange", 10)
        Max = .ReadProperty("Max", 255)
        Min = .ReadProperty("Min", 0)
        SmallChange = .ReadProperty("SmallChange", 1)
        Value = .ReadProperty("Value", 0)
        m_ColorScheme = .ReadProperty("ColorScheme", 0)
    End With
    InitBarColor m_ColorScheme, MyBarColor
    bLoaded = True
    UserControl_Resize

End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    UserControl.Width = 255
    If UserControl.Height < 510 Then
        UserControl.Height = 510
    End If
    bLoaded = False
    SetBarHeight
    DrawBasicScrollBar

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "LargeChange", LargeChange, 1
        .WriteProperty "Min", Min, 0
        .WriteProperty "Max", Max, 32767
        .WriteProperty "SmallChange", SmallChange, 1
        .WriteProperty "Value", Value, 0
        .WriteProperty "Colorscheme", m_ColorScheme, 0
    End With
    UserControl_Resize

End Sub

Public Property Get Value() As Long
Attribute Value.VB_HelpID = 8238
Attribute Value.VB_MemberFlags = "200"

    Value = m_Value

End Property

Public Property Let Value(vVal As Long)
    On Error Resume Next

    If vVal = m_Value Then Exit Property

    If vVal >= m_Min And vVal <= m_Max Then
        m_Value = vVal
    ElseIf vVal < m_Min Then
        m_Value = m_Min
    ElseIf vVal > m_Max Then
        m_Value = m_Max
    End If
    DrawBarValue
    RaiseEvent Change
    PropertyChanged "Value"

End Property


Private Sub DrawBackGround(Optional IsRedraw As Boolean)
    On Error Resume Next

    If m_ColorScheme = 2 Then
        If Not IsRedraw Then
            PaintTransMyBlt 0, 0, 17, ScaleHeight, LoadResPicture(9153, 0)
        Else
            PaintTransMyBlt 0, 17, 17, ScaleHeight - 34, LoadResPicture(9153, 0)
        End If
    Else
        If Not IsRedraw Then
            PaintTransMyBlt 0, 0, 17, ScaleHeight, LoadResPicture(9130, 0)
        Else
            PaintTransMyBlt 0, 17, 17, ScaleHeight - 34, LoadResPicture(9130, 0)
        End If
    End If
End Sub

Private Sub DrawClickBody()
    On Error Resume Next

    If Not m_IsDone Then
        If m_ColorScheme = 2 Then
            If m_Last_State = 4 Then
                PaintTransMyBlt 0, 17, 17, RcBar.lTop - 17, LoadResPicture(9265, 0)
            ElseIf m_Last_State = 5 Then
                PaintTransMyBlt 0, RcBar.LHeight + RcBar.lTop, 17, ScaleHeight - 17 - RcBar.LHeight - RcBar.lTop, LoadResPicture(9265, 0)
            End If
        Else
            If m_Last_State = 4 Then
                PaintTransMyBlt 0, 17, 17, RcBar.lTop - 17, LoadResPicture(9264, 0)
            ElseIf m_Last_State = 5 Then
                PaintTransMyBlt 0, RcBar.LHeight + RcBar.lTop, 17, ScaleHeight - 17 - RcBar.LHeight - RcBar.lTop, LoadResPicture(9264, 0)
            End If
        End If
    End If

End Sub

Private Sub DrawUPButton(Optional iPos As Integer)
    On Error Resume Next

    Dim idx As Integer

    If UserControl.Enabled Then
        Select Case m_ColorScheme
            Case 0
                idx = 9136 + iPos
            Case 1
                idx = 9147 + iPos
            Case 2
                idx = 9159 + iPos
        End Select
    Else
        idx = 9260
    End If
    PaintTransMyBlt 1, 0, 16, 17, LoadResPicture(idx, 0)

    m_POS = (5 * iPos) + 1

End Sub

Private Sub DrawBotButton(Optional iPos As Integer)
    On Error Resume Next

    Dim idx As Integer

    If UserControl.Enabled Then
        Select Case m_ColorScheme
            Case 0
                idx = 9139 + iPos
            Case 1
                idx = 9150 + iPos
            Case 2
                idx = 9162 + iPos
        End Select
    Else
        idx = 9261
    End If
    PaintTransMyBlt 1, ScaleHeight - 17, 16, 17, LoadResPicture(idx, 0)
    m_POS = (5 * iPos) + 3

End Sub

Private Sub DrawBasicScrollBar()
    On Error Resume Next

    UserControl.AutoRedraw = True
    DrawBackGround
    DrawUPButton
    DrawBotButton
    PaintBar
    UserControl.Refresh
    UserControl.AutoRedraw = False

End Sub

Private Sub DrawBarValue()
    On Error Resume Next

    UserControl.AutoRedraw = True
    DrawBackGround True
    CalculateBarTop
    PaintBar IIf(m_State = 2, 2, 0)
    DrawClickBody
    UserControl.Refresh
    UserControl.AutoRedraw = False

End Sub

Private Sub SetBarHeight()
    On Error Resume Next

    Dim h_Tmp As Long

    If UserControl.Height > 510 And m_Large > 0 And m_Max > 0 Then

        If m_Large <= m_Max Then
            h_Tmp = (m_Large / (m_Large + m_Max)) * (UserControl.ScaleHeight - 34)
        Else
            h_Tmp = (1 - (m_Max / (m_Large + m_Max))) * (UserControl.ScaleHeight - 34)
        End If

        If h_Tmp < 10 Then
            h_Tmp = 10
        End If

        RcBar.LHeight = h_Tmp
        m_RealSpace = UserControl.ScaleHeight - 34 - h_Tmp

    End If

End Sub

Private Sub CalculateBarTop()
    On Error Resume Next

    If m_Value > 0 Then
        RcBar.lTop = (Value * ((UserControl.ScaleHeight - 34 - RcBar.LHeight) / m_Max)) + 17
    Else
        RcBar.lTop = 17
    End If

End Sub

Private Sub PaintBar(Optional iPos As Integer)
    On Error Resume Next

    Dim Y As Long
    Dim Colr As Long, Colr2   As Long
    Dim StrtPt As Long

    With UserControl

        If .Enabled Then

            If .Height > 705 Then

                CalculateBarTop

                For Colr = 1 To 16
                    DrawXLine .hdc, Colr, RcBar.lTop + 2, Colr, RcBar.lTop + RcBar.LHeight - 2, MyBarColor(iPos, Colr - 1)
                Next Colr

                If m_ColorScheme = 0 Then
                    Colr = vbWhite
                    Colr2 = RGB(140, 176, 248)
                    ResID = 9133
                ElseIf m_ColorScheme = 1 Then
                    Colr = RGB(208, 223, 172)
                    Colr2 = RGB(140, 157, 115)
                    ResID = 9144
                Else
                    Colr = vbWhite
                    Colr2 = RGB(142, 149, 162)
                    ResID = 9156
                End If

                If RcBar.LHeight > 15 Then
                    StrtPt = ((RcBar.LHeight / 2) - 4) + RcBar.lTop
                    For Y = StrtPt To StrtPt + 7 Step 2
                        DrawXLine .hdc, 6, Y, 11, Y, Colr
                    Next Y
                    For Y = StrtPt + 1 To StrtPt + 8 Step 2
                        DrawXLine .hdc, 7, Y, 12, Y, Colr2
                    Next Y

                End If
                PaintTransMyBlt 1, RcBar.lTop, 16, 2, LoadResPicture(ResID - 2, 0)
                PaintTransMyBlt 1, RcBar.lTop + RcBar.LHeight - 3, 16, 3, LoadResPicture(ResID - 1, 0)
            End If

        End If
        m_POS = (5 * iPos) + 2

    End With

End Sub

