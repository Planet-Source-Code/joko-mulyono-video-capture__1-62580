Attribute VB_Name = "ButtonMod"
Option Explicit
'Cteated By Joko Mulyono
'Email:dantex_765@hotmail.com
Public Type TYPERECT
    Left                             As Long
    Top                              As Long
    Right                            As Long
    Bottom                           As Long
End Type
Public Enum Appearance
    Flat = 0
    HalfRaised = 1
    Raised = 2
    Sunken = 3
    Etched = 4
    Bump = 5
    Line = 6
    Push = 7
    PushDown = 8
End Enum
#If False Then
Private Flat, HalfRaised, Raised, Sunken, Etched, Bump, Line, Push, PushDown
#End If
Private Const BDR_RAISEDOUTER    As Long = &H1
Private Const BDR_SUNKENOUTER    As Long = &H2
Private Const BDR_RAISEDINNER    As Long = &H4
Private Const BDR_SUNKENINNER    As Long = &H8
Private Const EDGE_RAISED        As Double = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_ETCHED        As Double = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP          As Double = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT            As Long = &H1
Private Const BF_TOP             As Long = &H2
Private Const BF_RIGHT           As Long = &H4
Private Const BF_BOTTOM          As Long = &H8
Private Const BF_FLAT            As Long = &H4000
Private Const BF_RECT            As Double = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, _
                                                qrc As TYPERECT, _
                                                ByVal edge As Long, _
                                                ByVal grfFlags As Long) As Boolean

Public Sub PaintControl(picBox As PictureBox, _
                        Tampilan As Appearance, _
                        Optional ByVal prov_BackColor As Long, _
                        Optional ByVal prov_ForeColor As Long, _
                        Optional ByVal sCaption As String, _
                        Optional ByVal PDown As Boolean)


Dim typRect As TYPERECT
Dim origScaleMode As Integer
On Error Resume Next
    With picBox
        .BorderStyle = 0
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .Cls
        .BackColor = prov_BackColor
        .ForeColor = prov_ForeColor
    End With 'picBox
    With typRect
        .Right = picBox.ScaleWidth
        .Top = picBox.ScaleTop
        .Left = picBox.ScaleLeft     '    .Top = picBox.ScaleWidth
        .Bottom = picBox.ScaleHeight
    End With 'TYPRECT
    Select Case Tampilan 'm_Appearance
    Case 0
        DrawEdge picBox.hdc, typRect, EDGE_BUMP, BF_FLAT ' BF_FLAT
    Case 1 'HalfRaised
        DrawEdge picBox.hdc, typRect, BDR_RAISEDINNER, BF_RECT 'HalfRaised
    Case 2 'Raised
        With picBox
            DrawEdge .hdc, typRect, EDGE_RAISED, BF_RECT
        End With 'picBox
    Case 3 'sunken
        DrawEdge picBox.hdc, typRect, BDR_SUNKENOUTER, BF_RECT
    Case 4 'etched
        DrawEdge picBox.hdc, typRect, EDGE_ETCHED, BF_RECT
    Case 5 'Bump
        DrawEdge picBox.hdc, typRect, EDGE_BUMP, BF_RECT
    Case 7
        xPush picBox
    Case 8
        xPushDown picBox
    End Select
    picBox.ScaleMode = origScaleMode
    If PDown Then
        picBox.CurrentX = ((picBox.ScaleWidth - picBox.TextWidth(sCaption)) / 2) + 1
        picBox.CurrentY = ((picBox.ScaleHeight - picBox.TextHeight(sCaption)) / 2) + 1
    Else 'PDOWN = FALSE/0
        picBox.CurrentX = (picBox.ScaleWidth - picBox.TextWidth(sCaption)) / 2
        picBox.CurrentY = (picBox.ScaleHeight - picBox.TextHeight(sCaption)) / 2
    End If
    picBox.Print sCaption
    If picBox.AutoRedraw Then
        picBox.Refresh
    End If
    On Error GoTo 0

End Sub

Private Sub xPush(picBox As PictureBox)

    With picBox
        'Right
        picBox.Line (picBox.ScaleWidth - 1, picBox.ScaleHeight)-(picBox.ScaleWidth - 1, 0), RGB(170, 175, 179) ' RGB(48, 49, 51)'Right
        picBox.Line (picBox.ScaleWidth - 2, picBox.ScaleHeight - 1)-(picBox.ScaleWidth - 2, 1), RGB(48, 49, 51) ' RGB(48, 49, 51)'Right2
        picBox.Line (picBox.ScaleWidth - 3, picBox.ScaleHeight - 2)-(picBox.ScaleWidth - 3, 2), RGB(87, 91, 93) ' RGB(48, 49, 51)'Right3
        'Left
        picBox.Line (0, 0)-(0, picBox.ScaleHeight), RGB(75, 80, 84)
        'vb3DShadow ' bottomleft ' vbButtonFace 'Left1
        picBox.Line (1, 1)-(1, picBox.ScaleHeight - 1), RGB(48, 49, 51)
        'vb3DShadow ' bottomleft ' vbButtonFace 'left2
        picBox.Line (2, 2)-(2, picBox.ScaleHeight - 2), RGB(203, 206, 208)
        'vb3DShadow ' bottomleft ' vbButtonFace 'left3
        'Bottom
        picBox.Line (0, picBox.ScaleHeight - 1)-(picBox.ScaleWidth - 1, picBox.ScaleHeight - 1), RGB(170, 175, 179)  'Bottom
        picBox.Line (1, picBox.ScaleHeight - 2)-(picBox.ScaleWidth - 2, picBox.ScaleHeight - 2), RGB(48, 49, 51) 'RGB(87, 91, 93)  'Bottom
        picBox.Line (2, picBox.ScaleHeight - 3)-(picBox.ScaleWidth - 3, picBox.ScaleHeight - 3), RGB(87, 91, 93) 'RGB(87, 91, 93)  'Bottom
        'top
        picBox.Line (0, 0)-(picBox.ScaleWidth, 0), RGB(75, 80, 84) 'RGB(87, 91, 93) 'top side
        picBox.Line (1, 1)-(picBox.ScaleWidth - 1, 1), RGB(48, 49, 51) 'RGB(87, 91, 93) 'top side
        picBox.Line (2, 2)-(picBox.ScaleWidth - 2, 2), RGB(203, 206, 208)
        'RGB(87, 91, 93) 'top side
        'Edge top left 1
        picBox.Line (0, 0)-(1, 0), RGB(98, 103, 107)
        'Edge top left 2
        picBox.Line (2, 2)-(2, 1), RGB(234, 235, 236)
        picBox.Line (0, picBox.ScaleHeight - 1)-(0, picBox.ScaleHeight), RGB(129, 134, 138)
        picBox.Line (0, picBox.ScaleHeight - 2)-(0, picBox.ScaleHeight - 1), RGB(109, 114, 118)
        picBox.Line (1, picBox.ScaleHeight - 1)-(2, picBox.ScaleHeight), RGB(141, 146, 150)
        picBox.Line (2, picBox.ScaleHeight - 3)-(2, picBox.ScaleHeight - 2), RGB(135, 140, 144)
        'Edge top Right 1
        picBox.Line (picBox.ScaleWidth - 1, 0)-(picBox.ScaleWidth, 0), RGB(129, 134, 138)
        'Edge top Right 2
        picBox.Line (picBox.ScaleWidth - 2, 0)-(picBox.ScaleWidth - 1, 0), RGB(109, 114, 118)
        picBox.Line (picBox.ScaleWidth - 3, 2)-(picBox.ScaleWidth - 3, 3), RGB(135, 140, 144)
        picBox.Line (picBox.ScaleWidth - 1, 1)-(picBox.ScaleWidth - 1, 2), RGB(141, 146, 150)
        'Edge Bottom right
        picBox.Line (picBox.ScaleWidth - 1, picBox.ScaleHeight)-(picBox.ScaleWidth - 1, picBox.ScaleHeight - 2), RGB(169, 174, 178)
        picBox.Line (picBox.ScaleWidth - 1, picBox.ScaleHeight - 2)-(picBox.ScaleWidth - 1, picBox.ScaleHeight - 3), RGB(181, 186, 190)
        picBox.Line (picBox.ScaleWidth - 2, picBox.ScaleHeight)-(picBox.ScaleWidth - 2, picBox.ScaleHeight - 2), RGB(181, 186, 190)
        picBox.Line (picBox.ScaleWidth - 3, picBox.ScaleHeight - 3)-(picBox.ScaleWidth - 3, picBox.ScaleHeight - 2), RGB(72, 75, 77)
    End With 'PICBOX

End Sub

Private Sub xPushDown(picBox As PictureBox)

    With picBox
        'Right
        picBox.Line (picBox.ScaleWidth - 1, picBox.ScaleHeight)-(picBox.ScaleWidth - 1, 0), RGB(170, 175, 179) ' 'Right
        picBox.Line (picBox.ScaleWidth - 2, picBox.ScaleHeight - 1)-(picBox.ScaleWidth - 2, 1), RGB(48, 49, 51) ' 'Right2
        picBox.Line (picBox.ScaleWidth - 3, picBox.ScaleHeight - 2)-(picBox.ScaleWidth - 3, 2), RGB(203, 206, 208) ' 'Right3
        'Left
        picBox.Line (0, 0)-(0, picBox.ScaleHeight), RGB(75, 80, 84)
        'vb3DShadow ' bottomleft ' vbButtonFace 'Left1
        picBox.Line (1, 1)-(1, picBox.ScaleHeight - 1), RGB(48, 49, 51)
        'vb3DShadow ' bottomleft ' vbButtonFace 'left2
        picBox.Line (2, 2)-(2, picBox.ScaleHeight - 2), RGB(87, 91, 93)
        'vb3DShadow ' bottomleft ' vbButtonFace 'left3
        'Bottom
        picBox.Line (0, picBox.ScaleHeight - 1)-(picBox.ScaleWidth - 1, picBox.ScaleHeight - 1), RGB(170, 175, 179)  'Bottom
        picBox.Line (1, picBox.ScaleHeight - 2)-(picBox.ScaleWidth - 2, picBox.ScaleHeight - 2), RGB(48, 49, 51) 'Bottom
        picBox.Line (2, picBox.ScaleHeight - 3)-(picBox.ScaleWidth - 3, picBox.ScaleHeight - 3), RGB(203, 206, 208) 'Bottom 3
        'top
        picBox.Line (0, 0)-(picBox.ScaleWidth, 0), RGB(75, 80, 84) 'RGB(87, 91, 93) 'top side
        picBox.Line (1, 1)-(picBox.ScaleWidth - 1, 1), RGB(48, 49, 51) 'RGB(87, 91, 93) 'top side
        picBox.Line (2, 2)-(picBox.ScaleWidth - 2, 2), RGB(87, 91, 93)  'RGB(87, 91, 93) 'top side
        'Edge top left 1
        picBox.Line (0, 0)-(1, 0), RGB(98, 103, 107)
        'Edge top left 2
        picBox.Line (2, 2)-(2, 1), RGB(72, 75, 77)
        picBox.Line (0, picBox.ScaleHeight - 1)-(0, picBox.ScaleHeight), RGB(129, 134, 138)
        picBox.Line (0, picBox.ScaleHeight - 2)-(0, picBox.ScaleHeight - 1), RGB(109, 114, 118)
        picBox.Line (1, picBox.ScaleHeight - 1)-(2, picBox.ScaleHeight), RGB(141, 146, 150)
        picBox.Line (2, picBox.ScaleHeight - 3)-(2, picBox.ScaleHeight - 2), RGB(135, 140, 144)
        'Edge top Right 1
        picBox.Line (picBox.ScaleWidth - 1, 0)-(picBox.ScaleWidth, 0), RGB(129, 134, 138)
        'Edge top Right 2
        picBox.Line (picBox.ScaleWidth - 2, 0)-(picBox.ScaleWidth - 1, 0), RGB(109, 114, 118)
        picBox.Line (picBox.ScaleWidth - 3, 2)-(picBox.ScaleWidth - 3, 3), RGB(135, 140, 144)
        picBox.Line (picBox.ScaleWidth - 1, 1)-(picBox.ScaleWidth - 1, 2), RGB(141, 146, 150)
        'Edge Bottom right
        picBox.Line (picBox.ScaleWidth - 1, picBox.ScaleHeight)-(picBox.ScaleWidth - 1, picBox.ScaleHeight - 2), RGB(169, 174, 178)
        picBox.Line (picBox.ScaleWidth - 1, picBox.ScaleHeight - 2)-(picBox.ScaleWidth - 1, picBox.ScaleHeight - 3), RGB(181, 186, 190)
        picBox.Line (picBox.ScaleWidth - 2, picBox.ScaleHeight)-(picBox.ScaleWidth - 2, picBox.ScaleHeight - 2), RGB(181, 186, 190)
        picBox.Line (picBox.ScaleWidth - 3, picBox.ScaleHeight - 3)-(picBox.ScaleWidth - 3, picBox.ScaleHeight - 2), RGB(234, 235, 236) 'RGB(72, 75, 77)
    End With 'PICBOX

End Sub
'Public Sub Sleep(ByVal Seconds As Double)
'Dim TempTime As Double
'TempTime = Timer
'Do While Timer - TempTime < Seconds
'DoEvents
'If Timer < TempTime Then
'TempTime = TempTime - 24# * 3600#
'End If
'Loop
'End Sub
