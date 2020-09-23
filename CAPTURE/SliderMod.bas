Attribute VB_Name = "SliderMod"
Option Explicit
Public Enum SliderOrientation
    Vertically = 0
    Horizontally = 1
End Enum
#If False Then
Private Vertically, Horizontally
#End If
Public Type graphicSLIDER
    crPOS      As Long    'posisi up to date
    nwPOS      As Long    'posisi frame saat di drag
    oldpos     As Long    'posisi slider saat di drag
    pMAX       As Long    'property max
    pMIN       As Long    'property min
    sMax       As Long
    lVol       As Long
    vBal       As Long
    pLENGTH    As Long
    sOff       As Single
    XDrag      As Single
    YDrag      As Single
    Rev        As Boolean
    ChgVal     As Boolean
    ChgBal     As Boolean
    SdrMove    As Boolean
    CheckP     As Long
End Type
Public GS  As graphicSLIDER

Public Function getPos(ByVal nValue As Single, _
                       Obj As Object, _
                       ByVal BaseObj As Object, _
                       sldOrientation As SliderOrientation, _
                       Optional ByVal Reverse As Boolean) As Single


'Tested:OK

Dim Offset  As Single
Dim pLENGTH As Single

    If Not GS.SdrMove Then
        If sldOrientation = Vertically Then
            pLENGTH = BaseObj.ScaleHeight - Obj.Height
            If Not Reverse Then
                Offset = pLENGTH * nValue
                Obj.Top = -Obj.Height + Offset + Obj.Height
            Else
                Offset = -pLENGTH * nValue
                Obj.Top = BaseObj.ScaleHeight + Offset - Obj.Height
            End If
            getPos = Obj.Top
        Else
            pLENGTH = BaseObj.ScaleWidth - Obj.Width
            If Not Reverse Then
                Offset = pLENGTH * nValue
                Obj.Left = -Obj.Width + Offset + Obj.Width ' + 1
            Else
                Offset = -pLENGTH * nValue
                Obj.Left = BaseObj.ScaleWidth + Offset - Obj.Width
            End If
            getPos = Obj.Left
        End If
    End If

End Function

Private Sub GetValue(ByVal Max As Long, _
                     iRel As PictureBox, _
                     pctBtn As PictureBox, _
                     ByVal posisi As Long, _
                     sldOrientation As SliderOrientation, _
                     ByVal Reverse As Boolean)


'Tested:OK

    On Error Resume Next
    If sldOrientation = Horizontally Then
       
        With pctBtn
            GS.pLENGTH = iRel.ScaleWidth - .Width
            .Left = posisi
            If .Left < 0 Then
                .Left = 0
            End If
        End With 'pctBtn
        If pctBtn.Left > GS.pLENGTH Then
            pctBtn.Left = GS.pLENGTH
        End If
        pctBtn.Picture = pctBtn.Image
        iRel.Picture = iRel.Image
        If Not Reverse Then
            GS.crPOS = CInt((pctBtn.Left * GS.pMAX) / GS.pLENGTH) + GS.pMIN
        Else
            GS.crPOS = CInt(-(pctBtn.Left * GS.pMAX) / GS.pLENGTH) + GS.pMAX 'GS.pMIN
        End If
    Else
        
        With pctBtn
            GS.pLENGTH = iRel.ScaleHeight - .Height
            .Top = posisi
            If .Top < 0 Then
                .Top = 0
            End If
        End With 'pctBtn
        If pctBtn.Top > GS.pLENGTH Then
            pctBtn.Top = GS.pLENGTH
        End If
        pctBtn.Picture = pctBtn.Image
        iRel.Picture = iRel.Image
        If Reverse Then
            GS.crPOS = CInt(-(pctBtn.Top * GS.pMAX) / GS.pLENGTH) + GS.pMAX
        Else
            GS.crPOS = CInt((pctBtn.Top * GS.pMAX) / GS.pLENGTH)
        End If
    End If
    On Error GoTo 0

End Sub

Public Sub Sdr_Down(pctBtn As PictureBox, _
                    ByVal x As Single, _
                    ByVal y As Single, _
                    sldOrientation As SliderOrientation)


    If Not GS.SdrMove Then
        If sldOrientation = Horizontally Then
            GS.YDrag = y
            GS.XDrag = pctBtn.Left
        Else
            GS.YDrag = y
            GS.XDrag = pctBtn.Top
        End If
        GS.SdrMove = True
    End If

End Sub

Public Sub Sdr_Move(iRel As PictureBox, _
                    pctBtn As PictureBox, _
                    ByVal x As Single, _
                    ByVal y As Single, _
                    sldOrientation As SliderOrientation)


'Tested:OK

    On Error GoTo Pesan
    If GS.pMAX = 0 Then
        GS.SdrMove = False
    End If
    If sldOrientation = Horizontally Then
     
        With GS
            .pLENGTH = iRel.ScaleWidth - pctBtn.Width
            If .SdrMove Then
                .crPOS = .XDrag + (x - .YDrag)
                If .crPOS < 0 Then
                    .crPOS = 0
                End If
                If .crPOS > .pLENGTH Then
                    .crPOS = .pLENGTH
                End If
                .XDrag = .crPOS
                GetValue .pMAX, iRel, pctBtn, .crPOS, sldOrientation, .Rev
            End If
        End With
    Else
        GS.pLENGTH = iRel.ScaleHeight - pctBtn.Height
        If GS.pMAX = 0 Then
            GS.SdrMove = False
        End If
        With GS
            If .SdrMove Then
                .crPOS = .XDrag + (y - .YDrag)
                If .crPOS < 0 Then
                    .crPOS = 0
                End If
                If .crPOS > .pLENGTH Then
                    .crPOS = .pLENGTH
                End If
                pctBtn.Picture = pctBtn.Image
                iRel.Picture = iRel.Image
                .XDrag = .crPOS
                GetValue .pMAX, iRel, pctBtn, .crPOS, sldOrientation, .Rev
            End If
        End With
    End If

Exit Sub

Pesan:
    GS.SdrMove = False

End Sub
'Public Sub PaintProgress()
'With frmMain
'.Picture4.Cls
'.Picture4.Line (0, 0)-(.Picture4.ScaleWidth, .Picture5.ScaleHeight + 2), RGB(16, 0, 16), BF
''Picture4.Line (0, 0)-(Picture5.Left, Picture4.ScaleHeight), RGB(200, 227, 98), BF
'End With
'End Sub

