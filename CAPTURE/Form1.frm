VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "      VIDEO CAPTURE"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5550
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPEACE 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   360
      Picture         =   "Form1.frx":34CA
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   407
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   210
      Picture         =   "Form1.frx":DC64
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   86
      TabIndex        =   13
      Top             =   4080
      Width           =   1290
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   660
         Picture         =   "Form1.frx":1178E
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   14
         Top             =   480
         Width           =   480
         Begin VB.PictureBox Picture7 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   105
            Index           =   0
            Left            =   15
            Picture         =   "Form1.frx":11DD0
            ScaleHeight     =   7
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   16
            Top             =   15
            Width           =   135
         End
         Begin VB.PictureBox Picture7 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   105
            Index           =   3
            Left            =   15
            Picture         =   "Form1.frx":11ED6
            ScaleHeight     =   7
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   15
            Top             =   120
            Width           =   135
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   285
            TabIndex        =   17
            Top             =   15
            Width           =   105
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   375
         TabIndex        =   18
         Top             =   210
         Width           =   525
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   5760
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   5
      Left            =   600
      Picture         =   "Form1.frx":11FDC
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   4
      Left            =   480
      Picture         =   "Form1.frx":120E2
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   2
      Left            =   360
      Picture         =   "Form1.frx":121E8
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   1
      Left            =   240
      Picture         =   "Form1.frx":122EE
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   6240
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      Picture         =   "Form1.frx":123F4
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   7
      Top             =   4080
      Width           =   3735
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         Picture         =   "Form1.frx":152F6
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   8
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Index           =   3
      Left            =   4440
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Index           =   2
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Index           =   1
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Index           =   0
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      ForeColor       =   &H0000FF00&
      Height          =   3780
      Left            =   240
      ScaleHeight     =   3750
      ScaleWidth      =   5025
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   3780
      Left            =   240
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frame :"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Player         As DANTEplayer
Private VCapture         As DANTEplayer
Private nox              As Long
Private Film             As String
Private isPause          As Boolean
Private BtnIndex         As Integer
Private ChangePosition   As Boolean
Private X1               As Single
Private VLength          As Long
Private Slider_Down      As Boolean
Private idxcount         As Integer
Private btnidx           As Integer
Private isPlaying        As Boolean
Private isremain         As Boolean

Private Sub Cap()
Dim nb As Integer
Dim m  As Long

    nox = CLng(Val(m_Player.POSFORMAT(ByMS))) - 1 'retrieve 1 frame
    Picture2.Visible = True
    Picture2.ZOrder 0
    For nb = 0 To Val(Label2.Caption) - 1
        m = m + 1
        With VCapture
            .AliasName = "VideoCapture"
            .Filename = Film
            .hwndParent = Picture2
            .OpenCapture Picture2.hwnd, nox + m, .Filename, 0, 0, Picture2.Width / 15, Picture2.Height / 15
            .PutVideoCapture Picture2.hwnd, 0, 0, Picture2.Width, Picture2.Height
            .setCommand (PauseCD)
        End With
        Capture Picture2
        SavePicture Clipboard.GetData, App.Path & "\Capture" & nox + m - 1 & ".bmp"
        If m = Val(Label2.Caption) Then
            m = 0
            Picture2.Cls
            Picture2.Picture = Nothing
            Exit For
        End If
        DoEvents
        VCapture.setCommand (ResumeCD)
    Next nb
End Sub

Private Sub ClosePlayer()

    Timer1.Enabled = False
    With m_Player
        .setCommand StopCD
        .setCommand CloseCD
        isPlaying = False
    End With 'M_PLAYER

End Sub

Private Sub Form_Load()

Dim n2   As Integer
Dim xStr As String

    Set m_Player = New DANTEplayer
    Set VCapture = New DANTEplayer
    Picture2.BackColor = RGB(16, 0, 16)
    Me.BackColor = RGB(131, 135, 144)
    For n2 = 0 To 3
        Select Case n2
        Case 0
            xStr = "Capture"
        Case 1
            xStr = "Stop"
        Case 2
            xStr = "Open"
        Case 3
            xStr = "Exit"
        End Select
        PaintControl Picture3(n2), Push, RGB(75, 75, 80), vbWhite, xStr, False
    Next n2
    PaintControl Picture4, Push, RGB(75, 75, 80), vbWhite, "", False
    INSERT picPEACE.hwnd, Me.hwnd, 6, 247
    SetWidth = 307
    'PaintControl Picture8, Push, RGB(75, 75, 80), vbWhite, "", False

End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)

Dim xStr As String
    Select Case BtnIndex
            Case 0
                xStr = "Capture"
            Case 1
                xStr = "Stop"
            Case 2
                xStr = "Open"
            Case 3
                xStr = "Exit"
    End Select
    PaintControl Picture3(BtnIndex), Push, RGB(75, 75, 80), vbWhite, xStr, False
    If ChangePosition Then
        If Button = 1 Then
            x = x - X1
            If x < Picture4.Left + 1 Then
                x = Picture4.Left + 1
            End If
            If x > (Picture4.Left + Picture4.Width) - Picture5.Width - 1 Then
                x = (Picture4.Left + Picture4.Width) - Picture5.Width - 1
            End If
            Picture5.Left = x
        End If
        DoEvents
    End If

End Sub

Private Sub Label3_Click()

    isremain = IIf(isremain, False, True)
    
    If Not isremain Then
        isremain = False
    Else
        isremain = True
    End If

End Sub

Private Sub Label3_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    Label3.ToolTipText = "Click to change mode"

End Sub

Private Sub Picture2_Click()

    isPause = IIf(isPause, False, True)
   
    If isPause Then
        m_Player.setCommand (PauseCD)
    Else
        m_Player.setCommand (ResumeCD)
    End If

End Sub

Private Sub Picture2_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

    If Button = 1 Then
        Picture2.MousePointer = 99
        Picture2.MouseIcon = LoadResPicture(102, 1)
    End If

End Sub

Private Sub Picture2_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

    If Button <> 1 Then
        Picture2.MousePointer = 99
        Picture2.MouseIcon = LoadResPicture(101, 1)
    End If

End Sub

Private Sub Picture2_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)


'Picture2.MousePointer = 99
'Picture2.MouseIcon = LoadResPicture(102, 1)


End Sub

Private Sub Picture3_MouseDown(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

Dim xStr As String

    Select Case Index
    Case 0
        xStr = "Capture"
    Case 1
        xStr = "Stop"
    Case 2
        xStr = "Open"
    Case 3
        xStr = "Exit"
    End Select
    PaintControl Picture3(Index), PushDown, RGB(52, 71, 80), vbWhite, xStr, False

End Sub

Private Sub Picture3_MouseMove(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

Dim xStr As String
BtnIndex = Index
    Select Case Index
    Case 0
        xStr = "Capture"
    Case 1
        xStr = "Stop"
    Case 2
        xStr = "Open"
    Case 3
        xStr = "Exit"
    End Select
    PaintControl Picture3(Index), Push, RGB(33, 34, 39), vbWhite, xStr, False

End Sub

Private Sub Picture3_MouseUp(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

Dim exfname As Boolean
Dim spath As String
    Select Case Index
    Case 0
        Cap
        VCapture.CloseCapture
    Case 1
        ClosePlayer
    Case 2
        ClosePlayer
        Filter = "Movie (*.dat;*.mpg;*.avi;*.asf;*.wmv)" & Chr$(0) & "*.dat;*.mpg;*.avi;*.wmv" & Chr$(0) & "Other Mov(*.mov)" & Chr$(0) & "*.mov" & Chr$(0)
        exfname = OpenDialog(hwnd, vbNullString, 0, vbNullString, vbNullString, "  OPEN MEDIA FILE", spath)
        If LenB(ExFilename) Then
            Film = ExtractString(ExFilename)
        End If
        Play
    Case 3
        ClosePlayer
        Unload Me
    End Select

End Sub

Private Sub Picture5_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

'ChangePosition = True
'Picture5.Cls
'X1 = Int(ScaleX(X, 1, 3))

    GS.pMIN = 0
    GS.Rev = False
    If Button = vbLeftButton Then
        Sdr_Down Picture5, x, y, Horizontally
    End If
    With m_Player
        GS.pMAX = CLng(Val(.LENGTHFORMAT(ByMS)))
        GS.CheckP = .getStatusInfo(Position)
    End With 'M_PLAYER
    'tmrCONTROL.Enabled = False
    'Timer1.Enabled = False
    Slider_Down = True
    GS.SdrMove = True
Pesan:
    With Err
        If .Number <> 0 Then
            MsgBox .Description, vbCritical + vbOKOnly, "scrolldown"
            .Clear
        End If
    End With 'Err

End Sub

Private Sub Picture5_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

' Form_MouseMove Button, Shift, Int(ScaleX(X, 1, 3)) + Picture5.left, ScaleY(Y, 1, 3) + Picture5.top

    If Button = vbLeftButton Then
        Sdr_Move Picture4, Picture5, x, y, Horizontally
    End If
    GS.nwPOS = GS.crPOS

End Sub

Private Sub Picture5_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

Dim nVal2 As Long

    nVal2 = GS.CheckP - (GS.crPOS - (GS.nwPOS - GS.oldpos) * 1000)
    m_Player.setCommand seekCD, , nVal2
    'Timer1.Enabled = True
    Slider_Down = False
    GS.SdrMove = False

End Sub

Private Sub Picture7_MouseDown(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

    Picture7(Index).Picture = Picture7(Index + 2)
    btnidx = Index
    Timer2.Enabled = True

End Sub

Private Sub Picture7_MouseUp(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    Picture7(Index).Picture = Picture7(Index + 1)
    Timer2.Enabled = False

End Sub

Private Sub Play()

Dim fStr As String

    Picture2.Cls
    If LenB(Film) Then
        With m_Player
            .setCommand StopCD
            .Filename = Film
            .hwndParent = Picture1.hwnd
            .PlayMEDIAFILE
            VLength = .LENGTHFORMAT(ByMS)
            isPlaying = True
            Timer1.Enabled = True
        End With 'M_PLAYER
        fStr = GetFileName(Film, True)
        Picture2.CurrentX = Picture2.ScaleWidth - (TextWidth(fStr) * Screen.TwipsPerPixelX) - 5
        Picture2.CurrentY = Picture2.ScaleHeight - (TextHeight(fStr) * Screen.TwipsPerPixelY) - 5
        Picture2.Print fStr
    End If
    RefreshButton

End Sub

Private Sub RefreshButton()

Dim xStr As String
Dim nBT As Integer
    For nBT = 0 To 3
        Select Case nBT
        Case 0
            xStr = "Capture"
        Case 1
            xStr = "Stop"
        Case 2
            xStr = "Open"
        Case 3
            xStr = "Exit"
        End Select
        PaintControl Picture3(nBT), Push, RGB(75, 75, 80), vbWhite, xStr, False
    Next nBT

End Sub

Private Sub Timer1_Timer()

Dim str As String
Dim xV  As Single

    GS.sMax = m_Player.LENGTHFORMAT(ByMS)
    GS.oldpos = Val(m_Player.POSFORMAT(ByMS))
    If isPlaying Then
        xV = Val(m_Player.POSFORMAT(ByMS)) / Val(m_Player.getStatusInfo(Duration))
    End If
    If Not Slider_Down Then
        getPos xV, Picture5, Picture4, Horizontally, False
    End If
  
    If Not isremain Then
        str = m_Player.POSFORMAT(ByTMSF)
    Else
        str = m_Player.FormatRemain
    End If
    Label3.Caption = str

End Sub

Private Sub Timer2_Timer()

    If btnidx = 0 Then
        If idxcount < 99 Then
            idxcount = idxcount + 1
        End If
    ElseIf btnidx = 3 Then
        If idxcount > 1 Then
            idxcount = idxcount - 1
        Else
            idxcount = 1
        End If
    End If
    Label2.Caption = idxcount

End Sub


