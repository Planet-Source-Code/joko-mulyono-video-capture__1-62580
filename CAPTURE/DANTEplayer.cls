VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DANTEplayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Cteated By Joko Mulyono
'Email:dantex_765@hotmail.com
Option Explicit
Option Compare Text
Private m_AliasName     As String
Private sCommand        As Long
Private sDrive          As String
Private nReturn         As Long
Private m_LengthVideo   As String
Private sFileName       As String
Private m_HwndParent    As Long
Public Enum MCI_COMMAND
    DoorOpen
    DoorClose
    OpenCD
    StopCD
    PlayCD
    seekCD
    SpeedCD
    CloseCD
    PauseCD
    ResumeCD
    VideoOff
    VideoOn
End Enum
#If False Then
Private DoorOpen, DoorClose, OpenCD, StopCD, PlayCD, seekCD, SpeedCD, CloseCD, PauseCD, ResumeCD, VideoOff, VideoOn
#End If
Public Enum STATUS_INFO
    total_length
    total_frames
    Mode
    Position
    time_format
    Frame_Rate
    Speed_Rate
    WindowHDC
    Duration
End Enum
#If False Then
Private total_length, total_frames, Mode, Position, time_format, Frame_Rate, Speed_Rate, WindowHDC, Duration
#End If
Public Enum TimeFormat
    ByMS
    ByTMSF
    byFrames
End Enum
#If False Then
Private ByMS, ByTMSF, byFrames
#End If
Public Enum Vid_State
    vd_On
    vd_Off
End Enum
#If False Then
Private vd_On, vd_Off
#End If
Private Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type
Public Enum MOVIEDIMENTION
    mvd_Width
    mvd_Height
End Enum
#If False Then
Private mvd_Width, mvd_Height
#End If
Public Enum CHANELVOLUME
    Chan_Right
    Chan_Left
    Chan_All
End Enum
#If False Then
Private Chan_Right, Chan_Left, Chan_All
#End If
Public Enum xStyle
    Custom = 0
    Windows = 1
    Desktop = 2
    FScreen = 3
End Enum
#If False Then
Private Custom, Windows, Desktop, FScreen
#End If
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                      ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
                                                                                    ByVal lpszShortPath As String, _
                                                                                    ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, _
                                                                               ByVal lpstrReturnString As String, _
                                                                               ByVal uReturnLength As Long, _
                                                                               ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, _
                                                                                       ByVal lpstrBuffer As String, _
                                                                                       ByVal uLength As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
                                                      lpRect As Any, _
                                                      ByVal bErase As Long) As Long

Public Property Get AliasName() As String

    AliasName = m_AliasName

End Property

Public Property Let AliasName(ByVal new_AliasName As String)

    m_AliasName = new_AliasName

End Property

Private Sub Class_Initialize()

    If LenB(AliasName) = 0 Then
        AliasName = "MYPLAYER"
    End If

End Sub

Public Sub CloseCapture()

    mciSendString "close VideoCapture", 0, 0, 0

End Sub

Public Property Get DeviceType() As String

Dim EXT    As String
Dim Device As String

    EXT = UCase$(GetExtension(Filename))
    Select Case EXT
    Case "MID"
        Device = "Sequencer"
    Case "AIF", "AIFC", "ASX", "AU", "AVI", "DAT", "ENC", "M1V", "MID", "MOV", "MP2", "MPA", "MPE", "MPEG", "MPG", "MPM", "MPV", "MPV2", "QT", "RMI", "VOB", "WAV", "WAX", "WMA", "WMP", "WMV", "wmv", "WMX"
        Device = "MPEGVideo"
    Case Else
        Exit Property
    End Select
    DeviceType = Device

End Property

Public Property Get Filename() As String

Dim Tmp     As String * 255
Dim lenPath As Long

    Filename = sFileName
    lenPath = GetShortPathName(Filename, Tmp, 255)
    Filename = Left$(Tmp, lenPath)

End Property

Public Property Let Filename(ByVal xFile As String)

    sFileName = xFile

End Property

Public Function FormatRemain() As String

Dim nHour    As Integer
Dim nMinut   As Integer
Dim nSeconds As Integer
Dim newTimer As Long
Dim strTimer As String

    newTimer = CLng(Val(LENGTHFORMAT(ByMS)) - Val(POSFORMAT(ByMS)))
    nHour = Int(newTimer / 3600)
    newTimer = newTimer Mod 3600
    nMinut = Int(newTimer / 60)
    newTimer = newTimer Mod 60
    nSeconds = newTimer
    If nHour > 0 Then
        strTimer = Trim$(str(nHour)) & ":"
    Else
        strTimer = ""
    End If
    If nMinut >= 10 Then
        strTimer = strTimer & Trim$(str(nMinut))
    ElseIf nMinut > 0 Then
        strTimer = strTimer & Trim$(str(nMinut))
    Else
        strTimer = strTimer & "0"
    End If
    strTimer = Format$(strTimer, "00") & ":"
    If nSeconds >= 10 Then
        strTimer = strTimer & Trim$(str(nSeconds))
    ElseIf nSeconds > 0 Then
        strTimer = strTimer & "0" & Trim$(str(nSeconds))
    Else
        strTimer = strTimer & "00"
    End If
    FormatRemain = strTimer

End Function

Private Function GetExtension(ByVal FPath As String) As String

Dim p As Long

    If Len(FPath) > 0 Then
        p = InStrRev(FPath, ".")
        If p > 0 Then
            If p < Len(FPath) Then
                GetExtension = Mid$(FPath, p + 1)
            End If
        End If
    End If

End Function

Public Function getStatusInfo(ByRef Info As STATUS_INFO) As String

Dim strbuffer As String

    strbuffer = String$(256, vbNullChar)
    Select Case Info
    Case total_frames
        mciSendString "set " & AliasName & " time format frames", strbuffer, Len(strbuffer), 0&
        mciSendString "status " & AliasName & " length", strbuffer, Len(strbuffer), 0&
    Case Mode 'playing,paused
        mciSendString "status " & AliasName & " mode", strbuffer, Len(strbuffer), 0
    Case Position
        mciSendString "set " & AliasName & " time format ms", 0, 0, 0
        mciSendString "status " & AliasName & " position", strbuffer, Len(strbuffer), 0
        getString (strbuffer)
        strbuffer = CLng(Val(strbuffer))
    Case Frame_Rate
        mciSendString "Status " & AliasName & " frame rate", strbuffer, Len(strbuffer), 0
        strbuffer = Left$(strbuffer, InStr(strbuffer, vbNullChar) - 1)
    Case Speed_Rate
        mciSendString "status " & AliasName & " speed ", strbuffer, Len(strbuffer), 0
        strbuffer = Val(strbuffer) / 10
    Case time_format 'return : tmsf,ms
        mciSendString "status " & AliasName & " time format", strbuffer, Len(strbuffer), 0
    Case WindowHDC
        mciSendString "status " & AliasName & " window handle", strbuffer, Len(strbuffer), 0
        strbuffer = Mid$(strbuffer, 1, InStr(1, strbuffer, vbNullChar) - 1)
    Case Duration
        SetTimeFormat (ByMS)
        mciSendString "status " & AliasName & " length", strbuffer, Len(strbuffer), 0
        strbuffer = Round(Val(Mid$(strbuffer, 1, Len(strbuffer))) / 1000)
    End Select
    getStatusInfo = getString(strbuffer)

End Function

Public Function getString(str As String) As String

Dim a As Integer

    For a = 1 To Len(str)
        If Mid$(str, a, 1) = vbNullChar Then
            Exit For
        End If
    Next a
    getString = RTrim$(Left$(str, a - 1))

End Function

Public Property Get hwndParent() As Long

    hwndParent = m_HwndParent

End Property

Public Property Let hwndParent(ByVal new_HwndParent As Long)

    m_HwndParent = new_HwndParent

End Property

Public Function LENGTHFORMAT(sFormat As TimeFormat) As String

Dim strbuffer As String
Dim Sec       As Double
Dim Mins      As Integer
Dim str       As String * 128

    strbuffer = String$(256, vbNullChar)
    mciSendString "set " & AliasName & " time format milliseconds", 0, 0, 0
    Select Case sFormat
    Case ByMS
        mciSendString "status " & AliasName & " length", strbuffer, Len(strbuffer), 0
        getString strbuffer
        Sec = Round(Val(Mid$(strbuffer, 1, Len(strbuffer))) / 1000)
        strbuffer = Sec
    Case ByTMSF
        mciSendString "status " & AliasName & " length", strbuffer, Len(strbuffer), 0
        Sec = Round(Val(Mid$(strbuffer, 1, Len(strbuffer))) / 1000)
        'Round(CInt(Mid$(s, 1, Len(s))) / 1000)
        If Sec < 60 Then
            strbuffer = Format$(Sec, "00:00")
        End If
        If Sec > 59 Then
            Mins = Int(Sec / 60)
            Sec = Sec - (Mins * 60)
            strbuffer = Format$(Mins, "00") & ":" & Format$(Sec, "00")
        End If
    Case byFrames
        mciSendString "status " & AliasName & " length", str, Len(str), 0
        strbuffer = str
    End Select
    LENGTHFORMAT = getString(strbuffer)

End Function

Public Function MoveVideoTo(lngHWnd As Long, _
                            lngLeft As Long, _
                            lngTop As Long, _
                            lngWidth As Long, _
                            lngHeight As Long, _
                            Optional margin As Long) As String

Dim ret As String * 128
Dim rec As RECT

    If lngWidth = 0 Or lngHeight = 0 Then
        GetWindowRect lngHWnd, rec
        lngWidth = rec.Right - rec.Left
        lngHeight = rec.Bottom - rec.Top
    End If
    nReturn = mciSendString("put " & AliasName & " window at " & lngLeft + margin & " " & lngTop + margin & " " & lngWidth - (margin * 2) & " " & lngHeight - (margin * 2), 0&, 0&, 0&)
    If Not nReturn = 0 Then
        mciGetErrorString nReturn, ret, 128
        MoveVideoTo = ret
    Else
        MoveVideoTo = "Success"
    End If

End Function

Public Sub OpenCapture(lngHWnd As Long, _
                       where As Long, _
                       ByVal sName As String, _
                       lngLeft As Long, _
                       lngTop As Long, _
                       lngWidth As Long, _
                       lngHeight As Long)


Dim pos As String * 128

    AliasName = "VideoCapture"
    mciSendString "close VideoCapture", 0, 0, 0
    mciSendString "stop VideoCapture", 0, 0, 0
    OpenMediaFile lngHWnd, "VideoCapture", DeviceType
    mciSendString "set VideoCapture time format ms", pos, 128, 0& 'set time format
    mciSendString "set VideoCapture audio all off", 0, 0, 0 'no need sound
    mciSendString "window VideoCapture state hide", 0, 0, 0 'hide first frame
    mciSendString "seek VideoCapture to " & where * 1000, 0, 0, 0
    mciSendString "window VideoCapture state show", 0, 0, 0
    MoveVideoTo lngHWnd, lngLeft, lngTop, lngWidth, lngHeight

End Sub

Public Function OpenMediaFile(lngHWnd As Long, _
                              ByVal AliasName As String, _
                              ByVal typeDevice As String) As String

Dim sCommand   As String * 255
Dim dwReturn   As Long
Const WS_CHILD As Long = &H40000000

    sCommand = "open " & Filename & " type " & typeDevice & " Alias " & AliasName & " parent " & lngHWnd & " Style " & WS_CHILD
    dwReturn = mciSendString(sCommand, 0&, 0&, 0&)
    m_LengthVideo = LENGTHFORMAT(ByMS)
    If Not dwReturn = 0 Then
        mciGetErrorString dwReturn, sCommand, 128
        OpenMediaFile = sCommand
    End If
    OpenMediaFile = "Success"

End Function

Public Sub PlayMEDIAFILE(Optional where As Long)

    OpenMediaFile hwndParent, AliasName, DeviceType
    MoveVideoTo hwndParent, 0, 0, 0, 0
    mciSendString "seek " & AliasName & " to " & where * 1000, 0, 0, 0
    mciSendString "play " & AliasName, 0, 0, 0
    'WALLPAPERTHEME (ClearWall)
    'InvalidateRect 0&, ByVal 0, 1& 'refresh desktop

End Sub

'Public Sub SetVideoState(vfm_State As Vid_State)
'    Select Case vfm_State
'    Case vd_On
'        mciSendString "set all video on", 0, 0, 0
'    Case vd_Off
'        mciSendString "set all video off", 0, 0, 0
'    End Select
'End Sub
Public Function POSFORMAT(sPFormat As TimeFormat) As String

Dim strbuffer As String
Dim Sec       As Double
Dim Mins      As Integer

    On Error Resume Next
    strbuffer = String$(256, vbNullChar)
    Select Case sPFormat
    Case ByMS
        mciSendString "set " & AliasName & " time format tmsf ", 0, 0, 0
        mciSendString "status " & AliasName & " position", strbuffer, Len(strbuffer), 0
        strbuffer = getString(strbuffer)
        Sec = Round(Mid$(strbuffer, 1, Len(strbuffer)) / 1000) 'type mismatch
        strbuffer = Sec
    Case ByTMSF
        mciSendString "set " & AliasName & " time format milliseconds", 0, 0, 0
        mciSendString "status " & AliasName & " position", strbuffer, Len(strbuffer), 0
        Sec = Round(Mid$(strbuffer, 1, Len(strbuffer)) / 1000)
        If Sec < 60 Then
            strbuffer = "00:" & Format$(Sec, "00")
        End If
        If Sec > 59 Then
            Mins = Int(Sec / 60)
            Sec = Sec - (Mins * 60)
            strbuffer = Format$(Mins, "00") & ":" & Format$(Sec, "00")
        End If
    Case byFrames
        mciSendString "status " & AliasName & " position", strbuffer, Len(strbuffer), 0
    End Select
    POSFORMAT = getString(strbuffer)
    On Error GoTo 0

End Function

Public Function PutVideoCapture(lngHWnd As Long, _
                                lngLeft As Long, _
                                lngTop As Long, _
                                lngWidth As Long, _
                                lngHeight As Long) As String

Dim ret As String * 128
Dim rec As RECT

    If lngWidth = 0 Or lngHeight = 0 Then
        GetWindowRect lngHWnd, rec
        lngWidth = rec.Right - rec.Left
        lngHeight = rec.Bottom - rec.Top
    End If
    nReturn = mciSendString("put " & "VideoCapture" & " window at " & lngLeft & " " & lngTop & " " & lngWidth & " " & lngHeight, 0&, 0&, 0&)
    If Not nReturn = 0 Then  'tidak sukses
        mciGetErrorString nReturn, ret, 128 'Pesan Error
        PutVideoCapture = ret
    End If
    PutVideoCapture = "Success"

End Function

Public Sub setCommand(Cmd As MCI_COMMAND, _
                      Optional sDrive As String, _
                      Optional nValue As Long)

Dim pos As String * 128

    mciSendString "open " & sDrive & " Type cdaudio alias " & AliasName & " wait shareable", 0, 0, 0
    Select Case Cmd
    Case DoorOpen
        mciSendString "set " & AliasName & " door open", 0, 0, 0
        mciSendString "capability " & AliasName & " can eject notify", 0, 0, 0
    Case DoorClose
        mciSendString "set " & AliasName & " door closed", 0, 0, 0
    Case CloseCD
        mciSendString "close " & AliasName, 0, 0, 0
    Case PlayCD
        mciSendString "play " & AliasName, 0, 0, 0
    Case seekCD
        mciSendString "set " & AliasName & " time format ms", pos, 128, 0&
        mciSendString "seek " & AliasName & " to " & nValue, 0, 0, 0
        mciSendString "play " & AliasName & " from " & nValue, 0, 0, 0
    Case StopCD
        mciSendString "stop " & AliasName, 0, 0, 0
    Case SpeedCD
        mciSendString "set " & AliasName & " speed " & nValue, 0, 0, 0
    Case PauseCD
        mciSendString "pause " & AliasName & " wait", 0, 0, 0
    Case ResumeCD
        mciSendString "resume " & AliasName, 0, 0, 0
    End Select

End Sub

Public Sub SetTimeFormat(fm_Time As TimeFormat)

'TESTED:OK

    Select Case fm_Time
    Case ByMS
        mciSendString "set " & AliasName & " time format ms wait", 0, 0, 0
    Case ByTMSF
        mciSendString "set " & AliasName & " time format tmsf wait", 0, 0, 0
    Case byFrames
        mciSendString "set " & AliasName & " time format frames wait", 0, 0, 0
    End Select

End Sub

Public Sub SetVolumeState(st_Vol As CHANELVOLUME, _
                          new_Vol As Long)


    Select Case st_Vol
    Case Chan_All
        mciSendString "setaudio " & AliasName & " volume to " & new_Vol, 0, 0, 0
    Case Chan_Right
        mciSendString "setaudio " & AliasName & " right volume to " & new_Vol, 0, 0, 0
    Case Chan_Left
        mciSendString "setaudio " & AliasName & " left volume to " & new_Vol, 0, 0, 0
    End Select

End Sub

Public Sub FForward(ByVal nValue As Long)


Dim pos As String * 128
Dim I   As Long

    I = CLng(getStatusInfo(Position)) 'Current position
    mciSendString "set " & AliasName & " time format ms", pos, 128, 0&
    mciSendString "play " & AliasName & " from " & I + nValue * 1000, 0, 0, 0

End Sub

Public Sub FRewind(ByVal nValue As Long)


Dim pos As String * 128
Dim I   As Long

    I = CLng(getStatusInfo(Position)) 'Current position
    mciSendString "set " & AliasName & " time format ms", pos, 128, 0&
    mciSendString "play " & AliasName & " from " & I - nValue * 1000, 0, 0, 0

End Sub

Public Function GetHwndDesktop() As Long


Dim xhwnd As Long

    On Error Resume Next
    xhwnd = FindWindow(vbNullString, "Program Manager")
    GetHwndDesktop = xhwnd
    On Error GoTo 0

End Function

Public Sub SetAudioState(stChan As CHANELVOLUME, _
                         st_pos As Vid_State)


    Select Case stChan
    Case Chan_All
        If st_pos = vd_On Then
            mciSendString "set " & AliasName & " audio all on", 0, 0, 0
        Else 'NOT ST_POS...
            mciSendString "set " & AliasName & " audio all off", 0, 0, 0
        End If
    Case Chan_Right
        If st_pos = vd_On Then
            mciSendString "set " & AliasName & " audio right on", 0, 0, 0
        Else 'NOT ST_POS...
            mciSendString "set " & AliasName & " audio right off", 0, 0, 0
        End If
    Case Chan_Left
        If st_pos = vd_On Then
            mciSendString "set " & AliasName & " audio left on", 0, 0, 0
        Else 'NOT ST_POS...
            mciSendString "set " & AliasName & " audio left off", 0, 0, 0
        End If
    End Select

End Sub

'Public Function THE_ENDOFSONG(fTime As TimeFormat) As Boolean 'TESTED;OK

''TESTED:OK
'Dim curPos As String
'Dim EndPos As String
'Select Case fTime
'Case ByMS
'curPos = POSFORMAT(ByMS)
'EndPos = LENGTHFORMAT(ByMS)
'Case ByTMSF
'curPos = POSFORMAT(ByTMSF)
'EndPos = LENGTHFORMAT(ByTMSF)
'Case byFrames
'curPos = POSFORMAT(byFrames)
'EndPos = LENGTHFORMAT(byFrames)
'End Select
'THE_ENDOFSONG = curPos = EndPos
'End Function

