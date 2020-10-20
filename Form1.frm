VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "Gamepad Maus"
   ClientHeight    =   975
   ClientLeft      =   120
   ClientTop       =   705
   ClientWidth     =   3135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   3135
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture2 
      DragIcon        =   "Form1.frx":0CCA
      Height          =   315
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      DragIcon        =   "Form1.frx":1994
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFF80&
      Height          =   2535
      Left            =   10320
      TabIndex        =   3
      Top             =   0
      Width           =   2412
      Begin VB.Image img_Eras_Logo 
         Height          =   1065
         Left            =   840
         Picture         =   "Form1.frx":1C9E
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label lbl_copyright_Datum 
         BackStyle       =   0  'Transparent
         Caption         =   "08.05.2009"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lbl_copyright_Eras 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "alias Eras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbl_copyright_Tino 
         BackStyle       =   0  'Transparent
         Caption         =   "Tino Schuldt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl_copyright_by 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lbl_Programm 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Gamepad Maus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lbl_Version 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "v.1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.HScrollBar scr_speed 
      Height          =   255
      Left            =   120
      Max             =   14
      TabIndex        =   0
      Top             =   600
      Value           =   6
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lbl_bereit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kein Gamepad angeschlossen!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   30
      Width           =   2940
   End
   Begin VB.Label lbl_Speed 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:   7/15"
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1020
   End
   Begin VB.Menu mnu_datei 
      Caption         =   "&Datei"
      Begin VB.Menu mnu_einstellen 
         Caption         =   "Einstellungen"
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 1"
            Index           =   0
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 2"
            Index           =   1
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 3"
            Index           =   2
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 4"
            Index           =   3
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 5"
            Index           =   4
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 6"
            Index           =   5
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 7"
            Index           =   6
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 8"
            Index           =   7
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 9"
            Index           =   8
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 10"
            Index           =   9
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 11"
            Index           =   10
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 12"
            Index           =   11
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 13"
            Index           =   12
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 14"
            Index           =   13
         End
         Begin VB.Menu mnu_Speed 
            Caption         =   "Speed 15"
            Index           =   14
         End
      End
      Begin VB.Menu mnu_ende 
         Caption         =   "Beenden"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" _
                         Alias "Shell_NotifyIconA" ( _
                         ByVal dwMessage As Long, _
                         ByRef pnid As NOTIFYICONDATA) As Boolean
                         
Private Declare Function SetForegroundWindow Lib "user32" ( _
                         ByVal hwnd As Long) As Long
                         
Private Const NIM_ADD As Long = &H0&
Private Const NIM_MODIFY As Long = &H1&
Private Const NIM_DELETE As Long = &H2&

Private Const NIF_MESSAGE As Long = &H1&
Private Const NIF_ICON As Long = &H2&
Private Const NIF_TIP As Long = &H4&

Private Const WM_MOUSEMOVE As Long = &H200&
Private Const WM_LBUTTONDOWN As Long = &H201&
Private Const WM_LBUTTONUP As Long = &H202&
Private Const WM_LBUTTONDBLCLK As Long = &H203&
Private Const WM_RBUTTONDOWN As Long = &H204&
Private Const WM_RBUTTONUP As Long = &H205&
Private Const WM_RBUTTONDBLCLK As Long = &H206&

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private TIcon As NOTIFYICONDATA

'===============================Gamepad========================================


Private Const JOY_BUTTON1 As Long = &H1&
Private Const JOY_BUTTON2 As Long = &H2&
Private Const JOY_BUTTON3 As Long = &H4&
Private Const JOY_BUTTON4 As Long = &H8&

Private Const JOYERR_BASE As Long = 160&
Private Const JOYERR_NOERROR As Long = 0&
Private Const JOYERR_NOCANDO As Long = (JOYERR_BASE + 6&)
Private Const JOYERR_PARMS As Long = (JOYERR_BASE + 5&)
Private Const JOYERR_UNPLUGGED As Long = (JOYERR_BASE + 7&)

Private Const MAXPNAMELEN As Long = 32&

Private Const JOYSTICKID1 As Long = 0&
Private Const JOYSTICKID2 As Long = 1&

Private Type JOYINFO
    x As Long
    Y As Long
    Z As Long
    Buttons As Long
End Type

Private Type JOYCAPS
    wMid As Integer
    wPid As Integer
    szPname As String * MAXPNAMELEN
    wXmin As Long
    wXmax As Long
    wYmin As Long
    wYmax As Long
    wZmin As Long
    wZmax As Long
    wNumButtons As Long
    wPeriodMin As Long
    wPeriodMax As Long
End Type

Private Declare Function joyGetDevCaps Lib "winmm.dll" _
                         Alias "joyGetDevCapsA" ( _
                         ByVal id As Long, _
                         lpCaps As JOYCAPS, _
                         ByVal uSize As Long) As Long
                         
Private Declare Function joyGetNumDevs Lib "winmm.dll" () As Long

Private Declare Function joyGetPos Lib "winmm.dll" ( _
                         ByVal uJoyID As Long, _
                         pji As JOYINFO) As Long



'====================================Cursor======================================


'API-Funktion deklarieren
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'API-Funktion deklarieren
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long

'Nur bei Verwendung von CenterCursor nötig:
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, Rect As Rect) As Long

Private Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type


Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10

Dim Gamepad_status, i As Integer

'===

Private Type POINTAPI 'Variablentyp deklarieren
   x As Long
   Y As Long
End Type

Dim CursorPos As POINTAPI 'Variable deklarieren

'CenterCursor-Funktion
Sub CenterCursor(varObject As Control)
   Dim CtlSize As Rect
   GetWindowRect varObject.hwnd, CtlSize
   SetCursorPos CtlSize.Left + (CtlSize.Right - CtlSize.Left) / 2, CtlSize.Top + (CtlSize.Bottom - CtlSize.Top) / 2
End Sub

'CenterCursor-Funktion
Sub CenterCursor2(x, Y As Long)
   SetCursorPos x, Y
End Sub

'=============================================================================










Private Function GetJoyMax(ByVal joy As Integer, JI As JOYINFO) As Boolean
    Dim jc As JOYCAPS
    If joyGetDevCaps(joy, jc, Len(jc)) <> JOYERR_NOERROR Then
        GetJoyMax = False
    Else
        JI.x = jc.wXmax
        JI.Y = jc.wYmax
        JI.Z = jc.wZmax
        JI.Buttons = jc.wNumButtons
        GetJoyMax = True
    End If
End Function

Private Function GetJoyMin(ByVal joy As Integer, JI As JOYINFO) As Boolean
    Dim jc As JOYCAPS
    If joyGetDevCaps(joy, jc, Len(jc)) <> JOYERR_NOERROR Then
        GetJoyMin = False
    Else
        JI.x = jc.wXmin
        JI.Y = jc.wYmin
        JI.Z = jc.wZmin
        JI.Buttons = jc.wNumButtons
        GetJoyMin = True
    End If
End Function

Private Function GetJoystick(ByVal joy As Integer, JI As JOYINFO) As Boolean
    If joyGetPos(joy, JI) <> JOYERR_NOERROR Then
        GetJoystick = False
    Else
        GetJoystick = True
    End If
End Function

Private Function IsJoyPresent(Optional IsConnected As Variant) As Long
    Dim ic As Boolean
    Dim i As Long
    Dim j As Long
    Dim ret As Long
    Dim JI As JOYINFO
    ic = IIf(IsMissing(IsConnected), True, CBool(IsConnected))
    i = joyGetNumDevs
    If ic Then
        j = 0
        Do While i > 0
            i = i - 1
            If joyGetPos(i, JI) = JOYERR_NOERROR Then
                j = j + 1
            End If
        Loop
        IsJoyPresent = j
    Else
        IsJoyPresent = i
    End If
End Function







Private Sub Form_Load()
mnu_Speed(6).Checked = True

    Me.Hide
    App.TaskVisible = False
    mnu_datei.Visible = False
    
    TIcon.cbSize = Len(TIcon)
    TIcon.hwnd = Picture1.hwnd
    TIcon.uId = 1&
    TIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TIcon.ucallbackMessage = WM_MOUSEMOVE
    TIcon.hIcon = Picture1.DragIcon
    Gamepad_status = 2
    
    ' Hinzufügen des Icons in den Systemtray
    Call Shell_NotifyIcon(NIM_ADD, TIcon)
End Sub





Private Sub mnu_Speed_Click(Index As Integer)

scr_speed.Value = Index

End Sub

Private Sub Timer1_Timer()
Dim E_Buttons, E_X, E_Y, E_X_R, E_Y_R, E_X_Speed, E_Y_Speed, Max_Speed As Double
Dim i As Double
Dim E_Button(11) As Double
    Dim JInfo As JOYINFO


For i = 0 To 14
If (scr_speed.Value = i) Then
mnu_Speed(i).Checked = True
Else
mnu_Speed(i).Checked = False
End If
Next i


If ((IsJoyPresent(True)) > 0) Then

    If (Gamepad_status <> 1) Then
    For i = 0 To 14
    mnu_Speed(i).Enabled = True
    Next i
    mnu_einstellen.Enabled = True
    TIcon.hIcon = Picture2.DragIcon
    TIcon.szTip = "Gamepad bereit" & Chr$(0)
    Call Shell_NotifyIcon(NIM_MODIFY, TIcon)
    End If
Gamepad_status = 1

lbl_bereit.Caption = "Gamepad bereit!"
    If (scr_speed.Visible = False) Then
    scr_speed.Visible = True
    lbl_Speed.Visible = True
    End If

    GetJoyMax JOYSTICKID1, JInfo
    GetJoyMin JOYSTICKID1, JInfo
    GetJoystick JOYSTICKID1, JInfo

E_Buttons = Val(JInfo.Buttons)

'E_Buttons Zahl dekodieren in Buttons
If (E_Buttons >= 2048) Then
E_Button(11) = 1
E_Buttons = E_Buttons - 2048
Else
E_Button(11) = 0
End If
If (E_Buttons >= 1024) Then
E_Button(10) = 1
E_Buttons = E_Buttons - 1024
Else
E_Button(10) = 0
End If
If (E_Buttons >= 512) Then
E_Button(9) = 1
E_Buttons = E_Buttons - 512
Else
E_Button(9) = 0
End If
If (E_Buttons >= 256) Then
E_Button(8) = 1
E_Buttons = E_Buttons - 256
Else
E_Button(8) = 0
End If
If (E_Buttons >= 128) Then
E_Button(7) = 1
E_Buttons = E_Buttons - 128
Else
E_Button(7) = 0
End If
If (E_Buttons >= 64) Then
E_Button(6) = 1
E_Buttons = E_Buttons - 64
Else
E_Button(6) = 0
End If
If (E_Buttons >= 32) Then
E_Button(5) = 1
E_Buttons = E_Buttons - 32
Else
E_Button(5) = 0
End If
If (E_Buttons >= 16) Then
E_Button(4) = 1
E_Buttons = E_Buttons - 16
Else
E_Button(4) = 0
End If
If (E_Buttons >= 8) Then
E_Button(3) = 1
E_Buttons = E_Buttons - 8
Else
E_Button(3) = 0
End If
If (E_Buttons >= 4) Then
E_Button(2) = 1
E_Buttons = E_Buttons - 4
Else
E_Button(2) = 0
End If
If (E_Buttons >= 2) Then
E_Button(1) = 1
E_Buttons = E_Buttons - 2
Else
E_Button(1) = 0
End If
If (E_Buttons >= 1) Then
E_Button(0) = 1
E_Buttons = E_Buttons - 1
Else
E_Button(0) = 0
End If


'Speed und Richtung deklarieren
E_X = Val(JInfo.x)
E_Y = Val(JInfo.Y)

Max_Speed = scr_speed.Value + 1
lbl_Speed.Caption = "Speed:   " & Max_Speed & "/15"


'E_X_R    0=Links, 1=Mitte, 2=Rechts
'E_Y_R    0=Oben,  1=Mitte, 2=Unten

If (E_X >= 36511) Then 'Mitte 32511
E_X_R = 2
E_X_Speed = Max_Speed / 5
    If (E_X >= 45535) Then
    E_X_Speed = Max_Speed / 4
    End If
    If (E_X >= 55535) Then
    E_X_Speed = Max_Speed / 3
    End If
    If (E_X >= 60535) Then
    E_X_Speed = Max_Speed / 2
    End If
    If (E_X >= 65535) Then
    E_X_Speed = Max_Speed
    End If
ElseIf (E_X <= 28511) Then
E_X_R = 0
E_X_Speed = Max_Speed / 5
    If (E_X <= 15535) Then
    E_X_Speed = Max_Speed / 4
    End If
    If (E_X <= 10535) Then
    E_X_Speed = Max_Speed / 3
    End If
    If (E_X <= 5535) Then
    E_X_Speed = Max_Speed / 2
    End If
    If (E_X <= 0) Then
    E_X_Speed = Max_Speed
    End If
Else
E_X_R = 1
End If

If (E_Y >= 36511) Then 'Mitte 32511
E_Y_R = 2
E_Y_Speed = Max_Speed / 5
    If (E_Y >= 45535) Then
    E_Y_Speed = Max_Speed / 4
    End If
    If (E_Y >= 55535) Then
    E_Y_Speed = Max_Speed / 3
    End If
    If (E_Y >= 60535) Then
    E_Y_Speed = Max_Speed / 2
    End If
    If (E_Y >= 65535) Then
    E_Y_Speed = Max_Speed
    End If
ElseIf (E_Y <= 28511) Then
E_Y_R = 0
E_Y_Speed = Max_Speed / 5
    If (E_Y <= 15535) Then
    E_Y_Speed = Max_Speed / 4
    End If
    If (E_Y <= 10535) Then
    E_Y_Speed = Max_Speed / 3
    End If
    If (E_Y <= 5535) Then
    E_Y_Speed = Max_Speed / 2
    End If
    If (E_Y <= 0) Then
    E_Y_Speed = Max_Speed
    End If
Else
E_Y_R = 1
End If


'Aktuelle Cursor Pos. bestimmen
   Call GetCursorPos(CursorPos) 'API-Funktion aufrufen

'Cursor verschieben mit gewählten Speed
If (E_X_R = 0) Then
CursorPos.x = CursorPos.x - E_X_Speed
ElseIf (E_X_R = 2) Then
CursorPos.x = CursorPos.x + E_X_Speed
End If

If (E_Y_R = 0) Then
CursorPos.Y = CursorPos.Y - E_Y_Speed
ElseIf (E_Y_R = 2) Then
CursorPos.Y = CursorPos.Y + E_Y_Speed
End If
SetCursorPos CursorPos.x, CursorPos.Y


'Button(0) bis Button(11)
    'Dreieck, Kreis, X, Viereck
    'L2, R2, L1, R1
    'Select, Start, Analog_L, Analog_R
If (E_Button(0) = 1) Then
MouseClick (MiddleMouseButton)
End If
If (E_Button(1) = 1) Then
MouseClick (RightMouseButton)
End If
If (E_Button(2) = 1) Then
MouseClick (LeftMouseButton)
End If



Else

    If (Gamepad_status <> 0) Then
    For i = 0 To 14
    mnu_Speed(i).Enabled = False
    Next i
    mnu_einstellen.Enabled = False
    TIcon.hIcon = Picture1.DragIcon
    TIcon.szTip = "Gamepad nicht angeschlossen" & Chr$(0)
    Call Shell_NotifyIcon(NIM_MODIFY, TIcon)
    End If
Gamepad_status = 0

lbl_bereit.Caption = "Kein Gamepad angeschlossen!"
    If (scr_speed.Visible = True) Then
    scr_speed.Visible = False
    lbl_Speed.Visible = False
    End If
End If
End Sub











Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Form1.Hide
    
    If UnloadMode = vbAppWindows Or UnloadMode = vbFormCode Then
    
        ' Icon aus dem Systemtray entfernen
        Call Shell_NotifyIcon(NIM_DELETE, TIcon)
        
    Else
    
        Cancel = 1
        
    End If
    
End Sub

Private Sub mnu_ende_Click()

    ' Icon aus dem Systemtray entfernen
    Call Shell_NotifyIcon(NIM_DELETE, TIcon)
    
    Me.Refresh
    Unload Me
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As _
    Single, Y As Single)
    
    Dim Msg As Long
    
    Msg = x / Screen.TwipsPerPixelX
    
    Select Case Msg
    
    ' Beep
    Case WM_MOUSEMOVE:
    Case WM_LBUTTONDBLCLK: Me.Show
    Case WM_LBUTTONDOWN:
    Case WM_LBUTTONUP:
    Case WM_RBUTTONDBLCLK: Me.Show
    Case WM_RBUTTONDOWN:
    Case WM_RBUTTONUP
    
        ' Diese Funktion muss vor dem anzeigen des
        ' Menüs ausgeführt werden.
        ' weitere Informationen stehen im KB Artikel Q135788 auf
        ' http://support.microsoft.com/kb/q135788/
        Call SetForegroundWindow(Me.hwnd)
        
        ' Menü anzeigen
        Me.PopupMenu mnu_datei
        
        ' bei Verwendung von "TrackPopupMenu" muss noch
        ' die Funktion "PostMessage Me.hwnd, WM_USER, 0&, 0&"
        ' ausgeführt werden
        
    End Select
    
End Sub



