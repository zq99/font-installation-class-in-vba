Option Explicit

Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long

Private Declare Function SendMessageTimeoutA Lib "user32.dll" ( _
        ByVal hWnd As Long, _
        ByVal Msg As Long, _
        ByVal wParam As Long, _
        ByRef lParam As Any, _
        ByVal fuFlags As Long, _
        ByVal uTimeout As Long, _
        ByRef lpdwResult As Long) As Long

Private Const WM_FONTCHANGE     As Long = &H1D
Private Const HWND_BROADCAST    As Long = &HFFFF&
Private Const SMTO_NORMAL       As Long = &H0
Private Const CINT_MATCH        As Integer = 0

Private mstrFontName            As String
Private mstrFontFileName        As String
Private mlngMilliSeconds        As Long

Public Property Get FontName() As String
    FontName = mstrFontName
End Property

Public Property Let FontName(ByVal Value As String)
    mstrFontName = Value
End Property

Public Property Let FontFileName(ByVal Value As String)
     mstrFontFileName = Value
End Property

Public Property Let NotifyWindowsTimeOut(ByVal Value As Integer)
     mlngMilliSeconds = Value
End Property

Public Function UninstallFonts() As Boolean
On Error GoTo ERROR_HANDLER
    UninstallFonts = False
    If RemoveFontResource(mstrFontFileName) Then
        Dim lngReturn As Long
        Dim lResult   As Long

        lngReturn = SendMessageTimeoutA( _
            HWND_BROADCAST, _
            WM_FONTCHANGE, _
            0, _
            ByVal "windows", _
            SMTO_NORMAL, _
            1000, _
            lResult)
        UninstallFonts = True
    Else
        UninstallFonts = False
    End If
EXIT_HERE:
    Exit Function
ERROR_HANDLER:
    UninstallFonts = False
    GoTo EXIT_HERE
End Function


Public Function InstallFonts() As Boolean
On Error GoTo ERROR_HANDLER
    Dim lngReturn   As Long
    Dim lngResult   As Long
    InstallFonts = True

    If IsFontInstalled = False Then
        If AddFontResource(mstrFontFileName) Then
            lngReturn = SendMessageTimeoutA( _
                HWND_BROADCAST, _
                WM_FONTCHANGE, _
                0, _
                ByVal "windows", _
                SMTO_NORMAL, _
                mlngMilliSeconds, _
                lngResult)
        Else
            InstallFonts = False
        End If
    End If
EXIT_HERE:
    Exit Function
ERROR_HANDLER:
    InstallFonts = False
    GoTo EXIT_HERE
End Function

Public Function IsFontInstalled() As Boolean
On Error GoTo ERR_HANDLER:
    Dim objFont As New StdFont
    IsFontInstalled = False
    
    objFont.Name = mstrFontName
    If StrComp(mstrFontName, objFont.Name, vbTextCompare) = CINT_MATCH Then
        IsFontInstalled = True
    Else
        IsFontInstalled = False
    End If
EXIT_HERE:
    Set objFont = Nothing
    Exit Function
ERR_HANDLER:
    IsFontInstalled = False
    GoTo EXIT_HERE
End Function

Private Sub Class_Initialize()
    mlngMilliSeconds = 1000
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If IsFontInstalled = True Then
        Me.UninstallFonts
    End If
End Sub