Option Explicit

Dim oFont As New clsFontInstall

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
    Set oFont = Nothing
End Sub

Private Sub Workbook_Open()
    oFont.FontName = "My Font Name"
    oFont.FontFileName = ThisWorkbook.Path & "\MyFontFile.ttf"
    If oFont.InstallFonts = False Then
        MsgBox "Could not install the font(s): " & oFont.FontName
    End If
End Sub