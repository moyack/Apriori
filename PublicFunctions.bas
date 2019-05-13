Attribute VB_Name = "PublicFunctions"
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public GUIs As New Collection

Public Function OpenFileDlg(ByVal Frm As Form)
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Fila As Integer
    With OpenFile
        .lStructSize = Len(OpenFile)
        .hwndOwner = Frm.hWnd
        .hInstance = App.hInstance
        .lpstrFilter = "Archivo de texto separado por comas (*.csv)" & Chr(0) & "*.CSV" & Chr(0)
        .nFilterIndex = 1
        .lpstrFile = String(257, 0)
        .nMaxFile = Len(OpenFile.lpstrFile) - 1
        .lpstrFileTitle = OpenFile.lpstrFile
        .nMaxFileTitle = OpenFile.nMaxFile
        .lpstrInitialDir = App.Path
        .lpstrTitle = "Abrir archivo de tabla de datos..."
        .flags = 0
    End With
    lReturn = GetOpenFileName(OpenFile)
    If lReturn <> 0 Then
        On Error Resume Next
        GUIs.Add New FrmFileView, OpenFile.lpstrFile
        If Err.Number = 457 Then
            MsgBox "El archivo ya está cargado", vbInformation, Frm.Caption
            Exit Function ' detecta si el archivo ya está abierto...
        End If
        With GUIs(GUIs.Count)
            .Caption = OpenFile.lpstrFile
            .Show
        End With
    End If
End Function

