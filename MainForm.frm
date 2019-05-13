VERSION 5.00
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "Análisis A priori"
   ClientHeight    =   5850
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8370
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu MnuImport 
         Caption         =   "&Importar Datos"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "A&cerca de..."
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Terminate()
    For Each GUI In GUIs
        If IsObject(GUI) Then Unload (GUI)
    Next
End Sub

Private Sub MnuAbout_Click()
    Load FrmAbout
    FrmAbout.Show , Me
End Sub

Private Sub MnuImport_Click()
    OpenFileDlg Me
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub
