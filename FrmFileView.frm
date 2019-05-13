VERSION 5.00
Begin VB.Form FrmFileView 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9180
   Icon            =   "FrmFileView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   9180
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      Text            =   "0.5"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Atributos:"
      Height          =   1335
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valores:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resultados:"
      Height          =   4095
      Left            =   4200
      TabIndex        =   9
      Top             =   2520
      Width           =   4815
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Determinar relaciones"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5760
      TabIndex        =   3
      Text            =   "0.5"
      Top             =   1560
      Width           =   615
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6150
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Confianza:"
      Height          =   195
      Left            =   4320
      TabIndex        =   13
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Soporte mínimo:"
      Height          =   195
      Left            =   4320
      TabIndex        =   8
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Datos del archivo:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "FrmFileView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents DATAbase As Data
Attribute DATAbase.VB_VarHelpID = -1
Private HasLoaded As Boolean

Private Sub Combo1_Click()
    Dim i As Integer
    Combo2.Clear
    With DATAbase.Atributos(Combo1.ListIndex + 1)
        Combo2.AddItem "Cantidad de datos encontrados para el atributo '" & DATAbase.Atributos(Combo1.ListIndex + 1).Nombre & "'): " & .Count
        For i = 1 To .Count
            Combo2.AddItem .Valor(i)
        Next i
    End With
    Combo2.ListIndex = 0
End Sub

Private Sub Command1_Click()
    Dim Elem As Control
    Text2.Text = vbNullString
    For Each Elem In Me
        Elem.Enabled = False
    Next
    Command1.Caption = "Calculando..."
    DATAbase.Apriori Val(Text1.Text), Val(Text3.Text)
    For Each Elem In Me
        Elem.Enabled = True
    Next
    Text2.Text = DATAbase.AprioriResultados
    Command1.Caption = "Determinar relaciones"
End Sub

Private Sub DATAbase_OnApriori(DataProcesado As Double, Status As StatusProceso)
    DoEvents
    If Status = Calculando_Combinatoria Then
        Text2.Text = "Calculando las combinatorias... "
    ElseIf Status = Calculando_Relaciones Then
        Text2.Text = "Calculando las relaciones validas... "
    ElseIf Status = Optimizando Then
        Text2.Text = "Optimizando... "
    End If
    Text2.Text = Text2.Text & Round(DataProcesado * 100, 0) & "%"
End Sub

Private Sub Form_activate()
    Dim i, j As Long
    Dim Str As String
    Dim Sizes() As Long
    If Not HasLoaded Then
        Set DATAbase = New Data
        With DATAbase
            DoEvents
            .LoadFromCSVFile Me.Caption ' Carga los datos de un archivo CSV...
            ReDim Sizes(.Atributos.Count - 1) As Long
            For i = 1 To .Atributos.Count ' Carga los atributos (columnas) con sus posibles valores
                Combo1.AddItem .Atributos(i).Nombre
                Str = Str & (IIf(i <> 1, vbTab, vbNullString)) & .Atributos(i).Nombre
            Next i
            List1.AddItem Str
            Label1.Caption = "Datos del archivo (" & .Filas & " datos, " & .Atributos.Count & " atributos)"
            Combo1.ListIndex = 0
            For i = 1 To .Filas
                Str = vbNullString
                For j = 1 To .Atributos.Count
                    Str = Str & (IIf(j <> 1, vbTab, vbNullString)) & .GetDato((i), (j))
                Next j
                List1.AddItem Str
            Next i
        End With
    End If
    HasLoaded = True
    Erase Sizes
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width / 2 - List1.Left
    List1.Height = Me.Height - 2.5 * List1.Top
    Label2.Left = 2 * List1.Left + List1.Width
    Text1.Left = Label2.Left + Label2.Width + 100
    Label5.Left = Label2.Left
    Text3.Left = Text1.Left
    Command1.Left = Text1.Left + Text1.Width + 100
    Frame1.Left = Label2.Left
    Frame1.Width = List1.Width - 400
    Frame1.Height = Me.Height - (Command1.Top + Command1.Height + 800)
    Text2.Width = Frame1.Width - 300
    Text2.Height = Frame1.Height - 400
    Frame2.Left = Frame1.Left
    Frame2.Width = Frame1.Width
    Combo1.Width = Frame2.Width - (Label3.Width + 500)
    Combo2.Width = Frame2.Width - (Label4.Width + 500)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GUIs.Remove Me.Caption
    Set DATAbase = Nothing
End Sub

Private Sub Text1_Change()
    Dim NumVal As String
    If IsNumeric(Text1.Text) Then
        NumVal = (Text1.Text)
    Else
        Text1.Text = CStr(NumVal)
    End If
End Sub

Private Sub Text3_Change()
    Dim NumVal As String
    If IsNumeric(Text3.Text) Then
        NumVal = (Text3.Text)
    Else
        Text3.Text = CStr(NumVal)
    End If
End Sub
