VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Atributo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Nombre As String
Attribute Nombre.VB_VarUserMemId = 0
Private mvarValores() As String
Private mvarEmpty As Boolean
Public mvarIndex As Long


Public Function TieneValor(Atr As String) As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TieneAttrb
    TieneValor = False
    For i = 0 To UBound(mvarValores)
        TieneValor = TieneValor Or (mvarValores(i) = Atr)
        If TieneValor Then Exit For
    Next i
End Function

Public Property Get Count() As Long
    If Not mvarEmpty Then Count = UBound(mvarValores) + 1 Else Count = 0
End Property

Public Property Get Index() As Long
    Index = mvarIndex
End Property

Public Sub AdValores(Valores As String)
    If Valores = vbNullString Then Exit Sub
    Dim c As Integer
    c = UBound(mvarValores)
    If Not mvarEmpty Then ReDim Preserve mvarValores(Me.Count) As String
    If mvarEmpty Then mvarValores(0) = Valores Else mvarValores(Me.Count - 1) = Valores
    mvarEmpty = False
End Sub

Public Function Valor(Index As Integer) As String
    If 1 <= Index <= Me.Count Then Valor = mvarValores(Index - 1) Else Valor = "???"
End Function

Private Sub Class_Initialize()
    ReDim mvarValores(0) As String
    mvarEmpty = True
End Sub

Private Sub Class_Terminate()
    Erase mvarValores
End Sub

