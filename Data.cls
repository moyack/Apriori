VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Atributos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarAtributos As Atributos
'local variable(s) to hold property value(s)
Private mvarFilas As Long 'local copy
Private mvarDATA() As String ' Guarda los datos para analizar
Private mvarApriori() As String ' Guarda los resultados del algoritmo apriori
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent OnApriori[(arg1, arg2, ... , argn)]
Public Event OnApriori(DataProcesado As Double, Status As StatusProceso)

Public Enum StatusProceso
    Optimizando = 0
    Calculando_Combinatoria = 1
    Calculando_Relaciones = 2
End Enum

Public Sub Apriori(minSupp As Double, minConf As Double)
    Dim S As Double, c As Double
    Dim Valid() As String
    Dim List() As String
    With Me
        ' Primero determinamos el soporte de los atributos en forma independiente
        ' para asi hacer un primer descarte de posibilidades (optimización de algoritmo)
        ReDim Valid(0) As String
        ReDim mvarApriori(0) As String
        mvarApriori(0) = "Clasificación inicial:" & vbNewLine & "======================"
        For i = 1 To .Atributos.Count
            RaiseEvent OnApriori(i / .Atributos.Count, Optimizando)
            For j = 1 To .Atributos(i).Count
                S = Supp(.Atributos(i).Nombre & "|" & .Atributos(i).Valor(CInt(j)))
                If S > minSupp Then
                    ReDim Preserve Valid(UBound(Valid) + 1) As String
                    Valid(UBound(Valid)) = .Atributos(i).Nombre & "|" & .Atributos(i).Valor((j))
                    ReDim Preserve mvarApriori(UBound(mvarApriori) + 1) As String
                    mvarApriori(UBound(mvarApriori)) = "Atributo: " & .Atributos(i).Nombre & "  Valor: " & .Atributos(i).Valor((j)) & " (Soporte: " & Round(S, 2) & ")"
                End If
            Next j
        Next i
        ' Ahora se trabaja en la parte de hacer la combinatoria...
        ReDim List(0) As String
        For i = 1 To UBound(Valid)
            ADValor Valid(i), i, Valid, List ' función recursiva que determina las combinatorias de los valores...
        Next i
        
        ' Con las combinaciones, se procede a calcular los soportes y las confianzas para así saber cual es la
        ' relación que más nos sirve...
        ReDim Preserve mvarApriori(UBound(mvarApriori) + 1) As String
        mvarApriori(UBound(mvarApriori)) = vbNewLine & "Relaciones:" & vbNewLine & "==========="
        Dim Valor() As String, Contar As Long
        Contar = 0
        For i = 1 To UBound(List)
            RaiseEvent OnApriori(i / UBound(List), Calculando_Relaciones)
            S = Supp(List(i)) ' Calcula el soporte...
            If S > minSupp Then ' Si tiene un soporte mayor al soporte mínimo...
                ' Calcula la confiabilidad respecto a la mínima...
                Dim V As String
                c = Conf(List(i), V)
                If c > minConf Then
                    List(i) = Replace(V, "|", " = ")
                    List(i) = Replace(List(i), "¬", " ")
                    List(i) = List(i) & " [Soporte: " & Round(S, 2) & "] [Confiabilidad: " & Round(c, 2) & "]"
                    ReDim Preserve mvarApriori(UBound(mvarApriori) + 1) As String
                    mvarApriori(UBound(mvarApriori)) = "(" & (Contar + 1) & "). " & List(i)
                    Contar = Contar + 1
                End If
            End If
        Next i
        Erase Valid
        Erase List
        Erase Valor
    End With
End Sub

Private Sub ADValor(ByVal Str As String, ByVal Pos As Integer, ByRef Val() As String, ByRef Lista() As String)
    RaiseEvent OnApriori(Pos / UBound(Val), Calculando_Combinatoria)
    If Pos <= UBound(Val) Then
        For i = Pos + 1 To UBound(Val)
            ReDim Preserve Lista(UBound(Lista) + 1) As String
            Lista(UBound(Lista)) = Str & "¬" & Val(i)
            ADValor Lista(UBound(Lista)), i, Val, Lista
        Next i
    End If
End Sub

Private Function Conf(ByVal Valores As String, ByRef Valid As String) As Double
    Dim Tmp() As String
    Dim Val() As String
    Dim c() As Double
    ReDim Val(0) As String
    Tmp = Split(Valores, "¬")
    If UBound(Tmp) < 2 Then
        Valid = Replace(Valores, "¬", " ==> ")
        Conf = 1#
    Else
        For i = LBound(Tmp) To UBound(Tmp) - 1
            ADValor Tmp(i), i, Tmp, Val
        Next i
        ReDim c(UBound(Val)) As Double
        Conf = 0
        For i = 1 To UBound(Val)
            If Valores <> Val(i) Then
                c(i) = Supp(Valores) / Supp(Val(i))
                If c(i) > c(CLng(Conf)) Then Conf = i
            End If
        Next i
        Valid = Val(CLng(Conf))
        Val = Split(Valid, "¬")
        Valid = Valid & " ==> "
        Dim Test As Boolean
        Test = False
        For i = LBound(Tmp) To UBound(Tmp)
            For j = LBound(Val) To UBound(Val)
                Test = Test Or (Tmp(i) = Val(j))
            Next j
            If Not Test Then Valid = Valid & "¬" & Tmp(i)
        Next i
        Conf = c(CLng(Conf))
    End If
    Erase Tmp
    Erase Val
    Erase c
End Function

Private Function Supp(ByVal Valores As String) As Double
    Dim Tmp() As String
    Dim Val() As String
    Dim Cont As Double
    Dim Test As Boolean
    Tmp = Split(Valores, "¬")
    With Me
        Cont = 0
        For i = 1 To .Filas
            Test = True
            For j = LBound(Tmp) To UBound(Tmp)
                Val = Split(Tmp(j), "|")
                Test = Test And (.GetDato(CLng(i), Val(LBound(Val))) = Val(UBound(Val)))
            Next j
            If Test Then Cont = Cont + 1
        Next i
        Supp = Cont / CDbl(.Filas)
    End With
    Erase Tmp
    Erase Val
End Function

Public Property Get AprioriResultados() As String
    If UBound(mvarApriori) > 0 Then
        For i = 0 To UBound(mvarApriori)
            AprioriResultados = AprioriResultados & (IIf(i <> 0, vbNewLine, vbNullString)) & mvarApriori(i)
        Next i
    Else
        AprioriResultados = vbNullString
    End If
End Property

Public Property Get Filas() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Filas
    Filas = mvarFilas
End Property

Public Function GetDato(ByRef Fila As Long, Attr As Variant) As String
    Dim Dato() As String
    Dim Atr As Integer
    If IsNumeric(Attr) Then Atr = CInt(Attr) Else Atr = Me.Atributos(Attr).Index
    If 1 <= Fila <= UBound(mvarDATA) Then
        Dato = Split(mvarDATA(Fila), ",")
        GetDato = Dato(Me.Atributos(Attr).Index - 1)
    Else
        GetDato = "???"
    End If
    Erase Dato
End Function

Public Sub AdDato(Valores As String)
    ReDim Preserve mvarDATA(UBound(mvarDATA) + 1) As String
    mvarDATA(UBound(mvarDATA)) = Valores
    mvarFilas = UBound(mvarDATA)
End Sub

Public Sub LoadFromCSVFile(Filename As String)
    Dim FS As New FileSystemObject
    Dim TX As TextStream
    Dim HEADERS() As String
    Dim Fila As String
    Set TX = FS.OpenTextFile(Filename, ForReading)
    HEADERS = Split(TX.ReadLine, ",")
    For i = 0 To UBound(HEADERS)
        Me.Atributos.Add HEADERS(i)
    Next i
    Do While Not TX.AtEndOfLine
        With Me
            Fila = TX.ReadLine
            .AdDato Fila
            HEADERS = Split(Fila, ",")
            For i = 1 To .Atributos.Count
                If Not .Atributos(i).TieneValor(HEADERS(i - 1)) Then .Atributos(i).AdValores HEADERS(i - 1)
            Next i
        End With
    Loop
    Set FS = Nothing
    Set TX = Nothing
    Erase HEADERS
End Sub

Public Property Get Atributos() As Atributos
    If mvarAtributos Is Nothing Then
        Set mvarAtributos = New Atributos
    End If
    Set Atributos = mvarAtributos
End Property

Public Property Set Atributos(vData As Atributos)
    Set mvarAtributos = vData
End Property

Private Sub Class_Initialize()
    ReDim mvarDATA(0) As String
    ReDim mvarApriori(0) As String
    TieneResult = False
End Sub

Private Sub Class_Terminate()
    Set mvarAtributos = Nothing
    Erase mvarDATA
End Sub

