Attribute VB_Name = "Cuadre"
'==============================================================================
' Cuadre ST <-> FINREP/COREP : carga de ficheros en la hoja Datos
'
' - Fichero solo ASCII a proposito: el importador de VBA lee los .bas en ANSI.
' - Las formulas de las columnas calculadas de Datos NO estan aqui: se leen de la
'   hoja oculta _Formulas, que escribe el generador Python (fuente unica).
' - Deteccion de codigos igual que el loader Python excel_annex:
'   codigos de fila/columna como TEXTO ("0010", "r0010", "c0010"); se elige el par
'   (columna de codigos de fila, fila de cabecera) que explica mas codigos.
' - Tambien lee el layout de las plantillas oficiales ST de la EBA (columna "RowNum"
'   y fila de numeros de columna), igual que el loader Python st_eba.
' - Por defecto solo se cargan los datos que usa alguna regla (Parametros, P_SoloUsados).
' - Requiere Windows (Scripting.Dictionary).
'==============================================================================
Option Explicit

Private Const HOJA_DATOS As String = "Datos"
Private Const HOJA_CARGA As String = "Carga"
Private Const HOJA_MAPEO As String = "Mapeo_hojas"
Private Const HOJA_PARAM As String = "Parametros"
Private Const HOJA_FORM As String = "_Formulas"
Private Const NCOLS_DATOS As Long = 17
Private Const MAX_FILAS_DATOS As Long = 200000

Private gUsados As Object   ' claves usadas por alguna regla; Nothing = cargar todo

'------------------------------------------------------------------ botones
Public Sub CargarST()
    CargarFicheros "ST"
End Sub

Public Sub CargarFINREP()
    CargarFicheros "FINREP"
End Sub

Public Sub CargarCOREP()
    CargarFicheros "COREP"
End Sub

Public Sub CrearBotones()
    Dim ws As Worksheet, b As Object, i As Long
    Dim macros As Variant, textos As Variant
    Set ws = ThisWorkbook.Worksheets("Inicio")
    For Each b In ws.Buttons
        b.Delete
    Next b
    macros = Array("CargarST", "CargarFINREP", "CargarCOREP", "Recalcular", "LimpiarDatos")
    textos = Array("Cargar ST", "Cargar FINREP", "Cargar COREP", "Recalcular", "Borrar datos cargados")
    For i = LBound(macros) To UBound(macros)
        Set b = ws.Buttons.Add(ws.Range("A15").Left + i * 150, ws.Range("A15").Top, 140, 30)
        b.OnAction = macros(i)
        b.Caption = textos(i)
    Next i
End Sub

Public Sub Recalcular()
    Application.CalculateFull
    ThisWorkbook.Worksheets("Resumen").Activate
End Sub

Public Sub LimpiarDatos()
    If MsgBox("Se borraran todos los datos cargados y el registro de carga. Continuar?", _
              vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    VaciarFilas ThisWorkbook.Worksheets(HOJA_DATOS), NCOLS_DATOS
    VaciarFilas ThisWorkbook.Worksheets(HOJA_CARGA), 10
    Application.CalculateFull
End Sub

'------------------------------------------------------------------ carga
Public Sub CargarFicheros(ByVal marco As String)
    Dim sel As Variant, i As Long, resp As Variant
    Dim escala As Double, decimales As Variant
    Dim calcPrev As Long

    sel = Application.GetOpenFilename("Libros Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , _
                                      "Selecciona ficheros " & marco, , True)
    If Not IsArray(sel) Then Exit Sub

    escala = ParametroMarco(marco, 2, 1)
    resp = Application.InputBox("Escala a EUR de los ficheros " & marco & vbCrLf & _
                                "(1 = EUR, 1000 = miles, 1000000 = millones)", "Escala", escala, Type:=1)
    If VarType(resp) = vbBoolean Then Exit Sub
    escala = CDbl(resp)

    decimales = ParametroMarco(marco, 3, "")
    resp = Application.InputBox("Decimales con los que se reportan los importes" & vbCrLf & _
                                "(vacio = valores exactos, sin redondeo admitido)", "Decimales", CStr(decimales), Type:=2)
    If VarType(resp) = vbBoolean Then Exit Sub
    decimales = Trim$(CStr(resp))
    If decimales <> "" Then
        If Not IsNumeric(decimales) Then
            MsgBox "Decimales no numerico: " & decimales, vbCritical
            Exit Sub
        End If
        decimales = CLng(decimales)
    End If

    If ContarMarco(marco) > 0 Then
        Select Case MsgBox("Ya hay datos " & marco & " cargados. Borrarlos antes de cargar?" & vbCrLf & _
                           "(No = anadir a los existentes)", vbYesNoCancel + vbQuestion)
            Case vbYes: BorrarMarco marco
            Case vbCancel: Exit Sub
        End Select
    End If

    calcPrev = Application.Calculation
    Application.ScreenUpdating = False
    Set gUsados = ClavesUsadas()
    Application.Calculation = xlCalculationManual
    On Error GoTo fallo
    For i = LBound(sel) To UBound(sel)
        Application.StatusBar = "Cargando " & sel(i)
        CargarLibro CStr(sel(i)), marco, escala, decimales
    Next i
    On Error GoTo 0
    Application.StatusBar = False
    Application.Calculation = calcPrev
    Application.ScreenUpdating = True
    Application.CalculateFull
    ThisWorkbook.Worksheets("Resumen").Activate
    MsgBox "Carga terminada. Revisa avisos en la hoja Carga antes de interpretar los checks.", vbInformation
    Exit Sub
fallo:
    Application.StatusBar = False
    Application.Calculation = calcPrev
    Application.ScreenUpdating = True
    MsgBox "Error cargando: " & Err.Description, vbCritical
End Sub

Private Sub CargarLibro(ByVal ruta As String, ByVal marco As String, ByVal escala As Double, ByVal decimales As Variant)
    Dim wb As Workbook, ws As Worksheet, errNum As Long, errDesc As String
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    On Error GoTo cerrar
    For Each ws In wb.Worksheets
        CargarHoja ws, ruta, marco, escala, decimales
    Next ws
cerrar:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, , ruta & ": " & errDesc
End Sub

Private Sub CargarHoja(ByVal ws As Worksheet, ByVal ruta As String, ByVal marco As String, _
                       ByVal escala As Double, ByVal decimales As Variant)
    Dim plantilla As String, z As String, omitir As Boolean
    Dim v As Variant, r0 As Long, c0 As Long, nR As Long, nC As Long
    Dim i As Long, j As Long, k As Long, nCod As Long
    Dim codR() As Long, codC() As Long, codT() As String
    Dim avisos As String

    plantilla = ws.Name
    z = ""
    BuscarMapeo marco, ws.Name, plantilla, z, omitir
    If omitir Then Exit Sub
    If CargarHojaST(ws, ruta, marco, escala, decimales, plantilla, z) Then Exit Sub

    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
        Registrar marco, ruta, ws.Name, plantilla, z, 0, 0, 0, "hoja vacia"
        Exit Sub
    End If
    r0 = ws.UsedRange.Row
    c0 = ws.UsedRange.Column
    If ws.UsedRange.Cells.Count = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = ws.UsedRange.Value
    Else
        v = ws.UsedRange.Value
    End If
    nR = UBound(v, 1)
    nC = UBound(v, 2)

    ' 1) celdas de texto con forma de codigo
    nCod = 0
    For i = 1 To nR
        For j = 1 To nC
            If EsCodigo(v(i, j)) Then nCod = nCod + 1
        Next j
    Next i
    If nCod = 0 Then
        Registrar marco, ruta, ws.Name, plantilla, z, 0, 0, 0, "sin codigos de fila/columna detectados"
        Exit Sub
    End If
    ReDim codR(1 To nCod): ReDim codC(1 To nCod): ReDim codT(1 To nCod)
    k = 0
    For i = 1 To nR
        For j = 1 To nC
            If EsCodigo(v(i, j)) Then
                k = k + 1
                codR(k) = i: codC(k) = j: codT(k) = Trim$(CStr(v(i, j)))
            End If
        Next j
    Next i

    ' 2) codigos de fila (por fila del array) y de columna (por columna del array)
    Dim filaCod() As String, colCod() As String, nFil As Long, nCol As Long, nUsados As Long
    ReDim filaCod(1 To nR): ReDim colCod(1 To nC)
    Dim hayR As Boolean, hayC As Boolean
    For k = 1 To nCod
        If UCase$(Left$(codT(k), 1)) = "R" Then hayR = True
        If UCase$(Left$(codT(k), 1)) = "C" Then hayC = True
    Next k

    If hayR And hayC Then
        For k = 1 To nCod
            Select Case UCase$(Left$(codT(k), 1))
                Case "R": filaCod(codR(k)) = NormCodigo(codT(k)): nUsados = nUsados + 1
                Case "C": colCod(codC(k)) = NormCodigo(codT(k)): nUsados = nUsados + 1
            End Select
        Next k
    Else
        Dim bestCol As Long, bestRow As Long, bestScore As Long, score As Long
        Dim cc As Long, rr As Long, a As Long, b As Long
        bestScore = -1
        For a = 1 To nCod
            cc = codC(a)
            For b = 1 To nCod
                rr = codR(b)
                score = 0
                For k = 1 To nCod
                    If (codC(k) = cc And codR(k) > rr) Or (codR(k) = rr And codC(k) > cc) Then score = score + 1
                Next k
                If score > bestScore Or (score = bestScore And (rr < bestRow Or (rr = bestRow And cc < bestCol))) Then
                    bestScore = score: bestCol = cc: bestRow = rr
                End If
            Next b
        Next a
        If bestScore <= 0 Then
            Registrar marco, ruta, ws.Name, plantilla, z, 0, 0, 0, "no se encontro cabecera de columnas con codigos de fila debajo"
            Exit Sub
        End If
        For k = 1 To nCod
            If codC(k) = bestCol And codR(k) > bestRow Then
                filaCod(codR(k)) = NormCodigo(codT(k)): nUsados = nUsados + 1
            ElseIf codR(k) = bestRow And codC(k) > bestCol Then
                colCod(codC(k)) = NormCodigo(codT(k)): nUsados = nUsados + 1
            End If
        Next k
    End If
    If nCod > nUsados Then avisos = avisos & (nCod - nUsados) & " celdas con forma de codigo fuera de cabecera/columna de filas (ignoradas); "

    Dim dupR As Object, dupC As Object
    Set dupR = CreateObject("Scripting.Dictionary")
    Set dupC = CreateObject("Scripting.Dictionary")
    For i = 1 To nR
        If filaCod(i) <> "" Then
            nFil = nFil + 1
            If dupR.Exists(filaCod(i)) Then avisos = avisos & "codigo de fila repetido " & filaCod(i) & "; " Else dupR.Add filaCod(i), i
        End If
    Next i
    For j = 1 To nC
        If colCod(j) <> "" Then
            nCol = nCol + 1
            If dupC.Exists(colCod(j)) Then avisos = avisos & "codigo de columna repetido " & colCod(j) & "; " Else dupC.Add colCod(j), j
        End If
    Next j
    If nFil = 0 Or nCol = 0 Then
        Registrar marco, ruta, ws.Name, plantilla, z, nFil, nCol, 0, "falta la cabecera de columnas o la columna de codigos de fila"
        Exit Sub
    End If

    ' 3) valores en las intersecciones
    Dim salida() As Variant, n As Long, x As Variant, nNoNum As Long, nErr As Long, nOmit As Long
    ReDim salida(1 To nFil * nCol, 1 To NCOLS_DATOS)
    For i = 1 To nR
        If filaCod(i) <> "" Then
            For j = 1 To nC
                If colCod(j) <> "" Then
                    x = v(i, j)
                    If IsError(x) Then
                        nErr = nErr + 1
                    ElseIf IsEmpty(x) Then
                        ' celda vacia: no es un dato
                    ElseIf EsNumero(x) And Not Usado(Clave(marco, plantilla, filaCod(i), colCod(j), z)) Then
                        nOmit = nOmit + 1
                    ElseIf EsNumero(x) Then
                        n = n + 1
                        salida(n, 2) = marco
                        salida(n, 3) = plantilla
                        salida(n, 4) = filaCod(i)
                        salida(n, 5) = colCod(j)
                        salida(n, 6) = z
                        salida(n, 7) = CDbl(x)
                        salida(n, 8) = escala
                        salida(n, 10) = decimales
                        salida(n, 12) = ruta
                        salida(n, 13) = ws.Name
                        salida(n, 14) = ws.Cells(r0 + i - 1, c0 + j - 1).Address(False, False)
                    ElseIf Len(Trim$(CStr(x))) > 0 Then
                        nNoNum = nNoNum + 1
                    End If
                End If
            Next j
        End If
    Next i
    If nNoNum > 0 Then avisos = avisos & nNoNum & " valores no numericos (ignorados); "
    If nErr > 0 Then avisos = avisos & nErr & " celdas con error (#N/A, #REF!...) (ignoradas); "

    If n > 0 Then EscribirDatos salida, n
    Registrar marco, ruta, ws.Name, plantilla, z, nFil, nCol, n, avisos, InfoOmitidos(nOmit)
End Sub

Private Function CargarHojaST(ByVal ws As Worksheet, ByVal ruta As String, ByVal marco As String, _
                              ByVal escala As Double, ByVal decimales As Variant, _
                              ByVal plantilla As String, ByVal z As String) As Boolean
    ' Layout de las plantillas oficiales ST: cabecera "RowNum" y, encima, fila con los
    ' numeros de columna 1, 2, 3... Devuelve False si la hoja no tiene ese layout.
    Dim hdr As Long, rc As Long, r As Long, c As Long, r0 As Long, c0 As Long
    Dim lastR As Long, lastC As Long, v As Variant, i As Long
    Dim mejorFila As Long, mejorCnt As Long, cnt As Long
    Dim colNum() As Long, nCols As Long, nFilas As Long, rn As Variant
    Dim salida() As Variant, n As Long, nNoNum As Long, nErr As Long, nOmit As Long, avisos As String
    Dim x As Variant, fila As String, col As String

    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then Exit Function
    r0 = ws.UsedRange.Row
    c0 = ws.UsedRange.Column
    lastR = r0 + ws.UsedRange.Rows.Count - 1
    lastC = c0 + ws.UsedRange.Columns.Count - 1
    For r = 1 To 25
        For c = 1 To IIf(lastC < 60, lastC, 60)
            If VarType(ws.Cells(r, c).Value) = vbString Then
                If Trim$(ws.Cells(r, c).Value) = "RowNum" Then hdr = r: rc = c: Exit For
            End If
        Next c
        If hdr > 0 Then Exit For
    Next r
    If hdr = 0 Then Exit Function
    CargarHojaST = True

    ' fila con mas numeros de columna a la derecha de RowNum
    For r = 1 To hdr - 1
        cnt = 0
        For c = rc + 1 To lastC
            If EnteroColumna(ws.Cells(r, c).Value) > 0 Then cnt = cnt + 1
        Next c
        If cnt > mejorCnt Then mejorCnt = cnt: mejorFila = r
    Next r
    If mejorCnt = 0 Then
        Registrar marco, ruta, ws.Name, plantilla, z, 0, 0, 0, "layout RowNum sin fila de numeros de columna", ""
        Exit Function
    End If
    ReDim colNum(1 To lastC)
    For c = rc + 1 To lastC
        colNum(c) = EnteroColumna(ws.Cells(mejorFila, c).Value)
        If colNum(c) > 0 Then nCols = nCols + 1
    Next c
    If lastR <= hdr Then
        Registrar marco, ruta, ws.Name, plantilla, z, 0, nCols, 0, "sin filas de datos", ""
        Exit Function
    End If

    v = ws.Range(ws.Cells(hdr + 1, 1), ws.Cells(lastR, lastC)).Value
    ' primera pasada: contar datos numericos para dimensionar la salida
    For i = 1 To UBound(v, 1)
        rn = v(i, rc)
        If EnteroFila(rn) > 0 Then
            nFilas = nFilas + 1
            For c = rc + 1 To lastC
                If colNum(c) > 0 Then
                    If EsNumero(v(i, c)) And Not IsError(v(i, c)) Then n = n + 1
                End If
            Next c
        End If
    Next i
    If n = 0 Then
        Registrar marco, ruta, ws.Name, plantilla, z, nFilas, nCols, 0, "", "sin datos numericos"
        Exit Function
    End If
    ReDim salida(1 To n, 1 To NCOLS_DATOS)
    n = 0
    For i = 1 To UBound(v, 1)
        rn = v(i, rc)
        If EnteroFila(rn) > 0 Then
            fila = Format$(EnteroFila(rn), "0000")
            For c = rc + 1 To lastC
                If colNum(c) > 0 Then
                    x = v(i, c)
                    col = Format$(colNum(c), "0000")
                    If IsError(x) Then
                        nErr = nErr + 1
                    ElseIf IsEmpty(x) Then
                        ' vacia
                    ElseIf Not EsNumero(x) Then
                        If Len(Trim$(CStr(x))) > 0 Then nNoNum = nNoNum + 1
                    ElseIf Not Usado(Clave(marco, plantilla, fila, col, z)) Then
                        nOmit = nOmit + 1
                    Else
                        n = n + 1
                        salida(n, 2) = marco
                        salida(n, 3) = plantilla
                        salida(n, 4) = fila
                        salida(n, 5) = col
                        salida(n, 6) = z
                        salida(n, 7) = CDbl(x)
                        salida(n, 8) = escala
                        salida(n, 10) = decimales
                        salida(n, 12) = ruta
                        salida(n, 13) = ws.Name
                        salida(n, 14) = ws.Cells(hdr + i, c).Address(False, False)
                    End If
                End If
            Next c
        End If
    Next i
    If nNoNum > 0 Then avisos = avisos & nNoNum & " valores no numericos (ignorados); "
    If nErr > 0 Then avisos = avisos & nErr & " celdas con error (ignoradas); "
    If n > 0 Then EscribirDatos salida, n
    Registrar marco, ruta, ws.Name, plantilla, z, nFilas, nCols, n, avisos, "layout RowNum; " & InfoOmitidos(nOmit)
End Function

Private Sub EscribirDatos(ByRef salida() As Variant, ByVal n As Long)
    Dim dat As Worksheet, fila As Long, bloque() As Variant, i As Long, j As Long
    Dim f As Worksheet, r As Long
    Set dat = ThisWorkbook.Worksheets(HOJA_DATOS)
    fila = dat.Cells(dat.Rows.Count, 2).End(xlUp).Row + 1
    If fila < 2 Then fila = 2
    If fila + n - 1 > MAX_FILAS_DATOS + 1 Then Err.Raise vbObjectError + 1, , "Se supera el maximo de filas de Datos (" & MAX_FILAS_DATOS & ")"

    ReDim bloque(1 To n, 1 To NCOLS_DATOS)
    For i = 1 To n
        For j = 1 To NCOLS_DATOS
            bloque(i, j) = salida(i, j)
        Next j
    Next i
    dat.Range(dat.Cells(fila, 3), dat.Cells(fila + n - 1, 6)).NumberFormat = "@"   ' plantilla, fila, columna, z como texto
    dat.Range(dat.Cells(fila, 13), dat.Cells(fila + n - 1, 14)).NumberFormat = "@"
    dat.Range(dat.Cells(fila, 1), dat.Cells(fila + n - 1, NCOLS_DATOS)).Value = bloque
    dat.Range(dat.Cells(fila, 9), dat.Cells(fila + n - 1, 9)).NumberFormat = "#,##0.00;[Red]-#,##0.00"

    ' columnas calculadas: formulas R1C1 desde _Formulas
    Set f = ThisWorkbook.Worksheets(HOJA_FORM)
    r = 2
    Do While Len(CStr(f.Cells(r, 1).Value)) > 0
        dat.Range(dat.Cells(fila, CLng(f.Cells(r, 1).Value)), _
                  dat.Cells(fila + n - 1, CLng(f.Cells(r, 1).Value))).FormulaR1C1 = "=" & CStr(f.Cells(r, 2).Value)
        r = r + 1
    Loop
End Sub

'------------------------------------------------------------------ utilidades
Private Function EsCodigo(ByVal x As Variant) As Boolean
    Dim s As String
    If VarType(x) <> vbString Then Exit Function
    s = Trim$(x)
    EsCodigo = (s Like "###" Or s Like "####" Or s Like "[rRcC]###" Or s Like "[rRcC]####")
End Function

Private Function EnteroColumna(ByVal x As Variant) As Long
    ' Entero 1..999 como numero o como texto ("1", "12"); 0 si no lo es.
    Dim s As String
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If EsNumero(x) Then
        If x = Int(x) And x >= 1 And x <= 999 Then EnteroColumna = CLng(x)
    ElseIf VarType(x) = vbString Then
        s = Trim$(x)
        If s Like "#" Or s Like "##" Or s Like "###" Then
            If CLng(s) >= 1 Then EnteroColumna = CLng(s)
        End If
    End If
End Function

Private Function EnteroFila(ByVal x As Variant) As Long
    ' RowNum: entero >= 1 como numero o texto de hasta 6 cifras; 0 si no lo es.
    Dim s As String
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If EsNumero(x) Then
        If x = Int(x) And x >= 1 And x < 1000000 Then EnteroFila = CLng(x)
    ElseIf VarType(x) = vbString Then
        s = Trim$(x)
        If Len(s) >= 1 And Len(s) <= 6 And Not s Like "*[!0-9]*" Then
            If CLng(s) >= 1 Then EnteroFila = CLng(s)
        End If
    End If
End Function

Private Function Clave(ByVal marco As String, ByVal plantilla As String, ByVal fila As String, _
                       ByVal col As String, ByVal z As String) As String
    ' Misma construccion que la formula de la columna Clave de Datos.
    Clave = marco & ":" & plantilla & "!r" & fila & "c" & col
    If Len(z) > 0 Then Clave = Clave & "@" & z
End Function

Private Function Usado(ByVal k As String) As Boolean
    If gUsados Is Nothing Then
        Usado = True
    Else
        Usado = gUsados.Exists(k)
    End If
End Function

Private Function InfoOmitidos(ByVal nOmit As Long) As String
    If nOmit > 0 Then InfoOmitidos = nOmit & " datos no usados por ninguna regla (omitidos)"
End Function

Private Function ClavesUsadas() As Object
    ' Claves de la hoja Operandos. Nothing si P_SoloUsados = NO (se carga todo).
    Dim flag As String, d As Object, v As Variant, i As Long
    On Error Resume Next
    flag = UCase$(Trim$(CStr(ThisWorkbook.Names("P_SoloUsados").RefersToRange.Value)))
    On Error GoTo 0
    If flag = "NO" Then Exit Function
    ThisWorkbook.Worksheets("Operandos").Calculate
    Set d = CreateObject("Scripting.Dictionary")
    v = ThisWorkbook.Names("O_Clave").RefersToRange.Value
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If VarType(v(i, 1)) = vbString Then
                If Len(v(i, 1)) > 0 And Not d.Exists(v(i, 1)) Then d.Add v(i, 1), True
            End If
        Next i
    ElseIf VarType(v) = vbString Then
        d.Add v, True
    End If
    Set ClavesUsadas = d
End Function

Private Function NormCodigo(ByVal s As String) As String
    s = Trim$(s)
    If UCase$(Left$(s, 1)) = "R" Or UCase$(Left$(s, 1)) = "C" Then s = Mid$(s, 2)
    If Len(s) < 4 Then s = String$(4 - Len(s), "0") & s
    NormCodigo = s
End Function

Private Function EsNumero(ByVal x As Variant) As Boolean
    Select Case VarType(x)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            EsNumero = True
    End Select
End Function

Private Function ParametroMarco(ByVal marco As String, ByVal col As Long, ByVal defecto As Variant) As Variant
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(HOJA_PARAM)
    ParametroMarco = defecto
    For r = 13 To 30
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = UCase$(marco) Then
            If Len(CStr(ws.Cells(r, col).Value)) > 0 Then ParametroMarco = ws.Cells(r, col).Value
            Exit Function
        End If
    Next r
End Function

Private Sub BuscarMapeo(ByVal marco As String, ByVal hoja As String, ByRef plantilla As String, _
                        ByRef z As String, ByRef omitir As Boolean)
    Dim ws As Worksheet, r As Long, ultima As Long
    Set ws = ThisWorkbook.Worksheets(HOJA_MAPEO)
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To ultima
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = UCase$(marco) And Trim$(CStr(ws.Cells(r, 2).Value)) = hoja Then
            If Len(Trim$(CStr(ws.Cells(r, 3).Value))) > 0 Then plantilla = Trim$(CStr(ws.Cells(r, 3).Value))
            z = Trim$(CStr(ws.Cells(r, 4).Value))
            omitir = (LCase$(Trim$(CStr(ws.Cells(r, 5).Value))) = "si" Or LCase$(Trim$(CStr(ws.Cells(r, 5).Value))) = "s" & ChrW(237))
            Exit Sub
        End If
    Next r
End Sub

Private Sub Registrar(ByVal marco As String, ByVal ruta As String, ByVal hoja As String, ByVal plantilla As String, _
                      ByVal z As String, ByVal nFil As Long, ByVal nCol As Long, ByVal n As Long, ByVal avisos As String, _
                      Optional ByVal info As String = "")
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(HOJA_CARGA)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn")
    ws.Cells(r, 2).Value = marco
    ws.Cells(r, 3).Value = ruta
    ws.Cells(r, 4).NumberFormat = "@"
    ws.Cells(r, 4).Value = hoja
    ws.Cells(r, 5).NumberFormat = "@"
    ws.Cells(r, 5).Value = plantilla
    ws.Cells(r, 6).Value = z
    ws.Cells(r, 7).Value = nFil
    ws.Cells(r, 8).Value = nCol
    ws.Cells(r, 9).Value = n
    ws.Cells(r, 10).Value = Trim$(avisos & IIf(Len(info) > 0, "[info] " & info, ""))
    If Len(avisos) > 0 Then ws.Cells(r, 10).Interior.Color = RGB(255, 235, 156)
End Sub

Private Function ContarMarco(ByVal marco As String) As Long
    Dim dat As Worksheet, ultima As Long
    Set dat = ThisWorkbook.Worksheets(HOJA_DATOS)
    ultima = dat.Cells(dat.Rows.Count, 2).End(xlUp).Row
    If ultima < 2 Then Exit Function
    ContarMarco = Application.WorksheetFunction.CountIf(dat.Range(dat.Cells(2, 2), dat.Cells(ultima, 2)), marco)
End Function

Private Sub BorrarMarco(ByVal marco As String)
    ' No se eliminan filas: eso encogeria los rangos con nombre (D_Clave...) hasta #REF!.
    ' Se leen las formulas R1C1, se vacia el bloque y se reescriben compactadas las de otros marcos.
    Dim dat As Worksheet, ultima As Long, rng As Range, arr As Variant, nuevo() As Variant
    Dim i As Long, j As Long, k As Long, keep As Long
    Set dat = ThisWorkbook.Worksheets(HOJA_DATOS)
    ultima = dat.Cells(dat.Rows.Count, 2).End(xlUp).Row
    If ultima < 2 Then Exit Sub
    Set rng = dat.Range(dat.Cells(2, 1), dat.Cells(ultima, NCOLS_DATOS))
    arr = rng.FormulaR1C1
    For i = 1 To UBound(arr, 1)
        If CStr(arr(i, 2)) <> marco Then keep = keep + 1
    Next i
    rng.ClearContents
    If keep = 0 Then Exit Sub
    ReDim nuevo(1 To keep, 1 To NCOLS_DATOS)
    For i = 1 To UBound(arr, 1)
        If CStr(arr(i, 2)) <> marco Then
            k = k + 1
            For j = 1 To NCOLS_DATOS
                nuevo(k, j) = arr(i, j)
            Next j
        End If
    Next i
    dat.Range(dat.Cells(2, 1), dat.Cells(1 + keep, NCOLS_DATOS)).FormulaR1C1 = nuevo
End Sub

Private Sub VaciarFilas(ByVal ws As Worksheet, ByVal ncols As Long)
    Dim ultima As Long
    ultima = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If ultima < 2 Then Exit Sub
    With ws.Range(ws.Cells(2, 1), ws.Cells(ultima, ncols))
        .ClearContents
        .Interior.ColorIndex = xlColorIndexNone
    End With
End Sub
