

Dim estados() As String
Dim estadosC() As Integer

Dim totalO As Integer
Dim promDur As Double
Dim totalFasesO As Integer
Dim fasesProm As Double
Dim totalFases As Integer
Dim totalInc As Integer
Dim totalCom As Integer
Dim totalR As Integer

Dim porFase() As String
Dim porFaseC() As Integer

Dim durIni() As Integer
Dim durFin() As Integer
Dim durC() As Integer

Dim fasIni() As Integer
Dim fasFin() As Integer
Dim fasC() As Integer

Dim tipoFase() As String
Dim tipoFaseC() As Integer

Dim incFase() As String
Dim incFaseC() As Integer

Dim timeFase() As String
Dim timeFaseC() As Integer

Dim cond() As String
Dim condC() As Integer

Dim estCom() As String
Dim estComC() As Integer

Dim estRepo() As String
Dim estRepoC() As Integer

Dim causas() As String
Dim causasC() As Integer

Dim Events() As String

Dim regs As Integer


Function mesNum(num As Integer) As String

mesNum = "Error"

If num = 1 Then
mesNum = "Enero"
ElseIf num = 2 Then
mesNum = "Febrero"
ElseIf num = 3 Then
mesNum = "Marzo"
ElseIf num = 4 Then
mesNum = "Abril"
ElseIf num = 5 Then
mesNum = "Mayo"
ElseIf num = 6 Then
mesNum = "Junio"
ElseIf num = 7 Then
mesNum = "Julio"
ElseIf num = 8 Then
mesNum = "Agosto"
ElseIf num = 9 Then
mesNum = "Septiembre"
ElseIf num = 10 Then
mesNum = "Octubre"
ElseIf num = 11 Then
mesNum = "Noviembre"
ElseIf num = 12 Then
mesNum = "Diciembre"
Else
Stop '' Comentario: Error: el mes no fue encontrado
End If


End Function

Public Function BubbleSrt(ArrayIn, Ascending As Boolean)

Dim SrtTemp As Variant
Dim i As Long
Dim j As Long


If Ascending = True Then
    For i = LBound(ArrayIn) To UBound(ArrayIn)
         For j = i + 1 To UBound(ArrayIn)
             If ArrayIn(i) > ArrayIn(j) Then
                 SrtTemp = ArrayIn(j)
                 ArrayIn(j) = ArrayIn(i)
                 ArrayIn(i) = SrtTemp
             End If
         Next j
     Next i
Else
    For i = LBound(ArrayIn) To UBound(ArrayIn)
         For j = i + 1 To UBound(ArrayIn)
             If ArrayIn(i) < ArrayIn(j) Then
                 SrtTemp = ArrayIn(j)
                 ArrayIn(j) = ArrayIn(i)
                 ArrayIn(i) = SrtTemp
             End If
         Next j
     Next i
End If

BubbleSrt = ArrayIn

End Function

Function columnaCampo(hoja As String, name As String) As Integer

columnaCampo = 0
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Dim col As Integer
col = 1
Do While tabla.Cells(1, col) <> ""
If tabla.Cells(1, col) = name Then
columnaCampo = col
End If
col = col + 1
Loop

End Function

Function contarTabla(hoja As String, col As Integer) As Integer
contarTabla = 0
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Do While tabla.Cells(contarTabla + 1, col) <> ""
contarTabla = contarTabla + 1
Loop
End Function

Function columnasTabla(hoja As String) As Integer

columnasTabla = 0
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Do While tabla.Cells(1, columnasTabla + 1) <> ""
columnasTabla = columnasTabla + 1
Loop

End Function

Sub arreglarFechas(table As String, campo As String)

Dim tabla As Worksheet
Set tabla = Worksheets(table)
Dim top As Integer
top = contarTabla(table, 1)
Dim col As Integer
col = columnaCampo(table, campo)
For i = 2 To top
If tabla.Cells(i, col) <> "" Then
tabla.Cells(i, col) = fechaNew(tabla.Cells(i, col))
End If
Next i
End Sub


Function fechaNew(fecha As String) As Date

Dim Datos() As String
Datos = Split(fecha, "-")
Dim year As Integer
Dim mes As Integer
Dim dia As Integer
year = CInt(Datos(0))
mes = CInt(Datos(1))
dia = CInt(Datos(2))
fechaNew = DateSerial(year, mes, dia)
End Function

Sub historiaObra()

Call limpiarHistoria

Dim obras As Worksheet
Set obras = Worksheets("Obras")
Dim Actor As Worksheet
Set Actor = Worksheets("actor")
Dim Fase As Worksheet
Set Fase = Worksheets("faseObra")

Dim aux() As Integer
Call inicializarString("Param", 1, Events, aux)
Dim idO As String
idO = Cells(2, 3)
Dim fil As Integer
fil = 9
Dim filaO As Integer
filaO = buscarId(idO, "id", "Obras")
Cells(4, 3) = obras.Cells(filaO, 2)
Cells(5, 3) = obras.Cells(filaO, 3)
Dim filG As Integer
filG = buscarId(obras.Cells(filaO, 4), "id", "actor")
Dim filc As Integer
filc = buscarId(obras.Cells(filaO, 5), "id", "actor")
Cells(4, 6) = Actor.Cells(filG, 2)
Cells(5, 6) = Actor.Cells(filc, 2)

If obras.Cells(filaO, 7) <> "" Then
Cells(fil, 2) = "Ingreso al sistema"
Call fechaHistoria(obras.Cells(filaO, 7), fil)
fil = fil + 1
End If

If obras.Cells(filaO, 9) <> "" Then
Cells(fil, 2) = "Asignacion obra"
Call fechaHistoria(obras.Cells(filaO, 9), fil)
fil = fil + 1
End If
fil = fil + 1

Dim fases() As Integer
fases = pertenecen(idO, "faseObra", "idObra")

If fases(0) > 0 Then
    Dim up As Integer
    up = UBound(fases)
    For i = 0 To up
    Call escribirFase(fases(i), fil)
    fil = fil + 1
    Next i
End If

If obras.Cells(filaO, 8) <> "" Then
Cells(fil, 2) = "Cierre de obra"
Call fechaHistoria(obras.Cells(filaO, 8), fil)
fil = fil + 1
End If

End Sub

Sub inicializarString(hoja As String, col As Integer, arre() As String, count() As Integer)
Dim Param As Worksheet
Set Param = Worksheets(hoja)
Dim re As Integer
re = 0
ReDim arre(re)
ReDim count(re)
arre(re) = "vacio"
Dim fil As Integer
fil = 1
Do While Param.Cells(fil, col) <> ""
arre(re) = Param.Cells(fil, col)
count(re) = 0
re = re + 1
ReDim Preserve arre(re)
ReDim Preserve count(re)
fil = fil + 1
Loop
ReDim Preserve arre(re - 1)
ReDim Preserve count(re - 1)
End Sub

Function buscarId(id As String, idName As String, table As String) As Integer
buscarId = 0
Dim tabla As Worksheet
Set tabla = Worksheets(table)
Dim col As Integer
col = columnaCampo(table, idName)
Dim top As Integer
top = contarTabla(table, 1)
Dim parar As Boolean
parar = False
Dim fil As Integer
fil = 1

Do While parar = False
If tabla.Cells(fil, col) = id Then
buscarId = fil
parar = True
End If

If fil = top Then
parar = True
Else
fil = fil + 1
End If
Loop

End Function

Sub fechaHistoria(fecha As Date, fil As Integer)

Dim sem As Integer
Dim dia As Integer
Dim mes As Integer
Dim yea As Integer
sem = 5
dia = 6
mes = 7
yea = 8
Dim di As Integer
Dim mex As Integer
Dim ye As Integer
di = Day(fecha)
mex = Month(fecha)
ye = year(fecha)
Cells(fil, sem) = diaSemana(Weekday(fecha))
Cells(fil, dia) = di
Cells(fil, mes) = mesNum(mex)
Cells(fil, yea) = ye

End Sub

Function diaSemana(dia As Integer) As String
diaSemana = "Error"
If dia = 1 Then
diaSemana = "Domingo"
ElseIf dia = 2 Then
diaSemana = "Lunes"
ElseIf dia = 3 Then
diaSemana = "Martes"
ElseIf dia = 4 Then
diaSemana = "Miercoles"
ElseIf dia = 5 Then
diaSemana = "Jueves"
ElseIf dia = 6 Then
diaSemana = "Viernes"
ElseIf dia = 7 Then
diaSemana = "Sabado"
End If
End Function

Function pertenecen(idPapa As String, hija As String, idName As String) As Integer()
Dim tabla As Worksheet
Set tabla = Worksheets(hija)
Dim resul() As Integer
Dim re As Integer
re = 0
ReDim resul(re)
Dim col As Integer
col = columnaCampo(hija, idName)
Dim top As Integer
top = contarTabla(hija, 1)
For i = 2 To top
If tabla.Cells(i, col) = idPapa Then
resul(re) = i
re = re + 1
ReDim Preserve resul(re)
End If
Next i
If re > 0 Then
ReDim Preserve resul(re - 1)
Else
resul(re) = -1
End If
pertenecen = resul
End Function

Sub escribirFase(filFa As Integer, fil As Integer)
Dim Fase As Worksheet
Set Fase = Worksheets("faseObra")
Dim Prop As Worksheet
Set Prop = Worksheets("propuesta")
Dim Cont As Worksheet
Set Cont = Worksheets("contraprop")
Dim Com As Worksheet
Set Com = Worksheets("compromiso")
Dim Repo As Worksheet
Set Repo = Worksheets("reporte")

If Fase.Cells(filFa, 7) <> "" Then
Cells(fil, 3) = "Inicio fase " & Fase.Cells(filFa, 3)
Call fechaHistoria(Fase.Cells(filFa, 7), fil)
fil = fil + 1
End If

Dim fechas() As Date
Dim eventos() As String
Dim re As Integer
re = 0
ReDim fechas(re)
ReDim eventos(re)
eventos(re) = "Nada"

Dim filP As Integer
filP = buscarId(Fase.Cells(filFa, 1), "idFase", "propuesta")
If filP > 0 Then
If Prop.Cells(filP, 5) <> "" Then
fechas(re) = Prop.Cells(filP, 5)
eventos(re) = Events(0)
re = re + 1
ReDim Preserve fechas(re)
ReDim Preserve eventos(re)
End If
End If

Dim filc As Integer
filc = buscarId(Fase.Cells(filFa, 1), "idFase", "contraprop")
If filc > 0 Then
fechas(re) = Cont.Cells(filc, 10)
eventos(re) = Events(1)
re = re + 1
ReDim Preserve fechas(re)
ReDim Preserve eventos(re)
End If

Dim comps() As Integer
comps = pertenecen(Fase.Cells(filFa, 1), "compromiso", "idFase")
If comps(0) > 0 Then
    Dim up As Integer
    up = UBound(comps)
    Dim filR As Integer
    For i = 0 To up
    fechas(re) = Com.Cells(comps(i), 7)
    eventos(re) = Events(2)
    re = re + 1
    ReDim Preserve fechas(re)
    ReDim Preserve eventos(re)
    filR = buscarId(Com.Cells(comps(i), 1), "idCompromiso", "reporte")
    If filR > 0 Then
        fechas(re) = Repo.Cells(filR, 6)
        If Repo.Cells(filR, 3) <> "" Then
        eventos(re) = Events(4)
        Else
        eventos(re) = Events(3)
        End If
        re = re + 1
        ReDim Preserve fechas(re)
        ReDim Preserve eventos(re)
    End If
    Next i
End If

If eventos(0) <> "Nada" Then
    ReDim Preserve fechas(re - 1)
    ReDim Preserve eventos(re - 1)
    Dim fechasO() As Date
    Dim fechaux() As Date
    fechaux = fechas
    fechasO = BubbleSrt(fechaux, True)
    eventosO = eventosOrden(fechasO, fechas, eventos)
    Dim upp As Integer
    upp = UBound(fechasO)
    For i = 0 To upp
    Cells(fil, 4) = eventosO(i)
    Call fechaHistoria(fechasO(i), fil)
    fil = fil + 1
    Next i
End If

If Fase.Cells(filFa, 8) <> "" Then
Cells(fil, 3) = "Fin fase de obra"
Call fechaHistoria(Fase.Cells(filFa, 8), fil)
fil = fil + 1
End If

End Sub

Function eventosOrden(fechasO() As Date, fechasD() As Date, eventosD() As String) As String()
Dim up As Integer
up = UBound(fechasO)
Dim resul() As String
ReDim resul(up)
Dim i As Integer
For i = 0 To up
resul(i) = eventoFecha(fechasO(i), fechasD, eventosD)
Next i
eventosOrden = resul
End Function

Function eventoFecha(fecha As Date, fechasD() As Date, eventosD() As String) As String
Dim up As Integer
up = UBound(fechasD)
Dim Eve As String
Eve = "Evento Error"
Dim mark As Integer
mark = -1
Dim i As Integer
For i = 0 To up
If fecha = fechasD(i) Then
If eventosD(i) <> "ok" Then
If prioridadEvento(eventosD(i), Events) < prioridadEvento(Eve, Events) Then
Eve = eventosD(i)
mark = i
End If
End If
End If
Next i
eventosD(mark) = "ok"
eventoFecha = Eve

End Function

Function prioridadEvento(Eve As String, eventos() As String) As Integer
prioridadEvento = 10
Dim up As Integer
up = UBound(eventos)
Dim i As Integer
For i = 0 To up
If eventos(i) = Eve Then
prioridadEvento = i
End If
Next i
End Function

Sub limpiarHistoria()

Cells(4, 3) = ""
Cells(5, 3) = ""
Cells(4, 6) = ""
Cells(5, 6) = ""
Dim top As Integer
top = contarHistoria()
For i = 9 To top
For j = 2 To 8
Cells(i, j) = ""
Next j
Next i

End Sub

Function contarHistoria() As Integer
Dim fil As Integer
fil = 9
Do While Cells(fil, 8) <> "" Or Cells(fil + 1, 8) <> ""
fil = fil + 1
Loop
contarHistoria = fil + 3
End Function


Sub indicadoresPlanilla()
Dim Actor As Worksheet
Set Actor = Worksheets("actor")
Dim Obrass As Worksheet
Set Obrass = Worksheets("Obras")
Dim obras() As Integer
Dim top As Integer
top = contarTabla("Obras", 1)
ReDim obras(top - 2)
Dim i As Integer

For i = 2 To top
    obras(i - 2) = i
Next i
Dim ini As Date
Dim fin As Date
ini = Cells(5, 6)
fin = Cells(5, 7)
obras = filtroFechas("Obras", "fechaProg", ini, fin, obras)
If obras(0) > 0 Then
    If Cells(7, 3) <> "Todos" Then
        Dim filG As Integer
        filG = buscarId(Cells(7, 3), "nombre", "actor")
        Dim idG As String
        idG = Actor.Cells(filG, 1)
        obras = filtroCondicion("Obras", "idGo", idG, obras)
    End If
    If obras(0) > 0 Then
        If Cells(7, 5) <> "Todos" Then
            Dim filc As Integer
            filc = buscarId(Cells(7, 5), "nombre", "actor")
            Dim idC As String
            idC = Actor.Cells(filc, 1)
            obras = filtroCondicion("Obras", "idContrata", idC, obras)
        End If
        If obras(0) > 0 Then
            Call tabularArreglo(obras, "estado", "Obras", estados, estadosC)
            totalO = UBound(obras) + 1
            Dim obrasClosed() As Integer
            obrasClosed = filtroNoVacio("Obras", "fechaCierre", obras)
            Call indicadoresObrasFases2(obras)
            If obrasClosed(0) > 0 Then
                promDur = sumarTiempos("Obras", "fechaProg", "fechaCierre", obrasClosed) / (UBound(obrasClosed) + 1)
                Dim Time As Integer
                For i = 0 To UBound(obrasClosed)
                    Time = Obrass.Cells(obrasClosed(i), 8) - Obrass.Cells(obrasClosed(i), 9)
                    Call tabularRangos(Time, durIni, durFin, durC)
                Next i
            End If
        End If
    End If
End If

End Sub


Sub indicadoresObrasFases2(obras() As Integer)
Dim Obrass As Worksheet
Set Obrass = Worksheets("Obras")
Dim up As Integer
up = UBound(obras)
Dim fases() As Integer
ReDim fases(0)
fases(0) = -1
Dim fasesO() As Integer
Dim fasesTipo() As Integer
Dim idO As String
Dim numFa As Integer

For i = 0 To up
    idO = Obrass.Cells(obras(i), 1)
    fasesO = pertenecen(idO, "faseObra", "idObra")
    If fasesO(0) > 0 Then
        numFa = UBound(fasesO) + 1
        Call tabularRangos(numFa, fasIni, fasFin, fasC)
        If fases(0) > 0 Then
            fases = sumarArreglos(fases, fasesO)
        Else
            fases = fasesO
        End If
        For j = 0 To UBound(porFase)
            fasesTipo = filtroCondicion("faseObra", "fase", porFase(j), fasesO)
            If fasesTipo(0) > 0 Then
                Call tabularTipo(porFase(j), porFase, porFaseC)
            End If
        Next j
    Else
        Call tabularRangos(0, fasIni, fasFin, fasC)
    End If
Next i
If fases(0) > 0 Then
    fasesProm = (UBound(fases) + 1) / totalO
    Call indicadoresFases3(fases)
End If

End Sub

Sub indicadoresFases3(fasesA() As Integer)
Dim Fase As Worksheet
Set Fase = Worksheets("faseObra")
Dim Jefe As Worksheet
Set Jefe = Worksheets("jefeTrabajo")
Dim fases() As Integer
Dim idJ As String
Dim filJ As Integer

If Cells(35, 3) <> "Todos" Then
    filJ = buscarId(Cells(35, 3), "nombre", "jefeTrabajo")
    idJ = Jefe.Cells(filJ, 1)
    fases = filtroCondicion("faseObra", "idJt", idJ, fasesA)
Else
    fases = fasesA
End If

If fases(0) > 0 Then
    fases = filtroNoVacio("faseObra", "fechaCierre", fases)
    If fases(0) > 0 Then
        Call tabularArreglo(fases, "fase", "faseObra", tipoFase, tipoFaseC)
        Dim up As Integer
        up = UBound(fases)
        totalFases = up + 1
        Dim incump() As Integer
        Dim re As Integer
        re = 0
        ReDim incump(re)
        incump(0) = -1
        Dim fasesTipo() As Integer
        
        For i = 0 To UBound(timeFase)
            fasesTipo = filtroCondicion("faseObra", "fase", timeFase(i), fases)
            If fasesTipo(0) > 0 Then
                timeFaseC(i) = sumarTiempos("faseObra", "fechaCreacion", "fechaCierre", fasesTipo)
            End If
        Next i
        For i = 0 To up
            If faseIncumplida(fases(i)) = True Then
                incump(re) = fases(i)
                re = re + 1
                ReDim Preserve incump(re)
            End If
        Next i
        If incump(0) > 0 Then
            ReDim Preserve incump(re - 1)
            Call tabularArreglo(incump, "fase", "faseObra", incFase, incFaseC)
            totalInc = UBound(incump) + 1
        End If
        Call indicadoresFases4(fases)
    End If
End If
End Sub


Sub indicadoresFases4(fasesA() As Integer)
Dim Fase As Worksheet
Set Fase = Worksheets("faseObra")
Dim Com As Worksheet
Set Com = Worksheets("compromiso")
Dim Repo As Worksheet
Set Repo = Worksheets("reporte")
Dim fases() As Integer

If Cells(68, 3) <> "Todos" Then
    fases = filtroCondicion("faseObra", "fase", Cells(68, 3), fasesA)
Else
    fases = fasesA
End If

If fases(0) > 0 Then
    Dim up As Integer
    up = UBound(fases)
    Dim filc As Integer
    Dim idFase As String
    Dim compFase() As Integer
    For i = 0 To up
        idFase = Fase.Cells(fases(i), 1)
        filc = buscarId(idFase, "idFase", "contraprop")
        If filc > 0 Then
            Call tabularCondicion(filc)
        End If
        compFase = pertenecen(idFase, "compromiso", "idFase")
        If compFase(0) > 0 Then
            Call tabularArreglo(compFase, "estado", "compromiso", estCom, estComC)
            Dim upp As Integer
            upp = UBound(compFase)
            totalCom = totalCom + upp + 1
            Dim idCo As String
            Dim filR As Integer
            For j = 0 To upp
                idCo = Com.Cells(compFase(j), 1)
                filR = buscarId(idCo, "idCompromiso", "reporte")
                If filR > 0 Then
                    Call tabularReporte(filR)
                End If
            Next j
        End If
    Next i
End If

End Sub

Sub tabularReporte(filR As Integer)
Dim Repo As Worksheet
Set Repo = Worksheets("reporte")
totalR = totalR + 1

If Repo.Cells(filR, 4) <> "" Then
    estRepoC(1) = estRepoC(1) + 1
    Call tabularTipo(Repo.Cells(filR, 4), causas, causasC)
Else
    estRepoC(0) = estRepoC(0) + 1
End If

End Sub

Sub tabularCondicion(filc As Integer)
Dim Cont As Worksheet
Set Cont = Worksheets("contraprop")
If Left(Cont.Cells(filc, 4), 1) = "S" Then
    condC(0) = condC(0) + 1
End If
If Left(Cont.Cells(filc, 5), 1) = "S" Then
    condC(1) = condC(1) + 1
End If
If Left(Cont.Cells(filc, 6), 1) = "S" Then
    condC(2) = condC(2) + 1
End If
If Left(Cont.Cells(filc, 7), 1) = "S" Then
    condC(3) = condC(3) + 1
End If
If Left(Cont.Cells(filc, 8), 1) = "S" Then
    condC(4) = condC(4) + 1
End If

End Sub

Function faseIncumplida(Fase As Integer) As Boolean

Dim Faseh As Worksheet
Set Faseh = Worksheets("faseObra")
Dim Com As Worksheet
Set Com = Worksheets("compromiso")
Dim Repo As Worksheet
Set Repo = Worksheets("reporte")
faseIncumplida = False
Dim idFase As String
idFase = Faseh.Cells(Fase, 1)
Dim comps() As Integer
comps = pertenecen(idFase, "compromiso", "idFase")

If comps(0) > 0 Then
    Dim up As Integer
    up = UBound(comps)
    Dim idCom As String
    Dim filR As Integer
    For i = 0 To up
        idCom = Com.Cells(comps(i), 1)
        filR = buscarId(idCom, "idCompromiso", "reporte")
        If filR > 0 Then
            If Repo.Cells(filR, 4) <> "" Then
                faseIncumplida = True
            End If
        End If
    Next i
End If

End Function


Function sumarArreglos(arre1() As Integer, arre2() As Integer) As Integer()
Dim up1 As Integer
Dim up2 As Integer
up1 = UBound(arre1)
up2 = UBound(arre2)
Dim resul() As Integer
Dim re As Integer
re = up1 + up2 + 1
ReDim resul(re)
For i = 0 To up1
resul(i) = arre1(i)
Next i
For j = 0 To up2
resul(up1 + 1 + j) = arre2(j)
Next j
sumarArreglos = resul
End Function

Sub tabularRangos(num As Integer, inis() As Integer, fins() As Integer, count() As Integer)
Dim up As Integer
up = UBound(inis)
For i = 0 To up
If inis(i) <= num And num <= fins(i) Then
count(i) = count(i) + 1
End If
Next i
End Sub

Function sumarTiempos(hoja As String, fechai As String, fechaf As String, arre() As Integer) As Integer
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Dim coli As Integer
coli = columnaCampo(hoja, fechai)
Dim colf As Integer
colf = columnaCampo(hoja, fechaf)
sumarTiempos = 0
Dim up As Integer
up = UBound(arre)
Dim Time As Integer
For i = 0 To up
Time = tabla.Cells(arre(i), colf) - tabla.Cells(arre(i), coli)
sumarTiempos = sumarTiempos + Time
Next i
End Function

Sub inicializar()
totalO = 0
promDur = 0
totalFasesO = 0
fasesProm = 0
totalFases = 0
totalInc = 0
totalCom = 0
totalR = 0
Call inicializarString("Param", 7, estados, estadosC)
Call inicializarRango("Param", 2, 3, durIni, durFin, durC)
Call inicializarRango("Param", 4, 5, fasIni, fasFin, fasC)
Call inicializarString("Param", 6, porFase, porFaseC)
Call inicializarString("Param", 6, tipoFase, tipoFaseC)
Call inicializarString("Param", 6, incFase, incFaseC)
Call inicializarString("Param", 6, timeFase, timeFaseC)
Call inicializarString("Param", 9, cond, condC)
Call inicializarString("Param", 7, estCom, estComC)
Call inicializarString("Param", 8, causas, causasC)
Call inicializarString("Param", 10, estRepo, estRepoC)

End Sub

Sub inicializarRango(hoja As String, coli As Integer, colf As Integer, inis() As Integer, fins() As Integer, count() As Integer)
Dim Param As Worksheet
Set Param = Worksheets(hoja)
Dim re As Integer
re = 0
ReDim inis(re)
ReDim fins(re)
ReDim count(re)
Dim fil As Integer
fil = 1
Do While Param.Cells(fil, coli) <> ""
    inis(re) = Param.Cells(fil, coli)
    fins(re) = Param.Cells(fil, colf)
    count(re) = 0
    re = re + 1
    ReDim Preserve inis(re)
    ReDim Preserve fins(re)
    ReDim Preserve count(re)
    fil = fil + 1
Loop
ReDim Preserve inis(re - 1)
ReDim Preserve fins(re - 1)
ReDim Preserve count(re - 1)
End Sub

Sub mostrarIndicadores()

For i = 0 To 2
Cells(11 + i, 4) = estadosC(i)
Next i
Cells(15, 4) = totalO

For i = 0 To 4
Cells(11 + i, 10) = durC(i)
Cells(21 + i, 10) = fasC(i)
Next i
Cells(17, 10) = promDur
Cells(27, 10) = fasesProm

If totalO > 0 Then
For i = 0 To 9
Cells(19 + i, 4) = porFaseC(i) / totalO
Next i
End If

For i = 0 To 9
Cells(39 + i, 4) = tipoFaseC(i)
Cells(39 + i, 10) = incFaseC(i)
If tipoFaseC(i) > 0 Then
Cells(54 + i, 4) = timeFaseC(i) / tipoFaseC(i)
Else
Cells(54 + i, 4) = 0
End If
Next i

Cells(50, 4) = totalFases
Cells(50, 10) = totalInc

For i = 0 To 1
Cells(72 + i, 4) = estRepoC(i)
Next i
Cells(75, 4) = totalR

For i = 0 To 13
Cells(79 + i, 4) = causasC(i)
Next i

For i = 0 To 2
Cells(72 + i, 10) = estComC(i)
Next i
Cells(76, 10) = totalCom

For i = 0 To 4
Cells(80 + i, 10) = condC(i)
Next i

End Sub

Sub limpiar()
For i = 0 To 2
Cells(11 + i, 4) = ""
Next i
Cells(15, 4) = ""

For i = 0 To 4
Cells(11 + i, 10) = ""
Cells(21 + i, 10) = ""
Next i
Cells(17, 10) = ""
Cells(27, 10) = ""

For i = 0 To 9
Cells(19 + i, 4) = ""
Next i

For i = 0 To 9
Cells(39 + i, 4) = ""
Cells(39 + i, 10) = ""
Cells(54 + i, 4) = ""
Next i
Cells(50, 4) = ""
Cells(50, 10) = ""

For i = 0 To 1
Cells(72 + i, 4) = ""
Next i
Cells(75, 4) = ""

For i = 0 To 13
Cells(79 + i, 4) = ""
Next i

For i = 0 To 2
Cells(72 + i, 10) = ""
Next i
Cells(76, 10) = ""

For i = 0 To 4
Cells(80 + i, 10) = ""
Next i

End Sub

Function arregloInicial(hoja As String) As Integer()
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Dim top As Integer
top = contarTabla(hoja, 1)
Dim resul() As Integer
ReDim resul(0)
resul(0) = -1

If top > 1 Then
ReDim resul(top - 2)

Dim i As Integer
For i = 2 To top
resul(i - 2) = i
Next i

End If

arregloInicial = resul

End Function

Function filtroNoVacio(hoja As String, campo As String, arre() As Integer) As Integer()
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Dim col As Integer
col = columnaCampo(hoja, campo)
Dim resul() As Integer
Dim re As Integer
re = 0
ReDim resul(re)
Dim up As Integer
up = UBound(arre)

If arre(0) > 0 Then
For i = 0 To up
If tabla.Cells(arre(i), col) <> "" Then
resul(re) = arre(i)
re = re + 1
ReDim Preserve resul(re)
End If
Next i
End If

If re > 0 Then
ReDim Preserve resul(re - 1)
Else
resul(re) = -1
End If
filtroNoVacio = resul

End Function

Function filtroVacio(hoja As String, campo As String, arre() As Integer) As Integer()
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Dim col As Integer
col = columnaCampo(hoja, campo)
Dim resul() As Integer
Dim re As Integer
re = 0
ReDim resul(re)
Dim up As Integer
up = UBound(arre)

If arre(0) > 0 Then
For i = 0 To up
If tabla.Cells(arre(i), col) = "" Then
resul(re) = arre(i)
re = re + 1
ReDim Preserve resul(re)
End If
Next i
End If

If re > 0 Then
ReDim Preserve resul(re - 1)
Else
resul(re) = -1
End If
filtroVacio = resul

End Function



Sub tabularArreglo(arre() As Integer, campo As String, hoja As String, tipos() As String, count() As Integer)
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Dim col As Integer
col = columnaCampo(hoja, campo)
Dim up As Integer
up = UBound(arre)
Dim tipo As String
For i = 0 To up
tipo = tabla.Cells(arre(i), col)
Call tabularTipo(tipo, tipos, count)
Next i
End Sub


Sub tabularTipo(tipo As String, tipos() As String, count() As Integer)
Dim up As Integer
up = UBound(tipos)
Dim tabu As Boolean
tabu = False
For i = 0 To up
If tipos(i) = tipo Then
count(i) = count(i) + 1
tabu = True
End If
Next i
If tabu = False Then
    Dim otro As String
    otro = "Otros"
    For i = 0 To up
    If tipos(i) = otro Then
    count(i) = count(i) + 1
    tabu = True
    End If
    Next i
End If
End Sub

Function filtroExcluir(hoja As String, campo As String, cond As String, arre() As Integer) As Integer()
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Dim col As Integer
col = columnaCampo(hoja, campo)
Dim resul() As Integer
Dim re As Integer
re = 0
ReDim resul(re)
Dim up As Integer
up = UBound(arre)

If arre(0) > 0 Then
For i = 0 To up
If tabla.Cells(arre(i), col) <> cond Then
resul(re) = arre(i)
re = re + 1
ReDim Preserve resul(re)
End If
Next i
End If


If re > 0 Then
ReDim Preserve resul(re - 1)
Else
resul(re) = -1
End If
filtroExcluir = resul

End Function

Function filtroCondicion(hoja As String, campo As String, cond As String, arre() As Integer) As Integer()
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Dim col As Integer
col = columnaCampo(hoja, campo)
Dim resul() As Integer
Dim re As Integer
re = 0
ReDim resul(re)
Dim up As Integer
up = UBound(arre)

If arre(0) > 0 Then
For i = 0 To up
If tabla.Cells(arre(i), col) = cond Then
resul(re) = arre(i)
re = re + 1
ReDim Preserve resul(re)
End If
Next i
End If


If re > 0 Then
ReDim Preserve resul(re - 1)
Else
resul(re) = -1
End If
filtroCondicion = resul

End Function

Function filtroFechas(hoja As String, fecha As String, ini As Date, fin As Date, arre() As Integer) As Integer()
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Dim col As Integer
col = columnaCampo(hoja, fecha)
Dim resul() As Integer
Dim re As Integer
re = 0
ReDim resul(re)
Dim up As Integer
up = UBound(arre)
Dim fec As Date

If arre(0) > 0 Then
For i = 0 To up
If tabla.Cells(arre(i), col) <> "" Then
fec = tabla.Cells(arre(i), col)
If ini <= fec And fec <= fin Then
resul(re) = arre(i)
re = re + 1
ReDim Preserve resul(re)
End If
End If
Next i
End If

If re > 0 Then
ReDim Preserve resul(re - 1)
Else
resul(re) = -1
End If
filtroFechas = resul

End Function

Sub doIndi()

Call limpiar
Call inicializar
Call indicadoresPlanilla
Call mostrarIndicadores

End Sub


Sub analizarCamposTabla()

Dim Campos As Worksheet
Set Campos = Worksheets("Ancampos")
Dim hoja As String
hoja = Campos.Cells(1, 1)

Dim col As Integer
col = 1

Do While Campos.Cells(3, col) <> ""
Call analizarCampo(col, hoja)
col = col + 1
Loop

End Sub


Sub analizarCampo(col As Integer, hoja As String)

Dim values() As String
Dim count() As Integer
ReDim values(0)
ReDim count(0)

Dim valores As Integer
valores = 0

Dim Datos As Worksheet
Set Datos = Worksheets(hoja)
Dim Campos As Worksheet
Set Campos = Worksheets("Ancampos")
Dim colCampo As Integer
colCampo = columnaCampo(hoja, Campos.Cells(3, col))

Dim top As Integer
top = contarData(hoja, 1)

For fil = 2 To top

If Datos.Cells(fil, colCampo) <> "" Then
If noEsta(Datos.Cells(fil, colCampo), values) = True Then
values(valores) = Datos.Cells(fil, colCampo)
count(valores) = 1
valores = valores + 1
ReDim Preserve values(valores)
ReDim Preserve count(valores)

Else
Call buscarYContar(Datos.Cells(fil, colCampo), values, count)
End If
End If

Next fil

Dim cota As Integer
cota = col * 3 - 1

Campos.Cells(5, cota) = Campos.Cells(3, col)
Campos.Cells(5, cota + 1) = "Numero"

If valores > 0 Then
Dim up As Integer
up = valores - 1
For i = 0 To up
Campos.Cells(6 + i, cota) = values(i)
Campos.Cells(6 + i, cota + 1) = count(i)
Next i
Else
Campos.Cells(6 + i, cota) = "Ningun valor encontrado"
Campos.Cells(6 + i, cota + 1) = 0
End If

End Sub




Function contarData(hoja As String, col As Integer) As Integer

contarData = 0

Do While Worksheets(hoja).Cells(contarData + 1, col) <> ""
contarData = contarData + 1
Loop

End Function


Function noEsta(valor As String, values() As String) As Boolean

noEsta = True
Dim up As Integer
up = UBound(values) - 1
For i = 0 To up
If values(i) = valor Then
noEsta = False
End If
Next i

End Function

Sub buscarYContar(valor As String, values() As String, count() As Integer)

Dim up As Integer
up = UBound(values) - 1
For i = 0 To up
If values(i) = valor Then
count(i) = count(i) + 1
End If
Next i

End Sub




Function filtrar(fil As Integer, arre() As Integer) As Integer()

Dim campo As String
campo = Cells(fil, 2)
Dim hoja As String
hoja = Cells(1, 2)

Dim resul() As Integer
resul = arre

Dim cond As String

If Cells(fil, 1) = "Condicion" Then
cond = Cells(fil, 3)
resul = filtroCondicion(hoja, campo, cond, arre)
ElseIf Cells(fil, 1) = "Excluir" Then
cond = Cells(fil, 3)
resul = filtroExcluir(hoja, campo, cond, arre)
ElseIf Cells(fil, 1) = "Fechas" Then
resul = filtroFechas(hoja, campo, Cells(fil, 3), Cells(fil, 4), arre)
ElseIf Cells(fil, 1) = "NoVacio" Then
resul = filtroNoVacio(hoja, campo, arre)
ElseIf Cells(fil, 1) = "Vacio" Then
resul = filtroVacio(hoja, campo, arre)
End If

filtrar = resul

End Function


Sub hacerFiltros()

Dim hoja As String
hoja = Cells(1, 2)
Dim nueva As String
nueva = Cells(1, 4)

Dim arre() As Integer
arre = arregloInicial(hoja)

Dim i As Integer
For i = 4 To 14
If Cells(i, 1) <> "" And Cells(i, 1) <> "Nada" Then
arre = filtrar(i, arre)
End If
Next i

Call infoFiltro(hoja, nueva, arre)

End Sub

Sub copiarFila(filo As Integer, fild As Integer, ori As String, des As String, cols As Integer)

Dim origen As Worksheet
Set origen = Worksheets(ori)
Dim destino As Worksheet
Set destino = Worksheets(des)

For i = 1 To cols
destino.Cells(fild, i) = origen.Cells(filo, i)
Next i

End Sub


Sub infoFiltro(hoja As String, nueva As String, arre() As Integer)

If existeHoja(nueva) = False Then
Sheets.Add After:=Sheets(Sheets.count)
Sheets(Sheets.count).name = nueva
End If

Dim cols As Integer
cols = columnasTabla(hoja)

Call copiarFila(1, 1, hoja, nueva, cols)

If arre(0) > 0 Then
Dim up As Integer
up = UBound(arre)

Dim ini As Integer
ini = contarTabla(nueva, 1) + 1

Dim i As Integer
For i = 0 To up

Call copiarFila(arre(i), i + ini, hoja, nueva, cols)
Next i

End If

End Sub

Sub copiarHoja()

ActiveSheet.Copy After:=ActiveSheet

End Sub

Sub temporal()

Dim i As Integer
For i = 1 To Sheets.count
Cells(i, 15) = Sheets(i).name
Next i


End Sub

Sub hacerEventos()

Dim Obs As Worksheet
Set Obs = Worksheets("ObsC")

Dim top As Integer
top = contarTabla("ObsC", 1)
regs = 2

For i = 2 To top

Call registrarObra(Obs.Cells(i, 1))

Next i

End Sub


Sub registrarObra(idObra As String)


Dim obras As Worksheet
Set obras = Worksheets("Obras")
Dim Actor As Worksheet
Set Actor = Worksheets("actor")
Dim Fase As Worksheet
Set Fase = Worksheets("faseObra")

Dim aux() As Integer
Call inicializarString("Param", 1, Events, aux)
Dim idO As String
idO = idObra
Dim fil As Integer
fil = 9
Dim filaO As Integer
filaO = buscarId(idO, "id", "Obras")
Dim filG As Integer
filG = buscarId(obras.Cells(filaO, 4), "id", "actor")
Dim filc As Integer
filc = buscarId(obras.Cells(filaO, 5), "id", "actor")

If obras.Cells(filaO, 7) <> "" Then
Call registroIngreso(idO, obras.Cells(filaO, 7))
fil = fil + 1
End If

If obras.Cells(filaO, 9) <> "" Then
Call registroAsigna(idO, obras.Cells(filaO, 9))
fil = fil + 1
End If
fil = fil + 1

Dim fases() As Integer
fases = pertenecen(idO, "faseObra", "idObra")

If fases(0) > 0 Then
    Dim up As Integer
    up = UBound(fases)
    For i = 0 To up
    Call regsFase(fases(i), idO)
    Next i
End If

If obras.Cells(filaO, 8) <> "" Then
Call registroCierre(idO, obras.Cells(filaO, 8))
End If

End Sub

Sub registroIngreso(idO As String, fecha As Date)

Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")

Eve.Cells(regs, 1) = idO
Eve.Cells(regs, 2) = 0
Eve.Cells(regs, 3) = "NA"
Eve.Cells(regs, 4) = "Ingreso sistema"
Eve.Cells(regs, 5) = fecha

regs = regs + 1

End Sub

Sub registroAsigna(idO As String, fecha As Date)

Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")

Eve.Cells(regs, 1) = idO
Eve.Cells(regs, 2) = 0
Eve.Cells(regs, 3) = "NA"
Eve.Cells(regs, 4) = "Asignacion"
Eve.Cells(regs, 5) = fecha

regs = regs + 1

End Sub

Sub registroCierre(idO As String, fecha As Date)

Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")

Eve.Cells(regs, 1) = idO
Eve.Cells(regs, 2) = 0
Eve.Cells(regs, 3) = "NA"
Eve.Cells(regs, 4) = "Cierre"
Eve.Cells(regs, 5) = fecha

regs = regs + 1

End Sub

Sub registroEvento(idO As String, idFa As String, tifa As String, ev As String, fecha As Date)

Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")

Eve.Cells(regs, 1) = idO
Eve.Cells(regs, 2) = idFa
Eve.Cells(regs, 3) = tifa
Eve.Cells(regs, 4) = ev
Eve.Cells(regs, 5) = fecha

regs = regs + 1

End Sub

Sub regsFase(filFa As Integer, idO As String)
Dim Fase As Worksheet
Set Fase = Worksheets("faseObra")
Dim Prop As Worksheet
Set Prop = Worksheets("propuesta")
Dim Cont As Worksheet
Set Cont = Worksheets("contraprop")
Dim Com As Worksheet
Set Com = Worksheets("compromiso")
Dim Repo As Worksheet
Set Repo = Worksheets("reporte")

Dim idFa As String
idFa = Fase.Cells(filFa, 1)
Dim tifa As String
tifa = Fase.Cells(filFa, 3)

If Fase.Cells(filFa, 7) <> "" Then
Call registroEvento(idO, idFa, tifa, "Inicio Fase", Fase.Cells(filFa, 7))
End If

Dim fechas() As Date
Dim eventos() As String
Dim re As Integer
re = 0
ReDim fechas(re)
ReDim eventos(re)
eventos(re) = "Nada"

Dim filP As Integer
filP = buscarId(Fase.Cells(filFa, 1), "idFase", "propuesta")
If filP > 0 Then
If Prop.Cells(filP, 5) <> "" Then
fechas(re) = Prop.Cells(filP, 5)
eventos(re) = Events(0)
re = re + 1
ReDim Preserve fechas(re)
ReDim Preserve eventos(re)
End If
End If

Dim filc As Integer
filc = buscarId(Fase.Cells(filFa, 1), "idFase", "contraprop")
If filc > 0 Then
fechas(re) = Cont.Cells(filc, 10)
eventos(re) = Events(1)
re = re + 1
ReDim Preserve fechas(re)
ReDim Preserve eventos(re)
End If

Dim comps() As Integer
comps = pertenecen(Fase.Cells(filFa, 1), "compromiso", "idFase")
If comps(0) > 0 Then
    Dim up As Integer
    up = UBound(comps)
    Dim filR As Integer
    For i = 0 To up
    fechas(re) = Com.Cells(comps(i), 7)
    eventos(re) = Events(2)
    re = re + 1
    ReDim Preserve fechas(re)
    ReDim Preserve eventos(re)
    filR = buscarId(Com.Cells(comps(i), 1), "idCompromiso", "reporte")
    If filR > 0 Then
    fechas(re) = Repo.Cells(filR, 6)
    If Repo.Cells(filR, 3) <> "" Then
    eventos(re) = Events(4)
    Else
    eventos(re) = Events(3)
    End If
    re = re + 1
    ReDim Preserve fechas(re)
    ReDim Preserve eventos(re)
    End If
    Next i
End If

If eventos(0) <> "Nada" Then
    ReDim Preserve fechas(re - 1)
    ReDim Preserve eventos(re - 1)
    Dim fechasO() As Date
    Dim fechaux() As Date
    fechaux = fechas
    Dim eventosO() As String
    fechasO = BubbleSrt(fechaux, True)
    eventosO = eventosOrden(fechasO, fechas, eventos)
    Dim upp As Integer
    upp = UBound(fechasO)
    For i = 0 To upp
    Call registroEvento(idO, idFa, tifa, eventosO(i), fechasO(i))
    Next i
End If

If Fase.Cells(filFa, 8) <> "" Then
Call registroEvento(idO, idFa, tifa, "Fin fase", Fase.Cells(filFa, 8))
End If

End Sub

Sub buscar()

Dim hoj As String
hoj = Cells(2, 2)

Dim hoja As Worksheet
Set hoja = Worksheets(hoj)

Dim idB As String
Dim idName As String
idName = Cells(3, 2)
idB = Cells(4, 2)

Dim fil As Integer
fil = buscarId(idB, idName, hoj)

hoja.Activate
Cells(fil, 1).Activate


End Sub

Function existeHoja(hoja As String) As Boolean

existeHoja = False

Dim i As Integer
For i = 1 To Worksheets.count
If Worksheets(i).name = hoja Then
existeHoja = True
Exit Function
End If
Next i
     
End Function



Function colsTimeLine(hoja As String) As Integer

Dim Time As Worksheet
Set Time = Worksheets(hoja)

colsTimeLine = 2

Do While Time.Cells(4, colsTimeLine + 1) <> ""
colsTimeLine = colsTimeLine + 1
Loop

End Function

Function colTime(fecha As Date, hoja As String, cols As Integer) As Integer


Dim Time As Worksheet
Set Time = Worksheets(hoja)

colTime = 2

Dim parar As Boolean
parar = False
Dim col As Integer
col = 3

Do While parar = False
If Time.Cells(4, col) <= fecha And fecha <= Time.Cells(5, col) Then
colTime = col
parar = True
End If
If col = cols Then
parar = True
End If
col = col + 1
Loop

End Function

Sub findFilaTipo(hoja As String, tipo As String, ini As Integer, fin As Integer, filt As Integer)

Dim Time As Worksheet
Set Time = Worksheets(hoja)

Dim i As Integer
For i = ini To fin
If Time.Cells(i, 1) = tipo Then
filt = i
End If
Next i

End Sub

Sub primerasFases(hoja As String, cols As Integer)

Dim Time As Worksheet
Set Time = Worksheets(hoja)
Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")
Dim Obs As Worksheet
Set Obs = Worksheets("ObsC")

Dim top As Integer
top = contarTabla("ObsC", 1)

Dim arre() As Integer
Dim colti As Integer
Dim filt As Integer

For i = 2 To top
arre = pertenecen(Obs.Cells(i, 1), "Eventos", "idObra")
arre = filtroCondicion("Eventos", "Evento", "Inicio Fase", arre)

If arre(0) > 0 Then
colti = colTime(Eve.Cells(arre(0), 5), hoja, cols)
Time.Cells(10, colti) = Time.Cells(10, colti) + 1
filt = 59
Call findFilaTipo(hoja, Eve.Cells(arre(0), 3), 49, 58, filt)
Time.Cells(filt, colti) = Time.Cells(filt, colti) + 1
End If
Next i

End Sub

Sub hacerTimeLine(hoja As String)

Dim Time As Worksheet
Set Time = Worksheets(hoja)
Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")

Dim top As Integer
top = contarTabla("Eventos", 1)

Dim cols As Integer
cols = colsTimeLine(hoja)

Dim colti As Integer
Dim filt As Integer

For i = 2 To top

colti = colTime(Eve.Cells(i, 5), hoja, cols)
If Eve.Cells(i, 4) = "Ingreso sistema" Then
filt = 7
ElseIf Eve.Cells(i, 4) = "Asignacion" Then
filt = 8
ElseIf Eve.Cells(i, 4) = "Cierre" Then
filt = 9
ElseIf Eve.Cells(i, 4) = "Inicio Fase" Then
filt = 23
Call findFilaTipo(hoja, Eve.Cells(i, 3), 13, 22, filt)
ElseIf Eve.Cells(i, 4) = "Fin fase" Then
filt = 37
Call findFilaTipo(hoja, Eve.Cells(i, 3), 27, 36, filt)
Else
filt = 40
Call findFilaTipo(hoja, Eve.Cells(i, 4), 41, 45, filt)
End If
Time.Cells(filt, colti) = Time.Cells(filt, colti) + 1
Next i

Call primerasFases(hoja, cols)

Call calculos1(hoja)

Dim e As Integer
For e = 4 To cols
Call calculosN(hoja, e)
Next e

End Sub

Sub limpiarTimeLine(hoja As String)

Dim Time As Worksheet
Set Time = Worksheets(hoja)

Dim cols As Integer
cols = colsTimeLine(hoja)

For i = 6 To 61
For j = 2 To cols

Time.Cells(i, j) = ""

Next j
Next i

End Sub


Sub hacerDias()

Call fechasTimeLine
Call limpiarTimeLine("TimeLine")
Call hacerTimeLine("TimeLine")

End Sub

Sub hacerSemanas()

Call fechasSemanas
Call limpiarTimeLine("Semanas")
Call hacerTimeLine("Semanas")

End Sub


Sub hacerMeses()

Call fechasMeses
Call limpiarTimeLine("Meses")
Call hacerTimeLine("Meses")

End Sub


Sub calculos1(hoja As String)

Dim Time As Worksheet
Set Time = Worksheets(hoja)

Dim inif As Integer
Dim finf As Integer
Dim tem As Integer
inif = 0
finf = 0

For i = 13 To 23
inif = inif + Time.Cells(i, 3)
finf = finf + Time.Cells(i + 14, 3)
Next i

Time.Cells(24, 3) = inif
Time.Cells(38, 3) = finf

For i = 49 To 59
tem = 0
tem = tem + Time.Cells(i, 3)
Time.Cells(i + 17, 3) = tem
Next i

tem = 0
tem = tem + Time.Cells(10, 3)
Time.Cells(79, 3) = tem

tem = 0
tem = tem + Time.Cells(9, 3)
Time.Cells(80, 3) = tem

tem = 0
tem = tem + Time.Cells(79, 3) - Time.Cells(80, 3)
Time.Cells(78, 3) = tem


tem = 0
tem = tem + Time.Cells(24, 3)
Time.Cells(83, 3) = tem

tem = 0
tem = tem + Time.Cells(38, 3)
Time.Cells(84, 3) = tem

tem = 0
tem = tem + Time.Cells(83, 3) - Time.Cells(84, 3)
Time.Cells(82, 3) = tem

tem = 0

For i = 7 To 9
tem = tem + Time.Cells(i, 3)
Next i
tem = tem + Time.Cells(24, 3) + Time.Cells(38, 3)
For i = 41 To 45
tem = tem + Time.Cells(i, 3)
Next i
Time.Cells(86, 3) = tem

End Sub


Sub calculosN(hoja As String, col As Integer)

Dim Time As Worksheet
Set Time = Worksheets(hoja)

Dim inif As Integer
Dim finf As Integer
Dim tem As Integer
inif = 0
finf = 0

For i = 13 To 23
inif = inif + Time.Cells(i, col)
finf = finf + Time.Cells(i + 14, col)
Next i

Time.Cells(24, col) = inif
Time.Cells(38, col) = finf

For i = 49 To 59
tem = Time.Cells(i + 17, col - 1)
tem = tem + Time.Cells(i, col)
Time.Cells(i + 17, col) = tem
Next i

tem = 0
tem = tem + Time.Cells(10, col)
Time.Cells(79, col) = tem

tem = 0
tem = tem + Time.Cells(9, col)
Time.Cells(80, col) = tem

tem = Time.Cells(78, col - 1)
tem = tem + Time.Cells(79, col) - Time.Cells(80, col)
Time.Cells(78, col) = tem


tem = 0
tem = tem + Time.Cells(24, col)
Time.Cells(83, col) = tem

tem = 0
tem = tem + Time.Cells(38, col)
Time.Cells(84, col) = tem

tem = Time.Cells(82, col - 1)
tem = tem + Time.Cells(83, col) - Time.Cells(84, col)
Time.Cells(82, col) = tem

tem = 0

For i = 7 To 9
tem = tem + Time.Cells(i, col)
Next i
tem = tem + Time.Cells(24, col) + Time.Cells(38, col)
For i = 41 To 45
tem = tem + Time.Cells(i, col)
Next i
Time.Cells(86, col) = tem

End Sub

Sub fechasTimeLine()

Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")
Dim Time As Worksheet
Set Time = Worksheets("TimeLine")

Dim ini As Date
ini = Eve.Cells(2, 5)
Dim col As Integer
col = 3

Dim top As Integer
top = contarTabla("Eventos", 1)

Dim fin As Date
fin = Eve.Cells(top, 5)

Do While ini <= fin
Time.Cells(4, col) = ini
Time.Cells(5, col) = ini
Time.Cells(3, col) = Day(ini)
Time.Cells(1, col) = mesNum(Month(ini))
ini = DateAdd("d", 1, ini)
col = col + 1
Loop

End Sub

Sub fechasSemanas()

Dim Time As Worksheet
Set Time = Worksheets("Semanas")
Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")

Call iniSemanas

Dim col As Integer
col = 3

Dim top As Integer
top = contarTabla("Eventos", 1)

Dim fin As Date
fin = Eve.Cells(top, 5)

Dim inif As Date
Dim finf As Date
inif = Time.Cells(4, 2)
finf = Time.Cells(5, 2)

Do While finf <= fin
inif = DateAdd("d", 7, inif)
finf = DateAdd("d", 7, finf)
Time.Cells(4, col) = inif
Time.Cells(3, col) = Day(inif)
Time.Cells(1, col) = mesNum(Month(inif))
Time.Cells(5, col) = finf
col = col + 1
Loop


End Sub

Sub fechasMeses()

Dim Time As Worksheet
Set Time = Worksheets("Meses")
Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")

Dim col As Integer
col = 3

Dim top As Integer
top = contarTabla("Eventos", 1)

Dim fin As Date
fin = Eve.Cells(top, 5)
Dim ini As Date
ini = Eve.Cells(2, 5)

Dim mes As Integer
Dim ye As Integer
mes = Month(ini)
ye = year(ini)


Dim inif As Date
Dim finf As Date

Do While finf <= fin
inif = DateSerial(ye, mes, 1)
finf = CDate(Application.WorksheetFunction.EoMonth(inif, 0))
Time.Cells(4, col) = inif
Time.Cells(3, col) = Day(inif)
Time.Cells(1, col) = mesNum(Month(inif))
Time.Cells(5, col) = finf

If mes < 12 Then
mes = mes + 1
Else
mes = 1
ye = ye + 1
End If

col = col + 1
Loop



End Sub


Sub iniSemanas()

Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")
Dim Time As Worksheet
Set Time = Worksheets("Semanas")

Dim inicial As Date
inicial = Eve.Cells(2, 5)

Dim dia As Integer
dia = Weekday(inicial)

Dim menos As Integer

If dia >= 2 Then
menos = 7 + dia - 2
Else
menos = 13
End If

Time.Cells(4, 2) = DateAdd("d", -menos, inicial)
Time.Cells(5, 2) = DateAdd("d", -menos + 6, inicial)

End Sub


Sub fasesFinales()

Dim Obs As Worksheet
Set Obs = Worksheets("ObsC")
Dim Eve As Worksheet
Set Eve = Worksheets("Eventos")

Call inicializarString("Param", 6, tipoFase, tipoFaseC)

Dim obras() As Integer

obras = arregloInicial("ObsC")
obras = filtroNoVacio("ObsC", "fechaCierre", obras)

Dim up As Integer
up = UBound(obras)

Dim fines() As Integer
Dim idPapa As String

If obras(0) > 0 Then
For i = 0 To up

idPapa = Obs.Cells(obras(i), 1)
fines = pertenecen(idPapa, "Eventos", "idObra")
fines = filtroCondicion("Eventos", "Evento", "Fin fase", fines)

If fines(0) > 0 Then

Dim uf As Integer
uf = UBound(fines)
Call tabularTipo(Eve.Cells(fines(uf), 3), tipoFase, tipoFaseC)

End If

Next i
End If

Eve.Cells(1, 8) = "Fase"
Eve.Cells(1, 9) = "Conteo"

Dim unp As Integer
unp = UBound(tipoFase)
For e = 0 To unp

Eve.Cells(2 + e, 8) = tipoFase(e)
Eve.Cells(2 + e, 9) = tipoFaseC(e)

Next e

Eve.Cells(14, 8) = "Total"
Eve.Cells(14, 9).FormulaR1C1 = "=SUM(R2C:R12C)"

End Sub

Sub promedio(hoja As String, fil As Integer)

Dim Time As Worksheet
Set Time = Worksheets(hoja)

Dim col As Integer
col = colsTimeLine(hoja)

Time.Cells(fil - 1, col + 2) = "Promedio"
Time.Cells(fil, col + 2).FormulaR1C1 = "=AVERAGE(RC3:RC" & col & ")"

End Sub

Sub suma(hoja As String, fil As Integer)

Dim Time As Worksheet
Set Time = Worksheets(hoja)

Dim col As Integer
col = colsTimeLine(hoja)

Time.Cells(fil - 1, col + 3) = "Total"
Time.Cells(fil, col + 3).FormulaR1C1 = "=SUM(RC3:RC" & col & ")"

End Sub


Sub otrosIndica()

Dim hoja As String
hoja = "TimeLine"

Worksheets("Param").Cells(11, 6) = "Otros"

Call promedio("TimeLine", 86)
Call promedio("Semanas", 86)
Call promedio("Semanas", 78)

Call suma("Semanas", 80)
Call suma("Semanas", 79)

Call fasesFinales

End Sub


Function mesNume(num As Integer) As String

mesNume = "Error"

If num = 1 Then
mesNume = "Ene"
ElseIf num = 2 Then
mesNume = "Feb"
ElseIf num = 3 Then
mesNume = "Mar"
ElseIf num = 4 Then
mesNume = "Abr"
ElseIf num = 5 Then
mesNume = "May"
ElseIf num = 6 Then
mesNume = "Jun"
ElseIf num = 7 Then
mesNume = "Jul"
ElseIf num = 8 Then
mesNume = "Ago"
ElseIf num = 9 Then
mesNume = "Sep"
ElseIf num = 10 Then
mesNume = "Oct"
ElseIf num = 11 Then
mesNume = "Nov"
ElseIf num = 12 Then
mesNume = "Dic"
Else
Stop '' Comentario: Error: el mes no fue encontrado
End If


End Function

Sub aAreglarFechas()

Call fixFechas("TimeLine")
Call fixFechas("Semanas")
Call fixFechas("Meses")

End Sub

Sub fixFechas(hoja As String)

Dim Time As Worksheet
Set Time = Worksheets(hoja)

Dim ini As Integer
ini = 3

Dim dia As Integer
Dim mes As String

Do While Time.Cells(4, ini) <> ""
dia = Day(Time.Cells(4, ini))
mes = mesNume(Month(Time.Cells(4, ini)))
Time.Cells(4, ini) = dia & "." & mes
ini = ini + 1
Loop

End Sub




Sub GraficarDataDias()
'
' GraficarData Macro
'
'

Dim col As Integer
col = colsTimeLine("TimeLine")
Dim le As String
le = ColumnLetter(col)
    ActiveSheet.ChartObjects("2 Gr_fico").Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).name = "Obras activas"
    ActiveChart.SeriesCollection(1).values = "=TimeLine!$C$78:$" & le & "$78"
    ActiveChart.SeriesCollection(1).XValues = "=TimeLine!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(1).MarkerStyle = -4142
    ActiveSheet.ChartObjects("3 Gr_fico").Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).name = "Activacion"
    ActiveChart.SeriesCollection(1).values = "=TimeLine!$C$79:$" & le & "$79"
    ActiveChart.SeriesCollection(1).XValues = "=TimeLine!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(1).MarkerStyle = -4142
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).name = "Cierre"
    ActiveChart.SeriesCollection(2).values = "=TimeLine!$C$80:$" & le & "$80"
    ActiveChart.SeriesCollection(2).XValues = "=TimeLine!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(2).MarkerStyle = -4142
    
    ActiveSheet.ChartObjects("4 Gr_fico").Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).name = "Fases activas"
    ActiveChart.SeriesCollection(1).values = "=TimeLine!$C$82:$" & le & "$82"
    ActiveChart.SeriesCollection(1).XValues = "=TimeLine!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(1).MarkerStyle = -4142
    ActiveSheet.ChartObjects("5 Gr_fico").Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).name = "Inicio"
    ActiveChart.SeriesCollection(1).values = "=TimeLine!$C$83:$" & le & "$83"
    ActiveChart.SeriesCollection(1).XValues = "=TimeLine!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(1).MarkerStyle = -4142
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).name = "Cierre"
    ActiveChart.SeriesCollection(2).values = "=TimeLine!$C$84:$" & le & "$84"
    ActiveChart.SeriesCollection(2).XValues = "=TimeLine!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(2).MarkerStyle = -4142
    
    ActiveSheet.ChartObjects("6 Gr_fico").Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).name = "Actividad"
    ActiveChart.SeriesCollection(1).values = "=TimeLine!$C$86:$" & le & "$86"
    ActiveChart.SeriesCollection(1).XValues = "=TimeLine!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(1).MarkerStyle = -4142
    
    Dim fases() As Integer
    fases = ordenFases("TimeLine")
    Dim i As Integer
    
    ActiveSheet.ChartObjects("7 Gr_fico").Activate
    For i = 0 To 2
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(i + 1).name = "=TimeLine!$A$" & fases(i)
    ActiveChart.SeriesCollection(i + 1).values = "=TimeLine!$C$" & fases(i) & ":$" & le & "$" & fases(i)
    ActiveChart.SeriesCollection(i + 1).XValues = "=TimeLine!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(i + 1).MarkerStyle = -4142
    Next i
    
    ActiveSheet.ChartObjects("8 Gr_fico").Activate
    For i = 3 To 6
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(i - 2).name = "=TimeLine!$A$" & fases(i)
    ActiveChart.SeriesCollection(i - 2).values = "=TimeLine!$C$" & fases(i) & ":$" & le & "$" & fases(i)
    ActiveChart.SeriesCollection(i - 2).XValues = "=TimeLine!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(i - 2).MarkerStyle = -4142
    Next i
    
    ActiveSheet.ChartObjects("9 Gr_fico").Activate
    For i = 7 To 10
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(i - 6).name = "=TimeLine!$A$" & fases(i)
    ActiveChart.SeriesCollection(i - 6).values = "=TimeLine!$C$" & fases(i) & ":$" & le & "$" & fases(i)
    ActiveChart.SeriesCollection(i - 6).XValues = "=TimeLine!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(i - 6).MarkerStyle = -4142
    Next i
    
    
End Sub

Function ColumnLetter(col As Integer) As String

ColumnLetter = Split(Cells(1, col).Address, "$")(1)

End Function

Function ordenFases(hoja As String) As Integer()

Dim Time As Worksheet
Set Time = Worksheets(hoja)

Dim totals() As Integer
Dim ordin() As Integer
ReDim totals(10)
ReDim ordin(10)

Dim col As Integer
col = colsTimeLine(hoja)

Dim i As Integer
For i = 66 To 76
totals(i - 66) = Time.Cells(i, col)
ordin(i - 66) = i
Next i

totals = BubbleSrt(totals, False)

Dim resul(10) As Integer
Dim done As Boolean


For i = 0 To 10
done = False
For j = 0 To 10
If done = False Then
If ordin(j) > 0 Then
If totals(i) = Time.Cells(ordin(j), col) Then
resul(i) = ordin(j)
ordin(j) = 0
done = True
End If
End If
End If

Next j
Next i


ordenFases = resul

End Function


Sub GraficarData()
'
' GraficarData Macro
'
'

Dim hoja As String
hoja = ActiveSheet.name

Dim col As Integer
col = colsTimeLine(hoja)
Dim le As String
le = ColumnLetter(col)
    ActiveSheet.ChartObjects("1 Gr_fico").Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).name = "Obras activas"
    ActiveChart.SeriesCollection(1).values = "=" & hoja & "!$C$78:$" & le & "$78"
    ActiveChart.SeriesCollection(1).XValues = "=" & hoja & "!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(1).MarkerStyle = -4142
    ActiveSheet.ChartObjects("2 Gr_fico").Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).name = "Activacion"
    ActiveChart.SeriesCollection(1).values = "=" & hoja & "!$C$79:$" & le & "$79"
    ActiveChart.SeriesCollection(1).XValues = "=" & hoja & "!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(1).MarkerStyle = -4142
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).name = "Cierre"
    ActiveChart.SeriesCollection(2).values = "=" & hoja & "!$C$80:$" & le & "$80"
    ActiveChart.SeriesCollection(2).XValues = "=" & hoja & "!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(2).MarkerStyle = -4142
    
    ActiveSheet.ChartObjects("3 Gr_fico").Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).name = "Fases activas"
    ActiveChart.SeriesCollection(1).values = "=" & hoja & "!$C$82:$" & le & "$82"
    ActiveChart.SeriesCollection(1).XValues = "=" & hoja & "!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(1).MarkerStyle = -4142
    ActiveSheet.ChartObjects("4 Gr_fico").Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).name = "Inicio"
    ActiveChart.SeriesCollection(1).values = "=" & hoja & "!$C$83:$" & le & "$83"
    ActiveChart.SeriesCollection(1).XValues = "=" & hoja & "!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(1).MarkerStyle = -4142
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).name = "Cierre"
    ActiveChart.SeriesCollection(2).values = "=" & hoja & "!$C$84:$" & le & "$84"
    ActiveChart.SeriesCollection(2).XValues = "=" & hoja & "!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(2).MarkerStyle = -4142
    
    ActiveSheet.ChartObjects("5 Gr_fico").Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).name = "Actividad"
    ActiveChart.SeriesCollection(1).values = "=" & hoja & "!$C$86:$" & le & "$86"
    ActiveChart.SeriesCollection(1).XValues = "=" & hoja & "!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(1).MarkerStyle = -4142
    
    Dim fases() As Integer
    fases = ordenFases(hoja)
    Dim i As Integer
    
    ActiveSheet.ChartObjects("6 Gr_fico").Activate
    For i = 0 To 2
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(i + 1).name = "=" & hoja & "!$A$" & fases(i)
    ActiveChart.SeriesCollection(i + 1).values = "=" & hoja & "!$C$" & fases(i) & ":$" & le & "$" & fases(i)
    ActiveChart.SeriesCollection(i + 1).XValues = "=" & hoja & "!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(i + 1).MarkerStyle = -4142
    Next i
    
    ActiveSheet.ChartObjects("7 Gr_fico").Activate
    For i = 3 To 6
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(i - 2).name = "=" & hoja & "!$A$" & fases(i)
    ActiveChart.SeriesCollection(i - 2).values = "=" & hoja & "!$C$" & fases(i) & ":$" & le & "$" & fases(i)
    ActiveChart.SeriesCollection(i - 2).XValues = "=" & hoja & "!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(i - 2).MarkerStyle = -4142
    Next i
    
    ActiveSheet.ChartObjects("8 Gr_fico").Activate
    For i = 7 To 10
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(i - 6).name = "=" & hoja & "!$A$" & fases(i)
    ActiveChart.SeriesCollection(i - 6).values = "=" & hoja & "!$C$" & fases(i) & ":$" & le & "$" & fases(i)
    ActiveChart.SeriesCollection(i - 6).XValues = "=" & hoja & "!$C$4:$" & le & "$4"
    ActiveChart.SeriesCollection(i - 6).MarkerStyle = -4142
    Next i
    
    
End Sub

Sub unificarCampos()

Dim tabla As String
tabla = Cells(1, 2)

Dim Taab As Worksheet
Set Taab = Worksheets(tabla)

Dim viejos() As String
Dim nuevos() As String

Dim temp() As Integer

Call inicializarString("Unify", 3, viejos, temp)
Call inicializarString("Unify", 4, nuevos, temp)

Dim top As Integer
top = contarTabla(tabla, 1)

Dim campo As String
campo = Cells(2, 2)
Dim col As Integer
col = columnaCampo(tabla, campo)

Dim up As Integer
up = UBound(viejos)

For i = 1 To top

If Taab.Cells(i, col) <> "" Then

    For e = 0 To up
    
    If Taab.Cells(i, col) = viejos(e) Then
    Taab.Cells(i, col) = nuevos(e)
    End If
    
    Next e

End If

Next i

End Sub

Sub compararListadoRut()

Dim Lista As Worksheet
Set Lista = Worksheets("Listados RUT")

Dim top As Integer
top = contarTabla("Listados RUT", 1)
Dim id As String
Dim res As Integer

For i = 1 To top
id = Lista.Cells(i, 4)
res = buscarId(id, "RUT TRABAJADOR", "Personal")

If res > 0 Then
Lista.Cells(i, 15) = "Esta"
Else
Lista.Cells(i, 15) = "No Esta"
End If

Next i

End Sub

Sub hacerMapa()

Dim fil As Integer
Dim col As Integer
fil = 7
col = 1

For i = 1 To 86

Cells(fil, col) = Cells(i, 6)

If col = 3 Then
col = 1
fil = fil + 1
Else
col = col + 1
End If

Next i

End Sub

Function tipoLocal(locali As String) As String

Dim arre() As Integer
tipoLocal = "Otro"

arre = arregloInicial("Mapa")

Dim resu() As Integer

resu = filtroCondicion("Mapa", "ZONE", locali, arre)

If resu(0) > 0 Then
tipoLocal = "Zona"
End If

resu = filtroCondicion("Mapa", "ESTABLECIMIENTOS", locali, arre)

If resu(0) > 0 Then
tipoLocal = "Establecimiento"
End If

End Function

Function contarFila(hoja As String, fil As Integer) As Integer

contarFila = 0
Dim tabla As Worksheet
Set tabla = Worksheets(hoja)
Do While tabla.Cells(fil, contarFila + 1) <> ""
contarFila = contarFila + 1
Loop

End Function

Sub iniciarCampos(ori() As Integer, des() As Integer)

Dim Menu As Worksheet
Set Menu = ActiveSheet

Dim origen As String
Dim destino As String
origen = Menu.Cells(1, 2)
destino = Menu.Cells(2, 2)

Dim col As Integer
col = contarFila(ActiveSheet.name, 4)

ReDim ori(col - 1)
ReDim des(col - 1)

Dim c1 As Integer
Dim c2 As Integer

For i = 1 To col
c1 = columnaCampo(origen, Menu.Cells(4, i))
c2 = columnaCampo(destino, Menu.Cells(5, i))

If c1 = 0 Or c2 = 0 Then
Stop 'Algun campo no ha sido encontrado
End If

ori(i - 1) = c1
des(i - 1) = c2

Next i


End Sub

Sub copiarInfo(origen As String, destino As String, filo As Integer, fild As Integer, ori() As Integer, des() As Integer)

Dim up As Integer
up = UBound(ori)

For i = 0 To up

Worksheets(destino).Cells(fild, des(i)) = Worksheets(origen).Cells(filo, ori(i))

Next i

End Sub

Sub ampliarI()

Dim Menu As Worksheet
Set Menu = ActiveSheet

Dim origen As String
Dim destino As String
origen = Menu.Cells(1, 2)
destino = Menu.Cells(2, 2)

Dim idO As String
Dim idd As String
idO = Menu.Cells(1, 4)
idd = Menu.Cells(2, 4)

Dim ori() As Integer
Dim des() As Integer
Call iniciarCampos(ori, des)

Dim fild As Integer
Dim top As Integer
top = contarTabla(origen, 1)
Dim colido As Integer
colido = columnaCampo(origen, idO)

If colido > 0 Then

Dim idac As String
Dim i As Integer
For i = 2 To top

idac = Worksheets(origen).Cells(i, colido)
fild = buscarId(idac, idd, destino)
If fild > 0 Then
Call copiarInfo(origen, destino, i, fild, ori, des)
End If

Next i

Else
MsgBox "id origen no encontrado (el campo)"
End If

MsgBox "Listo!"

End Sub

Sub expandirReg()

Dim Menu As Worksheet
Set Menu = ActiveSheet

Dim origen As String
Dim destino As String
origen = Menu.Cells(1, 2)
destino = Menu.Cells(2, 2)

Dim idO As String
Dim idd As String
idO = Menu.Cells(1, 4)
idd = Menu.Cells(2, 4)

Dim ori() As Integer
Dim des() As Integer
Call iniciarCampos(ori, des)

Dim top As Integer
top = contarTabla(origen, 1)
Dim fild As Integer
fild = contarTabla(destino, 1) + 1

Dim i As Integer
For i = 2 To top

Call copiarInfo(origen, destino, i, fild, ori, des)
fild = fild + 1

Next i

End Sub

Function zonaEstab(locali As String) As String

zonaEstab = "Otro"

Dim zona As Integer
zona = buscarId(locali, "ZONE", "Mapa")

If zona > 0 Then
zonaEstab = locali
End If

Dim est As Integer
est = buscarId(locali, "ESTABLECIMIENTOS", "Mapa")
If est > 0 Then
zonaEstab = Worksheets("Mapa").Cells(est, 1)
End If

End Function

Sub AApruebebe()

Dim Cont As Worksheet
Set Cont = Worksheets("Contratista")

Dim Cuco As Worksheet
Set Cuco = Worksheets("Cuenta_contratista")

Dim Cuen As Worksheet
Set Cuen = Worksheets("Cuenta")

Dim top As Integer
top = contarTabla("Contratista", 1)

Dim idC As String
Dim filCC As Integer
Dim idCu As String
Dim filCu As Integer

For i = 2 To top
idC = Cont.Cells(i, 4)
filCC = buscarId(idC, "id_contratista", "Cuenta_contratista")
    If filCC > 0 Then
    idCu = Cuco.Cells(filCC, 2)
    filCu = buscarId(idCu, "id", "Cuenta")
        If filCu > 0 Then
        Cont.Cells(i, 5) = Cuen.Cells(filCu, 1)
        Cont.Cells(i, 6) = Cuen.Cells(filCu, 2)
        Else
        Cont.Cells(i, 5) = "No se encuentra la cuenta, idCu:" & idCu
        End If
    Else
    Cont.Cells(i, 5) = "No existe la relacion"
    End If
Next i

End Sub

Sub darzonas()

For i = 2 To 46

Cells(i, 2) = zonaEstab(Cells(i, 1))

Next i


End Sub

Sub llenarVacios(tabla As String, campo As String, value As String)

Dim hoja As Worksheet
Set hoja = Worksheets(tabla)

Dim top As Integer
top = contarTabla(tabla, 1)
Dim col As Integer
col = columnaCampo(tabla, campo)

For i = 2 To top

If hoja.Cells(i, col) = "" Then
hoja.Cells(i, col) = value
End If

Next i

End Sub

Sub llenarVaciosMenu()

Dim top As Integer
top = contarTabla(ActiveSheet.name, 1)

For i = 2 To top
Call llenarVacios(Cells(i, 1), Cells(i, 2), Cells(i, 3))
Next i

End Sub


Sub contratistaBrigada()

Dim Bri As Worksheet
Set Bri = Worksheets("Briga")

Dim Cont As Worksheet
Set Cont = Worksheets("Contra")

Dim top As Integer
top = contarTabla("Briga", 1)

Dim rut As String
Dim filc As Integer

For i = 2 To top

rut = Bri.Cells(i, 2)
filc = buscarId(rut, "RUT", "Contra")
If filc > 0 Then
Bri.Cells(i, 9) = Cont.Cells(filc, 5)
Else
Bri.Cells(i, 9) = "No esta"
End If

Next i

End Sub


Sub cuentaContratista()

Dim idC As String
Dim rutC As String
Dim idCo As String
Dim filco As Integer
 

For i = 2 To 76
idC = Worksheets("cuentas").Cells(i, 3)
rutC = Worksheets("cuentas").Cells(i, 1)
filco = buscarId(rutC, "RUT", "Contra")
idCo = Worksheets("Contra").Cells(filco, 5)

Worksheets("Data").Cells(i, 2) = idC
Worksheets("Data").Cells(i, 1) = idCo
Next i


End Sub

Sub contactoContratista()

Dim idCc As String
Dim idCt As String
Dim rutC As String
Dim filco As Integer

Dim top As Integer
top = contarTabla("Contac", 1)

For i = 2 To top
idCc = Worksheets("Contac").Cells(i, 14)
rutC = Worksheets("Contac").Cells(i, 3)

filco = buscarId(rutC, "RUT", "Contra")
idCt = Worksheets("Contra").Cells(filco, 5)
Worksheets("Data").Cells(i, 2) = idCc
Worksheets("Data").Cells(i, 1) = idCt

Next i

End Sub


Sub contratistaTrabajador()

Dim idT As String
Dim idCt As String
Dim filco As Integer
Dim rutC As String

Dim top As Integer
top = contarTabla("Pers", 1)

For i = 2 To top

idT = Worksheets("Pers").Cells(i, 11)
rutC = Worksheets("Pers").Cells(i, 1)

filco = buscarId(rutC, "RUT", "Contra")
idCt = Worksheets("Contra").Cells(filco, 5)
Worksheets("Data").Cells(i, 2) = idT
Worksheets("Data").Cells(i, 1) = idCt

Next i


End Sub


Sub vehiculoContratista()

Dim idV As String
Dim idCt As String
Dim filco As Integer
Dim rutC As String

Dim top As Integer
top = contarTabla("Vehic", 1)

For i = 2 To top

idV = Worksheets("Vehic").Cells(i, 8)
rutC = Worksheets("Vehic").Cells(i, 1)

filco = buscarId(rutC, "RUT", "Contra")
idCt = Worksheets("Contra").Cells(filco, 5)
Worksheets("Data").Cells(i, 1) = idV
Worksheets("Data").Cells(i, 2) = idCt

Next i


End Sub

Sub funcionesTrabajador()


Dim idT As String
Dim idF As String

Dim top As Integer
top = contarTabla("Pers", 1)

Dim file As Integer
file = 2

For i = 2 To top

idT = Worksheets("Pers").Cells(i, 11)

If Worksheets("Pers").Cells(i, 6) <> "" Then
idF = Worksheets("funciones").Cells(buscarId(Worksheets("Pers").Cells(i, 6), "Nombre", "funciones"), 3)
Worksheets("Data").Cells(file, 2) = idF
Worksheets("Data").Cells(file, 3) = idT
Worksheets("Data").Cells(file, 1) = "Principal"
file = file + 1
End If

If Worksheets("Pers").Cells(i, 7) <> "" Then
idF = Worksheets("funciones").Cells(buscarId(Worksheets("Pers").Cells(i, 7), "Nombre", "funciones"), 3)
Worksheets("Data").Cells(file, 2) = idF
Worksheets("Data").Cells(file, 3) = idT
Worksheets("Data").Cells(file, 1) = "Secundaria"
file = file + 1
End If

If Worksheets("Pers").Cells(i, 8) <> "" Then
idF = Worksheets("funciones").Cells(buscarId(Worksheets("Pers").Cells(i, 8), "Nombre", "funciones"), 3)
Worksheets("Data").Cells(file, 2) = idF
Worksheets("Data").Cells(file, 3) = idT
Worksheets("Data").Cells(file, 1) = "Tercera"
file = file + 1
End If

Next i

End Sub

Sub servicioBrigada()
Dim idB As String
Dim serv As String
Dim idS As String
Dim top As Integer
top = contarTabla("Briga", 1)
For i = 2 To top
idB = Worksheets("Briga").Cells(i, 8)
serv = Worksheets("Briga").Cells(i, 4)
idS = Worksheets("servicios").Cells(buscarId(serv, "Nombre", "servicios"), 3)
Worksheets("Data").Cells(i, 1) = idB
Worksheets("Data").Cells(i, 2) = idS
Next i
End Sub

Sub servicioVehiculo()

Dim idV As String
Dim serv As String
Dim idS As String
Dim top As Integer
top = contarTabla("Vehic", 1)
For i = 2 To top
idV = Worksheets("Vehic").Cells(i, 8)
serv = Worksheets("Vehic").Cells(i, 5)

idS = Worksheets("servicios_vehiculos").Cells(buscarId(serv, "Nombre", "servicios_vehiculos"), 3)
Worksheets("Data").Cells(i, 1) = idV
Worksheets("Data").Cells(i, 2) = idS
Next i


End Sub

Sub sevVehicuPruebaEditor()

Dim idV As String
Dim serv As String
Dim idS As String
Dim top As Integer
top = contarTabla("Vehic", 1)

For i = 2 To top
idV = Worksheets("Vehic").Cells(i, 8)
serv = Worksheets("Vehic").Cells(i, 8)
idS = Worksheets("servicios_vehiculos").Cells(buscarId(serv, "Nombre", "servicios_vehiculos"), 3)
Worksheets("Data").Cells(i, 1) = idV
Worksheets("Data").Cells(i, 2) = idS
Next i

End Sub

Sub relacionUbicacion(tabla As String)

Dim top As Integer
top = contarTabla(tabla, 1)

Dim idO As String
Dim idU As String

Dim colid As Integer
Dim colub As Integer
Dim fila As Integer
fila = 2
colid = columnaCampo(tabla, "id")
colub = columnaCampo(tabla, "Establecimiento")

For i = 2 To top

idO = Worksheets(tabla).Cells(i, colid)
If Worksheets(tabla).Cells(i, colub) <> "" Then
idU = Worksheets("localizaciones").Cells(buscarId(Worksheets(tabla).Cells(i, colub), "Nombre", "localizaciones"), 6)
Worksheets("Data").Cells(fila, 1) = idU
Worksheets("Data").Cells(fila, 2) = idO
fila = fila + 1
End If

Next i

End Sub


Sub hacerUbicacion()

'Felipe modifico esta macro


Call relacionUbicacion("Briga")

End Sub

