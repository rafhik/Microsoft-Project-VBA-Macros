Function RecursosEnFecha(Fecha As Date)
Dim T As Task
RecursosEnFecha = 0
If EsLaboral(Fecha) Then
    For Each T In ActiveProject.Tasks
        If T.Start <= Fecha And T.Finish >= Fecha Then
            RecursosEnFecha = RecursosEnFecha + NumRecTask(T)
        Else
            RecursosEnFecha = RecursosEnFecha
        End If
    Next T
End If
End Function
Function DatosFecha(Fecha As Date)

'Debug.Print "Día " & Day(Fecha)
'Debug.Print "Día semana " & Weekday(Fecha, vbMonday)
'Debug.Print "Mes " & Month(Fecha)
'Debug.Print "Año " & Year(Fecha)
'Debug.Print "----------"

End Function
Function NumRecTask(Tar As Task)
Dim Asig As Assignment
A = 0
    For Each Asig In Tar.Assignments
        If Asig.ResourceType = 0 Then
            A = A + Asig.Units
        End If
    Next Asig
    If A = Int(A) Then
        NumRecTask = A
    Else
        NumRecTask = Int(A) + 1
    End If
End Function


Function EsLunes(Fecha As Date) As Boolean
    If Weekday(Fecha, vbMonday) = 1 Then
        EsLunes = True
    Else
        EsLunes = False
    End If
End Function
Function EsDiaUno(Fecha As Date) As Boolean
    If Day(Fecha) = 1 Then
        EsDiaUno = True
    Else
        EsDiaUno = False
    End If
End Function
Function EsDiaUnoEnero(Fecha As Date) As Boolean
    If Day(Fecha) = 1 And Month(Fecha) = 1 Then
        EsDiaUnoEnero = True
    Else
        EsDiaUnoEnero = False
    End If
End Function
Function EsLaboral(Fecha As Date) As Boolean
EsLaboral = ActiveProject.Calendar.Period(Fecha).Working
End Function
Sub Recursos_Periodos_Excel()
Dim SD As Date, DF As Date, TD As Integer, ND As Date, Sema As Integer, Mes As Integer, Año As Integer
Dim RecEnSem As Integer, RecEnMes As Integer, RecEnAño As Integer

Dim listrecday As Object
Set listrecday = CreateObject("System.Collections.ArrayList")
Dim listrecweek As Object
Set listrecweek = CreateObject("System.Collections.ArrayList")
Dim listrecmonth As Object
Set listrecmonth = CreateObject("System.Collections.ArrayList")
Dim listrecyear As Object
Set listrecyear = CreateObject("System.Collections.ArrayList")

SD = ActiveProject.ProjectStart
FD = ActiveProject.ProjectFinish
TD = DateDiff("d", SD, FD)
Sema = 1
Mes = 1
Año = 1
RecEnSem = 0
RecEnMes = 0
RecEnAño = 0
For D = 0 To TD
    ''''Añadimos recurso dia
    ND = DateAdd("d", D, SD)
    listrecday.Add (RecursosEnFecha(ND))
    
    ''''Añadimos recuro  semana
    If EsLunes(ND) Then
        listrecweek.Add (RecEnSem)
        RecEnSem = RecursosEnFecha(ND)
        Sema = Sema + 1
    Else
        If RecursosEnFecha(ND) > RecEnSem Then
            RecEnSem = RecursosEnFecha(ND)
        Else
            RecEnSem = RecEnSem
            Sema = Sema
        End If
    End If
    
    ''''Añadimos recurso mes
    If EsDiaUno(ND) Then
        listrecmonth.Add (RecEnMes)
        RecEnMes = RecursosEnFecha(ND)
        Mes = Mes + 1
    Else
        
        If RecursosEnFecha(ND) > RecEnMes Then
            RecEnMes = RecursosEnFecha(ND)
        Else
            RecEnMes = RecEnMes
            Mes = Mes
        End If
    End If
  
    
    ''''Añadimos recurso año
    If EsDiaUnoEnero(ND) Then
        listrecyear.Add (RecEnAño)
        RecEnAño = RecursosEnFecha(ND)
        Año = Año + 1
    Else
    
        If RecursosEnFecha(ND) > RecEnAño Then
            RecEnAño = RecursosEnFecha(ND)
        Else
            RecEnAño = RecEnAño
            Año = Año
        End If
        
    End If

Next D

'Apertura excel
On Error Resume Next
Set Xl = GetObject(, "Excel.application")
If Err <> 0 Then
    On Error GoTo 0
    Set Xl = CreateObject("Excel.Application")
    If Err <> 0 Then
        MsgBox "Excel no está disponible para este equipo" _
            & vbCr & "Instalar o comprobar", vbCritical, _
            "Notes Text Export - Fatal Error"
        FilterApply Name:="all tasks"
        Set Xl = Nothing
        On Error GoTo 0     'clear error function
        Exit Sub
    End If
End If
On Error GoTo 0
Xl.Workbooks.Add
BookNam = Xl.ActiveWorkbook.Name



'Mantenemos Excel en segundo plano y minimizamos (speeds transfer)
'NOTA: Es necesaria la librería Excel object library
Xl.Visible = False
Xl.ScreenUpdating = False
Xl.DisplayAlerts = False
ActiveWindow.Caption = " Writing data to worksheet"
'Volcado de Datos
Set s = Xl.Workbooks(BookNam).Worksheets(1)
ActiveWindow.Caption = " do it again"

'Cabeceras
s.Range("H1").Value = "Dias"
s.Range("I1").Value = "Recursos_Días"
s.Range("J1").Value = "Semanas"
s.Range("K1").Value = "Recursos_Semanas"
s.Range("L1").Value = "Meses"
s.Range("M1").Value = "Recursos_Meses"
s.Range("N1").Value = "Años"
s.Range("O1").Value = "Recursos_Años"

'Datos
Set C = s.Range("H2")
Contador = 0
For Each ccc In listrecday
    C.Offset(Contador, 0).Value = Contador + 1
    C.Offset(Contador, 1).Value = ccc
    Contador = Contador + 1
Next ccc

Set C = s.Range("J2")
Contador = 0
For Each ccc In listrecweek
    C.Offset(Contador, 0).Value = Contador + 1
    C.Offset(Contador, 1).Value = ccc
    Contador = Contador + 1
Next ccc


Set C = s.Range("L2")
Contador = 0
For Each ccc In listrecmonth
    C.Offset(Contador, 0).Value = Contador + 1
    C.Offset(Contador, 1).Value = ccc
    Contador = Contador + 1
Next ccc

Set C = s.Range("N2")
Contador = 0
For Each ccc In listrecyear
    C.Offset(Contador, 0).Value = Contador + 1
    C.Offset(Contador, 1).Value = ccc
    Contador = Contador + 1
Next ccc

'Formatos
s.Rows(1).Font.Bold = True
s.Columns("H:O").ColumnWidth = 7
s.Range("H:O").HorizontalAlignment = xlCenter 'reference

'Final, cerrar y salir


''''Conversión a tablas
Dim tbl As ListObject, Copyrange As String

Let Copyrange = "H1" & ":" & "I" & listrecday.Count + 1
Set tbl = s.ListObjects.Add(xlSrcRange, s.Range(Copyrange), , xlYes)
tbl.TableStyle = "TableStyleLight9"


Let Copyrange = "J1" & ":" & "K" & listrecweek.Count + 1
Set tbl = s.ListObjects.Add(xlSrcRange, s.Range(Copyrange), , xlYes)
tbl.TableStyle = "TableStyleLight9"

Let Copyrange = "L1" & ":" & "M" & listrecmonth.Count + 1
Set tbl = s.ListObjects.Add(xlSrcRange, s.Range(Copyrange), , xlYes)
tbl.TableStyle = "TableStyleLight9"

Let Copyrange = "N1" & ":" & "O" & listrecyear.Count + 1
Set tbl = s.ListObjects.Add(xlSrcRange, s.Range(Copyrange), , xlYes)
tbl.TableStyle = "TableStyleLight9"

''''Creación de Gráficos
Let Copyrange = "I1" & ":" & "I" & listrecday.Count + 1
s.Range(Copyrange).Select

Set gr1 = s.Shapes.AddChart
With gr1
    .IncrementLeft -180
    .IncrementTop -50
    .Chart.FullSeriesCollection(1).Select
    .Chart.SetElement (msoElementDataLabelOutSideEnd)
    .Chart.ChartTitle.text = "Recursos por Día"
    .Chart.Legend.Delete
    .Chart.ChartType = xlColumnClustered
    .Chart.AutoScaling = False
    .Chart.Axes(xlValue).Select
    .Chart.Axes(xlValue).MinimumScale = 0
    .Chart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    .Chart.Axes(xlCategory, xlPrimary).AxisTitle.text = "Días"
    .Chart.FullSeriesCollection(1).DataLabels.Orientation = 45

End With

Let Copyrange = "K1" & ":" & "K" & listrecweek.Count + 1
s.Range(Copyrange).Select

Set gr2 = s.Shapes.AddChart
With gr2
    .IncrementLeft -180
    .IncrementTop 200
    .Chart.FullSeriesCollection(1).Select
    .Chart.SetElement (msoElementDataLabelOutSideEnd)
    .Chart.ChartTitle.text = "Nº Máx. Recursos por Semana"
    .Chart.Legend.Delete
    .Chart.ChartType = xlColumnClustered
    .Chart.AutoScaling = False
    .Chart.Axes(xlValue).Select
    .Chart.Axes(xlValue).MinimumScale = 0
    .Chart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    .Chart.Axes(xlCategory, xlPrimary).AxisTitle.text = "Semanas"
End With

If listrecmonth.Count > 1 Then
    Let Copyrange = "M1" & ":" & "M" & listrecmonth.Count + 1
    s.Range(Copyrange).Select
    
    Set gr3 = s.Shapes.AddChart
    With gr3
        .IncrementLeft -180
        .IncrementTop 450
        .Chart.FullSeriesCollection(1).Select
        .Chart.SetElement (msoElementDataLabelOutSideEnd)
        .Chart.ChartTitle.text = "Nº Máx. Recursos por Meses"
        .Chart.Legend.Delete
        .Chart.ChartType = xlColumnClustered
        .Chart.AutoScaling = False
        .Chart.Axes(xlValue).Select
        .Chart.Axes(xlValue).MinimumScale = 0
        .Chart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        .Chart.Axes(xlCategory, xlPrimary).AxisTitle.text = "Meses"
    End With
End If

If listrecyear.Count > 1 Then
    Let Copyrange = "O1" & ":" & "O" & listrecyear.Count + 1
    s.Range(Copyrange).Select
    
    
    Set gr4 = s.Shapes.AddChart
    With gr4
        .IncrementLeft -180
        .IncrementTop 700
        .Chart.FullSeriesCollection(1).Select
        .Chart.SetElement (msoElementDataLabelOutSideEnd)
        .Chart.ChartTitle.text = "Nº Máx. Recursos por Año"
        .Chart.Legend.Delete
        .Chart.ChartType = xlColumnClustered
        .Chart.AutoScaling = False
        .Chart.Axes(xlValue).Select
        .Chart.Axes(xlValue).MinimumScale = 0
        .Chart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        .Chart.Axes(xlCategory, xlPrimary).AxisTitle.text = "Años"
    
    End With
End If

MsgBox "Exportación Completada", vbOKOnly, "Exportación"
Application.Caption = ""
ActiveWindow.Caption = ""
Xl.Visible = True
Xl.ScreenUpdating = True
Set Xl = Nothing

End Sub


