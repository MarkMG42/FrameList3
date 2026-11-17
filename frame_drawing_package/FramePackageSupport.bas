Attribute VB_Name = "FramePackageSupport"
Option Explicit

Public Function HasItems(ByVal values As Variant) As Boolean
    On Error GoTo EmptyArray
    If IsArray(values) Then HasItems = (UBound(values) >= LBound(values))
EmptyArray:
End Function

Public Sub Require(ByVal condition As Boolean, ByVal message As String)
    If condition = False Then Err.Raise vbObjectError + 3400, "FrameDrawingPackage", message
End Sub

Public Function SameModel(ByVal host As SldWorks.SldWorks, ByVal first As SldWorks.ModelDoc2, _
                          ByVal second As SldWorks.ModelDoc2) As Boolean
    If first Is Nothing Then Exit Function
    If second Is Nothing Then Exit Function
    If first Is second Then SameModel = True: Exit Function
    SameModel = (host.IsSame(first, second) = swObjectSame)
End Function

Public Sub ActivateDocument(ByVal host As SldWorks.SldWorks, ByVal model As SldWorks.ModelDoc2)
    Dim active As SldWorks.ModelDoc2, errors As Long
    Require (model Is Nothing) = False, "Нет документа для активации."
    Set active = host.ActiveDoc
    If SameModel(host, active, model) Then Exit Sub
    Set active = host.ActivateDoc3(model.GetTitle, False, swDontRebuildActiveDoc, errors)
    Set active = host.ActiveDoc
    Require SameModel(host, active, model), "Не активирован документ: " & model.GetTitle & "; код=" & errors
End Sub

Public Sub ActivateSheet(ByVal host As SldWorks.SldWorks, ByVal model As SldWorks.ModelDoc2, ByVal sheetName As String)
    Dim drawing As SldWorks.DrawingDoc, sheet As SldWorks.Sheet
    ActivateDocument host, model
    Set drawing = model
    Require drawing.ActivateSheet(sheetName), "Не активирован лист: " & sheetName
    Set sheet = drawing.GetCurrentSheet
    Require (sheet Is Nothing) = False, "Нет текущего листа."
    Require StrComp(sheet.GetName, sheetName, vbBinaryCompare) = 0, "Не подтверждён текущий лист."
    drawing.EditSheet
    model.ClearSelection2 True
End Sub

Public Function DrawingComponentFor(ByVal view As SldWorks.View, ByVal requested As SldWorks.Component2) As SldWorks.Component2
    Dim values As Variant, item As SldWorks.DrawingComponent, candidate As SldWorks.Component2
    Dim root As SldWorks.DrawingComponent, rootComponent As SldWorks.Component2
    Dim rootName As String, rootAlias As String, i As Long, matches As Long, sameName As Boolean
    Set root = view.RootDrawingComponent
    If Not root Is Nothing Then
        rootName = root.Name
        Set rootComponent = root.Component
        If Not rootComponent Is Nothing Then rootAlias = rootComponent.Name2
    End If
    values = view.GetVisibleDrawingComponents
    Require HasItems(values), "Вид не вернул видимые компоненты. Проверьте конфигурацию и видимость рамы."
    For i = LBound(values) To UBound(values)
        Set item = values(i)
        Set candidate = item.Component
        If Not candidate Is Nothing Then
            sameName = (StrComp(candidate.Name2, requested.Name2, vbTextCompare) = 0)
            If Len(rootName) > 0 Then sameName = sameName Or _
                (StrComp(candidate.Name2, rootName & "/" & requested.Name2, vbTextCompare) = 0)
            If Len(rootAlias) > 0 Then sameName = sameName Or _
                (StrComp(candidate.Name2, rootAlias & "/" & requested.Name2, vbTextCompare) = 0)
            If sameName Then
                If StrComp(candidate.GetPathName, requested.GetPathName, vbTextCompare) = 0 Then
                    If StrComp(candidate.ReferencedConfiguration, requested.ReferencedConfiguration, vbTextCompare) = 0 Then
                        matches = matches + 1
                        Set DrawingComponentFor = candidate
                    End If
                End If
            End If
        End If
    Next i
    Require matches = 1, "Экземпляр рамы не найден однозначно в виде: " & requested.Name2 & "; совпадений=" & matches
End Function

Public Function StandardView(ByVal model As SldWorks.ModelDoc2, ByVal modelPath As String, _
                             ByVal config As String, ByVal direction As String, ByVal viewName As String) As SldWorks.View
    Dim drawing As SldWorks.DrawingDoc, view As SldWorks.View, aliases As Variant, alias As Variant
    Set drawing = model
    Select Case direction
        Case "ISO": aliases = Array("*Isometric", "*Изометрия")
        Case "FRONT": aliases = Array("*Front", "*Спереди")
        Case "TOP": aliases = Array("*Top", "*Сверху")
        Case Else: Require False, "Неизвестная ориентация: " & direction
    End Select
    For Each alias In aliases
        Set view = drawing.CreateDrawViewFromModelView3(modelPath, CStr(alias), 0.1, 0.1, 0#)
        If Not view Is Nothing Then Exit For
    Next alias
    Require (view Is Nothing) = False, "Не создан вид " & direction & ": " & modelPath
    view.ReferencedConfiguration = config
    Require StrComp(view.ReferencedConfiguration, config, vbTextCompare) = 0, "Не применена конфигурация вида: " & config
    Require view.SetName2(viewName), "Не присвоено имя вида: " & viewName
    view.UseSheetScale = False
    model.EditRebuild3
    Require (view.ReferencedDocument Is Nothing) = False, "Вид не загрузил связанную модель."
    Set StandardView = view
End Function

Public Sub FitView(ByVal model As SldWorks.ModelDoc2, ByVal view As SldWorks.View, _
                   ByVal left As Double, ByVal bottom As Double, ByVal right As Double, ByVal top As Double)
    Dim box As Variant, position As Variant, width As Double, height As Double, factor As Double, factorY As Double
    Require right > left And top > bottom, "Нет области для размещения вида."
    box = ViewBox(view)
    width = box(2) - box(0): height = box(3) - box(1)
    Require width > 0# And height > 0#, "Вырожденный габарит вида."
    factor = (right - left) / width: factorY = (top - bottom) / height
    If factorY < factor Then factor = factorY
    view.UseSheetScale = False
    view.ScaleDecimal = view.ScaleDecimal * factor
    model.EditRebuild3
    box = ViewBox(view): position = view.Position
    Require HasItems(position), "Не прочитано положение вида."
    position(0) = position(0) + (left + right - box(0) - box(2)) / 2#
    position(1) = position(1) + (bottom + top - box(1) - box(3)) / 2#
    view.Position = position
    model.EditRebuild3
    box = ViewBox(view)
    Require box(0) >= left - 0.001 And box(2) <= right + 0.001 And _
            box(1) >= bottom - 0.001 And box(3) <= top + 0.001, "Вид не поместился в выделенную область."
End Sub

Public Sub FitPair(ByVal model As SldWorks.ModelDoc2, ByVal front As SldWorks.View, ByVal top As SldWorks.View, _
                   ByVal width As Double, ByVal height As Double)
    Dim a As Variant, b As Variant, k As Double, current As Double, candidates As Variant, value As Variant
    Dim half As Double, usableWidth As Double, usableHeight As Double
    half = width / 2#: usableWidth = half - 0.065: usableHeight = height - 0.09
    Require usableWidth > 0.04 And usableHeight > 0.06, "Формат листа слишком мал для двух видов и размеров."
    front.UseSheetScale = False: top.UseSheetScale = False
    front.ScaleDecimal = 1#: top.ScaleDecimal = 1#
    model.EditRebuild3
    a = ViewBox(front): b = ViewBox(top)
    Require a(2) > a(0) And a(3) > a(1) And b(2) > b(0) And b(3) > b(1), "Нет габаритов видов."
    k = usableWidth / (a(2) - a(0))
    candidates = Array(usableHeight / (a(3) - a(1)), usableWidth / (b(2) - b(0)), usableHeight / (b(3) - b(1)))
    For Each value In candidates
        If CDbl(value) < k Then k = CDbl(value)
    Next value
    front.ScaleDecimal = k: top.ScaleDecimal = k
    model.EditRebuild3
    PlaceLowerLeft model, front, 0.025, 0.05
    PlaceLowerLeft model, top, half + 0.045, 0.05
    Require Abs(front.ScaleDecimal - top.ScaleDecimal) < 0.000000001, "Различаются масштабы переднего и верхнего видов."
End Sub

Private Sub PlaceLowerLeft(ByVal model As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal x As Double, ByVal y As Double)
    Dim box As Variant, position As Variant
    box = ViewBox(view): position = view.Position
    Require HasItems(position), "Нет положения вида."
    position(0) = position(0) + x - box(0): position(1) = position(1) + y - box(1)
    view.Position = position
    model.EditRebuild3
    box = ViewBox(view)
    Require Abs(box(0) - x) < 0.001 And Abs(box(1) - y) < 0.001, "Не подтверждено размещение вида."
End Sub

Public Function ViewBox(ByVal view As SldWorks.View) As Variant
    Dim value As Variant
    value = view.GetOutline
    Require HasItems(value), "Вид не вернул габарит."
    Require LBound(value) = 0 And UBound(value) >= 3, "Неполный габарит вида."
    ViewBox = value
End Function

Public Sub AuditCurrentSheet(ByVal model As SldWorks.ModelDoc2, ByVal audit As Collection)
    Dim drawing As SldWorks.DrawingDoc, sheetView As SldWorks.View, view As SldWorks.View
    Dim display As SldWorks.DisplayDimension, dimension As SldWorks.Dimension, record As Variant
    Set drawing = model
    Set sheetView = drawing.GetFirstView
    Require (sheetView Is Nothing) = False, "Не прочитан листовой вид."
    Set view = sheetView.GetNextView
    Do While Not view Is Nothing
        Set display = view.GetFirstDisplayDimension5
        Do While Not display Is Nothing
            Set dimension = display.GetDimension2(0)
            Require (dimension Is Nothing) = False, "Не прочитано значение созданного размера."
            ReDim record(0 To 1)
            Set record(0) = display
            record(1) = dimension.SystemValue
            audit.Add record
            Set display = display.GetNext5
        Loop
        Set view = view.GetNextView
    Loop
End Sub

Public Sub ValidateAudit(ByVal audit As Collection)
    Dim record As Variant, display As SldWorks.DisplayDimension, dimension As SldWorks.Dimension
    Dim annotation As SldWorks.Annotation
    For Each record In audit
        Set display = record(0)
        Set dimension = display.GetDimension2(0)
        Require (dimension Is Nothing) = False, "После перестроения не прочитан размер."
        Require Abs(dimension.SystemValue - CDbl(record(1))) <= 0.000001, "После восстановления конфигураций изменилось значение размера."
        Set annotation = display.GetAnnotation
        Require (annotation Is Nothing) = False, "Нет аннотации размера."
        Require annotation.IsDangling = False, "После восстановления конфигураций размер потерял привязку."
    Next record
End Sub
