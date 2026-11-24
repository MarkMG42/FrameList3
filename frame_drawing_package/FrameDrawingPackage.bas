Attribute VB_Name = "FrameDrawingPackage"
Option Explicit

Private Const VERSION As String = "FrameDrawingPackage P2.11"
Private Const FRAME_TOKEN As String = "рама"
Private Type FrameTask
    AssemblyConfig As String
    InstanceName As String
    PartPath As String
    PartConfig As String
    DisplayState As String
End Type
Private tasks() As FrameTask
Private taskCount As Long
Private host As SldWorks.SldWorks
Private assemblyModel As SldWorks.ModelDoc2
Private drawingModel As SldWorks.ModelDoc2
Private originalAssemblyConfig As String
Private originalDisplayState As String
Private overviewDisplayState As String
Private originalPartConfigs As Object
Private partModels As Collection
Private warnings As Object
Private incompleteOutput As Boolean
Private audit As Collection
Private pinnedViews As Collection
Private stage As String
Private completed As Long

Public Sub main()
    On Error GoTo Failed
    Dim selectedStates As Collection, stateName As Variant, template As String
    Dim firstSheet As SldWorks.Sheet, drawing As SldWorks.DrawingDoc, names As Variant
    Dim i As Long, failure As String, restored As String, resultState As String
    ResetState
    Set host = Application.SldWorks
    Set assemblyModel = host.ActiveDoc
    If Not assemblyModel Is Nothing Then
        If assemblyModel.GetType = swDocDRAWING Then
            RepairDrawingView host, assemblyModel
            ClearReferences
            Exit Sub
        End If
    End If
    Call FramePackageSupport.Require((assemblyModel Is Nothing) = False, "Откройте сборку рамы.")
    Call FramePackageSupport.Require(assemblyModel.GetType = swDocASSEMBLY, "Запускайте макрос из сборки.")
    Call FramePackageSupport.Require(Len(assemblyModel.GetPathName) > 0, "Сначала сохраните сборку.")
    originalAssemblyConfig = assemblyModel.ConfigurationManager.ActiveConfiguration.Name
    originalDisplayState = ActiveDisplayState()
    CaptureLoadedPartConfigurations
    overviewDisplayState = ChooseOverviewState()
    If Len(overviewDisplayState) = 0 Then ClearReferences: Exit Sub
    Set selectedStates = ChooseFrameStates()
    If selectedStates Is Nothing Then ClearReferences: Exit Sub
    stage = "Поиск рам в состояниях отображения"
    For Each stateName In selectedStates
        SetAssemblyState originalAssemblyConfig, CStr(stateName)
        DiscoverFrames originalAssemblyConfig, CStr(stateName)
    Next stateName
    Call FramePackageSupport.Require(taskCount > 0, "Не найдены детали рам в выбранных состояниях отображения.")
    stage = "Шаблон нового чертежа"
    template = DrawingTemplatePath()
    If Len(template) = 0 Then
        restored = RestoreSourceState()
        If Len(restored) > 0 Then Err.Raise vbObjectError + 3501, VERSION, restored
        ClearReferences
        Exit Sub
    End If
    Set drawingModel = NewFrameDrawing(template)
    Call FramePackageSupport.Require((drawingModel Is Nothing) = False, "Не создан чертёж по шаблону: " & template)
    Call FramePackageSupport.Require(drawingModel.GetType = swDocDRAWING, "Шаблон не создал документ чертежа.")
    Set drawing = drawingModel
    names = drawing.GetSheetNames
    Call FramePackageSupport.Require(FramePackageSupport.HasItems(names), "В новом чертеже нет листа.")
    Call FramePackageSupport.Require(UBound(names) - LBound(names) + 1 = 1, "Шаблон должен содержать ровно один исходный лист.")
    Set firstSheet = drawing.GetCurrentSheet
    Dim sheetProps As Variant, sheetFormat As String, width As Double, height As Double
    sheetProps = firstSheet.GetProperties: sheetFormat = firstSheet.GetTemplateName
    Call FramePackageSupport.Require(FramePackageSupport.HasItems(sheetProps), "Не прочитаны свойства листа шаблона.")
    Call FramePackageSupport.Require(LBound(sheetProps) = 0 And UBound(sheetProps) >= 6, "Неполные свойства листа шаблона.")
    If Len(sheetFormat) > 0 Or CLng(sheetProps(1)) <> swDwgTemplateNone Then
        sheetFormat = ResolveTemplateFile(sheetFormat, ".slddrt", _
            host.GetUserPreferenceStringValue(swFileLocationsSheetFormat) & ";" & _
            host.GetUserPreferenceStringValue(swFileLocationsDocumentTemplates))
    End If
    firstSheet.GetSize width, height
    Call FramePackageSupport.Require(width >= 0.419 And height >= 0.296, "Для комплекта нужен формат не меньше A3 в альбомной ориентации.")
    FramePackageSupport.HideConstructionGeometry drawingModel
    BuildOverview firstSheet, width, height
    For i = 1 To taskCount
        BuildTask i, sheetProps, sheetFormat, width, height
        completed = completed + 1
    Next i
    stage = "Восстановление исходных конфигураций"
    restored = RestoreSourceState()
    Call FramePackageSupport.Require(Len(restored) = 0, restored)
    FramePackageSupport.ActivateDocument host, drawingModel
    drawingModel.EditRebuild3
    stage = "Проверка размеров после восстановления конфигураций"
    FramePackageSupport.ValidateAudit audit
    ValidatePinnedViews
    names = drawing.GetSheetNames
    Call FramePackageSupport.Require(UBound(names) - LBound(names) + 1 = 1 + 3 * taskCount, "Число листов не совпадает с числом рам.")
    drawingModel.GraphicsRedraw2
    If warnings.Count > 0 Then PrintSheetIssues
    If Not incompleteOutput Then MsgBox "Успешно", vbInformation, VERSION
    ClearReferences
    Exit Sub
Failed:
    failure = CStr(Err.Number) & ": " & Err.Description
    restored = RestoreSourceState()
    ActivateResultForReview
    Debug.Print VERSION & "; stage=" & stage & "; " & failure & "; restore=" & restored
    If drawingModel Is Nothing Then
        resultState = "Новый чертёж ещё не создан."
    Else
        resultState = "Созданный чертёж оставлен для проверки; автоматически не сохранялся."
    End If
    MsgBox "Комплект остановлен. Этап: " & stage & vbCrLf & failure & vbCrLf & _
        "Завершено рам: " & completed & ". " & resultState & _
        IIf(Len(restored) > 0, vbCrLf & restored, ""), vbExclamation, VERSION
    ClearReferences
End Sub

Private Sub RepairDrawingView(ByVal host As SldWorks.SldWorks, ByVal model As SldWorks.ModelDoc2)
    On Error GoTo Failed
    Dim selection As SldWorks.SelectionMgr, view As SldWorks.View, report As String
    Set selection = model.SelectionManager
    Call FramePackageSupport.Require(selection.GetSelectedObjectCount2(-1) = 1, "Выберите один сверловочный вид DRILL_... на чертеже.")
    Call FramePackageSupport.Require(selection.GetSelectedObjectType3(1, -1) = swSelDRAWINGVIEWS, "Нужно выбрать сам сверловочный вид, щёлкнув его рамку.")
    Set view = selection.GetSelectedObject6(1, -1)
    report = InsertWeldWithHoles.RepairSelectedBreakForPackage(host, model, view)
    If Len(report) = 0 Then
        MsgBox "Успешно", vbInformation, VERSION
    Else
        MsgBox report, vbExclamation, VERSION
    End If
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, VERSION
End Sub

Private Sub ActivateResultForReview()
    On Error Resume Next
    If drawingModel Is Nothing Then Exit Sub
    FramePackageSupport.ActivateDocument host, drawingModel
End Sub

Private Sub ResetState()
    ClearReferences
    Erase tasks: taskCount = 0: completed = 0: stage = "Инициализация": originalAssemblyConfig = "": originalDisplayState = "": overviewDisplayState = ""
    Set originalPartConfigs = CreateObject("Scripting.Dictionary")
    originalPartConfigs.CompareMode = vbTextCompare
    incompleteOutput = False
    Set partModels = New Collection: Set warnings = CreateObject("Scripting.Dictionary"): Set audit = New Collection: Set pinnedViews = New Collection
End Sub

Private Sub ClearReferences()
    Set host = Nothing: Set assemblyModel = Nothing: Set drawingModel = Nothing
    Set originalPartConfigs = Nothing: Set partModels = Nothing: Set warnings = Nothing: Set audit = Nothing: Set pinnedViews = Nothing
End Sub

Private Sub CaptureLoadedPartConfigurations()
    Dim documents As Variant, model As SldWorks.ModelDoc2, i As Long, key As String
    documents = host.GetDocuments
    If FramePackageSupport.HasItems(documents) = False Then Exit Sub
    For i = LBound(documents) To UBound(documents)
        Set model = documents(i)
        If Not model Is Nothing Then
            If model.GetType = swDocPART Then
                key = model.GetPathName
                If Len(key) > 0 Then
                    If originalPartConfigs.Exists(key) = False Then
                        originalPartConfigs.Add key, model.ConfigurationManager.ActiveConfiguration.Name
                        partModels.Add model
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Function ActiveDisplayState() As String
    Dim values As Variant
    values = assemblyModel.ConfigurationManager.ActiveConfiguration.GetDisplayStates
    Call FramePackageSupport.Require(FramePackageSupport.HasItems(values), "Нет состояний отображения сборки.")
    ' SolidWorks documents the first element as the active display state.
    ActiveDisplayState = CStr(values(LBound(values)))
End Function

Private Function DisplayStateNames() As Variant
    Dim values As Variant
    values = assemblyModel.ConfigurationManager.ActiveConfiguration.GetDisplayStates
    Call FramePackageSupport.Require(FramePackageSupport.HasItems(values), "Нет состояний отображения.")
    DisplayStateNames = values
End Function

Private Function ChooseOverviewState() As String
    Dim values As Variant, text As String, answer As String, i As Long, choice As Long, suggested As String
    values = DisplayStateNames(): suggested = "1"
    text = "Выберите состояние для ПЕРВОГО общего листа: вся сборка, без размеров." & vbCrLf
    For i = LBound(values) To UBound(values)
        text = text & CStr(i - LBound(values) + 1) & ". " & CStr(values(i)) & vbCrLf
        If InStr(1, CStr(values(i)), "по умолчанию", vbTextCompare) > 0 Or _
           InStr(1, CStr(values(i)), "default", vbTextCompare) > 0 Then suggested = CStr(i - LBound(values) + 1)
    Next i
    answer = Trim$(InputBox(text, VERSION, suggested))
    If Len(answer) = 0 Then Exit Function
    Call FramePackageSupport.Require(IsNumeric(answer), "Нужен номер общего состояния.")
    Call FramePackageSupport.Require(CDbl(answer) = Fix(CDbl(answer)), "Номер должен быть целым.")
    choice = CLng(answer) + LBound(values) - 1
    Call FramePackageSupport.Require(choice >= LBound(values) And choice <= UBound(values), "Номер вне списка состояний.")
    ChooseOverviewState = CStr(values(choice))
End Function

Private Function ChooseFrameStates() As Collection
    Dim values As Variant, text As String, answer As String, suggested As String, pieces As Variant, token As Variant
    Dim result As Collection, seen As Object, i As Long, choice As Long, name As String
    values = DisplayStateNames()
    text = "Выберите состояния ОТДЕЛЬНЫХ рам: номера через запятую; * — все, кроме общего." & vbCrLf & _
        "В каждом должна быть видна ровно одна деталь рамы .sldprt с её оснащением." & vbCrLf
    For i = LBound(values) To UBound(values)
        name = CStr(values(i))
        text = text & CStr(i - LBound(values) + 1) & ". " & name & vbCrLf
        If StrComp(name, overviewDisplayState, vbTextCompare) <> 0 Then
            If InStr(1, name, FRAME_TOKEN, vbTextCompare) > 0 Or InStr(1, name, "frame", vbTextCompare) > 0 Then
                If Len(suggested) > 0 Then suggested = suggested & ","
                suggested = suggested & CStr(i - LBound(values) + 1)
            End If
        End If
    Next i
    answer = Trim$(InputBox(text, VERSION, suggested))
    If Len(answer) = 0 Then Exit Function
    Set result = New Collection: Set seen = CreateObject("Scripting.Dictionary")
    If answer = "*" Then
        For i = LBound(values) To UBound(values)
            If StrComp(CStr(values(i)), overviewDisplayState, vbTextCompare) <> 0 Then result.Add CStr(values(i))
        Next i
    Else
        pieces = Split(Replace(answer, ";", ","), ",")
        For Each token In pieces
            token = Trim$(CStr(token))
            Call FramePackageSupport.Require(IsNumeric(token), "Нужен номер состояния отображения.")
            Call FramePackageSupport.Require(CDbl(token) = Fix(CDbl(token)), "Номер должен быть целым.")
            choice = CLng(token) + LBound(values) - 1
            Call FramePackageSupport.Require(choice >= LBound(values) And choice <= UBound(values), "Номер состояния вне списка.")
            name = CStr(values(choice))
            If seen.Exists(name) = False Then seen.Add name, True: result.Add name
        Next token
    End If
    Call FramePackageSupport.Require(result.Count > 0, "Не выбраны отдельные рамы.")
    Set ChooseFrameStates = result
End Function

Private Sub SetAssemblyState(ByVal config As String, ByVal stateName As String)
    SetAssemblyConfiguration config
    Call FramePackageSupport.Require(assemblyModel.ConfigurationManager.ActiveConfiguration.ApplyDisplayState(stateName), _
        "Не применено состояние отображения: " & stateName)
    assemblyModel.EditRebuild3
    Call FramePackageSupport.Require(StrComp(ActiveDisplayState(), stateName, vbTextCompare) = 0, _
        "Не подтверждено состояние отображения: " & stateName)
End Sub

Private Sub SetAssemblyConfiguration(ByVal config As String)
    FramePackageSupport.ActivateDocument host, assemblyModel
    If StrComp(assemblyModel.ConfigurationManager.ActiveConfiguration.Name, config, vbTextCompare) <> 0 Then
        Call FramePackageSupport.Require(assemblyModel.ShowConfiguration2(config), "Не включена конфигурация сборки: " & config)
    End If
    Call FramePackageSupport.Require(StrComp(assemblyModel.ConfigurationManager.ActiveConfiguration.Name, config, vbTextCompare) = 0, _
        "Не подтверждена конфигурация сборки: " & config)
    assemblyModel.EditRebuild3
End Sub

Private Sub DiscoverFrames(ByVal config As String, ByVal stateName As String)
    Dim assembly As SldWorks.AssemblyDoc, values As Variant, component As SldWorks.Component2
    Dim model As SldWorks.ModelDoc2, i As Long, before As Long, path As String
    Set assembly = assemblyModel
    values = assembly.GetComponents(False)
    If FramePackageSupport.HasItems(values) = False Then
        Call FramePackageSupport.Require(False, "В состоянии нет компонентов: " & stateName)
    End If
    before = taskCount
    For i = LBound(values) To UBound(values)
        Set component = values(i)
        path = component.GetPathName
        If FrameName(path) And IsEffectivelyVisible(component) Then
            Set model = component.GetModelDoc2
            Call FramePackageSupport.Require((model Is Nothing) = False, "Полностью загрузите компонент рамы (Resolve): " & component.Name2)
            If model.GetType = swDocPART Then
                taskCount = taskCount + 1
                If taskCount = 1 Then ReDim tasks(1 To 16)
                If taskCount > UBound(tasks) Then ReDim Preserve tasks(1 To UBound(tasks) + 16)
                tasks(taskCount).DisplayState = stateName
                tasks(taskCount).AssemblyConfig = config
                tasks(taskCount).InstanceName = component.Name2
                tasks(taskCount).PartPath = path
                tasks(taskCount).PartConfig = component.ReferencedConfiguration
            End If
        End If
    Next i
    Call FramePackageSupport.Require(taskCount - before = 1, "Состояние [" & stateName & "] должно показывать ровно одну деталь рамы; найдено: " & CStr(taskCount - before))
End Sub

Private Function IsEffectivelyVisible(ByVal component As SldWorks.Component2) As Boolean
    Dim node As SldWorks.Component2, count As Long
    Set node = component
    Do While Not node Is Nothing
        If node.IsSuppressed Then Exit Function
        If node.IsHidden(True) Then Exit Function
        count = count + 1
        Call FramePackageSupport.Require(count < 1000, "Цикл в дереве компонентов.")
        Set node = node.GetParent
    Loop
    IsEffectivelyVisible = True
End Function

Private Function FrameName(ByVal path As String) As Boolean
    Dim stem As String, at As Long
    If LCase$(Right$(path, 7)) <> ".sldprt" Then Exit Function
    stem = Replace(path, "/", "\")
    at = InStrRev(stem, "\")
    If at > 0 Then stem = Mid$(stem, at + 1)
    FrameName = (InStr(1, stem, FRAME_TOKEN, vbTextCompare) > 0 Or InStr(1, stem, "frame", vbTextCompare) > 0)
End Function

Private Function ResolveTask(ByVal index As Long) As SldWorks.Component2
    Dim assembly As SldWorks.AssemblyDoc, values As Variant, candidate As SldWorks.Component2, i As Long, count As Long
    Set assembly = assemblyModel
    values = assembly.GetComponents(False)
    Call FramePackageSupport.Require(FramePackageSupport.HasItems(values), "Нет компонентов текущей конфигурации.")
    For i = LBound(values) To UBound(values)
        Set candidate = values(i)
        If StrComp(candidate.Name2, tasks(index).InstanceName, vbTextCompare) = 0 Then
            If StrComp(candidate.GetPathName, tasks(index).PartPath, vbTextCompare) = 0 Then
                If StrComp(candidate.ReferencedConfiguration, tasks(index).PartConfig, vbTextCompare) = 0 Then
                    count = count + 1: Set ResolveTask = candidate
                End If
            End If
        End If
    Next i
    Call FramePackageSupport.Require(count = 1, "Экземпляр рамы изменился: " & tasks(index).InstanceName)
End Function

Private Sub ActivatePartConfig(ByVal model As SldWorks.ModelDoc2, ByVal config As String)
    Dim key As String, active As String
    key = model.GetPathName
    active = model.ConfigurationManager.ActiveConfiguration.Name
    If originalPartConfigs.Exists(key) = False Then
        originalPartConfigs.Add key, active
        partModels.Add model
    End If
    If StrComp(active, config, vbTextCompare) <> 0 Then
        Call FramePackageSupport.Require(model.ShowConfiguration2(config), "Не включена конфигурация детали: " & config)
    End If
    Call FramePackageSupport.Require(StrComp(model.ConfigurationManager.ActiveConfiguration.Name, config, vbTextCompare) = 0, _
        "Не подтверждена конфигурация детали: " & config)
End Sub

Private Sub PinView(ByVal view As SldWorks.View, ByVal stateName As String)
    Dim record As Variant
    FramePackageSupport.ApplyViewDisplayState drawingModel, view, stateName
    ReDim record(0 To 1)
    Set record(0) = view: record(1) = stateName
    pinnedViews.Add record
End Sub

Private Sub ValidatePinnedViews()
    Dim record As Variant, view As SldWorks.View
    For Each record In pinnedViews
        Set view = record(0)
        Call FramePackageSupport.Require(StrComp(view.DisplayState, CStr(record(1)), vbTextCompare) = 0, _
            "Изменилось состояние отображения вида: " & view.GetName2)
    Next record
End Sub

Private Sub BuildOverview(ByVal firstSheet As SldWorks.Sheet, ByVal width As Double, ByVal height As Double)
    Dim view As SldWorks.View
    stage = "Общий вид всей сборки"
    SetAssemblyState originalAssemblyConfig, overviewDisplayState
    FramePackageSupport.ActivateDocument host, drawingModel
    firstSheet.SetName "Общий_вид"
    FramePackageSupport.ActivateSheet host, drawingModel, "Общий_вид"
    Set view = FramePackageSupport.StandardView(drawingModel, assemblyModel.GetPathName, originalAssemblyConfig, "ISO", "OVERVIEW")
    PinView view, overviewDisplayState
    PlaceFixedScale view, 12#, width, height
    CheckFixedView "Общий_вид", FramePackageSupport.ViewBox(view), width, height
End Sub

Private Sub BuildTask(ByVal index As Long, ByVal props As Variant, _
                      ByVal sheetFormat As String, ByVal width As Double, ByVal height As Double)
    Dim frame As SldWorks.Component2, part As SldWorks.ModelDoc2, drawing As SldWorks.DrawingDoc
    Dim iso As SldWorks.View, drilling As SldWorks.View, front As SldWorks.View, top As SldWorks.View
    Dim drillingIncomplete As Boolean
    Dim sheetName As String, prefix As String, sheetBase As String, warning As String, count As Long, dimensions As Long
    prefix = "F" & Format$(index, "000")
    sheetBase = prefix & "_" & FrameHint(tasks(index).PartPath)
    stage = prefix & ": конфигурация " & tasks(index).AssemblyConfig
    SetAssemblyState tasks(index).AssemblyConfig, tasks(index).DisplayState
    Set frame = ResolveTask(index)
    Set part = frame.GetModelDoc2
    Call FramePackageSupport.Require((part Is Nothing) = False, "Не загружена деталь рамы.")
    ActivatePartConfig part, tasks(index).PartConfig
    Set drawing = drawingModel
    sheetName = sheetBase & "_Общий"
    stage = sheetName & ": сборка и длины"
    FramePackageSupport.ActivateDocument host, drawingModel
    AddSheet sheetName, props, sheetFormat, width, height
    Set iso = FramePackageSupport.StandardView(drawingModel, assemblyModel.GetPathName, tasks(index).AssemblyConfig, "ISO", prefix & "_ISO")
    PinView iso, tasks(index).DisplayState
    iso.ProjectedDimensions = False
    Call FramePackageSupport.Require(iso.ProjectedDimensions = False, "Не включён истинный режим длин на изометрии.")
    PlaceFixedScale iso, 8#, width, height
    warning = FrameCutListDimensionsV3.RunForPackage(host, drawingModel, iso, frame, count)
    If Len(warning) > 0 Then incompleteOutput = True
    NoteSheetIssue sheetName, warning
    CheckFixedView sheetName, FramePackageSupport.AnnotatedViewBox(iso), width, height
    FramePackageSupport.AuditCurrentSheet drawingModel, audit
    sheetName = sheetBase & "_Сверловка"
    stage = sheetName & ": деталь рамы и сверловка"
    AddSheet sheetName, props, sheetFormat, width, height
    Set drilling = FramePackageSupport.StandardView(drawingModel, tasks(index).PartPath, tasks(index).PartConfig, "ISO", prefix & "_DRILL_SOURCE")
    FramePackageSupport.SetSheetScale drawingModel, 1#, 10#
    FramePackageSupport.SetScaleAndPlace drawingModel, drilling, 0.1, 0.03, 0.055
    warning = InsertWeldWithHoles.RunForPackage(host, drawingModel, drilling, count, drillingIncomplete)
    If drillingIncomplete Then incompleteOutput = True
    NoteSheetIssue sheetName, warning
    FramePackageSupport.ActivateSheet host, drawingModel, sheetName
    FramePackageSupport.AuditCurrentSheet drawingModel, audit
    sheetName = sheetBase & "_Монтаж"
    stage = sheetName & ": виды спереди/сверху и монтажные размеры"
    AddSheet sheetName, props, sheetFormat, width, height
    Set front = FramePackageSupport.StandardView(drawingModel, assemblyModel.GetPathName, tasks(index).AssemblyConfig, "FRONT", prefix & "_FRONT")
    Set top = FramePackageSupport.StandardView(drawingModel, assemblyModel.GetPathName, tasks(index).AssemblyConfig, "TOP", prefix & "_TOP")
    PinView front, tasks(index).DisplayState
    PinView top, tasks(index).DisplayState
    FramePackageSupport.MountingDisplay drawingModel, front
    FramePackageSupport.MountingDisplay drawingModel, top
    FramePackageSupport.FitMountingPair host, drawingModel, front, top, frame, width, height
    dimensions = FrameAutoSize.RunForPackage(host, drawingModel, front, frame, "FRONT")
    dimensions = dimensions + FrameAutoSize.RunForPackage(host, drawingModel, top, frame, "TOP")
    FramePackageSupport.CheckMountingPair drawingModel, front, top, width, height
    FramePackageSupport.AddViewCaption drawingModel, front, "Вид спереди"
    FramePackageSupport.AddViewCaption drawingModel, top, "Вид сверху"
    FramePackageSupport.AuditCurrentSheet drawingModel, audit
    Debug.Print prefix & "; " & tasks(index).InstanceName & "; config=" & tasks(index).AssemblyConfig & "; mounting dimensions=" & dimensions
End Sub

Private Function FrameHint(ByVal path As String) As String
    Dim text As String, at As Long
    text = Replace(path, "/", "\")
    at = InStrRev(text, "\")
    If at > 0 Then text = Mid$(text, at + 1)
    If LCase$(Right$(text, 7)) = ".sldprt" Then text = Left$(text, Len(text) - 7)
    at = InStrRev(text, FRAME_TOKEN, -1, vbTextCompare)
    If at = 0 Then at = InStrRev(text, "frame", -1, vbTextCompare)
    If at > 0 Then text = Mid$(text, at)
    FrameHint = Left$(text, 24)
End Function

Private Sub AddSheet(ByVal name As String, ByVal props As Variant, ByVal format As String, ByVal width As Double, ByVal height As Double)
    Dim drawing As SldWorks.DrawingDoc, templateType As Long, formatPath As String
    Set drawing = drawingModel
    FramePackageSupport.ActivateDocument host, drawingModel
    templateType = swDwgTemplateCustom: formatPath = format
    If Len(formatPath) = 0 Then templateType = swDwgTemplateNone
    Dim showDialog As Boolean, failure As String
    showDialog = host.GetUserPreferenceToggle(swDrawingShowSheetFormatDialog)
    On Error GoTo Failed
    If showDialog Then host.SetUserPreferenceToggle swDrawingShowSheetFormatDialog, False
    Call FramePackageSupport.Require(drawing.NewSheet3(name, CLng(props(0)), templateType, CDbl(props(2)), CDbl(props(3)), _
        CBool(props(4)), formatPath, width, height, "Default"), "Не создан лист: " & name)
    If showDialog Then host.SetUserPreferenceToggle swDrawingShowSheetFormatDialog, True
    On Error GoTo 0
    FramePackageSupport.ActivateSheet host, drawingModel, name
    Exit Sub
Failed:
    failure = Err.Description
    On Error Resume Next
    If showDialog Then host.SetUserPreferenceToggle swDrawingShowSheetFormatDialog, True
    On Error GoTo 0
    Err.Raise vbObjectError + 3510, VERSION, failure
End Sub

Private Function NewFrameDrawing(ByVal template As String) As SldWorks.ModelDoc2
    Dim wasShown As Boolean, failure As String
    wasShown = host.GetUserPreferenceToggle(swDrawingShowSheetFormatDialog)
    On Error GoTo Failed
    If wasShown Then host.SetUserPreferenceToggle swDrawingShowSheetFormatDialog, False
    Set NewFrameDrawing = host.NewDocument(template, 0, 0#, 0#)
    If wasShown Then host.SetUserPreferenceToggle swDrawingShowSheetFormatDialog, True
    Exit Function
Failed:
    failure = Err.Description
    On Error Resume Next
    If wasShown Then host.SetUserPreferenceToggle swDrawingShowSheetFormatDialog, True
    On Error GoTo 0
    Err.Raise vbObjectError + 3511, VERSION, failure
End Function

Private Function DrawingTemplatePath() As String
    DrawingTemplatePath = ResolveTemplateFile(host.GetUserPreferenceStringValue(swDefaultTemplateDrawing), _
        ".drwdot", host.GetUserPreferenceStringValue(swFileLocationsDocumentTemplates))
End Function

Private Function ResolveTemplateFile(ByVal saved As String, ByVal ext As String, ByVal roots As String) As String
    Dim fs As Object, matches As Object, folder As Variant, wanted As String, paths As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set matches = CreateObject("Scripting.Dictionary")
    matches.CompareMode = vbTextCompare
    saved = Trim$(saved)
    If LCase$(Right$(saved, Len(ext))) = ext Then
        If fs.FileExists(saved) Then
            ResolveTemplateFile = saved
            Exit Function
        End If
        wanted = fs.GetFileName(saved)
    End If
    For Each folder In Split(Replace(Replace(roots, vbCr, ";"), vbLf, ";"), ";")
        If fs.FolderExists(Trim$(CStr(folder))) Then
            FindTemplateFiles fs.GetFolder(Trim$(CStr(folder))), ext, wanted, matches
            If matches.Count > 0 And Len(wanted) > 0 Then Exit For
        End If
    Next
    Call FramePackageSupport.Require(matches.Count > 0, "Файл " & wanted & " (" & ext & ") не найден. Папки SOLIDWORKS: " & roots)
    Call FramePackageSupport.Require(matches.Count = 1, "Найдено несколько файлов " & wanted & " (" & ext & "). Уточните путь в настройках или свойствах листа.")
    paths = matches.Keys
    ResolveTemplateFile = CStr(paths(0))
End Function

Private Sub FindTemplateFiles(ByVal folder As Object, ByVal ext As String, ByVal wanted As String, ByVal matches As Object)
    Dim item As Object, subfolder As Object
    For Each item In folder.Files
        If LCase$(Right$(item.Name, Len(ext))) = ext Then
            If Len(wanted) = 0 Or StrComp(item.Name, wanted, vbTextCompare) = 0 Then matches(item.Path) = True
        End If
    Next
    For Each subfolder In folder.SubFolders
        If (subfolder.Attributes And 1024) = 0 Then FindTemplateFiles subfolder, ext, wanted, matches
    Next
End Sub

Private Function RestoreSourceState() As String
    On Error GoTo Failed
    Dim model As SldWorks.ModelDoc2, key As String, message As String
    If Not partModels Is Nothing Then
        For Each model In partModels
            key = model.GetPathName
            message = message & RestoreOneConfiguration(model, CStr(originalPartConfigs(key)))
        Next model
    End If
    If Not assemblyModel Is Nothing Then
        If Len(originalAssemblyConfig) > 0 Then message = message & RestoreOneConfiguration(assemblyModel, originalAssemblyConfig)
        If Len(originalDisplayState) > 0 Then
            If assemblyModel.ConfigurationManager.ActiveConfiguration.ApplyDisplayState(originalDisplayState) = False Then
                message = message & "Не восстановлено исходное состояние отображения." & vbCrLf
            ElseIf StrComp(ActiveDisplayState(), originalDisplayState, vbTextCompare) <> 0 Then
                message = message & "Не подтверждено восстановление состояния отображения." & vbCrLf
            End If
        End If
    End If
    RestoreSourceState = message
    Exit Function
Failed:
    RestoreSourceState = message & "Ошибка восстановления: " & Err.Description
End Function

Private Function RestoreOneConfiguration(ByVal model As SldWorks.ModelDoc2, ByVal config As String) As String
    On Error GoTo Failed
    If StrComp(model.ConfigurationManager.ActiveConfiguration.Name, config, vbTextCompare) <> 0 Then
        If model.ShowConfiguration2(config) = False Then
            RestoreOneConfiguration = "Не восстановлена конфигурация: " & model.GetTitle & vbCrLf
            Exit Function
        End If
    End If
    If StrComp(model.ConfigurationManager.ActiveConfiguration.Name, config, vbTextCompare) <> 0 Then _
        RestoreOneConfiguration = "Не подтверждена конфигурация: " & model.GetTitle & vbCrLf
    Exit Function
Failed:
    RestoreOneConfiguration = "Ошибка восстановления конфигурации: " & Err.Description & vbCrLf
End Function

Private Sub PlaceFixedScale(ByVal view As SldWorks.View, ByVal denominator As Double, _
                             ByVal width As Double, ByVal height As Double)
    Dim box As Variant, position As Variant
    view.UseSheetScale = False
    view.ScaleDecimal = 1# / denominator
    drawingModel.EditRebuild3
    box = FramePackageSupport.ViewBox(view)
    position = view.Position
    position(0) = position(0) + (0.03 + width - 0.01 - box(0) - box(2)) / 2#
    position(1) = position(1) + (0.045 + height - 0.015 - box(1) - box(3)) / 2#
    view.Position = position
    drawingModel.EditRebuild3
    Call FramePackageSupport.Require(Abs(view.ScaleDecimal - 1# / denominator) < 0.00000001, _
        "Не подтверждён масштаб вида: " & view.GetName2)
End Sub

Private Sub NoteSheetIssue(ByVal sheetName As String, ByVal detail As String)
    If Len(detail) = 0 Then Exit Sub
    If warnings.Exists(sheetName) Then
        warnings(sheetName) = CStr(warnings(sheetName)) & " | " & Replace(detail, vbCrLf, " | ")
    Else
        warnings.Add sheetName, Replace(detail, vbCrLf, " | ")
    End If
End Sub

Private Sub CheckFixedView(ByVal sheetName As String, ByVal box As Variant, ByVal width As Double, ByVal height As Double)
    If box(0) < 0.029 Or box(1) < 0.044 Or box(2) > width - 0.009 Or box(3) > height - 0.014 Then
        NoteSheetIssue sheetName, "Вид или размеры выходят за рабочую область при заданном масштабе."
    End If
End Sub

Private Sub PrintSheetIssues()
    Dim sheetName As Variant
    Debug.Print "Причины проверки листов"
    For Each sheetName In warnings.Keys
        Debug.Print CStr(sheetName) & ": " & CStr(warnings(sheetName))
    Next sheetName
End Sub
