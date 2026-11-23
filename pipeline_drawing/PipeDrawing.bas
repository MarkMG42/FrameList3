Attribute VB_Name = "PipeDrawing"
Option Explicit

Private stage As String
Private dimFailure As String
Private Const revision As String = "1.7"

Private Function HasItems(ByVal values As Variant) As Boolean
    On Error GoTo EmptyArray
    If IsArray(values) Then HasItems = UBound(values) >= LBound(values)
EmptyArray:
End Function

Private Function SourceOccurrence(ByVal displayed As SldWorks.Component2, ByRef occurrences As Variant) As SldWorks.Component2
    Dim candidate As SldWorks.Component2, matched As SldWorks.Component2
    Dim item As Variant, fullName As String, modelName As String, bestLength As Long
    fullName = "/" & displayed.Name2
    For Each item In occurrences
        Set candidate = item
        If Not candidate Is Nothing Then
            modelName = "/" & candidate.Name2
            If Len(modelName) >= bestLength And StrComp(Right$(fullName, Len(modelName)), modelName, vbTextCompare) = 0 Then
                If StrComp(candidate.GetPathName, displayed.GetPathName, vbTextCompare) = 0 Then
                    If Len(modelName) = bestLength Then Err.Raise 5, , "Ambiguous component: " & fullName
                    Set matched = candidate
                    bestLength = Len(modelName)
                End If
            End If
        End If
    Next
    Set SourceOccurrence = matched
End Function

Public Sub main()
    On Error GoTo Failed
    Dim sw As SldWorks.SldWorks, asm As SldWorks.ModelDoc2
    Dim doc As SldWorks.ModelDoc2, drw As SldWorks.DrawingDoc
    Dim front As SldWorks.View, iso As SldWorks.View
    Dim sheet As SldWorks.Sheet, w As Double, h As Double
    Dim template As String, path As String, config As String
    stage = "assembly"
    Set sw = Application.SldWorks
    Set asm = sw.ActiveDoc
    If asm Is Nothing Then Err.Raise 5, , "Open the piping assembly."
    If asm.GetType <> swDocASSEMBLY Then Err.Raise 5, , "Select an assembly."
    path = asm.GetPathName
    If Len(path) = 0 Then Err.Raise 5, , "Save the assembly first."
    config = asm.ConfigurationManager.ActiveConfiguration.Name
    template = sw.GetUserPreferenceStringValue(swDefaultTemplateDrawing)
    stage = "drawing template"
    If Len(template) = 0 Then Err.Raise 5, , "Choose a default drawing template in SOLIDWORKS."
    Set doc = sw.NewDocument(template, 0, 0#, 0#)
    If doc Is Nothing Then Err.Raise 5, , "Check the default drawing template."
    Set drw = doc
    stage = "drawing sheet"
    Set sheet = drw.GetCurrentSheet
    If sheet Is Nothing Then Err.Raise 5, , "No active drawing sheet."
    sheet.GetSize w, h
    If w <= 0 Or h <= 0 Then Err.Raise 5, , "Invalid sheet size."
    stage = "front view"
    Set front = InsertStandardView(drw, path, w * 0.68, h * 0.53, _
        Array("*Front", "*" & ChrW(1057) & ChrW(1087) & ChrW(1077) & ChrW(1088) & ChrW(1077) & ChrW(1076) & ChrW(1080)))
    stage = "front configuration"
    front.ReferencedConfiguration = config
    stage = "isometric view"
    Set iso = InsertStandardView(drw, path, w * 0.24, h * 0.53, _
        Array("*Isometric", "*" & ChrW(1048) & ChrW(1079) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1090) & ChrW(1088) & ChrW(1080) & ChrW(1103)))
    stage = "isometric configuration"
    iso.ReferencedConfiguration = config
    stage = "view scales 1:5"
    front.UseSheetScale = False
    front.ScaleDecimal = 0.2
    iso.UseSheetScale = False
    iso.ScaleDecimal = 0.2
    ' A False return means the mode was not reset (including an unchanged mode).
    stage = "view styles"
    Call front.SetDisplayMode4(False, swHIDDEN, False, False, True)
    Call iso.SetDisplayMode4(False, swSHADED, False, True, True)
    stage = "rebuild"
    doc.EditRebuild3
    stage = "front view geometry"
    If StrComp(front.ReferencedConfiguration, config, vbTextCompare) <> 0 Or StrComp(iso.ReferencedConfiguration, config, vbTextCompare) <> 0 Then Err.Raise 5, , "Drawing view configuration mismatch."
    If Abs(front.ScaleDecimal - 0.2) > 0.000001 Or Abs(iso.ScaleDecimal - 0.2) > 0.000001 Then Err.Raise 5, , "Expected both view scales to be 1:5."
    ' SelectRing supplies the target view through SelectData for each edge.
    AddPipeDimensions sw, doc, front, asm
    Exit Sub
Failed:
    MsgBox "PipeDrawing " & revision & " | " & stage & vbCrLf & "Error " & CStr(Err.Number) & ": " & Err.Description, vbExclamation, "PipeDrawing"
End Sub

Private Function InsertStandardView(ByVal drawing As SldWorks.DrawingDoc, ByVal modelPath As String, _
                                    ByVal x As Double, ByVal y As Double, ByVal names As Variant) As SldWorks.View
    Dim orientation As Variant, candidate As SldWorks.View, detail As String
    For Each orientation In names
        On Error Resume Next
        Err.Clear
        Set candidate = drawing.CreateDrawViewFromModelView3(modelPath, CStr(orientation), x, y, 0#)
        If Err.Number <> 0 Then detail = Err.Description
        On Error GoTo 0
        If Not candidate Is Nothing Then
            Set InsertStandardView = candidate
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 201, "PipeDrawing", "Cannot insert view: " & Join(names, " / ") & vbCrLf & modelPath & vbCrLf & detail
End Function

Private Function PartKind(ByVal c As SldWorks.Component2) As String
    Dim n As String
    n = LCase$(Mid$(c.GetPathName, InStrRev(c.GetPathName, "\") + 1))
    If Right$(n, 7) <> ".sldprt" Then Exit Function
    If Left$(n, 5) = (ChrW(1090) & ChrW(1088) & ChrW(1091) & ChrW(1073) & ChrW(1072)) Then PartKind = "pipe"
    If Left$(n, 6) = (ChrW(1092) & ChrW(1083) & ChrW(1072) & ChrW(1085) & ChrW(1077) & ChrW(1094)) Then PartKind = "flange"
    If Left$(n, 5) = (ChrW(1084) & ChrW(1091) & ChrW(1092) & ChrW(1090) & ChrW(1072)) Or Left$(n, 6) = (ChrW(1088) & ChrW(1077) & ChrW(1079) & ChrW(1100) & ChrW(1073) & ChrW(1072)) Then PartKind = "branch"
End Function

Private Sub AddPipeDimensions(ByVal sw As SldWorks.SldWorks, ByVal doc As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal assemblyModel As SldWorks.ModelDoc2)
    Dim parts As Variant, edges As Variant, i As Long, j As Long
    Dim c As SldWorks.Component2, e As SldWorks.Edge, curve As SldWorks.Curve
    Dim mu As SldWorks.MathUtility, pt As SldWorks.MathPoint, p As Variant, xyz As Variant
    Dim rings As New Collection, r As Collection, kind As String
    Dim placement As SldWorks.MathTransform, projection As SldWorks.MathTransform
    Dim sourceAssembly As SldWorks.AssemblyDoc, sourceComponent As SldWorks.Component2
    Dim occurrences As Variant
    Set sourceAssembly = assemblyModel
    occurrences = sourceAssembly.GetComponents(False)
    If Not HasItems(occurrences) Then Err.Raise 5, , "Assembly contains no accessible components."
    Set projection = v.ModelToViewTransform
    If projection Is Nothing Then Err.Raise 5, , "Front view has no model transform."
    Set mu = sw.GetMathUtility
    If mu Is Nothing Then Err.Raise 5, , "Math utility unavailable."
    parts = v.GetVisibleComponents
    If Not HasItems(parts) Then Err.Raise 5, , "No visible components in front view."
    For i = LBound(parts) To UBound(parts)
        Set c = parts(i)
        If c Is Nothing Then GoTo NextComponent
        stage = "front geometry | " & c.Name2
        kind = PartKind(c)
        If Len(kind) > 0 Then
            Set sourceComponent = SourceOccurrence(c, occurrences)
            Set placement = Nothing
            If Not sourceComponent Is Nothing Then Set placement = sourceComponent.Transform2
            If placement Is Nothing Then Set placement = c.Transform2
            edges = v.GetVisibleEntities2(c, swViewEntityType_Edge)
            Dim previousCount As Long
            previousCount = rings.Count
            If HasItems(edges) Then
                For j = LBound(edges) To UBound(edges)
                    Set e = edges(j)
                    If e Is Nothing Then GoTo NextEdge
                    Set curve = e.GetCurve
                    If curve Is Nothing Then GoTo NextEdge
                    If curve.IsCircle Then
                        p = curve.CircleParams
                        Dim coords(2) As Double
                        coords(0) = p(0)
                        coords(1) = p(1)
                        coords(2) = p(2)
                        Set pt = mu.CreatePoint(coords)
                        If placement Is Nothing Then Err.Raise 5, , "No component transform for circular edge."
                        If pt Is Nothing Then Err.Raise 5, , "Invalid edge center."
                        Set pt = pt.MultiplyTransform(placement)
                        If pt Is Nothing Then Err.Raise 5, , "Cannot transform edge center to assembly."
                        Set pt = pt.MultiplyTransform(projection)
                        If pt Is Nothing Then Err.Raise 5, , "Cannot project center into front view."
                        xyz = pt.ArrayData
                        Set r = New Collection
                        r.Add e, "edge"
                        r.Add CDbl(xyz(0)), "x"
                        r.Add CDbl(xyz(1)), "y"
                        r.Add CDbl(xyz(2)), "z"
                        r.Add CDbl(p(6)), "radius"
                        r.Add c.Name2, "component"
                        r.Add kind, "kind"
                        rings.Add r
                    End If
NextEdge:
                Next
            End If
            If rings.Count = previousCount Then Err.Raise 5, , "No visible circular edges on " & c.Name2
        End If
NextComponent:
    Next
    stage = "dimension placement"
    PlaceDimensions doc, v, rings
End Sub

Private Function EndRing(ByVal rings As Collection, ByVal seed As Collection, ByVal axis As String, ByVal direction As Long) As Collection
    Dim r As Collection, best As Collection
    Set best = seed
    For Each r In rings
        If r("component") = seed("component") Then
            If direction * (r(axis) - best(axis)) > 0.000001 Then Set best = r
        End If
    Next
    Set EndRing = best
End Function

Private Function FlangeEnd(ByVal rings As Collection, ByVal pipeEnd As Collection, ByVal axis As String, ByVal cross As String, ByVal viewScale As Double) As Collection
    Dim r As Collection, nearest As Collection, gap As Double, best As Double
    best = 1E+10
    For Each r In rings
        If r("kind") = "flange" Then
            gap = r(axis) - pipeEnd(axis)
            If Abs(r("z") - pipeEnd("z")) < 0.0002 * viewScale And Abs(r(cross) - pipeEnd(cross)) < 0.0002 * viewScale And gap >= -0.0002 * viewScale And gap < best Then
                best = gap
                Set nearest = r
            End If
        End If
    Next
    If nearest Is Nothing Then Exit Function
    ' Ignore bolt-hole circles when looking for the mating face.
    For Each r In rings
        If r("component") = nearest("component") And Abs(r("z") - pipeEnd("z")) < 0.0002 * viewScale And Abs(r(cross) - pipeEnd(cross)) < 0.0002 * viewScale Then
            If r(axis) > nearest(axis) Then Set nearest = r
        End If
    Next
    Set FlangeEnd = nearest
End Function

Private Function SelectRing(ByVal doc As SldWorks.ModelDoc2, ByVal v As SldWorks.View, _
                            ByVal ring As Collection, ByVal appendSelection As Boolean) As Boolean
    Dim ent As SldWorks.Entity, manager As SldWorks.SelectionMgr, data As SldWorks.SelectData
    Dim required As Long
    Set manager = doc.SelectionManager
    required = manager.GetSelectedObjectCount2(-1) + 1
    Set ent = ring("edge")
    Set data = manager.CreateSelectData
    data.View = v
    If Not ent.Select4(appendSelection, data) Then
        If Not v.SelectEntity(ent, appendSelection) Then Exit Function
    End If
    SelectRing = (manager.GetSelectedObjectCount2(-1) = required)
End Function

Private Function PutDim(ByVal doc As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal a As Collection, ByVal b As Collection, ByVal axis As String, ByVal x As Double, ByVal y As Double) As Boolean
    Dim disp As SldWorks.DisplayDimension, dimObj As SldWorks.Dimension, expected As Double
    dimFailure = "missing base"
    If a Is Nothing Then Exit Function
    doc.ClearSelection2 True
    dimFailure = "base selection"
    If Not SelectRing(doc, v, a, False) Then Exit Function
    If axis = "dia" Then
        expected = 2 * a("radius")
        Set disp = doc.AddDiameterDimension2(x, y, 0#)
    Else
        dimFailure = "missing target"
        If b Is Nothing Then Exit Function
        dimFailure = "target selection"
        If Not SelectRing(doc, v, b, True) Then Exit Function
        expected = Abs(b(axis) - a(axis)) / v.ScaleDecimal
        If axis = "x" Then
            Set disp = doc.AddHorizontalDimension2(x, y, 0#)
        Else
            Set disp = doc.AddVerticalDimension2(x, y, 0#)
        End If
    End If
    dimFailure = "SOLIDWORKS declined insertion"
    If disp Is Nothing Then Exit Function
    Set dimObj = disp.GetDimension2(0)
    dimFailure = "dimension value unavailable"
    If Not dimObj Is Nothing Then
        PutDim = Abs(dimObj.SystemValue - expected) < 0.0001
        If Not PutDim Then dimFailure = "expected " & Format$(1000 * expected, "0.##") & " mm, got " & Format$(1000 * dimObj.SystemValue, "0.##") & " mm"
    End If
    If Not PutDim Then
        Dim ann As SldWorks.Annotation
        Set ann = disp.GetAnnotation
        doc.ClearSelection2 True
        If ann Is Nothing Then Err.Raise 5, , "Cannot remove unverified dimension."
        If Not ann.Select3(False, Nothing) Then Err.Raise 5, , "Cannot select unverified dimension for removal."
        If Not doc.Extension.DeleteSelection2(0) Then Err.Raise 5, , "Cannot remove unverified dimension."
    End If
End Function

Private Sub PlaceDimensions(ByVal doc As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal rings As Collection)
    Dim r As Collection, lo As Collection, hi As Collection, leftEnd As Collection, rightEnd As Collection
    Dim topEnd As Collection, bottomEnd As Collection, flange As Collection, outer As Collection
    Dim seen As Object, outline As Variant, level As Long, missed As String
    Dim dx As Double, dy As Double
    For Each r In rings
        If r("kind") = "pipe" Then
            Set lo = EndRing(rings, r, "x", -1)
            Set hi = EndRing(rings, r, "x", 1)
            dx = hi("x") - lo("x")
            dy = Abs(hi("y") - lo("y"))
            If dx > 0.001 * v.ScaleDecimal And dy < 0.0002 * v.ScaleDecimal Then
                If Not leftEnd Is Nothing Then
                    If leftEnd("component") <> lo("component") Then Err.Raise 5, , "Multiple horizontal pipe runs in this view."
                End If
                Set leftEnd = lo
                Set rightEnd = hi
            Else
                Set lo = EndRing(rings, r, "y", -1)
                Set hi = EndRing(rings, r, "y", 1)
                If hi("y") - lo("y") > 0.001 * v.ScaleDecimal And Abs(hi("x") - lo("x")) < 0.0002 * v.ScaleDecimal Then
                    If Not bottomEnd Is Nothing Then
                        If bottomEnd("component") <> lo("component") Then Err.Raise 5, , "Multiple vertical pipe runs in this view."
                    End If
                    Set bottomEnd = lo
                    Set topEnd = hi
                End If
            End If
        End If
    Next
    If leftEnd Is Nothing Or topEnd Is Nothing Then Err.Raise 5, , "Could not identify both straight pipes in front view."
    outline = v.GetOutline
    If Not HasItems(outline) Then Err.Raise 5, , "Front view outline unavailable."
    Set seen = CreateObject("Scripting.Dictionary")
    Dim stations As New Collection, existing As Collection, slot As Long
    For Each r In rings
        If r("kind") = "branch" And r("x") > leftEnd("x") And r("x") < rightEnd("x") Then
            Dim key As String
            key = CStr(Round(r("x") / v.ScaleDecimal, 4))
            If Not seen.Exists(key) Then
                seen.Add key, True
                For slot = 1 To stations.Count
                    Set existing = stations(slot)
                    If existing("x") > r("x") Then Exit For
                Next
                If slot > stations.Count Then
                    stations.Add r
                Else
                    stations.Add r, Before:=slot
                End If
            End If
        End If
    Next
    For Each r In stations
        level = level + 1
        stage = "branch dimension | " & r("component")
        If Not PutDim(doc, v, leftEnd, r, "x", (leftEnd("x") + r("x")) / 2, outline(1) - 0.007 * level) Then missed = missed & vbCrLf & r("component") & " - " & dimFailure
    Next
    stage = "pipe and flange dimensions"
    If level = 0 Then missed = missed & " branches"
    If Not PutDim(doc, v, leftEnd, rightEnd, "x", (leftEnd("x") + rightEnd("x")) / 2, outline(1) - 0.007 * (level + 1)) Then missed = missed & vbCrLf & "horizontal pipe" & " - " & dimFailure
    If Not PutDim(doc, v, bottomEnd, topEnd, "y", outline(0) - 0.007, (bottomEnd("y") + topEnd("y")) / 2) Then missed = missed & vbCrLf & "vertical pipe" & " - " & dimFailure
    Set flange = FlangeEnd(rings, rightEnd, "x", "y", v.ScaleDecimal)
    If Not PutDim(doc, v, bottomEnd, flange, "x", (bottomEnd("x") + rightEnd("x")) / 2, outline(1) - 0.007 * (level + 2)) Then missed = missed & vbCrLf & "welded width" & " - " & dimFailure
    Set flange = FlangeEnd(rings, topEnd, "y", "x", v.ScaleDecimal)
    If Not PutDim(doc, v, leftEnd, flange, "y", outline(0) - 0.014, (leftEnd("y") + topEnd("y")) / 2) Then missed = missed & vbCrLf & "welded height" & " - " & dimFailure
    Set outer = rightEnd
    For Each r In rings
        If r("component") = rightEnd("component") And Abs(r("x") - rightEnd("x")) < 0.00001 Then
            If r("radius") > outer("radius") Then Set outer = r
        End If
    Next
    If Not PutDim(doc, v, outer, Nothing, "dia", outline(2) + 0.007, rightEnd("y")) Then missed = missed & vbCrLf & "OD" & " - " & dimFailure
    doc.ClearSelection2 True
    If Len(missed) > 0 Then Err.Raise 5, , "Missing dimensions:" & missed
End Sub
