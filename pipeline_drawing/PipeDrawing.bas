Attribute VB_Name = "PipeDrawing"
Option Explicit

Private stage As String

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
    Set doc = sw.NewDocument(template, 0, 0#, 0#)
    If doc Is Nothing Then Err.Raise 5, , "Check the default drawing template."
    Set drw = doc
    stage = "drawing sheet"
    Set sheet = drw.GetCurrentSheet
    If sheet Is Nothing Then Err.Raise 5, , "No active drawing sheet."
    sheet.GetSize w, h
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
    stage = "rebuild"
    doc.EditRebuild3
    stage = "front view geometry"
    AddPipeDimensions sw, doc, front
    Exit Sub
Failed:
    MsgBox "PipeDrawing | " & stage & vbCrLf & "Error " & CStr(Err.Number) & ": " & Err.Description, vbExclamation, "PipeDrawing"
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
    If Left$(n, 5) = (ChrW(1090) & ChrW(1088) & ChrW(1091) & ChrW(1073) & ChrW(1072)) Then PartKind = "pipe"
    If Left$(n, 6) = (ChrW(1092) & ChrW(1083) & ChrW(1072) & ChrW(1085) & ChrW(1077) & ChrW(1094)) Then PartKind = "flange"
    If Left$(n, 5) = (ChrW(1084) & ChrW(1091) & ChrW(1092) & ChrW(1090) & ChrW(1072)) Or Left$(n, 6) = (ChrW(1088) & ChrW(1077) & ChrW(1079) & ChrW(1100) & ChrW(1073) & ChrW(1072)) Then PartKind = "branch"
End Function

Private Sub AddPipeDimensions(ByVal sw As SldWorks.SldWorks, ByVal doc As SldWorks.ModelDoc2, ByVal v As SldWorks.View)
    Dim parts As Variant, edges As Variant, i As Long, j As Long
    Dim c As SldWorks.Component2, e As SldWorks.Edge, curve As SldWorks.Curve
    Dim mu As SldWorks.MathUtility, pt As SldWorks.MathPoint, p As Variant, xyz As Variant
    Dim rings As New Collection, r As Collection, kind As String
    Dim placement As SldWorks.MathTransform, projection As SldWorks.MathTransform
    Dim sourceAssembly As SldWorks.AssemblyDoc, sourceComponent As SldWorks.Component2
    Set sourceAssembly = v.ReferencedDocument
    Set projection = v.ModelToViewTransform
    If projection Is Nothing Then Err.Raise 5, , "Front view has no model transform."
    Set mu = sw.GetMathUtility
    If mu Is Nothing Then Err.Raise 5, , "Math utility unavailable."
    parts = v.GetVisibleComponents
    If IsEmpty(parts) Then Err.Raise 5, , "No visible components in front view."
    For i = LBound(parts) To UBound(parts)
        Set c = parts(i)
        If c Is Nothing Then GoTo NextComponent
        stage = "front geometry | " & c.Name2
        kind = PartKind(c)
        If Len(kind) > 0 Then
            Set placement = c.Transform2
            If placement Is Nothing Then
                If Not sourceAssembly Is Nothing Then
                    Set sourceComponent = sourceAssembly.GetComponentByName(c.Name2)
                    If Not sourceComponent Is Nothing Then Set placement = sourceComponent.Transform2
                End If
            End If
            edges = v.GetVisibleEntities2(c, swViewEntityType_Edge)
            If IsArray(edges) Then
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
                        r.Add CDbl(p(6)), "radius"
                        r.Add c.Name2, "component"
                        r.Add kind, "kind"
                        rings.Add r
                    End If
NextEdge:
                Next
            End If
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
            If Abs(r(cross) - pipeEnd(cross)) < 0.0002 * viewScale And gap >= -0.0002 * viewScale And gap < best Then
                best = gap
                Set nearest = r
            End If
        End If
    Next
    If nearest Is Nothing Then Exit Function
    ' Ignore bolt-hole circles when looking for the mating face.
    For Each r In rings
        If r("component") = nearest("component") And Abs(r(cross) - pipeEnd(cross)) < 0.0002 * viewScale Then
            If r(axis) > nearest(axis) Then Set nearest = r
        End If
    Next
    Set FlangeEnd = nearest
End Function

Private Function PutDim(ByVal doc As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal a As Collection, ByVal b As Collection, ByVal axis As String, ByVal x As Double, ByVal y As Double) As Boolean
    Dim disp As SldWorks.DisplayDimension, dimObj As SldWorks.Dimension, expected As Double
    If a Is Nothing Then Exit Function
    doc.ClearSelection2 True
    If Not v.SelectEntity(a("edge"), False) Then Exit Function
    If axis = "dia" Then
        expected = 2 * a("radius")
        Set disp = doc.AddDiameterDimension2(x, y, 0#)
    Else
        If b Is Nothing Then Exit Function
        If Not v.SelectEntity(b("edge"), True) Then Exit Function
        expected = Abs(b(axis) - a(axis)) / v.ScaleDecimal
        If axis = "x" Then
            Set disp = doc.AddHorizontalDimension2(x, y, 0#)
        Else
            Set disp = doc.AddVerticalDimension2(x, y, 0#)
        End If
    End If
    If disp Is Nothing Then Exit Function
    Set dimObj = disp.GetDimension2(0)
    If dimObj Is Nothing Then Err.Raise 5, , "Cannot read the created dimension."
    PutDim = Abs(dimObj.SystemValue - expected) < 0.0001
    If Not PutDim Then
        Dim ann As SldWorks.Annotation
        Set ann = disp.GetAnnotation
        doc.ClearSelection2 True
        If ann.Select3(False, Nothing) Then doc.EditDelete
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
                Set leftEnd = lo
                Set rightEnd = hi
            Else
                Set lo = EndRing(rings, r, "y", -1)
                Set hi = EndRing(rings, r, "y", 1)
                If hi("y") - lo("y") > 0.001 * v.ScaleDecimal And Abs(hi("x") - lo("x")) < 0.0002 * v.ScaleDecimal Then
                    Set bottomEnd = lo
                    Set topEnd = hi
                End If
            End If
        End If
    Next
    If leftEnd Is Nothing Or topEnd Is Nothing Then Err.Raise 5, , "Could not identify both straight pipes in front view."
    outline = v.GetOutline
    Set seen = CreateObject("Scripting.Dictionary")
    For Each r In rings
        If r("kind") = "branch" And r("x") > leftEnd("x") And r("x") < rightEnd("x") Then
            Dim key As String
            key = CStr(Round(r("x") / v.ScaleDecimal, 4))
            If Not seen.Exists(key) Then
                seen.Add key, True
                level = level + 1
                If Not PutDim(doc, v, leftEnd, r, "x", (leftEnd("x") + r("x")) / 2, outline(1) - 0.007 * level) Then missed = missed & " branch"
            End If
        End If
    Next
    If level = 0 Then missed = missed & " branches"
    If Not PutDim(doc, v, leftEnd, rightEnd, "x", (leftEnd("x") + rightEnd("x")) / 2, outline(1) - 0.007 * (level + 1)) Then missed = missed & " horizontal pipe"
    If Not PutDim(doc, v, bottomEnd, topEnd, "y", outline(0) - 0.007, (bottomEnd("y") + topEnd("y")) / 2) Then missed = missed & " vertical pipe"
    Set flange = FlangeEnd(rings, rightEnd, "x", "y", v.ScaleDecimal)
    If Not PutDim(doc, v, bottomEnd, flange, "x", (bottomEnd("x") + rightEnd("x")) / 2, outline(1) - 0.007 * (level + 2)) Then missed = missed & " welded width"
    Set flange = FlangeEnd(rings, topEnd, "y", "x", v.ScaleDecimal)
    If Not PutDim(doc, v, leftEnd, flange, "y", outline(0) - 0.014, (leftEnd("y") + topEnd("y")) / 2) Then missed = missed & " welded height"
    Set outer = rightEnd
    For Each r In rings
        If r("component") = rightEnd("component") And Abs(r("x") - rightEnd("x")) < 0.00001 Then
            If r("radius") > outer("radius") Then Set outer = r
        End If
    Next
    If Not PutDim(doc, v, outer, Nothing, "dia", outline(2) + 0.007, rightEnd("y")) Then missed = missed & " OD"
    doc.ClearSelection2 True
    If Len(missed) > 0 Then Err.Raise 5, , "Missing dimensions:" & missed
End Sub
