Attribute VB_Name = "PipeDrawing"
Option Explicit

Public Sub main()
    On Error GoTo Failed
    Dim sw As SldWorks.SldWorks, asm As SldWorks.ModelDoc2
    Dim doc As SldWorks.ModelDoc2, drw As SldWorks.DrawingDoc
    Dim front As SldWorks.View, iso As SldWorks.View
    Dim sheet As SldWorks.Sheet, w As Double, h As Double
    Dim template As String, path As String, config As String
    Set sw = Application.SldWorks
    Set asm = sw.ActiveDoc
    If asm Is Nothing Then Err.Raise 5, , "Open the piping assembly."
    If asm.GetType <> swDocASSEMBLY Then Err.Raise 5, , "Select an assembly."
    path = asm.GetPathName
    If Len(path) = 0 Then Err.Raise 5, , "Save the assembly first."
    config = asm.ConfigurationManager.ActiveConfiguration.Name
    template = sw.GetUserPreferenceStringValue(swDefaultTemplateDrawing)
    Set doc = sw.NewDocument(template, 0, 0#, 0#)
    If doc Is Nothing Then Err.Raise 5, , "Check the default drawing template."
    Set drw = doc
    Set sheet = drw.GetCurrentSheet
    sheet.GetSize w, h
    Set front = drw.CreateDrawViewFromModelView3(path, "*Front", w * 0.68, h * 0.53, 0#)
    If front Is Nothing Then Err.Raise 5, , "Front view creation failed."
    front.ReferencedConfiguration = config
    Set iso = drw.CreateDrawViewFromModelView3(path, "*Isometric", w * 0.24, h * 0.53, 0#)
    If iso Is Nothing Then Err.Raise 5, , "Isometric view creation failed."
    iso.ReferencedConfiguration = config
    doc.EditRebuild3
    AddPipeDimensions sw, doc, front
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "PipeDrawing"
End Sub

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
    Set mu = sw.GetMathUtility
    parts = v.GetVisibleComponents
    If IsEmpty(parts) Then Err.Raise 5, , "No visible components in front view."
    For i = LBound(parts) To UBound(parts)
        Set c = parts(i)
        kind = PartKind(c)
        If Len(kind) > 0 Then
            edges = v.GetVisibleEntities2(c, swViewEntityType_Edge)
            If IsArray(edges) Then
                For j = LBound(edges) To UBound(edges)
                    Set e = edges(j)
                    Set curve = e.GetCurve
                    If curve.IsCircle Then
                        p = curve.CircleParams
                        Dim coords(2) As Double
                        coords(0) = p(0)
                        coords(1) = p(1)
                        coords(2) = p(2)
                        Set pt = mu.CreatePoint(coords)
                        Set pt = pt.MultiplyTransform(c.GetTotalTransform(False))
                        Set pt = pt.MultiplyTransform(v.ModelToViewTransform)
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
                Next
            End If
        End If
    Next
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

Private Function FlangeEnd(ByVal rings As Collection, ByVal pipeEnd As Collection, ByVal axis As String, ByVal cross As String, ByVal scale As Double) As Collection
    Dim r As Collection, nearest As Collection, gap As Double, best As Double
    best = 1E+10
    For Each r In rings
        If r("kind") = "flange" Then
            gap = r(axis) - pipeEnd(axis)
            If Abs(r(cross) - pipeEnd(cross)) < 0.0002 * scale And gap >= -0.0002 * scale And gap < best Then
                best = gap
                Set nearest = r
            End If
        End If
    Next
    If nearest Is Nothing Then Exit Function
    ' Ignore bolt-hole circles when looking for the mating face.
    For Each r In rings
        If r("component") = nearest("component") And Abs(r(cross) - pipeEnd(cross)) < 0.0002 * scale Then
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
