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
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "PipeDrawing"
End Sub
