Attribute VB_Name = "AddNewCircleCtrRad"
Sub CATMain()

    On Error Resume Next
    
    Dim oPart As Part
    Dim oBodies As Bodies
    Dim oBody As Body
    Dim oHSF As HybridShapeFactory
    
    Set oPart = CATIA.ActiveDocument.Part
    
    If Err.Number <> 0 Then

        Dim oDoc As Document
        Set oDoc = CATIA.Documents.Add("Part")
        Set oPart = oDoc.Part

    End If

    On Error GoTo 0
    
    Set oBodies = oPart.Bodies
    Set oHSF = oPart.HybridShapeFactory
    Set oBody = oBodies.Item("零件几何体")

    Dim oPt As Point
    Set oPt = oHSF.AddNewPointCoord(0#, 0#, 0#)
    
    oBody.InsertHybridShape oPt


    Dim oRefPt As Reference
    Set oRefPt = oPart.CreateReferenceFromObject(oPt)

    Dim oPlane As Plane
    Set oPlane = oPart.OriginElements.PlaneXY

    Dim oRefplane As Reference
    Set oRefplane = oPart.CreateReferenceFromObject(oPlane)

    Dim oCircle As HybridShapeCircle
    Set oCircle = oHSF.AddNewCircleCtrRad(oRefPt, oRefplane, False, 20#)

    oBody.InsertHybridShape oCircle

    oPart.Update

End Sub

