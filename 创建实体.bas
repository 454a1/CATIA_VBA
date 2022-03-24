Attribute VB_Name = "����ʵ��"
Sub CATMain()
    
     On Error Resume Next
    
    Dim oPart As Part
    Dim oBodies As Bodies
    Dim oBody As Body
    Dim oSF As ShapeFactory
    Dim oPad As Pad
    

    '���õ�ǰ�ĵ�
    Set oPart = CATIA.ActiveDocument.Part
    
    '�����ǰû���ĵ���
    If Err.Number <> 0 Then

       '�½��ĵ�
        Dim oDoc As Document
        Set oDoc = CATIA.Documents.Add("Part")
        Set oPart = oDoc.Part

    End If

    On Error GoTo 0
    
    Set oBodies = oPart.Bodies
    Set oBody = oBodies.Item("���������")
      
    'ѡȡ�ο�ƽ��
    Dim plnXY As Plane
    Set plnXY = oPart.originElements.PlaneYZ
    
    '�ڲο�ƽ������Ӳ�ͼ
    Dim oSketch As Sketch
    Set oSketch = oBody.Sketches.Add(plnXY)
    
    '��ȡFactory2D������Factory2D���Դ�����ͼԪ�أ���Բ����ֱ�ߵȣ�
    Dim oFactory2D As Factory2D
    Set oFactory2D = oSketch.OpenEdition
    
    '��Բ
    Dim oCCircle As Circle2D
    Set oCCircle = oFactory2D.CreateClosedCircle(-30, -50, 15)
    
    oSketch.CloseEdition
    
    '����Pad�����ĸ߶�Ϊ20mm
    Set oSF = oPart.ShapeFactory
    Set oPad = oSF.AddNewPad(oSketch, 20)
    
    oPart.Update
    
End Sub

