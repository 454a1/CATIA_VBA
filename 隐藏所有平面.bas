Attribute VB_Name = "隐藏所有平面"
Sub CATMain()

        Dim oSel As Selection
        Set oSel = CATIA.ActiveDocument.Selection
        
        oSel.Search "Type = 平面,all"
        
        Dim Visprop As VisPropertySet
        Set Visprop = oSel.VisProperties
        Visprop.SetShow catVisPropertyNoShowAttr

End Sub

