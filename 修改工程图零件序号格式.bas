Attribute VB_Name = "Module1"
Sub CATMain()

    Dim Slct

    Set Slct = CATIA.ActiveDocument.Selection
    
    Dim view
    Set view = CATIA.ActiveDocument.Sheets.ActiveSheet.Views.ActiveView
    
    Slct.Clear

    For Each Text In view.Texts
        '英文环境下零件序号改为Balloon
        If InStr(Text.Name, "零件序号") <> O Then
        
        Dim MyStr

        MyStr = Text.Text

        Dim TextPosX, TextPosY, LeaderPosX, LeaderPosY
        TextPosX = Text.X
        TextPosY = Text.Y
        Text.Leaders.Item(1).GetPoint 1, LeaderPosX, LeaderPosY
        
        Slct.Add (Text)
        Set t = view.Texts.Add(MyStr, TextPosX, TextPosY)
        Set l = t.Leaders.Add(LeaderPosX, LeaderPosY)
        t.SetFontSize 0, 0, 10
        
    End If
    
Next

Slct.Delete

End Sub

