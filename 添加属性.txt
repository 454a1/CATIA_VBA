Attribute VB_Name = "添加属性"
Sub CATMain()
    
    Dim str1 As String, str2 As String, str3 As String, _
        str4 As String, str5 As String, str6 As String, _
        str7 As String
        
        str1 = "01"
        str2 = "JJ-01"
        str3 = "夹具体"
        str4 = "1"
        str5 = "HT150"
        str6 = ""
        str7 = "无"
    
    Dim oPartDoc As PartDocument
    Set oPartDoc = CATIA.ActiveDocument

    Dim oProduct As Product
    Set oProduct = oPartDoc.GetItem("Part")

    Dim par1 As Parameters, par2 As Parameters, par3 As Parameters, _
        par4 As Parameters, par5 As Parameters, par6 As Parameters, _
        par7 As Parameters
    
    Set par1 = oProduct.UserRefProperties
    Set par2 = oProduct.UserRefProperties
    Set par3 = oProduct.UserRefProperties
    Set par4 = oProduct.UserRefProperties
    Set par5 = oProduct.UserRefProperties
    Set par6 = oProduct.UserRefProperties
    Set par7 = oProduct.UserRefProperties
    
    Dim strParam1 As StrParam, strParam2 As StrParam, strParam3 As StrParam, _
        strParam4 As StrParam, strParam5 As StrParam, strParam6 As StrParam, _
        strParam7 As StrParam

    Set strParam1 = par1.CreateString("序号", "")
    Set strParam2 = par2.CreateString("代号", "")
    Set strParam3 = par3.CreateString("名称", "")
    Set strParam4 = par4.CreateString("数量", "")
    Set strParam5 = par5.CreateString("材料", "")
    Set strParam6 = par6.CreateString("重量", "")
    Set strParam7 = par7.CreateString("备注", "")
    
    strParam1.ValuateFromString str1
    strParam2.ValuateFromString str2
    strParam3.ValuateFromString str3
    strParam4.ValuateFromString str4
    strParam5.ValuateFromString str5
    strParam6.ValuateFromString str6
    strParam7.ValuateFromString str7
    
End Sub

