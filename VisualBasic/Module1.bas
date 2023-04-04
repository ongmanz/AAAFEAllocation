Attribute VB_Name = "Module1"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveWorkbook.Worksheets("ProductionOrders").ListObjects( _
        "ProductionOrders_Display").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ProductionOrders").ListObjects( _
        "ProductionOrders_Display").Sort.SortFields.Add2 Key:=Range( _
        "ProductionOrders_Display[IsLongRoute]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ProductionOrders").ListObjects( _
        "ProductionOrders_Display").Sort.SortFields.Add2 Key:=Range( _
        "ProductionOrders_Display[IsSparePart]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ProductionOrders").ListObjects( _
        "ProductionOrders_Display").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
