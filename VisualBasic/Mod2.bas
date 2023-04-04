Attribute VB_Name = "Mod2"
Sub AllocationLogic2()



Application.Calculation = xlCalculationManual


Dim twb As Workbook
Set twb = ThisWorkbook

Set ws_PreAllocation = twb.Sheets("PreAllocation")
Set ws_WaxCellUtil = twb.Sheets("WaxCellUtilization")
Set ws_ProductionOrders = twb.Sheets("ProductionOrders")
Set ws_PriorWk = twb.Sheets("PriorWk")
Set ws_Validation = twb.Sheets("Validation")

Set t_ItemAllocation = ws_Validation.ListObjects("ItemAllocation")
Set t_PriorWk = ws_PriorWk.ListObjects("PriorWk")


With ws_PreAllocation
    Set t_ProductionOrdersByCategory = .ListObjects("ProductionOrdersByCategory")
    Set t_ProductionOrdersByItem = .ListObjects("ProductionOrdersByItem_Display")
    r_TargetUtilization = .[r_TargetUtilization].Value
    r_MaxUtilByItem = .[r_MaxUtilByItem].Value
End With

Set t_WaxCellUtil = ws_WaxCellUtil.ListObjects("ActiveWaxCells")
Set t_ProductionOrders = ws_ProductionOrders.ListObjects("ProductionOrders_Display")


t_WaxCellUtil.QueryTable.Refresh BackgroundQuery:=False

TargetUtil = r_TargetUtilization

'---Dictionary for locating starting row of each catgory in Production Orders
Dim dict_CategoryRow As scripting.Dictionary
Set dict_CategoryRow = New scripting.Dictionary



CurCat = ""
With t_ProductionOrders
    For i = 1 To .DataBodyRange.Rows.count
        
        If .ListColumns("Category").DataBodyRange.Rows(i).Value <> CurCat Then
            dict_CategoryRow.Add .ListColumns("Category").DataBodyRange.Rows(i).Value, i
            CurCat = .ListColumns("Category").DataBodyRange.Rows(i).Value
        End If
        
    Next
End With
'---

'---Collection of Category
Dim set_CategoryMix As New Collection


With t_ProductionOrdersByCategory
    For i = 1 To .DataBodyRange.Rows.count
        set_CategoryMix.Add .ListColumns("Category").DataBodyRange.Rows(i).Value
    Next
End With


'---Target Category Mix

Dim dict_CategoryContribution As scripting.Dictionary
Set dict_CategoryContribution = New scripting.Dictionary

With t_ProductionOrdersByCategory
    For i = 1 To .DataBodyRange.Rows.count
        dict_CategoryContribution.Add .ListColumns("Category").DataBodyRange.Rows(i).Value, _
                                    .ListColumns("Contribution").DataBodyRange.Rows(i).Value
    Next
End With
'---


'---Wax Cell Constraint

    'Max
    Dim dict_MaxCell As scripting.Dictionary
    Set dict_MaxCell = New scripting.Dictionary
    'Used
    Dim dict_UsedCellString As scripting.Dictionary
    Set dict_UsedCellString = New scripting.Dictionary
    
    
    'Total Production Hour by Item
    Dim dict_ItemProdHour As scripting.Dictionary
    Set dict_ItemProdHour = New scripting.Dictionary
    
    With t_ProductionOrdersByItem
        
        For i = 1 To .DataBodyRange.Rows.count
            dict_MaxCell.Add .ListColumns("ItemId").DataBodyRange.Rows(i).Value, .ListColumns("MaximumWaxCellAllocation").DataBodyRange.Rows(i).Value
            dict_UsedCellString.Add .ListColumns("ItemId").DataBodyRange.Rows(i).Value, ""
            dict_ItemProdHour.Add .ListColumns("ItemId").DataBodyRange.Rows(i).Value, .ListColumns("ProductionHour").DataBodyRange.Rows(i).Value
            
        Next
    
    End With
    
    
'---


'---Prior Allocation
Dim dict_PreLoadedLine As scripting.Dictionary
Set dict_PreLoadedLine = New scripting.Dictionary

Dim dict_PreItemCap As scripting.Dictionary
Set dict_PreItemCap = New scripting.Dictionary

Dim dict_PreMaxLine As scripting.Dictionary
Set dict_PreMaxLine = New scripting.Dictionary

Dim dict_PreCountLine As scripting.Dictionary
Set dict_PreCountLine = New scripting.Dictionary

str_PreAlloKey = ""
With t_PriorWk
    If Not .DataBodyRange Is Nothing Then
        For i = 1 To .DataBodyRange.Rows.count
            If dict_ItemProdHour.Exists(.ListColumns("ItemId").DataBodyRange.Rows(i).Value) Then
                If Not dict_PreItemCap.Exists(.ListColumns("ItemId").DataBodyRange.Rows(i).Value) Then
                    the_divider = Application.WorksheetFunction.Min(.ListColumns("Lines").DataBodyRange.Rows(i).Value, dict_MaxCell(.ListColumns("ItemId").DataBodyRange.Rows(i).Value))
                    dict_PreItemCap.Add .ListColumns("ItemId").DataBodyRange.Rows(i).Value, dict_ItemProdHour(.ListColumns("ItemId").DataBodyRange.Rows(i).Value) / the_divider
                    dict_PreMaxLine.Add .ListColumns("ItemId").DataBodyRange.Rows(i).Value, Application.WorksheetFunction.Min(the_divider, 6)
                    dict_PreCountLine.Add .ListColumns("ItemId").DataBodyRange.Rows(i).Value, 0
                    
                End If
            End If
            
            arr_Lines = Split(.ListColumns("TargetWaxCell").DataBodyRange.Rows(i).Value, ",")
            For j = LBound(arr_Lines) To UBound(arr_Lines)
                str_PreAlloKey = arr_Lines(j) & "|" & .ListColumns("Category").DataBodyRange.Rows(i).Value
                If Not dict_PreLoadedLine.Exists(str_PreAlloKey) Then
                    dict_PreLoadedLine.Add Key:=str_PreAlloKey, Item:=.ListColumns("ItemId").DataBodyRange.Rows(i).Value
                Else
                    dict_PreLoadedLine(str_PreAlloKey) = dict_PreLoadedLine(str_PreAlloKey) & "," & .ListColumns("ItemId").DataBodyRange.Rows(i).Value
                End If
            Next
            
            
        Next
    
        For Each ky In dict_PreLoadedLine.Keys
            dict_PreLoadedLine(ky) = Split(dict_PreLoadedLine(ky), ",")
        Next
    End If
End With

'Item Row
Dim dict_ItemRow As scripting.Dictionary
Set dict_ItemRow = New scripting.Dictionary

str_ItemDummy = ""
With t_ProductionOrders
    For i = 1 To .DataBodyRange.Rows.count
        If i = 1 Then
            str_ItemDummy = .ListColumns("ItemId").DataBodyRange.Rows(i).Value
            dict_ItemRow.Add Key:=.ListColumns("ItemId").DataBodyRange.Rows(i).Value, Item:=1
        Else
            If str_ItemDummy <> .ListColumns("ItemId").DataBodyRange.Rows(i).Value Then
                str_ItemDummy = .ListColumns("ItemId").DataBodyRange.Rows(i).Value
                dict_ItemRow.Add Key:=str_ItemDummy, Item:=i
            End If
        End If
    
    Next
End With



'Cleaning previous Info
t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.ClearContents
t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.ClearContents

'-----


'---Allocate

ProdOrderLastRow = t_ProductionOrders.DataBodyRange.Rows.count

Dim dict_CatCap As scripting.Dictionary
Set dict_CatCap = New scripting.Dictionary

'Initialize CatCap
For i = 1 To t_WaxCellUtil.DataBodyRange.Rows.count
    WaxCell = t_WaxCellUtil.ListColumns("Wax Cell").DataBodyRange.Rows(i).Value
    WaxCap = t_WaxCellUtil.ListColumns("Total Hours/Week per cell").DataBodyRange.Rows(i).Value
    For Each cat In set_CategoryMix
        dict_CatCap(cat & "|" & WaxCell) = dict_CategoryContribution(cat) * WaxCap * TargetUtil
    Next
Next


'PRIOR ALLOCATION---------------------------------------------------------------------------------------------------------------------------------
'Loop Waxcell--------------
t_ProductionOrders.ListColumns("PriorWkLine").DataBodyRange.ClearContents

For i = 1 To t_WaxCellUtil.DataBodyRange.Rows.count
    WaxCell = t_WaxCellUtil.ListColumns("Wax Cell").DataBodyRange.Rows(i).Value

    WaxCap = t_WaxCellUtil.ListColumns("Total Hours/Week per cell").DataBodyRange.Rows(i).Value
    WaxRem = WaxCap
    
    'Loop Category------------
    For Each cat In set_CategoryMix
        
        startrow = dict_CategoryRow(cat)
        
        'Loop for Assigning item with prior allocation
        PreKey = WaxCell & "|" & cat
        If dict_PreLoadedLine.Exists(PreKey) Then
            PreArr = dict_PreLoadedLine(PreKey)
            For kk = LBound(PreArr) To UBound(PreArr)
                PreItem = PreArr(kk)

                If dict_ItemRow.Exists(PreItem) Then
                    'Loop Production Orders
                        '-------------------------------------------------------------------------------------------------------------------------------------
                        ll = dict_ItemRow(PreItem)
                        ItemCap = dict_PreItemCap(PreItem)
                        AccumAllo = 0
                        xx = 0
                        For j = ll To t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows.count 'drop only 1 production orders
                            If t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value = "S1CS05553-WS" Then
                                Debug.Print WaxCell
                            End If
                            If PreItem <> t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value Or dict_CatCap(cat & "|" & WaxCell) < 0 Or AccumAllo + t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value > ItemCap Or _
                            dict_PreCountLine(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) = dict_PreMaxLine(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) Or _
                            xx = 1 Then
                                
                                Exit For
                            End If
                            
                
                            If IsEmpty(t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(j)) And t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value <= WaxRem Then
                                
                                If Not dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) Like "*" & WaxCell & "*" Then
                                    
                                    If UBound(Split(dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value), "|")) + 1 < dict_MaxCell(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) Then
                                        If dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) = "" Then
                                            dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) = WaxCell
                                        Else
                                            dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) = _
                                            dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) & "|" & WaxCell
                
                                        End If
                                        
                                        dict_PreCountLine(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) = dict_PreCountLine(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) + 1
                                        xx = 1
                                        If t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value = "S1CS05553-WS" Then
                                            Debug.Print dict_PreMaxLine(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value)
                                            Debug.Print dict_PreCountLine(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value)
                                        End If
                                        t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(j).Value = WaxCell
                                        t_ProductionOrders.ListColumns("PriorWkLine").DataBodyRange.Rows(j).Value = "Yes"
                                        dict_CatCap(cat & "|" & WaxCell) = dict_CatCap(cat & "|" & WaxCell) - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                                        WaxRem = WaxRem - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                                        
                                    End If
                                Else
                                    t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(j).Value = WaxCell
                                    t_ProductionOrders.ListColumns("PriorWkLine").DataBodyRange.Rows(j).Value = "Yes"
                                    dict_CatCap(cat & "|" & WaxCell) = dict_CatCap(cat & "|" & WaxCell) - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                                    WaxRem = WaxRem - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                                    AccumAllo = AccumAllo + t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                                End If
                
                            End If
                            
                        Next
                        
                        '-------------------------------------------------------------------------------------------------------------------------------------
                End If
            Next
        End If
        
    Next
    
    t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(i).Value = WaxCap - WaxRem
Next
'---
'---------------------------------------------------------------------------------------------------------------------------------PRIOR ALLOCATION

'Loop Waxcell--------------
For i = 1 To t_WaxCellUtil.DataBodyRange.Rows.count
    WaxCell = t_WaxCellUtil.ListColumns("Wax Cell").DataBodyRange.Rows(i).Value

    WaxCap = t_WaxCellUtil.ListColumns("Total Hours/Week per cell").DataBodyRange.Rows(i).Value
    WaxRem = t_WaxCellUtil.ListColumns("Total Hours/Week per cell").DataBodyRange.Rows(i).Value - t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(i).Value
    
    'Loop Category------------
    For Each cat In set_CategoryMix
        
        startrow = dict_CategoryRow(cat)
        
        
        
        'Loop Production Orders
        For j = startrow To ProdOrderLastRow
            If cat <> t_ProductionOrders.ListColumns("Category").DataBodyRange.Rows(j).Value Or dict_CatCap(cat & "|" & WaxCell) < 0 Then
                Exit For
            End If
            

            If IsEmpty(t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(j)) And t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value <= WaxRem Then
                
                If Not dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) Like "*" & WaxCell & "*" Then
                    
                    If UBound(Split(dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value), "|")) + 1 < dict_MaxCell(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) Then
                        If dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) = "" Then
                            dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) = WaxCell
                        Else
                            dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) = _
                            dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(j).Value) & "|" & WaxCell

                        End If
                        t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(j).Value = WaxCell
                            
                        dict_CatCap(cat & "|" & WaxCell) = dict_CatCap(cat & "|" & WaxCell) - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                        WaxRem = WaxRem - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                        
                    End If
                Else
                    t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(j).Value = WaxCell
                        
                    dict_CatCap(cat & "|" & WaxCell) = dict_CatCap(cat & "|" & WaxCell) - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                    WaxRem = WaxRem - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                End If

            End If
            
        Next
        
    Next
    
    t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(i).Value = WaxCap - WaxRem
Next
'---

'Loop Wax cell for Create RemainingCapDict
Dim dict_WaxRemainingCap As scripting.Dictionary
Set dict_WaxRemainingCap = New scripting.Dictionary

Dim dict_WaxRow As scripting.Dictionary
Set dict_WaxRow = New scripting.Dictionary

For i = 1 To t_WaxCellUtil.DataBodyRange.Rows.count
    WaxCell = t_WaxCellUtil.ListColumns("Wax Cell").DataBodyRange.Rows(i).Value

    WaxCap = t_WaxCellUtil.ListColumns("Total Hours/Week per cell").DataBodyRange.Rows(i).Value - t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(i).Value
    dict_WaxRemainingCap.Add WaxCell, WaxCap
    dict_WaxRow.Add WaxCell, i
Next

WaxCellNum = t_WaxCellUtil.DataBodyRange.Rows.count


For i = 1 To t_ProductionOrders.DataBodyRange.Rows.count

    If IsEmpty(t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(i)) Then
        If t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value = "S1CS03542-WS" Then
            Debug.Print i
            Debug.Print dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value) = ""
        
        End If
        
        If dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value) = "" Then
            
            For k = 0 To WaxCellNum - 1
                WaxCell = t_WaxCellUtil.ListColumns("Wax Cell").DataBodyRange.Rows(WaxCellNum - k).Value
                
                'Test---------
                If t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value = "S1CS03542-WS" Then
                    Debug.Print WaxCell
                End If
                'Test---------
                
                
                If dict_WaxRemainingCap(WaxCell) >= t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value Then
                    t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(i).Value = WaxCell
                    dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value) = WaxCell
                    dict_WaxRemainingCap(WaxCell) = dict_WaxRemainingCap(WaxCell) - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value
                    
                    t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(WaxCellNum - k).Value = t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(WaxCellNum - k).Value + _
                                        t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value
                End If
            Next
        
        
        ElseIf UBound(Split(dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value), "|")) + 1 < dict_MaxCell(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value) Then
            arr_WaxCell = Split(dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value), "|")
            
            logiccount = 0
            For d = LBound(arr_WaxCell) To UBound(arr_WaxCell)
                If dict_WaxRemainingCap(arr_WaxCell(d)) >= t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value Then
                    t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(i).Value = arr_WaxCell(d)
                    t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(dict_WaxRow(arr_WaxCell(d))).Value = t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(dict_WaxRow(arr_WaxCell(d))).Value + _
                                        t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value
                    dict_WaxRemainingCap(arr_WaxCell(d)) = dict_WaxRemainingCap(arr_WaxCell(d)) - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value
                    logiccount = 1
                    Exit For
                End If
            Next
            
            If logiccount = 0 Then
                For k = 0 To WaxCelNum - 1
                    WaxCell = t_WaxCellUtil.ListColumns("Wax Cell").DataBodyRange.Rows(WaxCellNum - k).Value
                    If dict_WaxRemainingCap(WaxCell) >= t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value Then
                        t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(i).Value = WaxCell
                        dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value) = dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value) & "|" & WaxCell
                        dict_WaxRemainingCap(WaxCell) = dict_WaxRemainingCap(WaxCell) - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value
                        
                        t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(WaxCellNum - k).Value = t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(WaxCellNum - k).Value + _
                                            t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value
                    End If
                Next
            End If
        Else
            arr_WaxCell = Split(dict_UsedCellString(t_ProductionOrders.ListColumns("ItemId").DataBodyRange.Rows(i).Value), "|")
            
            
            For d = LBound(arr_WaxCell) To UBound(arr_WaxCell)
                If dict_WaxRemainingCap(arr_WaxCell(d)) >= t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value Then
                    t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(i).Value = arr_WaxCell(d)
                    t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(dict_WaxRow(arr_WaxCell(d))).Value = t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(dict_WaxRow(arr_WaxCell(d))).Value + _
                                        t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value
                    dict_WaxRemainingCap(arr_WaxCell(d)) = dict_WaxRemainingCap(arr_WaxCell(d)) - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value
                    
                    Exit For
                End If
            Next
            
        End If
    
    
    
    End If
Next

'---------
t_ItemAllocation.QueryTable.Refresh BackgroundQuery:=False
'ws_ProductionOrders.Select
Application.Calculation = xlCalculationAutomatic

End Sub



