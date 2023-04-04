Attribute VB_Name = "Mod1"
Sub AllocationLogic1()


Application.Calculation = xlCalculationManual


Dim twb As Workbook
Set twb = ThisWorkbook

Set ws_PreAllocation = twb.Sheets("PreAllocation")
Set ws_WaxCellUtil = twb.Sheets("WaxCellUtilization")
Set ws_ProductionOrders = twb.Sheets("ProductionOrders")

Set ws_Validation = twb.Sheets("Validation")

Set t_ItemAllocation = ws_Validation.ListObjects("ItemAllocation")



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
    
    With t_ProductionOrdersByItem
        
        For i = 1 To .DataBodyRange.Rows.count
            dict_MaxCell.Add .ListColumns("ItemId").DataBodyRange.Rows(i).Value, .ListColumns("MaximumWaxCellAllocation").DataBodyRange.Rows(i).Value
            dict_UsedCellString.Add .ListColumns("ItemId").DataBodyRange.Rows(i).Value, ""
        Next
    
    End With
    
    
'---

'Cleaning previous Info
t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.ClearContents
t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.ClearContents

'-----


'---Allocate

ProdOrderLastRow = t_ProductionOrders.DataBodyRange.Rows.count


'Loop Waxcell--------------
For i = 1 To t_WaxCellUtil.DataBodyRange.Rows.count
    WaxCell = t_WaxCellUtil.ListColumns("Wax Cell").DataBodyRange.Rows(i).Value

    WaxCap = t_WaxCellUtil.ListColumns("Total Hours/Week per cell").DataBodyRange.Rows(i).Value
    WaxRem = WaxCap
    
    'Loop Category------------
    For Each cat In set_CategoryMix
        CatCap = dict_CategoryContribution(cat) * WaxCap * TargetUtil
        startrow = dict_CategoryRow(cat)
        
        'Loop Production Orders
        For j = startrow To ProdOrderLastRow
            If cat <> t_ProductionOrders.ListColumns("Category").DataBodyRange.Rows(j).Value Or CatCap < 0 Then
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
                            
                        CatCap = CatCap - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                        WaxRem = WaxRem - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
                        
                    End If
                Else
                    t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(j).Value = WaxCell
                        
                    CatCap = CatCap - t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(j).Value
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
ws_ProductionOrders.Select
Application.Calculation = xlCalculationAutomatic

End Sub

Sub forcelogic1()

Application.Calculation = xlCalculationManual


Dim twb As Workbook
Set twb = ThisWorkbook

Set ws_PreAllocation = twb.Sheets("PreAllocation")
Set ws_WaxCellUtil = twb.Sheets("WaxCellUtilization")
Set ws_ProductionOrders = twb.Sheets("ProductionOrders")
Set ws_Validation = twb.Sheets("Validation")

Set t_ItemAllocation = ws_Validation.ListObjects("ItemAllocation")


Set t_WaxCellUtil = ws_WaxCellUtil.ListObjects("ActiveWaxCells")
Set t_ProductionOrders = ws_ProductionOrders.ListObjects("ProductionOrders_Display")

Dim dict_ItemWaxCell As scripting.Dictionary
Set dict_ItemWaxCell = New scripting.Dictionary

Dim dict_UnallocatedItem As scripting.Dictionary
Set dict_UnallocatedItem = New scripting.Dictionary

'-----------WaxCellRow
Dim dict_WaxRow As scripting.Dictionary
Set dict_WaxRow = New scripting.Dictionary

For i = 1 To t_WaxCellUtil.DataBodyRange.Rows.count
    WaxCell = t_WaxCellUtil.ListColumns("Wax Cell").DataBodyRange.Rows(i).Value
    dict_WaxRow.Add WaxCell, i
Next
'-------------------




With t_ProductionOrders
    For i = 1 To .DataBodyRange.Rows.count
        If Not IsEmpty(.ListColumns("TargetWaxCell").DataBodyRange.Rows(i)) Then
            If Not dict_ItemWaxCell.Exists(.ListColumns("ItemId").DataBodyRange.Rows(i).Value) Then
                dict_ItemWaxCell.Add .ListColumns("ItemId").DataBodyRange.Rows(i).Value, .ListColumns("TargetWaxCell").DataBodyRange.Rows(i).Value
            Else
                If Not dict_ItemWaxCell(.ListColumns("ItemId").DataBodyRange.Rows(i).Value) Like "*" & .ListColumns("TargetWaxCell").DataBodyRange.Rows(i).Value & "*" Then
                    dict_ItemWaxCell(.ListColumns("ItemId").DataBodyRange.Rows(i).Value) = dict_ItemWaxCell(.ListColumns("ItemId").DataBodyRange.Rows(i).Value) & "|" & .ListColumns("TargetWaxCell").DataBodyRange.Rows(i).Value
                End If
            End If
        Else
            If Not dict_UnallocatedItem.Exists(.ListColumns("ItemId").DataBodyRange.Rows(i).Value) Then
                dict_UnallocatedItem.Add .ListColumns("ItemId").DataBodyRange.Rows(i).Value, i
            Else
                dict_UnallocatedItem(.ListColumns("ItemId").DataBodyRange.Rows(i).Value) = dict_UnallocatedItem(.ListColumns("ItemId").DataBodyRange.Rows(i).Value) & "," & i
            End If
        End If
    Next
    
    
For Each c In dict_UnallocatedItem

    arr_row = Split(dict_UnallocatedItem(c), ",")
    arr_AvailWaxCell = Split(dict_ItemWaxCell(c), "|")
    NumWaxCell = UBound(arr_AvailWaxCell) + 1
    
    counter = 0
    For j = LBound(arr_row) To UBound(arr_row)
        ProdRow = arr_row(j)
        
        WaxRow = dict_WaxRow(arr_AvailWaxCell(counter))
        
        .ListColumns("TargetWaxCell").DataBodyRange.Rows(ProdRow).Value = arr_AvailWaxCell(counter)
        t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(WaxRow).Value = t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(WaxRow).Value + _
            .ListColumns("ProductionHour").DataBodyRange.Rows(ProdRow).Value
        
        counter = (counter + 1) Mod NumWaxCell
    Next
Next

End With


t_ItemAllocation.QueryTable.Refresh BackgroundQuery:=False
Application.Calculation = xlCalculationAutomatic

End Sub


Sub refreshAssignedline()

ThisWorkbook.Sheets("Preview Assigned Line").ListObjects("PreAssignedLine_FINAL").QueryTable.Refresh BackgroundQuery:=False
ThisWorkbook.Sheets("PriorWk_Pre").ListObjects("PriorWk_Pre").Refresh

End Sub
Sub twbrefresh()

ThisWorkbook.RefreshAll
ThisWorkbook.Sheets("PreAllocation").ListObjects("ProductionOrdersByCategory").ListColumns("Contribution").DataBodyRange.Style = "Percent"

End Sub


Sub PublishAssignedLine()




Set twb = ThisWorkbook
Set ws_Preview = twb.Sheets("Preview Assigned Line")
Set t_Preview = ws_Preview.ListObjects("PreAssignedLine_FINAL")

Set wb_Publish = Workbooks.Add
Set ws_Publish = wb_Publish.Sheets(1)


Set t_PriorWk_Pre = twb.Sheets("PriorWk_Pre").ListObjects("PriorWk_Pre")
Set t_PriorWk = twb.Sheets("PriorWk").ListObjects("PriorWk")
Set ws_PriorWk = twb.Sheets("PriorWk")

PriorRows = t_PriorWk_Pre.DataBodyRange.Rows.count
If Not t_PriorWk.DataBodyRange Is Nothing Then
    t_PriorWk.DataBodyRange.Delete
End If
t_PriorWk.ListRows.Add
t_PriorWk.Resize ws_PriorWk.Range("A1:D" & (1 + PriorRows))
t_PriorWk.DataBodyRange.Value = t_PriorWk_Pre.DataBodyRange.Value


ws_Publish.Name = "ProductionordersAPS"
ws_Publish.[A1].Value = "ProdId"
ws_Publish.[B1].Value = "AssignedLine"

numrows = t_Preview.DataBodyRange.Rows.count
numcols = t_Preview.DataBodyRange.Columns.count

ws_Publish.Range(ws_Publish.Cells(2, 1), ws_Publish.Cells(1 + numrows, numcols)).NumberFormat = "@"
ws_Publish.Range(ws_Publish.Cells(2, 1), ws_Publish.Cells(1 + numrows, numcols)).Value = t_Preview.DataBodyRange.Value

End Sub

Sub temp_ManualAdjustment()

Application.Calculation = xlCalculationManual

Dim twb As Workbook
Set twb = ThisWorkbook


Set ws_WaxCellUtil = twb.Sheets("WaxCellUtilization")
Set ws_ProductionOrders = twb.Sheets("ProductionOrders")

Set ws_Validation = twb.Sheets("Validation")

Set t_ItemAllocation = ws_Validation.ListObjects("ItemAllocation")


Set t_WaxCellUtil = ws_WaxCellUtil.ListObjects("ActiveWaxCells")
Set t_ProductionOrders = ws_ProductionOrders.ListObjects("ProductionOrders_Display")

t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.ClearContents


Dim dict_WaxRow As scripting.Dictionary
Set dict_WaxRow = New scripting.Dictionary

For i = 1 To t_WaxCellUtil.DataBodyRange.Rows.count
    WaxCell = t_WaxCellUtil.ListColumns("Wax Cell").DataBodyRange.Rows(i).Value
    dict_WaxRow.Add WaxCell, i
Next



For i = 1 To t_ProductionOrders.DataBodyRange.Rows.count
    WaxCell = t_ProductionOrders.ListColumns("TargetWaxCell").DataBodyRange.Rows(i).Value
    t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(dict_WaxRow(WaxCell)) = t_WaxCellUtil.ListColumns("Consumed Hour").DataBodyRange.Rows(dict_WaxRow(WaxCell)) + t_ProductionOrders.ListColumns("ProductionHour").DataBodyRange.Rows(i).Value

Next


t_ItemAllocation.QueryTable.Refresh BackgroundQuery:=False
Application.Calculation = xlCalculationAutomatic
ws_WaxCellUtil.Select
End Sub
