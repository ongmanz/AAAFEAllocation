Attribute VB_Name = "modAllocation"
Option Explicit
Public idxLine, idxCapHrs, idxWAHrs, idxLRHrs, idxSPHrs, idxGM, idxQty As Integer
Public idxFixedLine, idxScatter As Integer
Public idxAssLine, idxAssIdx As Integer
Public sDelimeter As String

Private Sub OZ_InitUtility()
    Dim oFSO As Object
    Dim oFolder As Object
    Dim directory As String
    Dim FileName As String
    
    directory = "\\THBAN1SRV008\PPT\039_Scheduling Planning\NS\lib"
    FileName = "OZ_modUtilz"
    
    'Remove if is exists
    Dim VBProj
    Dim VBComp

    Set VBProj = ActiveWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        If StrComp(UCase(VBComp.Name), UCase(FileName)) = 0 Then
            VBProj.VBComponents.Remove VBComp
        End If
    Next
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(directory)
    
    FileName = FileName & ".bas"
    directory = directory & "\" & FileName
    ActiveWorkbook.VBProject.VBComponents.Import directory
    If Err.Number <> 0 Then
        Call MsgBox("Failed to import " & FileName, vbCritical)
    End If
End Sub

Private Function vInitPara()
    '*** WAX CELL ***'
    idxLine = 1
    idxCapHrs = 2
    idxWAHrs = 3
    idxLRHrs = 4
    idxSPHrs = 5
    idxGM = 6
    idxQty = 7
    
    '*** ITEM INFO ***'
    idxFixedLine = 1
    idxScatter = 2
    sDelimeter = "|"
    
    '*** Allocated info ***'
    idxAssLine = 1
    idxAssIdx = 2
End Function

Sub btnAllocation_Click()
    ThisWorkbook.Save
    Dim arrWaxCells() As Variant
    Dim dictItemInfo As Dictionary
    
    Call vInitPara
    If Not bPrepareInfo(arrWaxCells, dictItemInfo) Then
        Call MsgBox("Failed to prepare necessary information. Contact scheduling team!!", vbCritical)
        End
    End If
    If Not bAllocation(arrWaxCells, dictItemInfo) Then
        Call MsgBox("Failed to allocation. Contact scheduling team!!", vbCritical)
        End
    End If
    Call MsgBox("Done", vbInformation)
End Sub

Private Function bAllocation(ByRef arrWaxCells As Variant, ByVal dictItemInfo As Dictionary) As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dLoop, dRows, dMaxCells, dOpenCells As Double
    Dim dCT, dIsLR, dIsSP, dGMQty, dQty As Double
    Dim dictItemAllocate As Dictionary
    Dim sItemID As String
    Dim arrAssLine(), arrAvaCells() As Variant
    Dim idxFind, idxLowest As Integer
    
    Set dictItemAllocate = New Dictionary
    Call vProdSort
    [ProductionOrders_Display[TargetWaxCell]].Value = ""
    Set ws = ThisWorkbook.Sheets("ProductionOrders")
    Set tbl = ws.ListObjects("ProductionOrders_Display")
    dRows = tbl.DataBodyRange.Rows.count
    ReDim arrAssLine(1 To dRows)
    dOpenCells = UBound(arrWaxCells)
    For dLoop = 1 To dRows
'        If dLoop = 259 Then
'            Debug.Print "AA"
'        End If
        sItemID = CStr(tbl.ListColumns("ItemId").DataBodyRange(dLoop).Value)
        dCT = CDbl(tbl.ListColumns("ProductionHour").DataBodyRange(dLoop).Value)
        dIsLR = CDbl(tbl.ListColumns("IsLongRoute").DataBodyRange(dLoop).Value)
        dIsSP = CDbl(tbl.ListColumns("IsSparePart").DataBodyRange(dLoop).Value)
        dGMQty = CDbl(tbl.ListColumns("GMQty").DataBodyRange(dLoop).Value)
        dQty = CDbl(tbl.ListColumns("QtySched").DataBodyRange(dLoop).Value)
        dMaxCells = CDbl(tbl.ListColumns("MaximumWaxCellAllocation").DataBodyRange(dLoop).Value)
        
'        If tbl.ListColumns("ProdId").DataBodyRange(dLoop).Value = "207198092" Then
'            Debug.Print "AA"
'        End If
        If dMaxCells > dOpenCells Then
            dMaxCells = dOpenCells
        End If
        If Not dictItemAllocate.Exists(sItemID) Then
            ReDim arrAvaCells(1 To dMaxCells, idxAssLine To idxAssIdx)
            'Check is it fixed line
            If dictItemInfo.Exists(sItemID) Then
                Call bBuildAvaCellsWithFixedLine(arrAvaCells, dictItemInfo(sItemID), arrWaxCells)
            End If
        Else
            arrAvaCells = dictItemAllocate(sItemID)
        End If
        'Get index for focus lowest load
        idxFind = iGetLoadIndex(dIsLR, dIsSP, dGMQty)
        If Not OZ_IsArrayEmpty(arrAvaCells) Then
            'If RM available and have room to assign to wax cell
            If bIsCellAvailable(arrAvaCells) Then
                idxLowest = iGetLowestLoad(arrWaxCells, idxFind, dCT)
            Else
                idxLowest = iGetLowestLoad(arrWaxCells, idxFind, dCT, True, arrAvaCells)
            End If
            If idxLowest > 0 Then
                'Update line load hours
                Call vUpdLineInfo(arrWaxCells, idxLowest, dCT, dGMQty, dQty, dIsLR, dIsSP)
                'Update assigned line
                Call vUpdAvaCells(arrAvaCells, arrWaxCells(idxLowest, idxLine), idxLowest)
                dictItemAllocate(sItemID) = arrAvaCells
                Erase arrAvaCells
                
                'Result line
                arrAssLine(dLoop) = arrWaxCells(idxLowest, idxLine)
            End If
        End If
    Next dLoop
    [ProductionOrders_Display[TargetWaxCell]].Value = Application.WorksheetFunction.Transpose(arrAssLine)
    ws.Select
    bAllocation = True
End Function

Private Function bBuildAvaCellsWithFixedLine(ByRef arrAvaCells As Variant, ByVal arrItemInfo As Variant, ByVal arrWaxCells As Variant) As Boolean
    Dim dNrFixedLine, dMaxCells, dLoop, idxFound, dAvaIdx As Double
    Dim sFixedLine As String
    Dim bRedim As Boolean
    
    sFixedLine = arrItemInfo(idxFixedLine)
    dNrFixedLine = OZ_dCountByDelimeter(sFixedLine, sDelimeter)
    dMaxCells = UBound(arrAvaCells)
    If dNrFixedLine > 0 And dNrFixedLine < dMaxCells Then
        dMaxCells = dNrFixedLine
        Erase arrAvaCells
        dAvaIdx = 1
        Dim arrFixed As Variant
        arrFixed = Split(sFixedLine, sDelimeter)
        For dLoop = 1 To dMaxCells
            idxFound = OZ_iGetMatchIndexIn2DArray(arrWaxCells, idxLine, arrFixed(dLoop - 1))
            If idxFound > 0 Then
                If Not bRedim Then
                    ReDim arrAvaCells(1 To 1, idxAssLine To idxAssIdx)
                    bRedim = True
                Else
                    dAvaIdx = UBound(arrAvaCells) + 1
                    arrAvaCells = OZ_ReDimPreserve(arrAvaCells, dAvaIdx, CDbl(idxAssIdx))
                End If
                arrAvaCells(dAvaIdx, idxAssLine) = arrFixed(dLoop - 1)
                arrAvaCells(dAvaIdx, idxAssIdx) = idxFound
            End If
        Next dLoop
    End If
    
End Function

Private Function vUpdAvaCells(ByRef arrAvaCells As Variant, ByVal sAssLine As String, ByVal iAssIdx As Integer)
    Dim dLoop, dBound As Double
    dBound = UBound(arrAvaCells)
    For dLoop = 1 To dBound
        If IsEmpty(arrAvaCells(dLoop, idxAssLine)) Or OZ_bStrComp(sAssLine, arrAvaCells(dLoop, idxAssLine)) Then
            arrAvaCells(dLoop, idxAssLine) = sAssLine
            arrAvaCells(dLoop, idxAssIdx) = iAssIdx
            Exit For
        End If
    Next dLoop
End Function

Private Function vUpdLineInfo(ByRef arrWaxCells As Variant, ByVal idxUpd As Integer, ByVal dCT As Double, ByVal dGMQty As Double, ByVal dQty As Double, ByVal dIsLR As Double, ByVal dIsSP As Double)
    arrWaxCells(idxUpd, idxWAHrs) = arrWaxCells(idxUpd, idxWAHrs) + dCT
    If dIsLR = 1 Then
        arrWaxCells(idxUpd, idxLRHrs) = arrWaxCells(idxUpd, idxLRHrs) + dCT
    End If
    If dIsSP = 1 Then
        arrWaxCells(idxUpd, idxSPHrs) = arrWaxCells(idxUpd, idxSPHrs) + dCT
    End If
    arrWaxCells(idxUpd, idxGM) = arrWaxCells(idxUpd, idxGM) + dGMQty
    arrWaxCells(idxUpd, idxQty) = arrWaxCells(idxUpd, idxQty) + dQty
End Function

Private Function iGetLowestLoad(ByVal arrWaxCells As Variant, ByVal idxFind As Integer, ByVal dWALoad As Double, Optional ByVal bIsRestrict As Boolean = False, Optional ByVal arrAvaCells As Variant = "") As Integer
    Dim dLoop, dBound As Double
    Dim dLoad, dLowest As Double
    Dim iRes As Integer
    Dim dNewLoad, dCapLim As Double
    
    dLoad = 0
    dLowest = 1000000
    If Not bIsRestrict Then
        dBound = UBound(arrWaxCells)
        For dLoop = 1 To dBound
            dLoad = arrWaxCells(dLoop, idxFind)
            dCapLim = arrWaxCells(dLoop, idxCapHrs)
            dNewLoad = arrWaxCells(dLoop, idxWAHrs) + dWALoad
            If dLowest > dLoad And dNewLoad <= dCapLim Then
                dLowest = dLoad
                iRes = dLoop
            End If
        Next dLoop
    Else
        dBound = UBound(arrAvaCells)
        For dLoop = 1 To dBound
            dLoad = arrWaxCells(arrAvaCells(dLoop, idxAssIdx), idxFind)
            dCapLim = arrWaxCells(arrAvaCells(dLoop, idxAssIdx), idxCapHrs)
            dNewLoad = arrWaxCells(arrAvaCells(dLoop, idxAssIdx), idxWAHrs) + dWALoad
            If dLowest > dLoad And dNewLoad <= dCapLim Then
                dLowest = dLoad
                iRes = arrAvaCells(dLoop, idxAssIdx)
            End If
        Next dLoop
    End If
    
    
    
    
    iGetLowestLoad = iRes
End Function

Private Function iGetLoadIndex(ByVal dIsLR As Double, ByVal dIsSP As Double, ByVal dGMQty As Double) As Integer
    Dim iRes As Integer
    iRes = idxWAHrs
    If dIsLR = 1 Then
        iRes = idxLRHrs
    ElseIf dIsSP = 1 Then
        iRes = idxSPHrs
    ElseIf dGMQty > 0 Then
        iRes = idxGM
    End If
    iGetLoadIndex = iRes
End Function

Private Function bIsCellAvailable(ByVal arrAvaCells As Variant) As Boolean
    Dim bRes As Boolean
    Dim dLoop, dBound As Double
    dBound = UBound(arrAvaCells)
    For dLoop = 1 To dBound
        If IsEmpty(arrAvaCells(dLoop, idxAssLine)) Then
            bRes = True
            Exit For
        End If
    Next dLoop
    bIsCellAvailable = bRes
End Function

Private Function bPrepareInfo(ByRef arrWaxCells As Variant, ByRef dictItemInfo As Dictionary) As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dRows, dLoop, dIdx As Double
    Dim dIsOpen, dVal As Double
    Dim sVal As String
    
    
    'Open wax cell
    Set ws = ThisWorkbook.Sheets("Configuration")
    Set tbl = ws.ListObjects("t_config_WaxCell")
    dRows = tbl.DataBodyRange.Rows.count
    For dLoop = 1 To dRows
        dIsOpen = CDbl(tbl.ListColumns("Active").DataBodyRange(dLoop).Value)
        If dIsOpen = 1 Then
            If OZ_IsArrayEmpty(arrWaxCells) Then
                dIdx = 1
                ReDim arrWaxCells(dIdx To dIdx, idxLine To idxQty)
            Else
                dIdx = dIdx + 1
                arrWaxCells = OZ_ReDimPreserve(arrWaxCells, dIdx, CDbl(idxQty))
            End If
            'Initial value
            sVal = Trim(CStr(tbl.ListColumns("Wax Cell").DataBodyRange(dLoop).Value))
            arrWaxCells(dIdx, idxLine) = sVal
            dVal = CDbl(tbl.ListColumns("Total Hours/Week per cell").DataBodyRange(dLoop).Value)
            arrWaxCells(dIdx, idxCapHrs) = dVal
            arrWaxCells(dIdx, idxWAHrs) = 0
            arrWaxCells(dIdx, idxLRHrs) = 0
            arrWaxCells(dIdx, idxSPHrs) = 0
            arrWaxCells(dIdx, idxGM) = 0
            arrWaxCells(dIdx, idxQty) = 0
        End If
    Next dLoop
    'Item fixed line
    Set ws = ThisWorkbook.Sheets("FixedLine&Scatter")
    Set tbl = ws.ListObjects("tblFixedLine")
    Set dictItemInfo = New Dictionary
    If Not OZ_bIsTableEmpty("tblFixedLine") Then
        dRows = tbl.DataBodyRange.Rows.count
        Dim arrItemInfo As Variant
        Dim sItemID, sLine As String
        For dLoop = 1 To dRows
            sItemID = CStr(tbl.ListColumns("ItemID").DataBodyRange(dLoop).Value)
            If Not dictItemInfo.Exists(sItemID) Then
                ReDim arrItemInfo(idxFixedLine To idxScatter)
                sLine = ""
            Else
                arrItemInfo = dictItemInfo(sItemID)
                sLine = arrItemInfo(idxFixedLine)
            End If
            sLine = sLine & sDelimeter & CStr(tbl.ListColumns("Line").DataBodyRange(dLoop).Value)
            arrItemInfo(idxFixedLine) = sLine
            arrItemInfo(idxScatter) = 0
            dictItemInfo(sItemID) = arrItemInfo
        Next dLoop
    End If
    'Item scatter
    Set tbl = ws.ListObjects("tblScatter")
    If Not OZ_bIsTableEmpty("tblScatter") Then
        dRows = tbl.DataBodyRange.Rows.count
        For dLoop = 1 To dRows
            sItemID = CStr(tbl.ListColumns("ItemID").DataBodyRange(dLoop).Value)
            If Not dictItemInfo.Exists(sItemID) Then
                ReDim arrItemInfo(idxFixedLine To idxScatter)
            Else
                arrItemInfo = dictItemInfo(sItemID)
            End If
            arrItemInfo(idxScatter) = 1
            arrItemInfo(idxFixedLine) = arrItemInfo(idxFixedLine)
            dictItemInfo(sItemID) = arrItemInfo
        Next dLoop
    End If
    
    'Clear duplicate
    Dim vKey As Variant
    For Each vKey In dictItemInfo.Keys()
        arrItemInfo = dictItemInfo(vKey)
        sLine = arrItemInfo(idxFixedLine)
        If Len(sLine) > 0 Then
            sLine = sRemoveDupFromDelimeter(sLine)
            arrItemInfo(idxFixedLine) = sLine
            dictItemInfo(vKey) = arrItemInfo
        End If
    Next
    
    bPrepareInfo = True
End Function

Private Function sRemoveDupFromDelimeter(ByVal sFixedLine As Variant) As String
    Dim sRes, sLine As String
    Dim arrSplit As Variant
    Dim dLoop As Double
    Dim dictLine As Dictionary
    Set dictLine = New Dictionary
    arrSplit = Split(sFixedLine, sDelimeter)
    For dLoop = LBound(arrSplit, 1) To UBound(arrSplit, 1)
        sLine = arrSplit(dLoop)
        If Len(sLine) > 0 And Not dictLine.Exists(sLine) Then
            If Len(sRes) > 0 Then
                sRes = sRes & sDelimeter
            End If
            sRes = sRes & sLine
            dictLine.Add sLine, 1
        End If
    Next dLoop
    sRemoveDupFromDelimeter = sRes
End Function

Private Function vProdSort()
    ActiveWorkbook.Worksheets("ProductionOrders").ListObjects("ProductionOrders_Display").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ProductionOrders").ListObjects("ProductionOrders_Display").Sort.SortFields.Add2 Key:=Range( _
        "ProductionOrders_Display[IsLongRoute]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ProductionOrders").ListObjects("ProductionOrders_Display").Sort.SortFields.Add2 Key:=Range( _
        "ProductionOrders_Display[IsSparePart]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ProductionOrders").ListObjects("ProductionOrders_Display").Sort.SortFields.Add2 Key:=Range( _
        "ProductionOrders_Display[GMQty]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ProductionOrders").ListObjects("ProductionOrders_Display").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Function






















