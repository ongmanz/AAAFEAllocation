Attribute VB_Name = "CustomMod"
Function assign_Arr(ByVal a As Range)

arr_size = a.count

Dim arr() As Variant
ReDim arr(1 To arr_size)

For i = LBound(arr) To UBound(arr)
    If a.Item(i) = " " Then
        arr(i) = Empty
    Else
        arr(i) = a.Item(i)
    End If
Next
assign_Arr = arr
End Function


Function TableToDictionary(ByVal tar_table As ListObject, Optional KeyCol = "")

Dim dict_temp As scripting.Dictionary
Set dict_temp = New scripting.Dictionary

On Error Resume Next
tar_table.AutoFilter.ShowAllData

If KeyCol <> "" Then
    With tar_table.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=tar_table.ListColumns(KeyCol).DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If

For Each hd In tar_table.HeaderRowRange
    arr = assign_Arr(tar_table.ListColumns(hd.Value).DataBodyRange)
    dict_temp.Add hd.Value, arr
Next

Set TableToDictionary = dict_temp
End Function


Function dictBinary_Search(ByVal tar_dict As scripting.Dictionary, ByVal tar_col, ByVal tar_var)

X = tar_dict(tar_col)

low = LBound(X)
high = UBound(X)

Do While low <= high
    midpoint = Application.WorksheetFunction.RoundDown((low + high) / 2, 0)
    If X(midpoint) > tar_var Then
        high = midpoint - 1
    ElseIf X(midpoint) = tar_var Then
        dictBinary_Search = midpoint
        Exit Do
    Else
        low = midpoint + 1
    End If
Loop

If low > high Then dictBinary_Search = -1

End Function

Function dictDescribe(ByVal tar_dict As scripting.Dictionary)

Dim dict_temp As scripting.Dictionary
Set dict_temp = New scripting.Dictionary

dict_temp.Add "STAT", Array("Data Type", "Count Rows", "Count Value", "Max", "Q3", "Q2", "Q1", "Min", "Range", "IQR", "Var_S")

Dim X(0 To 10) 'Describe,Count Rows, Count
Dim cArr() As Variant
For Each vKey In tar_dict.Keys
    'Check Datatype
    dateChk = True
    NumericChk = True
    
    For i = LBound(tar_dict(vKey)) To UBound(tar_dict(vKey))
        If IsEmpty(tar_dict(vKey)(i)) Then
            dateChk = dateChk * True
            NumericChk = NumericChk * True
        Else
            dateChk = dateChk * IsDate(tar_dict(vKey)(i))
            NumericChk = NumericChk * IsNumeric(tar_dict(vKey)(i))
        End If
        
    Next
    Erase X
    Erase cArr
    
    If dateChk <> 0 Or NumericChk <> 0 Then
        If dateChk <> 0 Then
            numrow = UBound(tar_dict(vKey)) - LBound(tar_dict(vKey)) + 1
            ReDim cArr(1 To numrow)
            
            For i = LBound(tar_dict(vKey)) To UBound(tar_dict(vKey))
                If Not IsEmpty(tar_dict(vKey)(i)) Then
                    cArr(i) = CDbl(tar_dict(vKey)(i))
                Else
                    cArr(i) = Empty
                End If
            Next
            X(0) = "Date"
            X(1) = UBound(tar_dict(vKey)) - LBound(cArr) + 1
            X(2) = Application.WorksheetFunction.count(cArr)
            X(3) = Application.WorksheetFunction.Max(cArr)
            On Error Resume Next
            X(4) = Application.WorksheetFunction.Quartile_Inc(cArr, 3)
            On Error Resume Next
            X(5) = Application.WorksheetFunction.Quartile_Inc(cArr, 2)
            On Error Resume Next
            X(6) = Application.WorksheetFunction.Quartile_Inc(cArr, 1)
            X(7) = Application.WorksheetFunction.Min(cArr)
            X(8) = CDbl(X(3) - X(7))
            X(9) = CDbl(X(4) - X(6))
            X(10) = CDbl(Application.WorksheetFunction.Var_S(cArr))
             dict_temp.Add vKey, X
        Else
            X(0) = "Numeric"
            X(1) = UBound(tar_dict(vKey)) - LBound(tar_dict(vKey)) + 1
            X(2) = Application.WorksheetFunction.count(tar_dict(vKey))
            X(3) = Application.WorksheetFunction.Max(tar_dict(vKey))
            On Error Resume Next
            X(4) = Application.WorksheetFunction.Quartile_Inc(tar_dict(vKey), 3)
            On Error Resume Next
            X(5) = Application.WorksheetFunction.Quartile_Inc(tar_dict(vKey), 2)
            On Error Resume Next
            X(6) = Application.WorksheetFunction.Quartile_Inc(tar_dict(vKey), 1)
            X(7) = Application.WorksheetFunction.Min(tar_dict(vKey))
            X(8) = X(3) - X(7)
            X(9) = X(4) - X(6)
            X(10) = Application.WorksheetFunction.Var_S(tar_dict(vKey))
             dict_temp.Add vKey, X
        End If
    Else
        X(0) = "String"
        X(1) = UBound(tar_dict(vKey)) - LBound(tar_dict(vKey)) + 1
        X(2) = Application.WorksheetFunction.CountA(tar_dict(vKey))

        dict_temp.Add vKey, X
    End If
Next

Set dictDescribe = dict_temp

End Function

Sub DictToRng(ByVal tar_dict As scripting.Dictionary, ByVal tar_rng As Range)

Set the_ws = tar_rng.Parent

numcol = tar_dict.count
numrow = UBound(tar_dict(tar_dict.Keys(0))) - LBound(tar_dict(tar_dict.Keys(0))) + 1 ' plusheader

j = 0
For Each vKey In tar_dict.Keys
    tar_rng.Offset(0, j).Value = vKey
    the_ws.Range(tar_rng.Offset(1, j), tar_rng.Offset(numrow, j)) = Application.WorksheetFunction.Transpose(tar_dict(vKey))
    j = j + 1
Next

End Sub

Function isoYearWeek(ByVal dte As Date)

isoyearnum = Year(dte)
isoweek = WorksheetFunction.IsoWeekNum(dte)

If isoweek = 1 Then
    If Month(dte) <> 1 Then
        isoyearnum = Year(dte) + 1
    End If
End If
If isoweek >= 52 Then
    If Month(dte) = 1 Then
        isoyearnum = Year(dte) - 1
    End If
End If
isoYearWeek = isoyearnum * 100 + isoweek

End Function

Sub test()

Set a = TableToDictionary(ThisWorkbook.Sheets(1).ListObjects(1))

Set b = dictDescribe(a)

Call DictToRng(b, [B10])


End Sub


