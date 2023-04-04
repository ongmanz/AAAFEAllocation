Attribute VB_Name = "OZ_modUtilz"
Public Const PATH = "\\THBAN1SRV008\PPT\039_Scheduling Planning\NS\lib"

Public Function OZ_sGetCurrentFileName(ByVal bExtension As Boolean) As String
    Dim twb As Workbook
    Dim sName  As String
    Set twb = ThisWorkbook
    
    sName = twb.Name
    If Not bExtension Then
    End If
    OZ_sGetCurrentFileName = sName
End Function

Public Function OZ_MsgBox(ByVal sText As String, ByVal sTitle As String, Optional ByVal iType As Integer = 1)
    Call MsgBox(sText, iType, sTitle)
End Function

Public Function OZ_dRound(ByRef dNumber As Double, ByRef dDigits As Double) As Double
    OZ_dRound = Round(dNumber, dDigits)
End Function

Public Function OZ_bTranspose(ByRef ws As Worksheet, ByVal dCol As Double, ByVal dStartRow As Double, ByVal arrVal As Variant, Optional ByVal idx2D As Integer = 0) As Boolean 'Sheet to apply, column (.databodyrange.column), row (.databodyrange.row), arrValue, [In case 2D array sent index to apply]
    Dim bRes As Boolean
    If IsArray(arrIn) Or IsEmpty(arrIn) Then
        Dim arrVal2 As Variant
        If idx2D > 0 Then 'Other index
            Call OZ_Copy2DArrayToArray(arrVal, arrVal2, idx2D)
        Else
            arrVal2 = arrVal
        End If
        Dim dLoop, dLimitIdx, dRound, dSRows, dLRows As Double
        Dim arrCopy As Variant
        dLimitIdx = 10000
        bRes = True
        dRound = Int(UBound(arrVal) / dLimitIdx)
        If UBound(arrVal) Mod dLimitIdx > 0 Then
            dRound = dRound + 1
        End If
        dSRows = dStartRow
        For dLoop = 1 To dRound
            If dLoop = dRound Then  'Last round
                dLRows = dSRows + (UBound(arrVal) Mod dLimitIdx) - 1
            Else
                dLRows = dSRows + dLimitIdx - 1
            End If
            If OZ_bArrRangeCopy(arrVal2, dSRows - dStartRow + 1, dLRows - dSRows + 1, arrCopy) Then
                ws.Select
                ws.Range(Cells(dSRows, dCol), Cells(dLRows, dCol)) = Application.Transpose(arrCopy)
                dSRows = dLRows + 1
            Else
                bRes = False
                Exit For
            End If
        Next dLoop
    End If
    OZ_bTranspose = bRes
End Function

Public Function OZ_bMatchString(ByVal sOri As String, ByVal sChk As String, Optional sDelimiter As String = "") As Boolean
    Dim bRes As Boolean
    If OZ_bStrComp(sDelimiter, "") Then 'No delimiter
        bRes = OZ_bStrComp(sOri, sChk)
    Else 'Need to split
        Dim arrSplit, arrVal As Variant
        arrSplit = Split(sOri, sDelimiter)
        For Each arrVal In arrSplit
            If OZ_bStrComp(arrVal, sChk) Then
                bRes = True
                Exit For
            End If
        Next
    End If
    OZ_bMatchString = bRes
End Function

Public Function OZ_bArrRangeCopy(ByRef arrIn As Variant, ByVal dSIdx As Double, ByVal dLength As Double, ByRef arrOut As Variant) As Boolean
    Dim bRes As Boolean
    If IsArray(arrIn) Or IsEmpty(arrIn) Then
        Dim dLoop, dRows As Double
        dRows = dLength
        ReDim arrOut(1 To dLength)
        For dLoop = 1 To dRows
            arrOut(dLoop) = arrIn(dSIdx + dLoop - 1)
        Next dLoop
        bRes = True
    End If
    OZ_bArrRangeCopy = bRes
End Function

Public Function OZ_InitArrayDefVal(ByRef arrIn As Variant, Optional ByVal sVal As String = "")
    If IsEmpty(arrIn) Then
        Exit Function
    End If
    
    Dim dLoop, dRows As Double
    dRows = UBound(arrIn)
    For dLoop = 1 To dRows
        arrIn(dLoop) = sVal
    Next dLoop
End Function

Public Function OZ_RemoveAllFilters()
    Dim lo As ListObject
    Dim af As AutoFilter
    'Removes all active filters from tables
    For Each sht In ActiveWorkbook.Sheets
        For Each lo In sht.ListObjects
            currFilter = lo.ShowAutoFilter
            lo.ShowAutoFilter = False
            lo.ShowAutoFilter = True
            lo.ShowAutoFilter = currFilter
        Next
    Next
End Function

Public Function OZ_Copy2DArrayToArray(ByRef arr2D As Variant, ByRef arr1D As Variant, ByVal idx2DCopy As Integer)
    If IsEmpty(arr2D) Then
        Exit Function
    End If
    
    Dim dLoop, dRows As Double
    dRows = UBound(arr2D)
    ReDim arr1D(1 To dRows)
    For dLoop = 1 To dRows
        arr1D(dLoop) = arr2D(dLoop, idx2DCopy)
    Next dLoop
 End Function
 
 Public Function OZ_dGetLower(ByVal dVal1 As Double, ByVal dVal2 As Double) As Double
    Dim dRes As Double
    If dVal1 < dVal2 Then
        dRes = dVal1
    Else
        dRes = dVal2
    End If
    OZ_dGetLower = dRes
 End Function
 
  Public Function OZ_dGetHigher(ByVal dVal1 As Double, ByVal dVal2 As Double) As Double
    Dim dRes As Double
    If dVal1 > dVal2 Then
        dRes = dVal1
    Else
        dRes = dVal2
    End If
    OZ_dGetHigher = dRes
 End Function
 
 Public Function OZ_dGetLowerNearest(ByVal dTarget As Double, ByVal dStart As Double, ByVal dStep As Double) As Double
    Dim dRes As Double
    Dim dLoop, dCount As Double
    dCount = Application.WorksheetFunction.RoundDown(dTarget / dStep, 0)
    For dLoop = 1 To dCount
        If dRes + dStep > dStart Then
            Exit For
        End If
        dRes = dRes + dStep
    Next dLoop
    OZ_dGetLowerNearest = dRes
 End Function
 
 Public Function OZ_DeleteFile(ByVal sFilePath As String)
    ' First remove readonly attribute, if set
    SetAttr sFilePath, vbNormal
    ' Then delete the file
    Kill sFilePath
 End Function
 
 Public Function OZ_RenameFile(ByVal sFrom As String, ByVal sTo As String, Optional ByVal bOverWrite As Boolean = False)
    If OZ_IsFileExists(sTo) Then
        If bOverWrite Then
            Call OZ_DeleteFile(sTo)
        End If
    End If
    Name sFrom As sTo
 End Function
 
 Public Function OZ_IsFileExists(ByVal sFilePath As String) As Boolean
    Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(sFilePath)
    On Error GoTo 0
    If TestStr = "" Then
        OZ_IsFileExists = False
    Else
        OZ_IsFileExists = True
    End If
End Function
 
Public Function OZ_QuickSort(ByRef SortArray As Variant, ByVal col As Double, ByVal bAscending As Boolean, Optional L As Long = -1, Optional R As Long = -1)
    On Error Resume Next
    'Sort a 2-Dimensional array
    'Originally Posted by Jim Rech 10/20/98 Excel.Programming
    'Modified to sort on first column of a two dimensional array
    'Modified to handle a second dimension greater than 1 (or zero)
    'Modified to do Ascending or Descending

    Dim i As Long
    Dim j As Long
    Dim X, Y, mm

    If L = -1 Then
        L = LBound(SortArray, 1)
    End If
    If R = -1 Then
        R = UBound(SortArray, 1)
    End If
    If L >= R Then    ' no sorting required
        Exit Function
    End If
    
    i = L
    j = R
    X = SortArray((L + R) / 2, col)
    If bAscending Then
        While (i <= j)
            While (SortArray(i, col) < X And i < R)
                i = i + 1
            Wend
            While (X < SortArray(j, col) And j > L)
                j = j - 1
            Wend
            If (i <= j) Then
                For mm = LBound(SortArray, 2) To UBound(SortArray, 2)
                    Y = SortArray(i, mm)
                    SortArray(i, mm) = SortArray(j, mm)
                    SortArray(j, mm) = Y
                Next mm
                i = i + 1
                j = j - 1
            End If
        Wend
    Else
        While (i <= j)
            While (SortArray(i, col) > X And i < R)
                i = i + 1
            Wend
            While (X > SortArray(j, col) And j > L)
                j = j - 1
            Wend
            If (i <= j) Then
                For mm = LBound(SortArray, 2) To UBound(SortArray, 2)
                    Y = SortArray(i, mm)
                    SortArray(i, mm) = SortArray(j, mm)
                    SortArray(j, mm) = Y
                Next mm
                i = i + 1
                j = j - 1
            End If
        Wend
    End If
    If (L < j) Then Call OZ_QuickSort(SortArray, col, bAscending, L, j)
    If (i < R) Then Call OZ_QuickSort(SortArray, col, bAscending, i, R)
End Function

Public Function OZ_bStrComp(ByVal sVal1 As String, ByVal sVal2 As String) As Boolean
    Dim bRes  As Boolean
    bRes = False
    If StrComp(UCase(Trim(sVal1)), UCase(Trim(sVal2)), vbTextCompare) = 0 Then
        bRes = True
    End If
    OZ_bStrComp = bRes
End Function

Public Function OZ_bSetDictDataFromTable(ByVal objTable As ListObject, ByVal sKeyField As String, ByVal sValField As String, ByRef dictIn As scripting.Dictionary) As Boolean
    Dim bRes As Boolean
    Dim dictRes As scripting.Dictionary
    Set dictRes = New scripting.Dictionary
    bRes = False
    If Not OZ_bIsTableEmpty(objTable.Name) Then
        Dim dLoop, dRows As Double
        Dim sKey, sVal As String
        dRows = objTable.DataBodyRange.Rows.count
        For dLoop = 1 To dRows
            sKey = Trim(CStr(objTable.ListColumns(sKeyField).DataBodyRange(dLoop).Value))
            sVal = Trim(CStr(objTable.ListColumns(sValField).DataBodyRange(dLoop).Value))
            If Not dictRes.Exists(sKey) Then
                'Add
                dictRes.Add sKey, sVal
            End If
        Next dLoop
        Set dictIn = dictRes
    End If
    OZ_bSetDictDataFromTable = dRes
End Function

Public Function OZ_vSetAuthorName(ByVal sName As String)
    ActiveWorkbook.BuiltinDocumentProperties("Author") = (sName)
End Function

Public Function OZ_vDeleteTableRows(ByRef Table As ListObject)
    On Error Resume Next
    '~~> Clear Header Row `IF` it exists
    Table.DataBodyRange.Rows(1).ClearContents
    '~~> Delete all the other rows `IF `they exist
    Table.DataBodyRange.Offset(1, 0).Resize(Table.DataBodyRange.Rows.count - 1, Table.DataBodyRange.Columns.count).Rows.Delete
    On Error GoTo 0
End Function

Function OZ_QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs
    
    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Function
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Function
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Function
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call OZ_QuickSortArray(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call OZ_QuickSortArray(SortArray, i, lngMax, lngColumn)

End Function

Public Function OZ_bIsExistIn2DArray(ByRef arrIn As Variant, ByVal iChkIdx As Integer, ByVal vVal As Variant) As Boolean
    Dim bRes As Boolean
    Dim dLoop As Double
    
    bRes = False
    If IsArray(arrIn) Then
        For dLoop = LBound(arrIn) To UBound(arrIn)
            If OZ_bStrComp(arrIn(dLoop, iChkIdx), vVal) Then
                bRes = True
                Exit For
            End If
        Next dLoop
    End If
       
    OZ_bIsExistIn2DArray = bRes
End Function

Public Function OZ_bIsInArray(ByVal stringToBeFound As String, ByRef arr As Variant, Optional ByRef iIdxFound As Integer) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If OZ_bStrComp(arr(i), stringToBeFound) Then
            iIdxFound = i
            OZ_bIsInArray = True
            Exit Function
        End If
    Next i
    OZ_bIsInArray = False
End Function

Public Function OZ_iGetMatchIndexIn2DArray(ByRef arrIn As Variant, ByVal iChkIdx As Integer, ByVal vVal As Variant) As Integer
    Dim iRes As Integer
    Dim dLoop As Double
    If IsArray(arrIn) Then
        For dLoop = LBound(arrIn) To UBound(arrIn)
            If OZ_bStrComp(arrIn(dLoop, iChkIdx), vVal) Then
                iRes = dLoop
                Exit For
            End If
        Next dLoop
    End If
    OZ_iGetMatchIndexIn2DArray = iRes
End Function

Public Function OZ_sGetWorkingShift(ByVal dTime As Double) As String
    Dim sRes As String
    
    If dTime >= 0.3333 And dTime <= 0.5 Then
        sRes = "01"
    ElseIf dTime >= 0.5417 And dTime <= 0.7292 Then
        sRes = "02"
    ElseIf dTime >= 0.75 And dTime <= 0.875 Then
        sRes = "03"
    ElseIf dTime >= 0.9167 And dTime <= 0.0625 Then
        sRes = "04"
    Else
        sRes = "05"
    End If
    
    OZ_sGetWorkingShift = sRes
End Function

Public Function OZ_bIsDayShift(ByVal dDateTime As Double) As Boolean
    Dim dTime As Double
    Dim bRes As Boolean
    Dim dDayShiftTime, dNightShiftTime As Double
    
    dDayShiftTime = 0.3332
    dNightShiftTime = 0.9165
'    dTime = dDateTime - Int(dDateTime)
    dTime = (dDateTime) - Application.WorksheetFunction.RoundDown(dDateTime, 0)
    bRes = True
    If dTime >= dNightShiftTime Or dTime < dDayShiftTime Then
        bRes = False
    End If
    
    OZ_bIsDayShift = bRes
End Function

Public Function OZ_bIsNightShift(ByVal dDateTime As Double) As Boolean
    OZ_bIsNightShift = Not OZ_bIsDayShift(dDateTime)
End Function

' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Function OZ_ExportVisualBasicCode(Optional ByVal bDateFolder As Boolean = True)
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim PATH As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    directory = ActiveWorkbook.PATH & "\VisualBasic"
    count = 0
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    
    If bDateFolder Then
        directory = directory & "\" & Format(Date, "YYYYMMDD")
        If Not fso.FolderExists(directory) Then
            Call fso.CreateFolder(directory)
        End If
    End If
    
    'Copy original Excel file to backup path
    
    Call fso.CopyFile(ActiveWorkbook.FullName, directory & "\" & ActiveWorkbook.Name, True)
    
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
        On Error Resume Next
        Err.Clear
        
        PATH = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(PATH)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & PATH, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & PATH
        End If
        On Error GoTo 0
    Next
    
'    MsgBox "Successfully exported " & CStr(count) & " VBA files to " & directory
    Call OZ_vSetAuthorName(ThisWorkbook.WriteReservedBy)
End Function

Public Function OZ_ImportVisualBasicCode()
 
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    Dim directory As String
     
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(ActiveWorkbook.PATH & "\VisualBasic")
    For Each oFile In oFolder.Files
        directory = ActiveWorkbook.PATH & "\VisualBasic\" & oFile.Name
        ActiveWorkbook.VBProject.VBComponents.Import directory
        If Err.Number <> 0 Then
            Call MsgBox("Failed to import " & oFile.Name, vbCritical)
        End If
    Next oFile
 
End Function

Public Function OZ_bIsTableEmpty(ByVal sTblName As String) As Boolean
    Dim bRes As Boolean
    bRes = True
    If WorksheetFunction.CountA(Range(sTblName)) > 1 Then
        bRes = False
    End If
    OZ_bIsTableEmpty = bRes
End Function

Public Function OZ_IsArrayEmpty(arr As Variant) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsArrayEmpty
    ' This Function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
    '
    ' The VBA IsArray Function indicates whether a variable is an array, but it does not
    ' distinguish between allocated and unallocated arrays. It will return TRUE for both
    ' allocated and unallocated arrays. This Function tests whether the array has actually
    ' been allocated.
    '
    ' This Function is really the reverse of IsArrayAllocated.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim LB As Long
    Dim UB As Long
    
    Err.Clear
    On Error Resume Next
    If IsArray(arr) = False Then
        ' we weren't passed an array, return True
        OZ_IsArrayEmpty = True
    End If
    
    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(arr, 1)
    If (Err.Number <> 0) Then
        OZ_IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occassion, under circumstances I
        ' cannot reliably replictate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occassions, LBound is 0 and
        ' UBoung is -1.
        ' To accomodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        LB = LBound(arr)
        If LB > UB Then
            OZ_IsArrayEmpty = True
        Else
            OZ_IsArrayEmpty = False
        End If
    End If

End Function
 
Public Function OZ_ReDimPreserve(MyArray As Variant, nNewFirstUBound As Double, nNewLastUBound As Double) As Variant
    Dim i, j As Long
    Dim nOldFirstUBound, nOldLastUBound, nOldFirstLBound, nOldLastLBound As Long
    Dim TempArray() As Variant 'Change this to "String" or any other data type if want it to work for arrays other than Variants. MsgBox UCase(TypeName(MyArray))
'---------------------------------------------------------------
'COMMENT THIS BLOCK OUT IF YOU CHANGE THE DATA TYPE OF TempArray
'    If InStr(1, UCase(TypeName(MyArray)), "VARIANT") = 0 Then
'        MsgBox "This Function only works if your array is a Variant Data Type." & vbNewLine & _
'               "You have two choice:" & vbNewLine & _
'               " 1) Change your array to a Variant and try again." & vbNewLine & _
'               " 2) Change the DataType of TempArray to match your array and comment the top block out of the Function OZ_ReDimPreserve" _
'                , vbCritical, "Invalid Array Data Type"
'        End
'    End If
'---------------------------------------------------------------
    OZ_ReDimPreserve = False
    'check if its in array first
    If Not IsArray(MyArray) Then MsgBox "You didn't pass the Function an array.", vbCritical, "No Array Detected": End
    
    'get old lBound/uBound
    nOldFirstUBound = UBound(MyArray, 1): nOldLastUBound = UBound(MyArray, 2)
    nOldFirstLBound = LBound(MyArray, 1): nOldLastLBound = LBound(MyArray, 2)
    'create new array
    ReDim TempArray(nOldFirstLBound To nNewFirstUBound, nOldLastLBound To nNewLastUBound)
    'loop through first
    For i = LBound(MyArray, 1) To nNewFirstUBound
        For j = LBound(MyArray, 2) To nNewLastUBound
            'if its in range, then append to new array the same way
            If nOldFirstUBound >= i And nOldLastUBound >= j Then
                If TypeOf MyArray(i, j) Is Object  Then
                    Set TempArray(i, j) = MyArray(i, j)
                Else
                    TempArray(i, j) = MyArray(i, j)
                End If
            End If
        Next
    Next
    'return the array redimmed
    If IsArray(TempArray) Then OZ_ReDimPreserve = TempArray
End Function

Public Function OZ_WeekNum(someDate, isoWeekYear, isoWeekNumber, isoWeekDay, isoYearWeek)
  Dim nearestThursday
  isoWeekDay = Weekday(someDate, vbMonday)
  nearestThursday = DateAdd("d", 4 - Int(isoWeekDay), someDate)
  isoWeekYear = Year(nearestThursday)
  isoWeekNumber = Int((nearestThursday - DateSerial(isoWeekYear, 1, 1)) / 7) + 1
  isoYearWeek = WorksheetFunction.Text(isoWeekYear, "0000") & WorksheetFunction.Text(isoWeekNumber, "00")
End Function

Public Function OZ_DeleteModule(ByVal sModuleName As String)
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(sModuleName)
    VBProj.VBComponents.Remove VBComp
End Function

Public Function OZ_ListFileInFolder(ByVal sPath As String)
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    
    For Each oFile In oFolder.Files
        Debug.Print oFile.Name
    Next oFile

End Function

Public Function OZ_StringCountOccurrences(strText As String, strFind As String, Optional lngCompare As VbCompareMethod) As Long
    ' Counts occurrences of a particular character or characters.
    ' If lngCompare argument is omitted, procedure performs binary comparison.
    'Testcases:
    '?StringCountOccurrences("","") = 0
    '?StringCountOccurrences("","a") = 0
    '?StringCountOccurrences("aaa","a") = 3
    '?StringCountOccurrences("aaa","b") = 0
    '?StringCountOccurrences("aaa","aa") = 1
    Dim lngPos As Long
    Dim lngTemp As Long
    Dim lngCount As Long
    If Len(strText) = 0 Then Exit Function
    If Len(strFind) = 0 Then Exit Function
    lngPos = 1
    Do
        lngPos = InStr(lngPos, strText, strFind, lngCompare)
        lngTemp = lngPos
        If lngPos > 0 Then
            lngCount = lngCount + 1
            lngPos = lngPos + Len(strFind)
        End If
    Loop Until lngPos = 0
    OZ_StringCountOccurrences = lngCount
End Function


Public Function OZ_dSumByDelimeter(ByVal sIn As String, ByVal iLen As Integer) As Double
    Dim aList() As String
    Dim dRes As Double
    Dim sEle As Variant
    Dim iCount As Integer
    
    
    If iLen <> 0 Then
        aList = Split(sIn, "|")
        For Each sEle In aList
            If iCount + 1 > iLen Then
                Exit For
            Else
                dRes = dRes + CDbl(sEle)
            End If
            iCount = iCount + 1
        Next sEle
        
        OZ_dSumByDelimeter = dRes
    Else
        OZ_dSumByDelimeter = 0
    End If
End Function

Public Function OZ_dCountByDelimeter(ByVal sIn As String, Optional ByVal sDelimeter As String = "|") As Double
    Dim aList() As String
    Dim dRes As Double
    aList = Split(sIn, sDelimeter)
    dRes = UBound(aList) + 1
    OZ_dCountByDelimeter = dRes
End Function

Public Function OZ_bIsMatchDPF(ByVal sOri As String, ByVal sCheck As String) As Boolean
    Dim dLoop, dLen As Double
    Dim sChar As String
    Dim bRes As Boolean
    
    If Len(sCheck) > Len(sOri) Then
        bRes = False
    Else
        For dLoop = 1 To Len(sCheck)
            sChar = Mid(sCheck, dLoop, 1)
            bRes = True
            If InStr(sOri, sChar) < 1 Then
                bRes = False
                Exit For
            End If
        Next
    End If
    OZ_bIsMatchDPF = bRes
End Function

Public Function OZ_sGetDifferenceDPF(ByVal sOri As String, ByVal sCheck As String, Optional sDelimiter As String = ", ", Optional bIsReplaceSkill As Boolean = True) As String
    Dim sRes, sChar As String
    Dim dLoop, dLen As Double
    Dim arrSpecialSkill, arrReplaceSkill As Variant
    Dim iIdxFound As Integer
    Dim bIsMultiple As Boolean
    
    arrSpecialSkill = Array("V")
    arrReplaceSkill = Array("EV")
    
    dLen = Len(sCheck)
    For dLoop = 1 To dLen
        sChar = Mid(sCheck, dLoop, 1)
        If InStr(sOri, sChar) < 1 Then
            If bIsReplaceSkill And OZ_bIsInArray(sChar, arrSpecialSkill, iIdxFound) Then
                sChar = arrReplaceSkill(iIdxFound)
            End If
            
            If bIsMultiple Then
                sChar = sDelimiter & sChar
            Else
                bIsMultiple = True
            End If
            sRes = sRes & sChar
        End If
    Next dLoop
    
    OZ_sGetDifferenceDPF = sRes
End Function

Public Function OZ_dCountFileInFolder(ByVal sPath As String) As Double
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim dRes As Double
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    dRes = 0
    For Each oFile In oFolder.Files
        dRes = dRes + 1
    Next oFile
    OZ_dCountFileInFolder = dRes
End Function

Public Function OZ_ExportModule(ByVal sModuleName As String, ByVal sExportPart As String, Optional ByVal bGenDate As Boolean = False)
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim PATH As String
    Dim verFilePath As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    directory = sExportPart
    count = 0
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
        
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
        On Error Resume Next
        Err.Clear
        
        If StrComp(sModuleName, VBComponent.Name, vbTextCompare) = 0 Then
            PATH = directory & "\" & VBComponent.Name & extension
            If bGenDate Then
                If OZ_IsFileExists(PATH) Then
                    verFilePath = directory & "\" & Format(Date, "YYYYMMDD") & "_" & VBComponent.Name & extension
                    'Version
                    Call OZ_RenameFile(PATH, verFilePath, True)
                End If
            End If
            Call VBComponent.Export(PATH)
            
            If Err.Number <> 0 Then
                Call MsgBox("Failed to export " & VBComponent.Name & " to " & PATH, vbCritical)
                Exit Function
            Else
                count = count + 1
                Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & PATH
            End If
            Exit For
        End If
        On Error GoTo 0
    Next
    
    MsgBox "Successfully exported " & CStr(count) & " VBA files to " & directory
End Function

Public Function OZ_ImportModule(ByVal sModuleName As String, ByVal sImportPart As String)
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    Dim directory As String
     
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sImportPart)
     
    For Each oFile In oFolder.Files
        If InStr(sModuleName, oFile.Name) = 1 Then
            directory = sImportPart & oFile.Name
            ActiveWorkbook.VBProject.VBComponents.Import directory
            If Err.Number <> 0 Then
                Call MsgBox("Failed to import " & oFile.Name, vbCritical)
            End If
            Exit For
        End If
    Next oFile
End Function

Private Function OZ_ExportUtilz()
    Call OZ_ExportModule("OZ_modUtilz", PATH, True)
End Function

Private Function OZ_ImportUtils()
    Call OZ_ImportModule("OZ_modUtilz.bas", PATH)
End Function

Sub OZ_ImportUtility()
    Call OZ_ImportUtils
End Sub

Sub OZ_ExportUtility()
    'execute this sub to backup this file to repository
    Call OZ_ExportUtilz
End Sub

Sub OZ_Test()
    Call OZ_ExportUtilz
End Sub

Private Function OZ_vXLSBToXLSX(ByVal sPath As String)
    WorkingDir = sPath
    extension = "xlsb"
    Dim fso, myFolder, fileColl, aFile, FileName, SaveName
    Dim objExcel, objWorkbook
    
    Set fso = CreateObject("Scripting.FilesystemObject")
    Set myFolder = fso.GetFolder(WorkingDir)
    Set fileColl = myFolder.Files
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = False
    objExcel.DisplayAlerts = False
    
    For Each aFile In fileColl
        ext = Right(aFile.Name, 4)
        If UCase(ext) = UCase(extension) And InStr(aFile.Name, "~$") = 0 Then
            'open excel
            FileName = Left(aFile, InStrRev(aFile, "."))
            Set objWorkbook = objExcel.Workbooks.Open(aFile)
            SaveName = FileName & "xlsx"
            objWorkbook.SaveAs SaveName, 51
            objWorkbook.Close
        End If
    Next
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    Set fso = Nothing
    Set myFolder = Nothing
    Set fileColl = Nothing
End Function

Public Function OZ_vUpdateVersionAndBuild(ByVal sRangeTxt As String)
    Dim major, minor, revision, build As Double
    Dim rngVersion As String
    'Init version
    major = 1: minor = 0: revision = 0
    build = 1
    
    rngVersion = Range(sRangeTxt)
    If Not OZ_bStrComp(rngVersion, "") Then
        Dim version, rngBuild As Variant
        version = Split(rngVersion, ".")
        major = CDbl(version(0))
        minor = CDbl(version(1))
        revision = CDbl(Left(version(2), 2))
        rngBuild = Split(version(2), "build")
        build = CDbl(rngBuild(1)) + 1
    End If
    Range(sRangeTxt) = major & "." & IIf(minor < 10, "0" & minor, minor) & "." & IIf(revision < 10, "0" & revision, revision) & " build " & build
End Function

Public Function OZ_dCount2DArrayIndex(ByVal arrIn As Variant, ByVal idxChk As Double, ByVal sVal As String) As Double
    Dim dRes As Double
    Dim dLoop, dBound As Double
    If IsArray(arrIn) Then
        dBound = UBound(arrIn)
        For dLoop = 0 To dBound
            On Error GoTo NextLoop
            If OZ_bStrComp(arrIn(dLoop, idxChk), sVal) Then
                dRes = dRes + 1
            End If
NextLoop:
        Next dLoop
    End If
    Err.Clear
    OZ_dCount2DArrayIndex = dRes
End Function
























