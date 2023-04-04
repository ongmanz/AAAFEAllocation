VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If UCase(Range("rngDEVMode")) = "TRUE" Then
        Call vUpdateVersionAndBuild
        Call OZ_ExportVisualBasicCode(False)
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If UCase(Range("rngDEVMode")) = "TRUE" Then
        Call vUpdateVersionAndBuild
        Call OZ_ExportVisualBasicCode(False)
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    If UCase(Range("rngDEVMode")) = "TRUE" Then
'        Dim answer As Integer
'        answer = MsgBox("Switch to user mode?", vbQuestion + vbYesNo + vbDefaultButton2, "Switch to user mode")
'        If answer = vbYes Then
'          Range("rngDEVMode").Value = "FALSE"
'        End If
'    End If
End Sub

Sub vUpdateVersionAndBuild()
    Dim major, minor, revision, build As Double
    Dim rngVersion As String
    'Init version
    major = 1: minor = 23: revision = 0
    build = 1
    
    rngVersion = Range("rngVersion")
    If Not OZ_bStrComp(rngVersion, "") Then
        Dim version, rngBuild As Variant
        version = Split(rngVersion, ".")
        major = CDbl(version(0))
        minor = CDbl(version(1))
        revision = CDbl(Left(version(2), 2))
        rngBuild = Split(version(2), "build")
        build = CDbl(rngBuild(1)) + 1
    End If
    Range("rngVersion") = major & "." & IIf(minor < 10, "0" & minor, minor) & "." & IIf(revision < 10, "0" & revision, revision) & " build " & build
    ActiveWorkbook.BuiltinDocumentProperties("Document version") = Range("rngVersion")
End Sub






