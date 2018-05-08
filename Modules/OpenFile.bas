Attribute VB_Name = "OpenFile"
Option Explicit
'It starts here by creation of an object of the class FileHandler
Dim FileObject As New File
Const cFilepath As String = "F5"
Public Const cRunSheet As String = "Run"
Public Const cSummarySheet As String = "Consolidated Systems"
Public Const cSystemWiseSheet As String = "System-Wise Information"
Public Const cSystemListSheet As String = "System List"
Public Const cEngineListSheet As String = "Engine List"
Public Const cUserListSheet As String = "User List"
Public Const cMetricNameSheet As String = "Metric Names"
Public Const cUserTypeNameSheet As String = "User Type Names"
Public Const cLabelLength As Integer = 20      'Length of the initial Label section of a record
Public Const cDelimiter As String * 1 = "#"    'Delimiter that separates data
Public Const cBeginCell As String = "A2"
Public Const cConsEngineResults As String = "Meas. LAW Results"
Public Const cConsEngineResultsLength As Integer = 17   'Length of the above string
Public Const cMeasLawCombination As String = "Meas. LAW Combinatio"
Public Const cMeasLawCombinationLength As Integer = 20   'Length of the above string
Public gEngineListCursor As Range
Public gUserListCursor As Range
Public gMetricName As Range
Public gUserTypeName As Range
Public gUserLawCombination As Range

Sub OpenFile()
    FileObject.GetFileChooser
    Range(cFilepath).Value = FileObject.Filepath
End Sub

Sub ProcessFile()
    Initialize
    FileObject.Filepath = Worksheets(cRunSheet).Range(cFilepath).Value
    FileObject.FetchFileContent
    FileObject.SplitIt
    FileObject.SeparateData
    ClearSheets
    FileObject.DisplayResults
    AutoFit
    Finalize
End Sub


Sub ClearSheets()
    Worksheets(cSummarySheet).Rows("2:" & Rows.Count).Delete
    Worksheets(cSystemWiseSheet).Rows("2:" & Rows.Count).Delete
    Worksheets(cEngineListSheet).Rows("2:" & Rows.Count).Delete
    Worksheets(cUserListSheet).Rows("2:" & Rows.Count).Delete
    Worksheets(cSystemListSheet).Rows("2:" & Rows.Count).Delete
    AutoFit
End Sub

Sub AutoFit()
    Worksheets(cSummarySheet).UsedRange.Columns.AutoFit
    Worksheets(cSystemWiseSheet).UsedRange.Columns.AutoFit
    Worksheets(cEngineListSheet).UsedRange.Columns.AutoFit
    Worksheets(cUserListSheet).UsedRange.Columns.AutoFit
    Worksheets(cSystemListSheet).UsedRange.Columns.AutoFit
End Sub

Sub IntializeRanges()
'Intialize Metric name range for Vlookup
    Sheets(cMetricNameSheet).Select
    Range(cBeginCell).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set gMetricName = Selection

'Initialize User Type name range for Vlookup
    Sheets(cUserTypeNameSheet).Select
    Range(cBeginCell).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set gUserTypeName = Selection
End Sub

Sub CleanMemory()
    Set gMetricName = Nothing
    Set gUserTypeName = Nothing
    Set FileObject = Nothing
End Sub

Sub Finalize()
    CleanMemory
'    Sheets(cMetricNameSheet).Visible = False
'    Sheets(cUserTypeNameSheet).Visible = False
    Worksheets(cRunSheet).Activate
    Application.ScreenUpdating = True
    ClearStatusBar
End Sub

Sub Initialize()
    Application.ScreenUpdating = False
'    Sheets(cMetricNameSheet).Visible = True
'    Sheets(cUserTypeNameSheet).Visible = True
    IntializeRanges
End Sub

Sub ClearStatusBar(Optional ClearStatusBar As Boolean)
    If ClearStatusBar Then
        Application.StatusBar = False
    Else
        Application.OnTime Now + TimeValue("00:00:03"), "'ClearStatusBar True'"
    End If
End Sub
