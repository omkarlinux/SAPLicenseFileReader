VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''The objective of this class is to handle the summary within a file '''
'''Members'''
Private pContent() As String
Private pData() As New Record
Private pUserIDArray() As New Record
Dim pFirstUserIDItem As Integer
Dim pLastUserIDItem As Integer
Private Const cMetricIDOffset As Integer = 3     'Offset to be used to get metric id
Private Const cUserTypeIDOffset As Integer = 3      'Offset to be used to get user type id
Private Const cMeasResultOffset As Integer = 4     'Offset to be used to get measurement result
Private Const cMeasStartDateOffset As Integer = 5     'Offset to be used to get measurement start date
Private Const cMeasEndDateOffset As Integer = 6     'Offset to be used to get measurement end date
Private Const cMetricIDLength As Integer = 4     'Length of metric ID
Private Const cEngineKey As String = "IS"        'Value denotes engine measurement record
Private Const cUserKey As String = "US"          'Value denotes user measurement record

'''Functions'''
Public Property Set Content(ValueCollection As Collection)
    ReDim pContent(0 To ValueCollection.Count - 1)
    For counter = 1 To ValueCollection.Count
            pContent(counter - 1) = ValueCollection(counter)
    Next
End Property

Public Sub SeparateData()
    Dim counter As Integer
    ReDim pData(UBound(pContent))
    For counter = 0 To UBound(pContent)
        pData(counter).Label = Left(pContent(counter), cLabelLength)
        pData(counter).ValueList = Split(Mid(pContent(counter), cLabelLength + 1), cDelimiter)
    Next
'Clean older variables
    Erase pContent
End Sub

Public Sub DisplayResults()
    Dim Destination As Range
    Dim pFirstTime As Integer
    Set Destination = Worksheets(cSummarySheet).Range(cBeginCell)
    Set gEngineListCursor = Worksheets(cEngineListSheet).Range(cBeginCell)
    Set gUserListCursor = Worksheets(cUserListSheet).Range(cBeginCell)
    pFirstTime = 1    'To mark the first time a row in the userID range is encountered
    
    For counter = 0 To UBound(pData)
        Destination.Offset(counter).Value = pData(counter).Label
        For internalCounter = 0 To UBound(pData(counter).ValueList)
            Destination.Offset(counter, internalCounter + 1).Value = pData(counter).ValueList(internalCounter)
        Next
        If Left(pData(counter).Label, cConsEngineResultsLength) = cConsEngineResults Then
            If pData(counter).ValueList(2) = cEngineKey Then
                PopulateEngineListSheet pData(counter)
            ElseIf pData(counter).ValueList(2) = cUserKey Then
                PopulateUserListSheet pData(counter)
            End If
        ElseIf Left(pData(counter).Label, cMeasLawCombinationLength) = cMeasLawCombination Then
            If pData(counter).ValueList(0) = cUserKey Then
                If pFirstTime = 1 Then
                    pFirstTime = 0
                    pFirstUserIDItem = counter      'Contains position of first occurance of User ID row
                End If
                pLastUserIDItem = counter           'Contains position of last occurance of User ID row
            End If
        End If
    Next
End Sub

Private Sub PopulateEngineListSheet(MeasurementRecord As Record)
    Dim StartYear As Integer
    Dim StartMonth As Integer
    Dim StartDay As Integer
    Dim EndYear As Integer
    Dim EndMonth As Integer
    Dim EndDay As Integer
    Dim StartDate As String
    Dim EndDate As String

    gEngineListCursor.Value = Left(MeasurementRecord.ValueList(cMetricIDOffset), cMetricIDLength)
    gEngineListCursor.Offset(0, 1).Value = Mid(MeasurementRecord.ValueList(cMetricIDOffset), 5, cMetricIDLength)
    gEngineListCursor.Offset(0, 2).Value = Application.VLookup(gEngineListCursor.Value, gMetricName, 2, False)
    gEngineListCursor.Offset(0, 3).Value = Application.VLookup(gEngineListCursor.Offset(0, 1).Value, gMetricName, 2, False)
    gEngineListCursor.Offset(0, 4).Value = MeasurementRecord.ValueList(cMeasResultOffset)
    StartDate = Trim(Replace(Replace(Replace(MeasurementRecord.ValueList(cMeasStartDateOffset), vbCrLf, ""), vbLf, ""), vbCr, ""))
    EndDate = Trim(Replace(Replace(Replace(MeasurementRecord.ValueList(cMeasEndDateOffset), vbCrLf, ""), vbLf, ""), vbCr, ""))
    If StartDate <> vbNullString Then
        StartYear = CInt(Left(StartDate, 4))
        StartMonth = CInt(Mid(StartDate, 5, 2))
        StartDay = CInt(Right(StartDate, 2))
        gEngineListCursor.Offset(0, 5).Value = DateSerial(StartYear, StartMonth, StartDay)
    End If
    If EndDate <> vbNullString Then
        EndYear = CInt(Left(EndDate, 4))
        EndMonth = CInt(Mid(EndDate, 5, 2))
        EndDay = CInt(Right(EndDate, 2))
        gEngineListCursor.Offset(0, 6).Value = DateSerial(EndYear, EndMonth, EndDay)
    End If
    
    Set gEngineListCursor = gEngineListCursor.Offset(1)
End Sub

Private Sub PopulateUserListSheet(MeasurementRecord As Record)
    Dim counter As Integer
    Dim NotFound As Integer
    
    NotFound = 1
        
    For counter = pFirstUserIDItem To pLastUserIDItem
        If pData(counter).ValueList(1) = MeasurementRecord.ValueList(cUserTypeIDOffset) Then
            gUserListCursor.Value = MeasurementRecord.ValueList(cUserTypeIDOffset)
            gUserListCursor.Offset(0, 1).Value = Left(pData(counter).ValueList(2), 2)
            gUserListCursor.Offset(0, 2).Value = Application.VLookup(gUserListCursor.Offset(0, 1).Value, gUserTypeName, 2, False)
            gUserListCursor.Offset(0, 3).Value = MeasurementRecord.ValueList(cMeasResultOffset)
            gUserListCursor.Offset(0, 3).NumberFormat = "#,##0"
            gUserListCursor.Offset(0, 3).TextToColumns
            Set gUserListCursor = gUserListCursor.Offset(1)
            NotFound = 0
        End If
    Next
    If NotFound = 1 Then
        gUserListCursor.Value = MeasurementRecord.ValueList(cUserTypeIDOffset)
        gUserListCursor.Offset(0, 3).Value = MeasurementRecord.ValueList(cMeasResultOffset)
        Set gUserListCursor = gUserListCursor.Offset(1)
    End If
End Sub
