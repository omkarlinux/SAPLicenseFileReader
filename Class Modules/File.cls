VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''The objective of this class is to manage one SLAW Consolidated File'''
'''Members'''
Public Filepath As String
Private pFileContent As String
Private pSystems() As New System        'Array of system objects in the file
Private pSummary As New Summary         'Summary data object

'''Functions and Procedures'''

Public Sub FetchFileContent()
    Dim filePointer As Integer
    Dim fileContent() As Byte
    
    filePointer = FreeFile
    Open Filepath For Binary Access Read As filePointer
    ReDim fileContent(LOF(filePointer) - 1)
    Get filePointer, , fileContent
    Close filePointer

'Convert from byte array to string
    pFileContent = StrConv(fileContent, vbUnicode)
End Sub


Public Sub GetFileChooser()
'''Chooses a file and returns the name of the file
    Dim lSaveDriveDir As String
    Dim lMyPath As String

' Save the current directory.
    lSaveDriveDir = CurDir
    
' Set the path to the folder that you want to open.
    lMyPath = Application.DefaultFilePath
    
' You can also use a fixed path.
    'MyPath = "C:\Users\Ron de Bruin\Test"
    
' Change drive/directory to MyPath.
    ChDrive lMyPath
    ChDir lMyPath
    
' Open GetOpenFilename with the file filters.
    Filepath = Application.GetOpenFilename( _
            Title:="Select SLAW file", _
            MultiSelect:=False)
    
' Change drive/directory back to SaveDriveDir.
    ChDrive lSaveDriveDir
    ChDir lSaveDriveDir
End Sub

Public Sub PrintIt()
    Debug.Print (pFileContent)
End Sub

Public Sub SplitIt()
    Dim pFileContentLines() As String
    Const cStart As String = "BEGIN     SYSTEM"     'The start identifier of a system
    Const cEnd As String = "END       SYSTEM"       'The end identifier of a system
    Const cSummaryStart As String = "BEGIN     SUMME"     'The start identifier of a summary
    Const cSummaryEnd As String = "END       SUMME"       'The end identifier of a summary
    Dim pStartLen As Integer: pStartLen = Len(cStart)
    Dim pEndLen As Integer: pEndLen = Len(cEnd)
    Dim pSummaryStartLen As Integer: pSummaryStartLen = Len(cSummaryStart)
    Dim pSummaryEndLen As Integer: pSummaryEndLen = Len(cSummaryEnd)
    Dim pActivateDataCollection As String * 1: pActivateDataCollection = "N"    'Flag to denote whether we are between the start and end of a system
    Dim pTempCollection As Collection               'Temporarily stores each system
    Dim pSystemCollection As New Collection              'Stores all the systems from the file
    Dim SystemStringsTemp() As String           'Temp array of lines in a string
    Dim counter As Integer
    Dim pSummaryDataCollected As String * 1: pSummaryDataCollected = "N"    'Flag to denote whether summary data has previously been collected or not
    Dim pSystemNames As New Collection                      'To store the names of the systems
    
    
    
    pFileContentLines = Split(pFileContent, vbLf)

    For Each line In pFileContentLines
        If Left(line, pStartLen) = cStart Then
' Start of a new system, turn on system active flag, create new temp collection
            pActivateDataCollection = "Y"
            Set pTempCollection = Nothing
            Set pTempCollection = New Collection
            pSystemNames.Add (Mid(line, 21, 3))         'Fetch the 3 digit system name in front of the Begin system
        ElseIf Left(line, pEndLen) = cEnd Then
' End of a system, turn off system active flag, add temp collection to collection of collections
            pActivateDataCollection = "N"
            pSystemCollection.Add pTempCollection
        ElseIf Left(line, pSummaryStartLen) = cSummaryStart And pSummaryDataCollected = "N" Then
' Start of summary data, turn on summary data collected flag, create collection of summary data
            pSummaryDataCollected = "Y"
            pActivateDataCollection = "Y"
            Set pTempCollection = Nothing
            Set pTempCollection = New Collection
        ElseIf Left(line, pSummaryEndLen) = cSummaryEnd Then
'End of summary data, create summary object
            pActivateDataCollection = "N"
            Set pSummary.Content = pTempCollection
        ElseIf pActivateDataCollection = "Y" Then
' Log all of this system's data into a temporary collection
                pTempCollection.Add (line)
        ElseIf Not line = "" Then
' We got ourselves an exception here!
                Debug.Print ("Found some lines that are not between start and end")
        End If
    Next line
    
'Convert collection to array of system objects
    ReDim pSystems(0 To pSystemCollection.Count - 1)
    For counter = 1 To pSystemCollection.Count
        Set pSystems(counter - 1).Content = pSystemCollection(counter)
        pSystems(counter - 1).Name = pSystemNames(counter)
    Next
    
'Clean older data to free up memory
    pFileContent = vbNullString
End Sub

Public Sub SeparateData()
'Split data into label of the record and a list of values for that records
    pSummary.SeparateData
    For Each individualSystem In pSystems
        individualSystem.SeparateData
    Next individualSystem
End Sub

Public Sub DisplayResults()
    Dim CellCursor As Range
    Dim SystemsCursor As Range
    Dim offsetRows As Integer
    Dim counter As Integer
'Display Results into each appropriate sheet
    pSummary.DisplayResults
    
    Set CellCursor = Worksheets(cSystemWiseSheet).Range(cBeginCell)
    Set SystemsCursor = Worksheets(cSystemListSheet).Range(cBeginCell)
    For Each individualSystem In pSystems
        offsetRows = individualSystem.DisplayResults(CellCursor)
        Set CellCursor = CellCursor.Offset(offsetRows)              'Move current cell down by number of lines in system wise sheet
        individualSystem.PopulateSystemList SystemsCursor
        Set SystemsCursor = SystemsCursor.Offset(1)                 'Move current cell down by 1 line in system list sheet
        Application.StatusBar = "Progress: " & counter & " of " & UBound(pSystems) & ": " & Format(counter / UBound(pSystems), "0%")    'Shows progress bar in status
        counter = counter + 1
    Next individualSystem
End Sub
