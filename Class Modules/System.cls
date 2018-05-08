VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''The objective of this class is to handle each system within a file '''
'''Members'''
Private pContent() As String
Public Name As String
Private Status As String
Private ReleaseNo As String
Private Database As String
Private InstallationNo As String
Private CustomerKey As String
Private SystemNo As String
Private Const cStatus As String = "System    Status    "
Private Const cRelease As String = "System    R/3 Rel.  "
Private Const cDatabase As String = "System    Database  "
Private Const cInstallationNo As String = "System    Inst. No. "
Private Const cCustomerKey As String = "System    Cust. key "
Private Const cSystemNo As String = "System    System No."
Private pData() As New Record
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
        If counter < 30 Then
            If pData(counter).Label = cStatus Then
                Status = pData(counter).ValueList(0)
            ElseIf pData(counter).Label = cRelease Then
                ReleaseNo = pData(counter).ValueList(0)
            ElseIf pData(counter).Label = cDatabase Then
                Database = pData(counter).ValueList(0)
            ElseIf pData(counter).Label = cInstallationNo Then
                InstallationNo = pData(counter).ValueList(0)
            ElseIf pData(counter).Label = cCustomerKey Then
                CustomerKey = pData(counter).ValueList(0)
            ElseIf pData(counter).Label = cSystemNo Then
                SystemNo = pData(counter).ValueList(0)
            End If
        End If
    Next
'Clean older variables
    Erase pContent
End Sub

Public Function DisplayResults(StartingCell As Range) As Integer
    Dim Destination As Range
'    Set Destination = Worksheets(cSystemWiseSheet).Range(cBeginCell)
    Set Destination = StartingCell
    For counter = 0 To UBound(pData)
        Destination.Offset(counter).Value = Name
        Destination.Offset(counter, 1).Value = pData(counter).Label
        For internalCounter = 0 To UBound(pData(counter).ValueList)
            Destination.Offset(counter, internalCounter + 2).Value = pData(counter).ValueList(internalCounter)
        Next
    Next
    DisplayResults = UBound(pData) + 1
End Function

Public Sub PopulateSystemList(Cursor As Range)
    Cursor.Value = Name
    Cursor.Offset(0, 1) = Status
    Cursor.Offset(0, 2) = ReleaseNo
    Cursor.Offset(0, 3) = Database
    Cursor.Offset(0, 4) = InstallationNo
    Cursor.Offset(0, 5) = CustomerKey
    Cursor.Offset(0, 6) = SystemNo
End Sub
