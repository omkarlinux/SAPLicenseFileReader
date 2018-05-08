VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''Members '''
Private pLabel As String * 20
Private pValueList() As String
'''Functions
Public Property Let Label(Value As String)
    pLabel = Left(Value, cLabelLength)
End Property

Public Property Get Label() As String
    Label = pLabel
End Property

Public Property Let ValueList(ValueVariant As Variant)
    ReDim pValueList(0 To UBound(ValueVariant))
    pValueList = ValueVariant
End Property

Public Property Get ValueList() As Variant
    ValueList = pValueList
End Property
