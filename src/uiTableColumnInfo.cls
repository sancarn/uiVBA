VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiTableColumnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Implements uiIMessagable

'Events to encode and decode a column value. E.G.
'@example
'Value in dataset = "yes"
'Private Sub col_EncodeControlValue(v):  v = v = "yes": End Sub
'Private Sub col_DecodeControlValue(v):  v = iif(v, "yes", "no"): End Sub
Public Event EncodeControlValue(ByRef v as Variant)
Public Event DecodeControlValue(ByRef v as Variant)

Private pName as string
Private pType as EUIElementType
Private pLeft as Double
Private pWidth as Double
Private pCells as new Collection 'Collection<uiElement>

Public Function Create(ByVal sHeaderName as string, ByVal iType as EUIElementType, Optional ByVal fWidth as double = 100) as uiTableColumnInfo
  set Create = new uiTableColumnInfo
  Call Create.protInit(sHeaderName, iType, fWidth)
End Function
Public Sub protInit(ByVal sHeaderName as string, ByVal iType as EUIElementType, ByVal fWidth as double)
  pName = sHeaderName
  pType = iType
  pWidth = fWidth
End Sub

'Get/Let name
Public Property Get Name() as string
  Name = pName
End Property
Public Property Let Name(ByVal v as string)
  pName = v
End Property

'Get type of element
Public Property Get iType() as EUIElementType
  iType = pType
End Property

'Get/Let left of cells
Public Property Get left() as Double
  Left = pLeft
End Property
Friend Property Let left(fLeft as Double)
  pLeft = fLeft
  
  'Change cell's left
  Dim cell as uiElement
  For each cell in pCells
    cell.left = fLeft
  next
End Property

'Get/Let width of cell
Public Property Get Width() as Double
  Width = pWidth
End Property
Public Property Let Width(fWidth as Double)
  pWidth = fWidth

  'Change cell's width
  Dim cell as uiElement
  For each cell in pCells
    cell.width = fWidth
  next
End Property


'Get/Set value from element
'@example
'ElementData(control) = "Yes"                            i.e. "Yes" ==> True
'data = ElementData(control)    'data now contains "Yes" i.e. True  ==> "Yes"
Public Property Get ElementData(ByRef element as uiElement) as variant
  Dim v: v = element.value
  RaiseEvent EncodeControlValue(v)
  ElementData = v
End Property
Public Property Let ElementData(ByRef element as uiElement, ByVal v as Variant)
  RaiseEvent DecodeControlValue(v)
  element.value = v
End Property

'Get collection
Public Property Get Cells() as Collection
  set Cells = pCells
End Property

Private Function uiIMessagable_Message(ByVal Caller as Object, iMessage as Long, Optional params as Variant = Empty) as Variant
  
End Function