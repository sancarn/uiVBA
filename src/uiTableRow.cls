VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiTableRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Implements uiIMessagable

Public Enum EUITableRowMessage
  uiTableRowEventChange
  uiTableRowEventMove
  uiTableRowEventResize
  uiTableRowEventMouseClick
  uiTableRowEventMouseFocus
  uiTableRowEventMouseMove
  uiTableRowEventTabSelected
  uiTableRowEventAddChild
End Enum

Private pParent as uiTable
Private pMessageParent as uiIMessagable
Private pColumns as Collection
Private pHeight as Double
Private pCells as new Collection 'Collection<uiElement>
Private pTop as Double
Private pRowIndex as long
Private pMeta as Object
Private pID as string

Public Function Create(ByRef oParent as uiTable, ByVal iIndex as long , ByRef vValues as variant, Optional ByVal fTop as Double, Optional ByVal fHeight as Double=30) as uiTableRow
  set Create = new uiTableRow
  Call Create.protInit(oParent, iIndex, vValues,fTop, fHeight)
End Function
Public Sub protInit(ByRef oParent as uiTable, ByVal iIndex as long, ByRef vValues as variant, Optional ByVal fTop as Double, Optional ByVal fHeight as Double=30)
  set pParent = oParent
  set pMessageParent = oParent
  set pColumns = oParent.ColumnInfo
  pRowIndex = iIndex

  Dim uiFrame as Frame: set uiFrame = pParent.uiFrame
  Dim iLB as Long: iLB = lbound(vValues)
  Dim iUB as Long: iUB = ubound(vValues)
  For i = iLB to iUB
    Dim index as long: index = 1 + i - iLB
    Dim colInfo as uiTableColumnInfo: set colInfo = pColumns(index)
    Dim element as uiElement: set element = uiElement.CreateFromType(Me, uiFrame.Controls, colInfo.iType,,"", colInfo.left, fTop, colInfo.Width, fHeight)
    Call colInfo.cells.add(element)
    Call pCells.add(element)
    colInfo.ElementData(element) = vValues(i)
  next
End Sub

'Get/Set Height of all cells
Friend Property Get Height() as Double
  Height = pHeight
End Function
Friend Property Let Height(fValue as Double)
  pHeight   = fValue

  'Change cell's height
  Dim cell as uiElement
  for each cell in pCells
    cell.height = fValue
  next
End Property

'Get/Set the top of all cells
Friend Property Get Top() as Double
  Top = pTop
End Property
Friend Property Let Top(fValue as Double)
  pTop = fValue

  'Change cell's top
  Dim cell as uiElement
  for each cell in pCells
    cell.top = fValue
  next
End Property

'Get/Set cells of the row
Friend Property Get Cells() as Collection
  set Cells = pCells
End Property
Friend Property Set Cells(oValue as Collection)
  set pCells = oValue
End Property

Public Property Get Data() as Variant
  if pCells.Count = 0 then Exit Property
  Dim v()
  Redim v(1 to pCells.Count)
  For i = 1 to pCells.Count
    v(i) = pColumns(i).ElementData(pCells(i))
  Next
  Data = v
End Property

Public Property Get RowIndex() as Long
  RowIndex = pRowIndex
End Property

Public Property Get Meta() as Object
  if pMeta is nothing then set pMeta = CreateObject("Scripting.Dictionary")
  set Meta = pMeta
End Property

Public Property Get ID() as string
  if pID = "" then pID = getGUID()
  ID = pID
End Property






Private Function getGUID() As String
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(CLng(Rnd() * 15.9999)), 1, 1)
  Next
End Function

Private Function uiIMessagable_Message(ByVal Caller as Object, iMessage as Long, Optional params as Variant = Empty) as Variant
  Call pMessageParent.Message(Caller, iMessage, params)
  select case TypeName(caller)
    case "uiElement"
      select case iMessage
        Case uiElementEventChange
          Call pMessageParent.Message(Me, uiTableRowEventChange, data)
        Case uiElementEventMove
          Call pMessageParent.Message(Me, uiTableRowEventMove, data)
        Case uiElementEventResize
          Call pMessageParent.Message(Me, uiTableRowEventResize, data)
        Case uiElementEventMouseClick
          Call pMessageParent.Message(Me, uiTableRowEventMouseClick, data)
        Case uiElementEventMouseFocus
          Call pMessageParent.Message(Me, uiTableRowEventMouseFocus, data)
        Case uiElementEventMouseMove
          Call pMessageParent.Message(Me, uiTableRowEventMouseMove, data)
        Case uiElementEventTabSelected
          Call pMessageParent.Message(Me, uiTableRowEventTabSelected, data)
        Case uiElementEventAddChild
          Call pMessageParent.Message(Me, uiTableRowEventAddChild, data)
      end select
  end select
End Function
