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

Private pParent as uiTable
Private pColumns as Collection
Private pHeight as Double
Private pCells as new Collection 'Collection<uiElement>
Private pTop as Double
Private pMeta as Object

Public Function Create(ByRef oParent as uiTable, ByRef vValues as variant) as uiTableRow
  set Create = new uiTableRow
  Call Create.protInit(oParent, vValues)
End Function
Public Sub protInit(ByRef oParent as uiTable, ByRef vValues as variant)
  set pParent = oParent
  set pColumns = oParent.ColumnInfo
  Dim uiFrame as Frame: set uiFrame = pParent.uiFrame
  Dim iLB as Long: iLB = lbound(vValues)
  Dim iUB as Long: iUB = ubound(vValues)
  For i = iLB to iUB
    Dim index as long: index = 1 + i - iLB
    Dim colInfo as uiTableColumnInfo: set colInfo = pColumns(index)
    Dim element as uiElement: set element = uiElement.CreateFromType(Me, uiFrame.Controls, colInfo.iType, fLeft:=colInfo.left, fWidth:=colInfo.Width)
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
  Dim v()
  Redim v(1 to pCells.Count)
  For i = 1 to pCells.Count
    v(i) = pColumns(i).ElementData(pCells(i))
  Next
  Data = v
End Property

'TODO: Get ID()


Private Function uiIMessagable_Message(ByVal Caller As Object, iMessage As Long, ParamArray params() As Variant) As Variant: End Function