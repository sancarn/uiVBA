VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Implements uiIMessagable

Private pHeaders as Collection   'Collection<uiTableColumnInfo>
Private pRows as new Collection  'Collection<uiTableRow>
Private pRowHeight as Double
Private pFrame as Frame

Public Function Create(ByRef cHeaders as Collection, ByRef cRows as Collection, ByRef fr as Frame) as uiTable
  set Create = new uiTable
  Call Create.protInit(cHeaders, cRows, fr)
End Function
Friend Sub protInit(ByRef cHeaders as Collection, ByRef cRows as Collection, ByRef fr as Frame)
  set pHeaders = cHeaders
  set pFrame = fr

  Dim vRowData as variant
  For each vRowData in cRows
    Call pRows.add(uiTableRow.Create(Me, vRowData))
  next
End Sub

Public Property Get uiFrame() as Frame
  set uiFrame = pFrame
End Property


Public Property Get ColumnInfo() as Collection
  set ColumnInfo = pHeaders
End Property
Public Property Get Rows() as collection
  set Rows = pRows
End Property

Public Property Get RowHeight() as Double
  RowHeight = pRowHeight
End Property
Public Property Let RowHeight(fValue as double)
  pRowHeight = fValue

  Dim row as uiTableRow
  For each row in pRows
    row.height = fValue
  next
End Property



Public Property Get data() As Variant
  Dim v():
  ReDim v(1 To pRows.count, 1 To pHeaders.count)
  Dim row As uiTableRow, iRow As Long: iRow = 0
  For Each row In pRows
    iRow = iRow + 1
    Dim vRow: vRow = row.data
    Dim j As Long: For j = LBound(vRow) To UBound(vRow)
      v(i, j) = vRow(j)
    Next
  Next
  data = v
End Property




Private Function uiIMessagable_Message(ByVal Caller As Object, iMessage As Long, ParamArray params() As Variant) As Variant: End Function