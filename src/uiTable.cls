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

Public Event Change(ByVal Caller as object, ByVal iRowIndex as long, params As Variant)

Private pHeaders as Collection        'Collection<uiTableColumnInfo>
Private pHeaderElements as collection 'Collection<uiElement>
Private pRows as new Collection       'Collection<uiTableRow>
Private pRowHeight as Double
Private pFrame as Frame
Private pPadding as double

Public Function Create(ByRef cHeaders as Collection, ByRef vRows as Variant, ByRef fr as Frame, Optional ByVal fPadding as Double = 10) as uiTable
  set Create = new uiTable
  Call Create.protInit(cHeaders, vRows, fr, fPadding)
End Function
Friend Sub protInit(ByRef cHeaders as Collection, ByRef vRows as Variant, ByRef fr as Frame, ByVal fPadding as Double)
  set pHeaders = cHeaders
  set pFrame = fr
  pPadding = fPadding
  pRowHeight = 20

  'Set left of each column based on widths and padding
  Dim col as uiTableColumnInfo, fSumLeft as Double
  fSumLeft = fPadding
  for each col in pHeaders
    col.left = fSumLeft
    fSumLeft = fSumLeft + col.width
  next
  fr.ScrollWidth = fSumLeft

  Dim iRow as Long
  For iRow = 1 to ubound(vRows,1)
    Dim fTop as Long: fTop = pRowHeight*(iRow-1) + pPadding
    Call pRows.add(uiTableRow.Create(Me, iRow, GetRowArray(vRows,iRow), fTop, pRowHeight))
  next
  fr.scrollHeight = (ubound(vRows,1)+1) * pRowHeight + pPadding
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
      v(iRow, j) = vRow(j)
    Next
  Next
  data = v
End Property
Private Function GetRowArray(ByRef vTable as variant, ByRef iRow as long) as Variant
  Dim iUB as long: iUB = uBound(vTable,2)
  Dim v(): Redim v(1 to iUB)

  Dim i as long
  For i = 1 to iUB
    v(i) = vTable(iRow, i)
  next

  GetRowArray = v
End Function



Private Function uiIMessagable_Message(ByVal Caller as Object, iMessage as Long, Optional params as Variant = Empty) as Variant
  Select Case TypeName(Caller)
    Case "uiTableRow"
      Select Case iMessage
        Case uiTableRowEventChange
          RaiseEvent Change(Caller, Caller.rowIndex, Array(Caller, iMessage, params))
      End Select
  End Select
End Function