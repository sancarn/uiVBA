VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents ciRequired As uiTableColumnInfo
Attribute ciRequired.VB_VarHelpID = -1
Private WithEvents tbl As uiTable
Attribute tbl.VB_VarHelpID = -1

Public Sub TestAll()
  Dim headers As Collection: Set headers = New Collection
  headers.Add uiTableColumnInfo.Create("Type", uiLabel, 70)
  headers.Add uiTableColumnInfo.Create("Required", uiCheckBox, 40)
  headers.Add uiTableColumnInfo.Create("Count", uiTextBox, 40)
  headers.Add uiTableColumnInfo.Create("Int", uiCheckBox, 40)
  headers.Add uiTableColumnInfo.Create("Ext", uiCheckBox, 40)
  '... add additional columns
  
  Set ciRequired = headers(2)
  
  Set tbl = uiTable.Create(headers, TestData.Range("MitTypes").value, UF.frFrame)
  UF.Show False
End Sub

'Decode a value from input data type to checkbox value i.e. true/false
Private Sub ciRequired_DecodeControlValue(ByRef v As Variant)
  v = v = "Yes"
End Sub
'Encode a checkbox value i.e. (true/false) into a data type ("Yes"/"No")
Private Sub ciRequired_EncodeControlValue(ByRef v As Variant)
  v = IIf(v, "Yes", "No")
End Sub


Private Sub tbl_Change(ByVal Caller As Object, ByVal iRowIndex As Long, params As Variant)
  TestData.Range("MitTypes").value = tbl.data
End Sub
