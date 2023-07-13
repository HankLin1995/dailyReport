VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBudgetDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'useless

Private shtDB As Object
Private sr As Integer
Private lr As Integer

Private Sub Class_Initialize()

Set shtDB = Sheets("Budget_DB")

'sr = 2
'lr = shtDB.Cells(shtDB.Rows.count, 1).End(xlUp).Row

End Sub

Sub AppendData(arr)

With Sheets("Budget_DB")

lr = .Cells(.Rows.count, 1).End(xlUp).Row

r = lr + 1

.Cells(r, 1) = r - 1

For i = LBound(arr) To UBound(arr)

    .Cells(r, 2 + i) = arr(i)

Next

End With

End Sub

Function getUniqueItems(ByVal find_col_eng As String)

Dim coll As New Collection

With shtDB

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = 2 To lr
    
        s = .Cells(r, find_col_eng)
    
        On Error Resume Next
        
        coll.Add s, s
        
        On Error GoTo 0
    
    Next

End With

Set getUniqueItems = coll

End Function

Sub clearRows(ByVal find_col_eng As String, ByVal find_str As String)

With shtDB

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = lr To 2 Step -1

      If .Cells(r, find_col_eng) = find_str Then .Rows(r).Delete

    Next

End With

End Sub

Function IsExisted(ByVal find_col_eng As String, ByVal find_str As String)

IsExisted = True

With shtDB

    Set rng = .Columns(find_col_eng).Find(find_str)
    
    If rng Is Nothing Then IsExisted = False
    
End With

End Function
