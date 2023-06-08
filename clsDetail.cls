VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private shtDetail As Object
Private arrDetail
Private sr As Integer
Private er As Integer

Private Sub Class_Initialize()

Set shtDetail = Sheets("Detail")

End Sub

Function getSeperateRows() '20210824 取得分割Mix資料列位們

Dim coll As New Collection

With shtDetail

    '.UsedRange.EntireRow.Hidden = False '摺疊會出問題!
    
    lr = .Cells(Rows.count, 3).End(xlUp).Row
    
    For r = 3 To lr
    
        If .Cells(r, 1) <> "" Then coll.Add r
    
    Next

    coll.Add lr + 1

End With

Set getSeperateRows = coll

End Function

Sub getPropertiesByName(ByVal myName As String)

Set coll = getSeperateRows()

With shtDetail

For i = 1 To coll.count - 1

    sr = coll(i)
    er = coll(i + 1) - 1

    targetName = .Cells(sr, 1)
    
    If targetName = myName Then
    
        arrDetail = .Range("B" & sr & ":F" & er)
        Exit Sub
    
    End If

Next

End With

End Sub

Function getSum()

For i = LBound(arrDetail, 1) To UBound(arrDetail, 1)

    getSum = getSum + arrDetail(i, 3) * arrDetail(i, 4)

Next

End Function

Sub setItemToCbo()

For i = LBound(arrDetail, 1) To UBound(arrDetail, 1)

    frm_Detail.cboItem.AddItem arrDetail(i, 1)

Next

End Sub

Function getUnit(ByVal itemName As String)

For i = LBound(arrDetail, 1) To UBound(arrDetail, 1)

    If itemName = arrDetail(i, 1) Then
    
        getUnit = arrDetail(i, 2)
        Exit For
        
    End If
    
Next

End Function

Function getLast(ByVal itemName As String)

For i = LBound(arrDetail, 1) To UBound(arrDetail, 1)

    If itemName = arrDetail(i, 1) Then
    
        getLast = arrDetail(i, 3) - arrDetail(i, 5)
        Exit For
        
    End If
    
Next

End Function

Sub setAmount(ByVal s As String) '設定已做數量

tmp = Split(mid(s, 2), ",")

For i = LBound(tmp) To UBound(tmp) Step 3

    myitem = tmp(i)
    myAmount = tmp(i + 1)

    With shtDetail

        For r = sr To er
        
            If .Cells(r, 2) = myitem Then .Cells(r, 6) = .Cells(r, 6) + myAmount: Exit For
        
        Next

    End With

Next

'calcRatio = cost / getSum

End Sub

Function calcRatio(ByVal s As String)

tmp = Split(mid(s, 2), ",")

For i = LBound(tmp) To UBound(tmp) Step 3

    myitem = tmp(i)
    myAmount = tmp(i + 1)

    For j = LBound(arrDetail, 1) To UBound(arrDetail, 1)

       If myitem = arrDetail(j, 1) Then cost = cost + myAmount * arrDetail(j, 4): Exit For

    Next

Next

calcRatio = cost / getSum

End Function

