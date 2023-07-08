VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBudgetItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public collSep As New Collection
Public collItem As New Collection
Public collItemIndex As New Collection
Public collUnit As New Collection
Public collPrice As New Collection
Public collNum As New Collection
Public collCPrice As New Collection
Public collCNum As New Collection

Public collSepEach As New Collection
Public collSepEachIndex As New Collection
Public collsumSep As New Collection
Public collsumCSep As New Collection

Sub getData(Optional c As Integer = 4)

With Sheets("Budget")
    
    lr = .Cells(3, 1).End(xlDown).Row
    
    For r = 3 To lr
    
        'If r = 43 Then Stop
    
        Sep = .Cells(r, "A")
        Item = .Cells(r, "B")
        unit = .Cells(r, "C")
        Num = .Cells(r, c) 'D
        Price = .Cells(r, c).Offset(0, 1)
        CNum = .Cells(r, c + 4) 'H
        CPrice = .Cells(r, c + 4).Offset(0, 1)
        
        collSep.Add Sep '主要分類項
        collItem.Add Item, Item '工程項目
        collUnit.Add unit '單位
        collPrice.Add Price '單價
        collNum.Add Num '數量
        collCPrice.Add CPrice '更改後單價
        collCNum.Add CNum '更改後數量
    
    Next

End With

Call getSepName
Call getItemIndex

End Sub

Private Sub getItemIndex() '將原本的collection中的值與count互換,可以用it去查i

For Each it In collItem

    i = i + 1
    
    collItemIndex.Add i, it

Next

End Sub

Private Sub getSepName()

For Each mSep In collSep

    On Error Resume Next

    collSepEach.Add mSep, mSep

    On Error GoTo 0

Next

For Each it In collSepEach '將原本的collection中的值與count互換,可以用it去查i

    i = i + 1
    
    collSepEachIndex.Add i, it

Next

End Sub

Sub sumEachSepPrice()

For Each mSepEach In collSepEach

    Sum = 0
    CSum = 0
    
    For i = 1 To collSep.count
    
        mSep = collSep(i)
    
        If mSep = mSepEach Then
        
            Sum = Sum + collPrice(i) * collNum(i)
            
            CSum = CSum + collCPrice(i) * collCNum(i)
            
        End If
        
    Next

    Debug.Print mSepEach & ":" & Sum

    collsumSep.Add Sum '各工程項目累計總價
    collsumCSep.Add CSum '各工程項目累計變更後總價

    tSum = tSum + Sum

Next

Debug.Print "Total:" & tSum

End Sub

Function sumSpecificPrice(ByVal s As Integer, ByVal e As Integer)

For i = s To e
    sumSpecificPrice = sumSpecificPrice + collsumSep(i)
Next

End Function

Function sumSpecificCPrice(ByVal s As Integer, ByVal e As Integer)

For i = s To e
    sumSpecificCPrice = sumSpecificCPrice + collsumCSep(i)
Next

End Function



