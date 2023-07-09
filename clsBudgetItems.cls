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
        num = .Cells(r, c) 'D
        price = .Cells(r, c).Offset(0, 1)
        CNum = .Cells(r, c + 4) 'H
        CPrice = .Cells(r, c + 4).Offset(0, 1)
        
        If price <> 0 Then
        
            collSep.Add Sep '�D�n������
            'collSep.Add test_getParentItem(Sep) '�D�n������
            collItem.Add Item, Item '�u�{����
            collUnit.Add unit '���
            collPrice.Add price '���
            collNum.Add num '�ƶq
            collCPrice.Add CPrice '������
            collCNum.Add CNum '����ƶq
    
        End If
    
    Next

End With

Call getSepName
Call getItemIndex

End Sub

Private Sub getItemIndex() '�N�쥻��collection�����ȻPcount����,�i�H��it�h�di

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

For Each it In collSepEach '�N�쥻��collection�����ȻPcount����,�i�H��it�h�di

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

    collsumSep.Add Sum '�U�u�{���ز֭p�`��
    collsumCSep.Add CSum '�U�u�{���ز֭p�ܧ���`��

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

Function CBudget_changeIndex()

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = lr To 3 Step -1
    
        item_index = .Cells(r, 1)
        
        On Error Resume Next
        .Cells(r, 1).Comment.Delete
        On Error GoTo 0
        
        .Cells(r, 1).AddComment
        .Cells(r, 1).Comment.Text Text:=item_index
        .Cells(r, 1) = getParentItem(item_index)
        
    Next
    

End With

End Function

Function CBudget_recover()

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = 3 To lr
    
        item_index = .Cells(r, 1).Comment.Text
        
        On Error Resume Next
        .Cells(r, 1).Comment.Delete
        On Error GoTo 0
        

        .Cells(r, 1) = item_index
    
    Next
    
End With

End Function

Function getParentItem(ByVal item_index As String)

With Sheets("Budget")
    
    s = item_index ' .Range("A19")
    
    tmp = Split(s, ".")
    
    If UBound(tmp) > 1 Then
    
        For i = 0 To UBound(tmp) - 1
        
            ch = ch & "." & tmp(i)
        
        Next
        
        item_index = mid(ch, 2, Len(ch) - 1)
        
        Set rng = .Columns("A").Find(item_index)
        
        getParentItem = rng.Offset(0, 1).Value
    
    ElseIf UBound(tmp) = 1 Or UBound(tmp) = 0 Then
    
        Set rng = .Columns("A").Find(item_index)
        getParentItem = rng.Offset(0, 1).Value
        
    End If

End With

End Function


