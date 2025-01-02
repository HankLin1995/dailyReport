Attribute VB_Name = "progress_plot"
'20230301 side project
'測試如何透過Record上所記錄之施作位置與組合工項繪製線條圖取代紙本記載並偵錯重複區段

'todo:
'1.取得不重複組合工項
'2.依照組合工項名稱取得施作位置集合
'3.將施作位置集合切分為小單元預備畫圖
'4.依照小單元前後順序依序畫圖

'==============================

'rely on [module]="tranFunction"

Public specific_date As String

Sub clearAllShape()

With Sheets("plot")

    For Each shp In .Shapes
    
        shp.Delete
    
    Next

End With

End Sub


Sub plotBarProgressByChname() 'ByVal chname As String)

chname = InputBox("請輸入渠道名稱，全部則空白")

Dim myfunc As New clsMyfunction

specific_date = Format(Now(), "yyyy/m/d")

Call clearAllShape

Call AddText(X0, 15, 15, 100, chname, 2)

Set collMixItems = getMixItemsByChname(chname)
Set collPropIndex = myfunc.changeOrder(getSepIndexByChname(chname))

X_gap = 120
Y_origin = CDbl(InputBox("請輸入起始樁號", , 0)) - 50

For Each i In collMixItems
    
    targetMix = i
    
    prop = getPropByMixName(targetMix) '取得分類
    '取得分類所處Index
    myindex = myfunc.getCollIndex(collPropIndex, prop)
    
    Set collDoLoc = getDoLocationsByMix(targetMix)
    
    X0 = myindex * X_gap + 100 'X0 + X_gap '這裡會每跳一次就累加120
    X1 = X0
    
    Debug.Print X0
    
    Call AddText(X0 - 50 / 2, 15, 15, 70, prop, 2)
    
    For Each it2 In collDoLoc
    
        Debug.Print targetMix & ":" & it2
        
        tmp = Split(it2, ";")

        tmp_date = tmp(1)
        Call getSLocAndELoc(tmp(0), sloc, eloc)
        
        Y0 = sloc - Y_origin
        Y1 = eloc - Y_origin
    
        Call AddLine(X0, Y0, X1, Y1)
        
        Call AddText(X0 + 10, (Y0 + Y1) / 2 - 15 / 2, 15, 50, tmp_date)
        
        Call AddText(X0 - 40, Y0 - 15 / 2, 15, 30, Split(tmp(0), "~")(0), 1)
        Call AddText(X0 - 40, Y1 - 15 / 2, 15, 30, Split(tmp(0), "~")(1), 1)

    Next

Next

'Call AddPaper(X0 + X_gap, 1000)

Dim o As New clsPrintOut

o.SpecificShtToXLS ("plot")

End Sub


Function getSepIndexByChname(ByVal chname As String)

    'chanme = "土厝小排2-5"
    
    If chname = "" Then
    
        Dim myfunc As New clsMyfunction
        
        Set getSepIndexByChname = myfunc.getUniqueItems("Mix", 3, , "分類")
        
        Exit Function
    
    End If
    
    Dim coll As New Collection
    
    With Sheets("Mix")
    
        lr = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        For r = 2 To lr
        
            If .Cells(r, "I") = chname Then
            
                prop = .Cells(r, "K")
                
                If prop <> "" Then
                
                    On Error Resume Next
                    
                    coll.Add prop, CStr(prop)
                    
                    On Error GoTo 0
                
                End If
                
            End If
        
        Next
    
    End With
    
    Set getSepIndexByChname = coll


End Function

Function getPropByMixName(ByVal mix_name As String)

With Sheets("Mix")

Set rng = .Columns("A").Find(mix_name)

r = rng.Row

getPropByMixName = .Cells(r, "J")

End With

End Function

Sub plotBarProgress()

specific_date = Format(Now(), "yyyy/m/d")

Call clearAllShape

Set collMixItems = getMixItems

myIndexs = getShowIndex(collMixItems)

X_gap = 120
Y_origin = CDbl(InputBox("請輸入起始樁號")) - 50

For Each i In myIndexs

    'targetMix = it
    
    targetMix = collMixItems(CInt(i))
    
    Set collDoLoc = getDoLocationsByMix(targetMix)
    
    X0 = X0 + X_gap '這裡會每跳一次就累加120
    X1 = X0
    
    Call AddText(X0 - 50 / 2, 15, 15, 70, targetMix, 2)
    
    For Each it2 In collDoLoc
    
        Debug.Print targetMix & ":" & it2
        
        tmp = Split(it2, ";")

        tmp_date = tmp(1)
        Call getSLocAndELoc(tmp(0), sloc, eloc)
        
        Y0 = sloc - Y_origin
        Y1 = eloc - Y_origin
    
        Call AddLine(X0, Y0, X1, Y1)
        
        Call AddText(X0 + 10, (Y0 + Y1) / 2 - 15 / 2, 15, 50, tmp_date)
        
        Call AddText(X0 - 40, Y0 - 15 / 2, 15, 30, Split(tmp(0), "~")(0), 1)
        Call AddText(X0 - 40, Y1 - 15 / 2, 15, 30, Split(tmp(0), "~")(1), 1)

    Next

Next

'Call AddPaper(X0 + X_gap, 1000)

Dim o As New clsPrintOut

o.SpecificShtToXLS ("plot")

End Sub

Function getX0_ByPropIndex(ByVal prop As String)


End Function

Function getShowIndex(ByVal collMixItems)

'TODO:create a tmp plot orders

plot_order = Sheets("Records").Range("E1")

For Each it In collMixItems

    j = j + 1

    prompt = prompt & j & "." & it & vbNewLine

    k = k & "," & j

Next

If plot_order = "" Then
    indexDefault = mid(k, 2)
Else
    indexDefault = plot_order
End If

getShowIndex = Split(InputBox(prompt, "顯示序位選擇器", indexDefault), ",")

If UBound(getShowIndex) = -1 Then MsgBox "取消操作!", vbCritical: End

For Each it In getShowIndex

    prompt2 = prompt2 & it & "." & collMixItems(CInt(it)) & vbNewLine

Next

MsgBox "預定繪製順序：" & vbNewLine & prompt2, , "Plot_Progress"

End Function

Sub AddPaper(W, H)

    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, _
        W, H).Select
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    
    Selection.ShapeRange.ZOrder msoSendToBack

    ActiveSheet.Columns("A").Delete

End Sub


Sub AddText(ByVal txtX, ByVal txtY, ByVal txtHeight, ByVal txtWidth, ByVal txtStr, Optional ByVal RGB_Selector As Integer = 0)

If txtStr = specific_date Then RGB_Selector = 3

With Sheets("plot")

    .Activate

    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtX, txtY, txtHeight * Len(txtStr), txtHeight).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = txtStr
        
    Selection.ShapeRange.Fill.Visible = msoFalse
    Selection.ShapeRange.Line.Visible = msoFalse
        
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Font
'        .NameComplexScript = "+mn-cs"
'        .NameFarEast = "+mn-ea"
'        .Fill.Visible = msoTrue
'        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
'        .Fill.ForeColor.TintAndShade = 0
'        .Fill.ForeColor.Brightness = 0
'        .Fill.Transparency = 0
'        .Fill.Solid
        
        Select Case RGB_Selector
        
        Case 0: .Fill.ForeColor.RGB = RGB(0, 0, 0)
        Case 1: .Fill.ForeColor.RGB = RGB(0, 204, 255)
        Case 2: .Fill.ForeColor.RGB = RGB(255, 0, 0)
        Case 3: .Fill.ForeColor.RGB = RGB(255, 102, 0)
        
        End Select
        
        '.Fill.Visible = msoFalse
        .Size = txtHeight * 0.5
'        .Name = "+mn-lt"
    End With

End With



End Sub

Sub AddLine(X0, Y0, X1, Y1)

With Sheets("plot")

    .Activate

    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, X0, Y0, X1, Y1).Select
    Selection.ShapeRange.ScaleWidth 1.6145833333, msoFalse, msoScaleFromTopLeft
    'Selection.ShapeRange.ScaleHeight 0, msoFalse, msoScaleFromTopLeft
    
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 10
        '.ForeColor.RGB = RGB(50, 0, 128)
    End With

    Call ActiveSheet.Shapes.AddShape(msoShapeRightBrace, X0 + 5, Y0, 10, Y1 - Y0)

End With

End Sub

Function getDoLocationsByMix(ByVal MixName As String)

Dim coll As New Collection

With Sheets("Records")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        myMixName = .Cells(r, "J")
        recDate = .Cells(r, "B")
        myContent = .Cells(r, "D")
    
        If myMixName = MixName Then
        
        If Not myContent Like "*~*" Then
        
            error_prompt = error_prompt & "第" & r & "列:樁號區間" & myContent & "未含~" & vbNewLine
        
        End If
    
            tmp = Split(myContent, "、")
            
            For i = 0 To UBound(tmp)
        
                myContent_split = tmp(i)
                
                On Error GoTo ERRORHANDLE
        
                coll.Add myContent_split & ";" & recDate, CStr(myContent_split)
        
            Next
            
        End If
    
    Next
    
    Set getDoLocationsByMix = coll

End With

If error_prompt <> "" Then

    MsgBox error_prompt, vbCritical
    End

End If

Exit Function

ERRORHANDLE:

MsgBox ("第" & r & "列," & MixName & "項目的" & myContent_split & "<重複>"), vbCritical

End Function

Function getMixItems()

Dim coll As New Collection

With Sheets("Mix") 'get item orders by records

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        s = .Cells(r, "A")
    
        Set rng = Sheets("Records").Columns("J").Find(s)
    
        If s <> "" And Not rng Is Nothing Then
    
        On Error Resume Next
        coll.Add s, CStr(s)
        On Error GoTo 0
    
        End If
    
    Next
    
    Set getMixItems = coll

End With

End Function

Function getMixItemsByChname(ByVal chname As String)

If chname = "" Then

    Set getMixItemsByChname = getMixItems
    
    Exit Function

End If

Dim coll As New Collection

With Sheets("Mix") 'get item orders by records

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        s = .Cells(r, "A")
        ch = .Cells(r, "I")
    
        Set rng = Sheets("Records").Columns("J").Find(s)
    
        If s <> "" And Not rng Is Nothing And ch = chname Then
    
        On Error Resume Next
        coll.Add s, CStr(s)
        On Error GoTo 0
    
        End If
    
    Next
    
    Set getMixItemsByChname = coll

End With

End Function

Function getMixItems2()

Dim coll As New Collection

With Sheets("Records") 'get item orders by records

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        s = .Cells(r, "J")
    
        If s <> "" Then
    
        On Error Resume Next
        coll.Add s, CStr(s)
        On Error GoTo 0
    
        End If
    
    Next
    
    Set getMixItems = coll

End With

End Function

'TODO:create a tmp sheet to keyin plot order

Sub plotS_Curve()

Dim Inf_obj As New clsInformation
Dim shp As Shape

For Each shp In Sheets("S-CURVE").Shapes

    If shp.OnAction = "" Then shp.Delete

Next


With Sheets("Diary")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

End With

    Sheets("S-CURVE").Activate

    Sheets("S-CURVE").Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    ActiveChart.SetSourceData Source:=Range("Diary!$B$2:$B$147,Diary!$D$2:$D$147,Diary!$I$2:$I$48")
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Diary!$B$2:$B$" & lr
    ActiveChart.FullSeriesCollection(1).Values = "=Diary!$D$2:$D$" & lr
    ActiveChart.FullSeriesCollection(1).Name = "=""預定進度"""
    ActiveChart.FullSeriesCollection(2).XValues = "=Diary!$B$2:$B$" & lr
    ActiveChart.FullSeriesCollection(2).Values = "=Diary!$I$2:$I$" & lr
    ActiveChart.FullSeriesCollection(2).Name = "=""實際進度"""
    ActiveChart.FullSeriesCollection(3).Delete
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 0
    ActiveChart.Axes(xlValue).MaximumScale = 1
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = Sheets("Diary").Range("B2")
    ActiveChart.Axes(xlCategory).MaximumScale = Sheets("Diary").Range("B" & lr)
    
    For Each shp In Sheets("S-CURVE").Shapes
    
        If shp.OnAction = "" Then ShpName = shp.Name
    
    Next
    
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = Inf_obj.conName
    Selection.Format.TextFrame2.TextRange.Characters.Text = Inf_obj.conName
    With Selection.Format.TextFrame2.TextRange.Characters(1, 2).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 2).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementLegendRight)
    
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes(ShpName).IncrementLeft -410
    ActiveSheet.Shapes(ShpName).IncrementTop -140
    ActiveSheet.Shapes(ShpName).ScaleWidth 2, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes(ShpName).ScaleHeight 2, msoFalse, _
        msoScaleFromTopLeft
        

End Sub



