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

Sub plotBarProgress()

specific_date = Format(Now(), "yyyy/m/d")

Call clearAllShape

X_gap = 120
Y_origin = 28500

Set collMixItems = getMixItems

myIndexs = getShowIndex(collMixItems)

'For Each it In collMixItems

For Each i In myIndexs

    'targetMix = it
    
    targetMix = collMixItems(CInt(i))
    
    Set collDoLoc = getDoLocationsByMix(targetMix)
    
    X0 = X0 + X_gap
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

Call AddPaper(X0 + X_gap, 1000)

Dim o As New clsPrintOut

o.SpecificShtToXLS ("plot")

End Sub

Function getShowIndex(ByVal collMixItems)

For Each it In collMixItems

    j = j + 1

    prompt = prompt & j & "." & it & vbNewLine

    k = k & "," & j

Next

getShowIndex = Split(InputBox(prompt, "顯示序位選擇器", mid(k, 2)), ",")

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

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        myMixName = .Cells(r, "J")
        recDate = .Cells(r, "B")
        myContent = .Cells(r, "D")
    
        If myMixName = MixName Then
    
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

Exit Function

ERRORHANDLE:

MsgBox ("第" & r & "列," & MixName & "項目的" & myContent_split & "<重複>"), vbCritical

End Function

Function getMixItems()

Dim coll As New Collection

With Sheets("Records")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
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

