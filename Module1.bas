Attribute VB_Name = "Module1"
Sub ����1()
Attribute ����1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����1 ����
'

'
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 78, 70.5, 150, 142.5). _
        Select
    Selection.ShapeRange.ScaleWidth 1.6145833333, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0, msoFalse, msoScaleFromTopLeft
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 6
    End With
End Sub
Sub ����2()
Attribute ����2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����2 ����
'

'
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 11.25, 24.75, 50.25 _
        , 18.75).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "vu;4ja4" & Chr(13) & ""
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 8).ParagraphFormat. _
        FirstLineIndent = 0
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 8).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With
    Range("E12").Select
End Sub
Sub ����3()
Attribute ����3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����3 ����
'

'
    ActiveSheet.Shapes.AddShape(msoShapeRightBrace, 114.75, 105, 14.25, 141).Select
    'Selection.ShapeRange.IncrementRotation 270
End Sub
Sub ����4()
Attribute ����4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����4 ����
'

'
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 91.3636220472, _
        49.7727559055, 24.8863779528, 16.3636220472).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "2"
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 1).ParagraphFormat. _
        FirstLineIndent = 0
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 1).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 8
        .Name = "+mn-lt"
    End With
    Range("C4").Select
    ActiveSheet.Shapes.Range(Array("TextBox 972")).Select
    Selection.ShapeRange.Fill.Visible = msoFalse
End Sub
Sub ����5()
Attribute ����5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����5 ����
'

'
    ActiveSheet.Shapes.Range(Array("TextBox 11141")).Select
    Selection.ShapeRange.Fill.Visible = msoFalse
    Selection.ShapeRange.Line.Visible = msoFalse
End Sub
Sub ����6()
Attribute ����6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����6 ����
'

'
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 1.1538582677, 1.1538582677, _
        1047.1153543307, 306.9230708661).Select
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With

End Sub
Sub ����7()
Attribute ����7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����7 ����
'

'
    Selection.ShapeRange.ZOrder msoSendToBack
End Sub
Sub ����8()
Attribute ����8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����8 ����
'

'
    ActiveSheet.Shapes.Range(Array("TextBox 12389")).Select
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6).Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
        .Solid
    End With
End Sub
