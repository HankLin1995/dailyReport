Attribute VB_Name = "word_test"
Sub ReplaceText()

Dim rng As Word.Range
Set rng = Word.ActiveDocument.Range

rng.Find.ClearFormatting
rng.Find.Replacement.ClearFormatting
rng.Find.Text = "<施工位置>"
rng.Find.Replacement.Text = "測試"
rng.Find.Execute Replace:=wdReplaceAll

'rng.Find.Text = "要被替?的文本2"
'rng.Find.Replacement.Text = "替?文本2"
'rng.Find.Execute Replace:=wdReplaceAll

End Sub

Sub PrintSpecificPages()
'Declare variables
Dim wdApp As Word.Application
Dim wdDoc As Word.Document
Dim strPageRange As String

'Start a new instance of Word
Set wdApp = New Word.Application
wdApp.Visible = True

'Open the document you want to print
Set wdDoc = wdApp.Documents.Open("d:\Users\USER\Desktop\word_vba_test\監造計畫書Ver20.docx")

Call ReplaceText

'Prompt the user for the page range to print
strPageRange = InputBox("Enter the page range to print (e.g. 1-3, 5):", "Print Specific Pages")

'Print the specified page range
wdDoc.PrintOut Range:=wdPrintFromTo, From:=strPageRange, To:=strPageRange

'Close the document
wdDoc.Close False

'Quit the instance of Word
wdApp.Quit

End Sub

Sub ExtractPages()

Dim startPage As Integer
Dim endPage As Integer
Dim sourceDoc As Document
Dim targetDoc As Document

startPage = InputBox("請輸入要開始的頁數：")
endPage = InputBox("請輸入要結束的頁數：")

Set sourceDoc = Word.ActiveDocument

sourceDoc.Range(Start:=sourceDoc.Range.Start + (startPage - 1) * _
sourceDoc.Range.Information(wdActiveEndAdjustedPageNumber), _
End:=sourceDoc.Range.Start + (endPage - 1) * _
sourceDoc.Range.Information(wdActiveEndAdjustedPageNumber) + _
sourceDoc.Range.ComputeStatistics(wdStatisticPages)).Copy

Set targetDoc = Documents.Add
targetDoc.Range.Paste
targetDoc.Range.ParagraphFormat.SpaceAfter = sourceDoc.Range.ParagraphFormat.SpaceAfter
targetDoc.Range.Font.Name = sourceDoc.Range.Font.Name
targetDoc.Range.Font.Size = sourceDoc.Range.Font.Size
targetDoc.Range.ParagraphFormat.Alignment = sourceDoc.Range.ParagraphFormat.Alignment
targetDoc.SaveAs Filename:="提取的頁面.docx", FileFormat:=wdFormatXMLDocument

End Sub
