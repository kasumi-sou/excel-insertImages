Attribute VB_Name = "Module1"
Option Explicit
Sub insertImage()
  Dim ws As Worksheet
  Dim selectedFiles As Variant
  Dim file As Variant
  Dim pic As Shape
  Dim startCell As Range
  Dim i As Long
  Dim cellWidth As Double
  Dim cellHeight As Double
  Dim picRatio As Double
  Dim nextCell As Long
  Dim filecount As Integer
  Dim rowSize As Integer
  Dim colSize As Integer
  
  
  'シート名を入力
  Set ws = ActiveSheet
  
  '画像を挿入する最初のセル
  Set startCell = ws.Range("B7")
  
  '挿入間隔
  nextCell = 3
  
  '画像ファイル(複数)を選択
  selectedFiles = Application.GetOpenFilename(MultiSelect:=True)
  
  If IsArray(selectedFiles) Then
    '画像ファイル数を取得
    filecount = UBound(selectedFiles) - LBound(selectedFiles) + 1
    
    If filecount <= 20 Then
      For Each file In selectedFiles
        '画像挿入
        Set pic = ws.Shapes.AddPicture(file, msoFalse, msoCTrue, 0, 0, -1, -1)
        
        'セルサイズ取得
        If startCell.MergeCells Then
          
          rowSize = startCell.MergeArea.Rows.Count
          colSize = startCell.MergeArea.Columns.Count
          
          cellWidth = startCell.Width * colSize
          cellHeight = startCell.Height * rowSize
         Else
          cellWidth = startCell.Width
          cellHeight = startCell.Height
        End If
        
        '画像リサイズ
        'picRatio = pic.Width / pic.Height
        pic.Width = cellWidth
        pic.Height = cellHeight
        
        '画像を指定セルに配置
        pic.Top = startCell.Top
        pic.Left = startCell.Left + cellWidth / 2
        
        'セルs駆除などで移動するが、リサイズはしない
        pic.Placement = xlMove
        
        '次セルに移動
        Set startCell = startCell.Offset(nextCell, 0)
      Next file
    ElseIf filecount > 20 Then
      MsgBox "選択ファイル数は20以下", vbInformation
    End If
    MsgBox "処理完了", vbInformation
  Else
  End If
End Sub

'2025/02/11 作成
