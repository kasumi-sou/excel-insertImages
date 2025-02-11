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
  
  
  '�V�[�g�������
  Set ws = ActiveSheet
  
  '�摜��}������ŏ��̃Z��
  Set startCell = ws.Range("B7")
  
  '�}���Ԋu
  nextCell = 3
  
  '�摜�t�@�C��(����)��I��
  selectedFiles = Application.GetOpenFilename(MultiSelect:=True)
  
  If IsArray(selectedFiles) Then
    '�摜�t�@�C�������擾
    filecount = UBound(selectedFiles) - LBound(selectedFiles) + 1
    
    If filecount <= 20 Then
      For Each file In selectedFiles
        '�摜�}��
        Set pic = ws.Shapes.AddPicture(file, msoFalse, msoCTrue, 0, 0, -1, -1)
        
        '�Z���T�C�Y�擾
        If startCell.MergeCells Then
          
          rowSize = startCell.MergeArea.Rows.Count
          colSize = startCell.MergeArea.Columns.Count
          
          cellWidth = startCell.Width * colSize
          cellHeight = startCell.Height * rowSize
         Else
          cellWidth = startCell.Width
          cellHeight = startCell.Height
        End If
        
        '�摜���T�C�Y
        'picRatio = pic.Width / pic.Height
        pic.Width = cellWidth
        pic.Height = cellHeight
        
        '�摜���w��Z���ɔz�u
        pic.Top = startCell.Top
        pic.Left = startCell.Left + cellWidth / 2
        
        '�Z��s�쏜�Ȃǂňړ����邪�A���T�C�Y�͂��Ȃ�
        pic.Placement = xlMove
        
        '���Z���Ɉړ�
        Set startCell = startCell.Offset(nextCell, 0)
      Next file
    ElseIf filecount > 20 Then
      MsgBox "�I���t�@�C������20�ȉ�", vbInformation
    End If
    MsgBox "��������", vbInformation
  Else
  End If
End Sub

'2025/02/11 �쐬
