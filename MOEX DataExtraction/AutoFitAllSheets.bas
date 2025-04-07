Attribute VB_Name = "Module1"
Sub AutoFitAllSheets()
    Dim ws As Worksheet
    Dim lastCol As Long
    
    ' ������� ���� ������ � �������� �����
    For Each ws In ThisWorkbook.Worksheets
        ' ����� ��������� ����������� ������� �� �����
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' ���������� ������ ���� ����������� ��������
        ws.Columns("A:" & Split(ws.Cells(1, lastCol).Address, "$")(1)).AutoFit
    Next ws
End Sub

