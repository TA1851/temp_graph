Attribute VB_Name = "main2"
Option Explicit
Sub single_macro()
    Application.ScreenUpdating = False

    '�V�[�g�Ƀf�[�^������ꍇ�͍폜����
    Call del
    
    '�O���t�쐬�p�̃f�[�^�𐶐�
    Call file_open
    Call writing2

End Sub
Sub del()
    Cells.clear
    If ActiveSheet.ChartObjects.Count = 1 Then
        ActiveSheet.ChartObjects.delete
        Else
    End If
End Sub
Sub file_open()
    Application.ScreenUpdating = False
    
    Dim txtName As String  '�_�C�A���O�ŔC�ӂ̃t�@�C�����J��
    Dim buf As String
    
    txtName = Application.GetOpenFilename("�e�L�X�g�t�@�C��,*.csv")
    
    If txtName <> "False" Then  '�ǂݍ��݃��[�h�� file open
        Open txtName For Input As #1
    End If
    
    Call writing1
    
End Sub
Sub writing1()
    Application.ScreenUpdating = False
    
    Dim r As Long

    r = 1 '1�s�ڂ��珑���o��

    'EOF�֐��Ńt�@�C����ǂݍ���
    Do Until EOF(1)
        Dim buf As String
        Line Input #1, buf
        
        Dim arry1 As Variant
        Dim arry2 As Variant

        arry1 = Replace(buf, ";", ",")   '�u��
        arry2 = Split(arry1, ",")


        Dim i As Long
            For i = LBound(arry2) To UBound(arry2)
            Cells(r, i + 1) = arry2(i)
            Next
            r = r + 1
    Loop
            
    Close #1
    
End Sub
Sub writing2()
    Application.ScreenUpdating = False
    
    '�s�v�f�[�^�̍폜
    Columns("C").delete
    
    Dim j As Variant
    j = 1
    
    '���x�̌v�Z
    Do While Cells(j, 3) <> ""
        Cells(j, 4) = Cells(j, 3) * 0.1
        j = j + 1
    Loop
    
    Columns("C").delete
    Columns("C").Select
    Selection.NumberFormatLocal = "0.0"
    
    '1�s�ڂɍs��}������
    Cells(1, 1).EntireRow.Insert

    Range("B1").Value = "Time"
    Range("C1").Value = "Runnung"
    Range("D1").Value = "Not Runnung"

    Columns("A:D").Select
    Selection.AutoFilter
    
End Sub


