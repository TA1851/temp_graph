Attribute VB_Name = "main2"
Option Explicit
Sub single_macro()
    Application.ScreenUpdating = False

    'シートにデータがある場合は削除する
    Call del
    
    'グラフ作成用のデータを生成
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
    
    Dim txtName As String  'ダイアログで任意のファイルを開く
    Dim buf As String
    
    txtName = Application.GetOpenFilename("テキストファイル,*.csv")
    
    If txtName <> "False" Then  '読み込みモードで file open
        Open txtName For Input As #1
    End If
    
    Call writing1
    
End Sub
Sub writing1()
    Application.ScreenUpdating = False
    
    Dim r As Long

    r = 1 '1行目から書き出す

    'EOF関数でファイルを読み込む
    Do Until EOF(1)
        Dim buf As String
        Line Input #1, buf
        
        Dim arry1 As Variant
        Dim arry2 As Variant

        arry1 = Replace(buf, ";", ",")   '置換
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
    
    '不要データの削除
    Columns("C").delete
    
    Dim j As Variant
    j = 1
    
    '温度の計算
    Do While Cells(j, 3) <> ""
        Cells(j, 4) = Cells(j, 3) * 0.1
        j = j + 1
    Loop
    
    Columns("C").delete
    Columns("C").Select
    Selection.NumberFormatLocal = "0.0"
    
    '1行目に行を挿入する
    Cells(1, 1).EntireRow.Insert

    Range("B1").Value = "Time"
    Range("C1").Value = "Runnung"
    Range("D1").Value = "Not Runnung"

    Columns("A:D").Select
    Selection.AutoFilter
    
End Sub


