Attribute VB_Name = "main1"
Option Explicit
Sub data_set()
    Application.ScreenUpdating = False

    'シートにデータがある場合は削除する
    Call del
    
    'グラフ作成用のデータを生成
    Call file_open
    Call writing2
    Application.Quit
    ThisWorkbook.Close

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
    
    txtName = Application.GetOpenFilename("テキストファイル,*.log")
    
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

        arry1 = Split(buf, ",")


        Dim i As Long
            For i = LBound(arry1) To UBound(arry1)
            Cells(r, i + 1) = arry1(i)
            Next
            r = r + 1
    Loop
       
    Close #1
    
End Sub
Sub writing2()
    Application.ScreenUpdating = False
    
    '不要データの削除
    Columns("C:F").delete
    Columns("D").delete
    
    Dim j As Variant
    j = 1
    
    Do While Cells(j, 3) <> ""
         Cells(j, 4) = Replace(Cells(j, 3), ":", ";")
          j = j + 1
    Loop
    
    Columns("C").delete
    
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\Marge.csv", FileFormat:=xlCSV
    ThisWorkbook.Close
    
End Sub


