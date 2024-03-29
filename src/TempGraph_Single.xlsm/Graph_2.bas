Attribute VB_Name = "Graph_2"
Option Explicit
Sub Graph2()
        
        Range("B1:D1").Select
        Range(Selection, Selection.End(xlDown)).Select
        
        'Graph_Create
        ActiveSheet.Shapes.AddChart2.Select
        'Graph_Type
        ActiveChart.ChartType = xlXYScatter '散布図
        
        'Add_Title
        ActiveChart.HasTitle = True
        'Title_Name
        ActiveChart.ChartTitle.Text = "TEMP"
        '時間軸の位置を下にする
        ActiveChart.Axes(xlCategory).Select
        Selection.TickLabelPosition = xlLow
        
        With ActiveChart
        .Axes(xlCategory).Select
        '.Axes(xlCategory).MajorUnit = 0.041667 '1hr
        '.Axes(xlCategory).MajorUnit = 0.0125 '3hr
        .Axes(xlCategory).MajorUnit = 0.5 '12hr
        .Axes(xlCategory).MinorUnit = 0.01
End With

End Sub
