Attribute VB_Name = "Graph_2"
Option Explicit
Sub Graph2()
        
        Range("B1:D1").Select
        Range(Selection, Selection.End(xlDown)).Select
        
        'Graph_Create
        ActiveSheet.Shapes.AddChart2.Select
        'Graph_Type
        ActiveChart.ChartType = xlXYScatter 'éUïzê}
        
        'Add_Title
        ActiveChart.HasTitle = True
        'Title_Name
        ActiveChart.ChartTitle.Text = "TEMP"
        'éûä‘é≤ÇÃà íuÇâ∫Ç…Ç∑ÇÈ
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
