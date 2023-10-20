Attribute VB_Name = "Graph_1"
Option Explicit
Sub Graph1()
        
        Range("B1:D1").Select
        Range(Selection, Selection.End(xlDown)).Select
        
        'Graph_Create
        ActiveSheet.Shapes.AddChart2.Select
        'Graph_Type
        ActiveChart.ChartType = xlLine
        'Add_Title
        ActiveChart.HasTitle = True
        'Title_Name
        ActiveChart.ChartTitle.Text = "TEMP"

End Sub
