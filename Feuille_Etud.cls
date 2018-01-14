VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub CheckBox1_Click()
    If CheckBox1.Value = False Then
        ActiveSheet.ChartObjects("graph_radar").Activate
        ActiveChart.FullSeriesCollection(2).Select
        Selection.Format.Line.Visible = msoFalse
    Else
        ActiveSheet.ChartObjects("graph_radar").Activate
        ActiveChart.FullSeriesCollection(2).Select
        Selection.Format.Line.Visible = msoTrue
    End If
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2.Value = False Then
        ActiveSheet.ChartObjects("graph_radar").Activate
        ActiveChart.FullSeriesCollection(3).Select
        Selection.Format.Line.Visible = msoFalse
    Else
        ActiveSheet.ChartObjects("graph_radar").Activate
        ActiveChart.FullSeriesCollection(3).Select
        Selection.Format.Line.Visible = msoTrue
    End If
End Sub

Private Sub CheckBox4_Click()
    If CheckBox4.Value = False Then
        ActiveSheet.ChartObjects("graph_radar").Activate
        ActiveChart.FullSeriesCollection(4).Select
        Selection.Format.Line.Visible = msoFalse
    Else
        ActiveSheet.ChartObjects("graph_radar").Activate
        ActiveChart.FullSeriesCollection(4).Select
        Selection.Format.Line.Visible = msoTrue
    End If
End Sub
