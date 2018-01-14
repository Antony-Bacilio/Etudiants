Attribute VB_Name = "Module1"
Sub semaine_clic()
    Set F = Worksheets("etudiant")
    'semaine courante
    sem = F.range("B2").Value
    'activation du graphique
    F.ChartObjects("graph_abs").Activate
    'étiquettes des abscisses
    ActiveChart.SeriesCollection(1).XValues = F.range("G2:G" & (sem + 1))
    'données des 4 séries
    ActiveChart.SeriesCollection(3).Values = F.range("I2:I" & (sem + 1))
    ActiveChart.SeriesCollection(2).Values = F.range("J2:J" & (sem + 1))
    ActiveChart.SeriesCollection(1).Values = F.range("K2:K" & (sem + 1))
    ActiveChart.SeriesCollection(4).Values = F.range("H2:H" & (sem + 1))

End Sub
Sub graphique_aires()
    Set F = Worksheets("etudiant")
    F.ChartObjects("graph_abs").Activate
    '   aires empilées 100 % (j'ai trouvé la valeur de la constante avec l'enregistreur de macro)
    ActiveChart.ChartType = xlAreaStacked100
End Sub

Sub graphique_barres()
    Set F = Worksheets("etudiant")
    F.ChartObjects("graph_abs").Activate
    'barres empilées 100 % (j'ai trouvé la valeur de la constante avec l'enregistreur de macro)
    ActiveChart.ChartType = xlColumnStacked100
End Sub
