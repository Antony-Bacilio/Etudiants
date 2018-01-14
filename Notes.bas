Attribute VB_Name = "Module2"
Sub chercherNotes()

    Dim range As Variant
    Set F_etud = Worksheets("etudiant")
    Set F_insc = Worksheets("inscrits")
    Set F_not = Worksheets("notes")
    Set F_mod = Worksheets("modules")
    
    Set no_etud = F_etud.Cells(2, 1)
    Set nom_etud = F_insc.Cells(no_etud, 1)
    Set cell_etud = F_not.Cells.Find(nom_etud, LookIn:=xlValues, lookat:=xlWhole)

   For lig = 35 To 60
       F_etud.range("G35:L49") = ""
    Next lig
    
    
    ligMoy = 34
    
    For lig = 1 To 15
   
        Set nom_mod = F_mod.Cells(lig, 1)
        Set no_cof = F_mod.Cells(lig, 2)
        Set cell_mod = F_not.Cells.Find(nom_mod, LookIn:=xlValues, lookat:=xlWhole)
        
        Set note = F_not.Cells(cell_etud.Row, cell_mod.Column)
        Set promo = F_not.Cells(23, cell_mod.Column)
        Set Min = F_not.Cells(24, cell_mod.Column)
        Set Max = F_not.Cells(25, cell_mod.Column)
        
            If note <> "" Then
                F_etud.Cells(lig + ligMoy, 7) = no_cof
                F_etud.Cells(lig + ligMoy, 8) = nom_mod
                F_etud.Cells(lig + ligMoy, 9) = note
                F_etud.Cells(lig + ligMoy, 10) = promo
                F_etud.Cells(lig + ligMoy, 11) = Min
                F_etud.Cells(lig + ligMoy, 12) = Max
                
    
            Else
                ligMoy = ligMoy - 1
                
            End If
            
            
    Next lig
    
    ActiveSheet.ChartObjects("graph_radar").Activate
    ActiveChart.SetSourceData Source:=F_etud.range(Cells(34, 8), Cells(15 + ligMoy, 12))
    
    For lig = 35 To 49

        ligEval = F_etud.range("G" & lig)
        
        If ligEval <> "" Then
        
            somme_cof = somme_cof + F_etud.range("G" & (lig))

            sommes_notes_etud = sommes_notes_etud + (ligEval * F_etud.range("I" & lig))
            sommes_notes_promo = sommes_notes_promo + (ligEval * F_etud.range("J" & lig))
            sommes_notes_min = sommes_notes_min + (ligEval * F_etud.range("K" & lig))
            sommes_notes_max = sommes_notes_max + (ligEval * F_etud.range("L" & lig))
            
            moyenne_etud = (sommes_notes_etud / somme_cof)
            moyenne_promo = (sommes_notes_promo / somme_cof)
            moyenne_min = (sommes_notes_min / somme_cof)
            moyenne_max = (sommes_notes_max / somme_cof)


        Else
            F_etud.range("I" & lig) = moyenne_etud
            F_etud.range("J" & lig) = moyenne_promo
            F_etud.range("K" & lig) = moyenne_min
            F_etud.range("L" & lig) = moyenne_max
            
            F_etud.ChartObjects("graphe_anneaux").Activate
            ActiveChart.SetElement (msoElementDataLabelNone)
            
            'F_etud.ChartObjects("graphe_thr").Activate
            'ActiveChart.SetElement (msoElementDataLabelNone)
            
            note_etud = F_etud.range("I" & lig)
            
            If note_etud < 8 Then
               F_etud.range("C68:C73") = 0
               F_etud.range("R61:R66") = 0
               F_etud.range("C68") = note_etud
               F_etud.range("R61") = note_etud
               ActiveChart.FullSeriesCollection(1).Points(1).ApplyDataLabels
               
            ElseIf note_etud >= 8 And note_etud < 10 Then
                F_etud.range("C68:C73") = 0
                F_etud.range("R61:R66") = 0
                F_etud.range("C69") = note_etud
                F_etud.range("R62") = note_etud
                ActiveChart.FullSeriesCollection(1).Points(2).ApplyDataLabels
                
            ElseIf note_etud >= 10 And note_etud < 12 Then
                F_etud.range("C68:C73") = 0
                F_etud.range("R61:R66") = 0
                F_etud.range("C70") = note_etud
                F_etud.range("R63") = note_etud
                ActiveChart.FullSeriesCollection(1).Points(3).ApplyDataLabels
                
            ElseIf note_etud >= 12 And note_etud < 14 Then
                F_etud.range("C68:C73") = 0
                F_etud.range("R61:R66") = 0
                F_etud.range("C71") = note_etud
                F_etud.range("R64") = note_etud
                ActiveChart.FullSeriesCollection(1).Points(4).ApplyDataLabels
                
            ElseIf note_etud >= 14 And note_etud < 16 Then
                F_etud.range("C68:C73") = 0
                F_etud.range("R61:R66") = 0
                F_etud.range("C72") = note_etud
                F_etud.range("R65") = note_etud
                ActiveChart.FullSeriesCollection(1).Points(5).ApplyDataLabels
            
            ElseIf note_etud >= 16 Then
                F_etud.range("C68:C73") = 0
                F_etud.range("R61:R66") = 0
                F_etud.range("C73") = note_etud
                F_etud.range("R66") = note_etud
                ActiveChart.FullSeriesCollection(1).Points(6).ApplyDataLabels
            End If
            
            Exit For
            
        End If
        
    Next lig
    
End Sub
    
