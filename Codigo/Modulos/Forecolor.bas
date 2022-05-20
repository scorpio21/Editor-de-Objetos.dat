Attribute VB_Name = "Forecolor"

Public Sub ForeColorColumn(Color, Columna As Integer, Lv As ListView)

    Dim item As ListItem

    Dim i    As Integer
    
    ' Verifica que el control contenga items
    If Lv.ListItems.Count = 0 Then
        Exit Sub

    End If
      
    ' Verifica que el Listview esté en vista de reporte
    If Lv.View <> lvwReport Then
        MsgBox "el listview debe estar en vista reporte", vbQuestion
        Exit Sub

    End If
      
    ' chequea el índice de la columna que sea válido
    If (Columna + 1) > Lv.ColumnHeaders.Count Then
        MsgBox "El número de columna está fuera del intervalo"
        Exit Sub

    End If
     
    ' Color de fuente Para los items
    If Columna = 0 Then

        ' recorre la lista
        For i = 1 To Lv.ListItems.Count
            ' cambia el color de la fuente de la columna indicada
            Lv.ListItems(i).Forecolor = Color
        Next
        ' Color de fuente para los Subitems
    Else

        ' recorre
        For i = 1 To Lv.ListItems.Count
            ' cambia el color de la fuente de la columna indicada
            Lv.ListItems(i).ListSubItems(Columna).Forecolor = Color
        Next

    End If
    
    If Columna = 1 Then
    
        For i = 1 To Lv.ListItems.Count
      
            If UCase$(Lv.ListItems(i).ListSubItems(Columna).Text) = "LIBRE" Then
                Lv.ListItems(i).ListSubItems(Columna).Forecolor = vbRed
            Else
                Lv.ListItems(i).ListSubItems(Columna).Forecolor = vbYellow

            End If
      
        Next i
    
    End If
    
    ' refresca el control
    Lv.Refresh
      
End Sub
