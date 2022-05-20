Attribute VB_Name = "Autosize"

'constantes para usar con SendMessage
Private Const LVM_FIRST          As Long = &H1000

Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
  
' Declaración del Api SendMessage para hacer el AutoSize de las columnas
Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal Hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Public Sub redimencionar()

    Dim La_Columna As Long

    With ListView

        'Recorre los encabezados
        For La_Columna = 2 To 2
            ' Aplica el Autoajuste ( -1 mediante el item, -2 mediante el caption de la columna)
            Call ListView_AutoSize(FrmMain.ListView, La_Columna, -1)
        Next
      
        'Recorre los encabezados
        For La_Columna = 1 To 161
            ' Aplica el Autoajuste ( -1 mediante el item, -2 mediante el caption de la columna)
            Call ListView_AutoSize(FrmMain.ListView, La_Columna, -2)
        Next
     
        'Recorre los encabezados
        For La_Columna = 100 To 100
            ' Aplica el Autoajuste ( -1 mediante el item, -2 mediante el caption de la columna)
            Call ListView_AutoSize(FrmMain.ListView, La_Columna, -1)
        Next

    End With
   
End Sub

Public Sub redimenciona()

    Dim La_Columna As Long

    With FrmMain

        'Recorre los encabezados
        For La_Columna = 2 To 2
            ' Aplica el Autoajuste ( -1 mediante el item, -2 mediante el caption de la columna)
            Call ListView_AutoSize(.Listnpcs, La_Columna, -1)
        Next
      
        'Recorre los encabezados
        For La_Columna = 1 To 73
            ' Aplica el Autoajuste ( -1 mediante el item, -2 mediante el caption de la columna)
            Call ListView_AutoSize(.Listnpcs, La_Columna, -2)
        Next
     
        'Recorre los encabezados
        For La_Columna = 100 To 100
            ' Aplica el Autoajuste ( -1 mediante el item, -2 mediante el caption de la columna)
            Call ListView_AutoSize(.Listnpcs, La_Columna, -1)
        Next

    End With
   
End Sub

Private Sub ListView_AutoSize(El_ListView As ListView, _
                              ByVal La_Columna As Long, _
                              ByVal Modo_De_Ajuste As Long)
  
    With El_ListView
        
        'si no está el ListView en modo reporte sale
        If .View = lvwReport Then

            'Si hay columnas
            If La_Columna >= 1 And La_Columna <= .ColumnHeaders.Count Then
                'Establece el Autosize
                Call SendMessage(.Hwnd, LVM_SETCOLUMNWIDTH, La_Columna - 1, ByVal Modo_De_Ajuste)

            End If

        End If

    End With

End Sub

