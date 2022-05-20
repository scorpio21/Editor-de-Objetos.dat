Attribute VB_Name = "Guardarobj"

Public Function Guardar()

    'FrmMain.Barra.Visible = False
    FrmMain.ucProgress.Visible = False
    'FrmMain.Bgrabar.Visible = True
    ucProgress.Visible = True
    
    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal

    Open SavePath & "Objo" & ".dat" For Output Shared As #nfile 'Abrimos el Fichero
    Call DatosObj(nfile)
    Print #nfile, "[INIT]"
    Print #nfile, "NumObjs=" & FrmMain.ListView.ListItems.Count
    'FrmMain.Bgrabar.max = FrmMain.ListView.ListItems.Count
    FrmMain.ucBgrabar.max = FrmMain.ListView.ListItems.Count
    Call Espacio(nfile)

    'Print #nfile, vbCrLf
    For i = 1 To FrmMain.ListView.ListItems.Count

        'Comienzo Pack Premium (Dat by Lugus, Programado por Helios)
        If FrmMain.ListView.ListItems.item(i).SubItems(1) = "Pack 1 (Mago humano)" Then
            Call DatosPack(nfile)
            Call Espacio(nfile)
            'Print #nfile, vbCrLf
            'Fin Pack Premium (Dat by Lugus, Programado por Helios)
        ElseIf FrmMain.ListView.ListItems.item(i).SubItems(1) = "Alas Sagradas" Then
            Call DatosfinPack(nfile)
            Call Espacio(nfile)

            'Print #nfile, vbCrLf
        End If

        'Añadimos el Numero de OBJ"
        Print #nfile, "[OBJ" & i & "]"

        For X = 1 To FrmMain.ListView.ColumnHeaders.Count - 1

            If FrmMain.ListView.ListItems.item(i).SubItems(X) <> "" Then
                Print #nfile, FrmMain.ListView.ColumnHeaders(X + 1).Text & "=" & FrmMain.ListView.ListItems.item(i).SubItems(X)

            End If
            
        Next X

        Call Espacio(nfile)
        'Print #nfile, vbCrLf
        'FrmMain.Bgrabar.value = i
        FrmMain.ucBgrabar.value = i
    Next i

    Close #nfile
    FrmMain.Lbltexto.Visible = True
    FrmMain.Lbltexto.Caption = "Objetos guardados correctamente."
    FrmMain.Timer.Enabled = True
    FrmMain.Timer2.Enabled = True
    FrmMain.Timer3.Enabled = True

End Function

Sub SaveOBJData()

    On Error GoTo ErrHandler

    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal

    If FileExist(App.path & "\" & "Obj.dat", vbNormal) Then
        Kill App.path & "\" & "Obj.dat"

    End If
    
    If NumObjDatas = 0 Then Exit Sub
    ' FrmMain.Barra.Visible = False
    FrmMain.ucProgress.Visible = False
    'FrmMain.Bgrabar.Visible = True
    FrmMain.ucBgrabar.Visible = True

    Dim Object As Integer, fObj As String
    
    Call CabezeraBlockObj
    
    fObj = App.path & "\" & "Obj.dat"
    
    Call FP_Append(fObj, "[INIT]" & vbCrLf)
    Call FP_Append(fObj, "NumsObjs=" & NumObjDatas & vbCrLf & vbCrLf)
    'FrmMain.Bgrabar.max = NumObjDatas
    FrmMain.ucBgrabar.max = NumObjDatas

    'Llena la lista
    For Object = 1 To NumObjDatas
        
        If ObjData(Object).Name <> "" Then
            Call FP_Append(fObj, "[OBJ" & Object & "]" & vbCrLf)
            Call FP_Append(fObj, "Name=" & ObjData(Object).Name & vbCrLf)

        End If
        
        If ObjData(Object).Caos > 0 Then Call FP_Append(fObj, "Caos=" & Val(ObjData(Object).Caos) & vbCrLf)
        If ObjData(Object).Real > 0 Then Call FP_Append(fObj, "Real=" & Val(ObjData(Object).Real) & vbCrLf)
        If ObjData(Object).templ > 0 Then Call FP_Append(fObj, "Templ=" & Val(ObjData(Object).templ) & vbCrLf)
        If ObjData(Object).Nemes > 0 Then Call FP_Append(fObj, "Nemes=" & Val(ObjData(Object).Nemes) & vbCrLf)
        If ObjData(Object).Facc > 0 Then Call FP_Append(fObj, "Facc=" & Val(ObjData(Object).Facc) & vbCrLf)
        If ObjData(Object).Cae > 0 Then Call FP_Append(fObj, "Cae=" & Val(ObjData(Object).Cae) & vbCrLf)
        If ObjData(Object).TiRaRObj > 0 Then Call FP_Append(fObj, "Tirar=" & Val(ObjData(Object).TiRaRObj) & vbCrLf)

        If ObjData(Object).Robable > 0 Then Call FP_Append(fObj, "NoRobable=" & Val(ObjData(Object).Robable) & vbCrLf)
        If ObjData(Object).bonifica > 0 Then Call FP_Append(fObj, "bonifica=" & Val(ObjData(Object).bonifica) & vbCrLf)
        If ObjData(Object).tipobonifica > 0 Then Call FP_Append(fObj, "tipobonifica=" & Val(ObjData(Object).tipobonifica) & vbCrLf)
        If ObjData(Object).Donacion > 0 Then Call FP_Append(fObj, "Donacion=" & Val(ObjData(Object).Donacion) & vbCrLf)

        If ObjData(Object).GrhIndex > 0 Then Call FP_Append(fObj, "GrhIndex=" & Val(ObjData(Object).GrhIndex) & vbCrLf)
        
        If ObjData(Object).ObjType > 0 Then Call FP_Append(fObj, "ObjType=" & Val(ObjData(Object).ObjType) & vbCrLf)
        If ObjData(Object).SubTipo > 0 Then Call FP_Append(fObj, "Subtipo=" & Val(ObjData(Object).SubTipo) & vbCrLf)
        
        If ObjData(Object).Newbie > 0 Then Call FP_Append(fObj, "Newbie=" & Val(ObjData(Object).Newbie) & vbCrLf)
        
        If ObjData(Object).Destruir > 0 Then Call FP_Append(fObj, "Destruir=" & Val(ObjData(Object).Destruir) & vbCrLf)

        If ObjData(Object).SubTipo = sObjtype.sotEscudo Then
        
            If ObjData(Object).ShieldAnim > 0 Then Call FP_Append(fObj, "Anim=" & Val(ObjData(Object).ShieldAnim) & vbCrLf)
            If ObjData(Object).LingH > 0 Then Call FP_Append(fObj, "LingH=" & Val(ObjData(Object).LingH) & vbCrLf)
            If ObjData(Object).LingP > 0 Then Call FP_Append(fObj, "LingP=" & Val(ObjData(Object).LingP) & vbCrLf)
            If ObjData(Object).LingO > 0 Then Call FP_Append(fObj, "LingO=" & Val(ObjData(Object).LingO) & vbCrLf)
            If ObjData(Object).LingM > 0 Then Call FP_Append(fObj, "LingM=" & Val(ObjData(Object).LingM) & vbCrLf)
            If ObjData(Object).Madera > 0 Then Call FP_Append(fObj, "Madera=" & Val(ObjData(Object).Madera) & vbCrLf)
            If ObjData(Object).MaderaElfica > 0 Then Call FP_Append(fObj, "MaderaElfica=" & Val(ObjData(Object).MaderaElfica) & vbCrLf)
            If ObjData(Object).SkHerreria > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreria) & vbCrLf)
            If ObjData(Object).Drena > 0 Then Call FP_Append(fObj, "Drena=" & Val(ObjData(Object).Drena) & vbCrLf)
            
            If ObjData(Object).SkHerreriaMagica > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreriaMagica) & vbCrLf)
            If ObjData(Object).Gemas > 0 Then Call FP_Append(fObj, "Gemas=" & Val(ObjData(Object).Gemas) & vbCrLf)
            If ObjData(Object).Diamantes > 0 Then Call FP_Append(fObj, "Diamantes=" & Val(ObjData(Object).Diamantes) & vbCrLf)
            
        End If

        If ObjData(Object).SubTipo = sObjtype.sotCasco Then    ' OBJTYPE_CASCO Then
                                   
            If ObjData(Object).CascoAnim > 0 Then Call FP_Append(fObj, "Anim=" & Val(ObjData(Object).CascoAnim) & vbCrLf)
            If ObjData(Object).LingH > 0 Then Call FP_Append(fObj, "LingH=" & Val(ObjData(Object).LingH) & vbCrLf)
            If ObjData(Object).LingP > 0 Then Call FP_Append(fObj, "LingP=" & Val(ObjData(Object).LingP) & vbCrLf)
            If ObjData(Object).LingO > 0 Then Call FP_Append(fObj, "LingO=" & Val(ObjData(Object).LingO) & vbCrLf)
            If ObjData(Object).LingM > 0 Then Call FP_Append(fObj, "LingM=" & Val(ObjData(Object).LingM) & vbCrLf)
            If ObjData(Object).Madera > 0 Then Call FP_Append(fObj, "Madera=" & Val(ObjData(Object).Madera) & vbCrLf)
            If ObjData(Object).MaderaElfica > 0 Then Call FP_Append(fObj, "MaderaElfica=" & Val(ObjData(Object).MaderaElfica) & vbCrLf)
            If ObjData(Object).SkHerreria > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreria) & vbCrLf)
            
            If ObjData(Object).DefensaMagicaMax > 0 Then Call FP_Append(fObj, "DefensaMagicaMax=" & Val(ObjData(Object).DefensaMagicaMax) & vbCrLf)
            If ObjData(Object).DefensaMagicaMin > 0 Then Call FP_Append(fObj, "DefensaMagicaMin=" & Val(ObjData(Object).DefensaMagicaMin) & vbCrLf)
            
            If ObjData(Object).SkHerreriaMagica > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreriaMagica) & vbCrLf)
            If ObjData(Object).Gemas > 0 Then Call FP_Append(fObj, "Gemas=" & Val(ObjData(Object).Gemas) & vbCrLf)
            If ObjData(Object).Diamantes > 0 Then Call FP_Append(fObj, "Diamantes=" & Val(ObjData(Object).Diamantes) & vbCrLf)
            
        End If
        
        If ObjData(Object).SubTipo = sObjtype.sotBotas Then
            
            If ObjData(Object).Botas > 0 Then Call FP_Append(fObj, "Anim=" & Val(ObjData(Object).Botas) & vbCrLf)
            If ObjData(Object).LingH > 0 Then Call FP_Append(fObj, "LingH=" & Val(ObjData(Object).LingH) & vbCrLf)
            If ObjData(Object).LingP > 0 Then Call FP_Append(fObj, "LingP=" & Val(ObjData(Object).LingP) & vbCrLf)
            If ObjData(Object).LingO > 0 Then Call FP_Append(fObj, "LingO=" & Val(ObjData(Object).LingO) & vbCrLf)
            If ObjData(Object).LingM > 0 Then Call FP_Append(fObj, "LingM=" & Val(ObjData(Object).LingM) & vbCrLf)
            If ObjData(Object).SkHerreria > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreria) & vbCrLf)
            
        End If

        If ObjData(Object).SubTipo = sObjtype.sotAlas Then
            
            If ObjData(Object).Alas > 0 Then Call FP_Append(fObj, "AlasAnim=" & Val(ObjData(Object).Alas) & vbCrLf)
            If ObjData(Object).LingH > 0 Then Call FP_Append(fObj, "LingH=" & Val(ObjData(Object).LingH) & vbCrLf)
            If ObjData(Object).LingP > 0 Then Call FP_Append(fObj, "LingP=" & Val(ObjData(Object).LingP) & vbCrLf)
            If ObjData(Object).LingO > 0 Then Call FP_Append(fObj, "LingO=" & Val(ObjData(Object).LingO) & vbCrLf)
            If ObjData(Object).LingM > 0 Then Call FP_Append(fObj, "LingM=" & Val(ObjData(Object).LingM) & vbCrLf)
            If ObjData(Object).Caos > 0 Then Call FP_Append(fObj, "Caos=" & Val(ObjData(Object).Caos) & vbCrLf)
            If ObjData(Object).Real > 0 Then Call FP_Append(fObj, "Real=" & Val(ObjData(Object).Real) & vbCrLf)
            If ObjData(Object).templ > 0 Then Call FP_Append(fObj, "Templ=" & Val(ObjData(Object).templ) & vbCrLf)
            If ObjData(Object).Nemes > 0 Then Call FP_Append(fObj, "Nemes=" & Val(ObjData(Object).Nemes) & vbCrLf)
            
            If ObjData(Object).SkHerreria > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreria) & vbCrLf)
            If ObjData(Object).Gemas > 0 Then Call FP_Append(fObj, "Gemas=" & Val(ObjData(Object).Gemas) & vbCrLf)
            If ObjData(Object).Diamantes > 0 Then Call FP_Append(fObj, "Diamantes=" & Val(ObjData(Object).Diamantes) & vbCrLf)
            If ObjData(Object).Resistencia > 0 Then Call FP_Append(fObj, "resistencia=" & Val(ObjData(Object).Resistencia) & vbCrLf)
         
        End If
        
        If ObjData(Object).Ropaje > 0 Then Call FP_Append(fObj, "NumRopaje=" & Val(ObjData(Object).Ropaje) & vbCrLf)
        If ObjData(Object).HechizoIndex > 0 Then Call FP_Append(fObj, "HechizoIndex=" & Val(ObjData(Object).HechizoIndex) & vbCrLf)
        If ObjData(Object).Clase > "" Then Call FP_Append(fObj, "Clase=" & ObjData(Object).Clase)
    
        If ObjData(Object).ObjType = eOBJType.otWeapon Then
            If ObjData(Object).WeaponAnim > 0 Then Call FP_Append(fObj, "Anim=" & Val(ObjData(Object).WeaponAnim) & vbCrLf)
            If ObjData(Object).Apuñala > 0 Then Call FP_Append(fObj, "Apuñala=" & Val(ObjData(Object).Apuñala) & vbCrLf)
            If ObjData(Object).Paraliza > 0 Then Call FP_Append(fObj, "Paraliza=" & Val(ObjData(Object).Paraliza) & vbCrLf)
            If ObjData(Object).Ceguera > 0 Then Call FP_Append(fObj, "Ceguera=" & Val(ObjData(Object).Ceguera) & vbCrLf)
            If ObjData(Object).Estupidez > 0 Then Call FP_Append(fObj, "Estupidez=" & Val(ObjData(Object).Estupidez) & vbCrLf)
            If ObjData(Object).Vida > 0 Then Call FP_Append(fObj, "Vida=" & Val(ObjData(Object).Vida) & vbCrLf)
            If ObjData(Object).Paraliza > 0 Then Call FP_Append(fObj, "Paraliza=" & Val(ObjData(Object).Paraliza) & vbCrLf)
            If ObjData(Object).Mana > 0 Then Call FP_Append(fObj, "Mana=" & Val(ObjData(Object).Mana) & vbCrLf)
            If ObjData(Object).Envenena > 0 Then Call FP_Append(fObj, "Envenena=" & Val(ObjData(Object).Envenena) & vbCrLf)
            If ObjData(Object).MaxHit > 0 Then Call FP_Append(fObj, "MaxHit=" & Val(ObjData(Object).MaxHit) & vbCrLf)
            If ObjData(Object).MinHit > 0 Then Call FP_Append(fObj, "MinHit=" & Val(ObjData(Object).MinHit) & vbCrLf)
            If ObjData(Object).LvlMin > 0 Then Call FP_Append(fObj, "LvlMin=" & Val(ObjData(Object).LvlMin) & vbCrLf)
            If ObjData(Object).LvlMax > 0 Then Call FP_Append(fObj, "LvlMax=" & Val(ObjData(Object).LvlMax) & vbCrLf)
            If ObjData(Object).Heroe > 0 Then Call FP_Append(fObj, "Heroe=" & Val(ObjData(Object).Heroe) & vbCrLf)
            If ObjData(Object).LingH > 0 Then Call FP_Append(fObj, "LingH=" & Val(ObjData(Object).LingH) & vbCrLf)
            If ObjData(Object).LingP > 0 Then Call FP_Append(fObj, "LingP=" & Val(ObjData(Object).LingP) & vbCrLf)
            If ObjData(Object).LingO > 0 Then Call FP_Append(fObj, "LingO=" & Val(ObjData(Object).LingO) & vbCrLf)
            If ObjData(Object).LingM > 0 Then Call FP_Append(fObj, "LingM=" & Val(ObjData(Object).LingM) & vbCrLf)
            If ObjData(Object).Caos > 0 Then Call FP_Append(fObj, "Caos=" & Val(ObjData(Object).Caos) & vbCrLf)
            If ObjData(Object).Real > 0 Then Call FP_Append(fObj, "Real=" & Val(ObjData(Object).Real) & vbCrLf)
            If ObjData(Object).templ > 0 Then Call FP_Append(fObj, "Templ=" & Val(ObjData(Object).templ) & vbCrLf)
            If ObjData(Object).Nemes > 0 Then Call FP_Append(fObj, "Nemes=" & Val(ObjData(Object).Nemes) & vbCrLf)
            If ObjData(Object).SkHerreria > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreria) & vbCrLf)
            If ObjData(Object).proyectil > 0 Then Call FP_Append(fObj, "Proyectil=" & Val(ObjData(Object).proyectil) & vbCrLf)
            If ObjData(Object).Municion > 0 Then Call FP_Append(fObj, "Municion=" & Val(ObjData(Object).Municion) & vbCrLf)
            If ObjData(Object).TipoMunicion > 0 Then Call FP_Append(fObj, "Municiones=" & Val(ObjData(Object).TipoMunicion) & vbCrLf)
            If ObjData(Object).DosManos > 0 Then Call FP_Append(fObj, "DosManos=" & Val(ObjData(Object).DosManos) & vbCrLf)
            If ObjData(Object).SkHerreriaMagica > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreriaMagica) & vbCrLf)
            If ObjData(Object).Gemas > 0 Then Call FP_Append(fObj, "Gemas=" & Val(ObjData(Object).Gemas) & vbCrLf)
            If ObjData(Object).Diamantes > 0 Then Call FP_Append(fObj, "Diamantes=" & Val(ObjData(Object).Diamantes) & vbCrLf)
            If ObjData(Object).Nivel > 0 Then Call FP_Append(fObj, "Nivel=" & Val(ObjData(Object).Nivel) & vbCrLf)
            If ObjData(Object).GameMaster > 0 Then Call FP_Append(fObj, "GM=" & Val(ObjData(Object).GameMaster) & vbCrLf)
            If ObjData(Object).UsoNpc > 0 Then Call FP_Append(fObj, "NPC=" & Val(ObjData(Object).UsoNpc) & vbCrLf)
            If ObjData(Object).Madera > 0 Then Call FP_Append(fObj, "Madera=" & Val(ObjData(Object).Madera) & vbCrLf)
            If ObjData(Object).Caos > 0 Then Call FP_Append(fObj, "MaderaElfica=" & Val(ObjData(Object).MaderaElfica) & vbCrLf)
    
        End If
        
        If ObjData(Object).ObjType = eOBJType.otArmadura Then
        
            If ObjData(Object).LingH > 0 Then Call FP_Append(fObj, "LingH=" & Val(ObjData(Object).LingH) & vbCrLf)
            If ObjData(Object).LingP > 0 Then Call FP_Append(fObj, "LingP=" & Val(ObjData(Object).LingP) & vbCrLf)
            If ObjData(Object).LingO > 0 Then Call FP_Append(fObj, "LingO=" & Val(ObjData(Object).LingO) & vbCrLf)
            If ObjData(Object).LingM > 0 Then Call FP_Append(fObj, "LingM=" & Val(ObjData(Object).LingM) & vbCrLf)
            If ObjData(Object).SkHerreria > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreria) & vbCrLf)
            If ObjData(Object).Caos > 0 Then Call FP_Append(fObj, "Caos=" & Val(ObjData(Object).Caos) & vbCrLf)
            If ObjData(Object).Real > 0 Then Call FP_Append(fObj, "Real=" & Val(ObjData(Object).Real) & vbCrLf)
            If ObjData(Object).templ > 0 Then Call FP_Append(fObj, "Templ=" & Val(ObjData(Object).templ) & vbCrLf)
            If ObjData(Object).Nemes > 0 Then Call FP_Append(fObj, "Nemes=" & Val(ObjData(Object).Nemes) & vbCrLf)
            If ObjData(Object).GameMaster > 0 Then Call FP_Append(fObj, "GM=" & Val(ObjData(Object).GameMaster) & vbCrLf)
            
            If ObjData(Object).MinDef > 0 Then Call FP_Append(fObj, "MINDEF=" & Val(ObjData(Object).MinDef) & vbCrLf)
            If ObjData(Object).MaxDef > 0 Then Call FP_Append(fObj, "MAXDEF=" & Val(ObjData(Object).MaxDef) & vbCrLf)
            If ObjData(Object).Resistencia > 0 Then Call FP_Append(fObj, "Resistencia=" & Val(ObjData(Object).Resistencia) & vbCrLf)
            
            If ObjData(Object).SkHerreria > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreriaMagica) & vbCrLf)
            If ObjData(Object).Gemas > 0 Then Call FP_Append(fObj, "Gemas=" & Val(ObjData(Object).Gemas) & vbCrLf)
            If ObjData(Object).Diamantes > 0 Then Call FP_Append(fObj, "Diamantes=" & Val(ObjData(Object).Diamantes) & vbCrLf)
            If ObjData(Object).Madera > 0 Then Call FP_Append(fObj, "Madera=" & Val(ObjData(Object).Madera) & vbCrLf)
            If ObjData(Object).MaderaElfica > 0 Then Call FP_Append(fObj, "MaderaElfica=" & Val(ObjData(Object).MaderaElfica) & vbCrLf)
            
        End If
        
        If ObjData(Object).ObjType = eOBJType.otHerramientas Then    ' OBJTYPE_HERRAMIENTAS Then
            
            If ObjData(Object).LingH > 0 Then Call FP_Append(fObj, "LingH=" & Val(ObjData(Object).LingH) & vbCrLf)
            If ObjData(Object).LingP > 0 Then Call FP_Append(fObj, "LingP=" & Val(ObjData(Object).LingP) & vbCrLf)
            If ObjData(Object).LingO > 0 Then Call FP_Append(fObj, "LingO=" & Val(ObjData(Object).LingO) & vbCrLf)
            If ObjData(Object).LingM > 0 Then Call FP_Append(fObj, "LingM=" & Val(ObjData(Object).LingM) & vbCrLf)
            If ObjData(Object).SkHerreria > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreria) & vbCrLf)
            If ObjData(Object).SkHerreriaMagica > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreriaMagica) & vbCrLf)
            If ObjData(Object).Gemas > 0 Then Call FP_Append(fObj, "Gemas=" & Val(ObjData(Object).Gemas) & vbCrLf)
            If ObjData(Object).Diamantes > 0 Then Call FP_Append(fObj, "Diamantes=" & Val(ObjData(Object).Diamantes) & vbCrLf)
            If ObjData(Object).MaderaElfica > 0 Then Call FP_Append(fObj, "MaderaElfica=" & Val(ObjData(Object).MaderaElfica) & vbCrLf)

        End If

        If ObjData(Object).ObjType = eOBJType.otInstrumentos Then
        
            If ObjData(Object).Snd1 > 0 Then Call FP_Append(fObj, "SND1=" & Val(ObjData(Object).Snd1) & vbCrLf)
            If ObjData(Object).Snd2 > 0 Then Call FP_Append(fObj, "SND2=" & Val(ObjData(Object).Snd2) & vbCrLf)
            If ObjData(Object).Snd3 > 0 Then Call FP_Append(fObj, "SND3=" & Val(ObjData(Object).Snd3) & vbCrLf)
            If ObjData(Object).MinInt > 0 Then Call FP_Append(fObj, "MinInt=" & Val(ObjData(Object).MinInt) & vbCrLf)
            If ObjData(Object).SkHerreriaMagica > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreriaMagica) & vbCrLf)
            If ObjData(Object).Gemas > 0 Then Call FP_Append(fObj, "Gemas=" & Val(ObjData(Object).Gemas) & vbCrLf)
            If ObjData(Object).Diamantes > 0 Then Call FP_Append(fObj, "Diamantes=" & Val(ObjData(Object).Diamantes) & vbCrLf)
            If ObjData(Object).MaderaElfica > 0 Then Call FP_Append(fObj, "MaderaElfica=" & Val(ObjData(Object).MaderaElfica) & vbCrLf)
            
        End If
        
        If ObjData(Object).Caos > 0 Then Call FP_Append(fObj, "LingoteIndex=" & Val(ObjData(Object).LingoteIndex) & vbCrLf)

        If ObjData(Object).ObjType = 31 Or ObjData(Object).ObjType = 23 Then
            If ObjData(Object).Caos > 0 Then Call FP_Append(fObj, "MinSkill=" & Val(ObjData(Object).MinSkill) & vbCrLf)

        End If
        
        If ObjData(Object).sagrado > 0 Then Call FP_Append(fObj, "Sagrado=" & Val(ObjData(Object).sagrado) & vbCrLf)
        If ObjData(Object).MineralIndex > 0 Then Call FP_Append(fObj, "MineralIndex=" & Val(ObjData(Object).MineralIndex) & vbCrLf)
        
        If ObjData(Object).MaxHP > 0 Then Call FP_Append(fObj, "MaxHP=" & Val(ObjData(Object).MaxHP) & vbCrLf)
        If ObjData(Object).MinHP > 0 Then Call FP_Append(fObj, "MinHP=" & Val(ObjData(Object).MinHP) & vbCrLf)
        
        If ObjData(Object).Mujer > 0 Then Call FP_Append(fObj, "Mujer=" & Val(ObjData(Object).Mujer) & vbCrLf)
        If ObjData(Object).Hombre > 0 Then Call FP_Append(fObj, "Hombre=" & Val(ObjData(Object).Hombre) & vbCrLf)
        
        If ObjData(Object).MinHam > 0 Then Call FP_Append(fObj, "MinHam=" & Val(ObjData(Object).MinHam) & vbCrLf)
        If ObjData(Object).MinSed > 0 Then Call FP_Append(fObj, "MinAgu=" & Val(ObjData(Object).MinSed) & vbCrLf)
        
        If ObjData(Object).MinDef > 0 Then Call FP_Append(fObj, "MinDef=" & Val(ObjData(Object).MinDef) & vbCrLf)
        If ObjData(Object).MaxDef > 0 Then Call FP_Append(fObj, "MaxDef=" & Val(ObjData(Object).MaxDef) & vbCrLf)
        
        If ObjData(Object).Vendible > 0 Then Call FP_Append(fObj, "Vendible=" & Val(ObjData(Object).Vendible) & vbCrLf)
        
        If ObjData(Object).Respawn > 0 Then Call FP_Append(fObj, "Respawn=" & Val(ObjData(Object).Respawn) & vbCrLf)
        
        If ObjData(Object).RazaEnana > 0 Then Call FP_Append(fObj, "RazaEnana=" & Val(ObjData(Object).RazaEnana) & vbCrLf)
        If ObjData(Object).RazaElfa > 0 Then Call FP_Append(fObj, "RazaElfa=" & Val(ObjData(Object).RazaElfa) & vbCrLf)
        If ObjData(Object).RazaVampiro > 0 Then Call FP_Append(fObj, "RazaVampiro=" & Val(ObjData(Object).RazaVampiro) & vbCrLf)
        If ObjData(Object).RazaOrco > 0 Then Call FP_Append(fObj, "RazaOrco=" & Val(ObjData(Object).RazaOrco) & vbCrLf)
        
        If ObjData(Object).ClaseAsesino > 0 Then Call FP_Append(fObj, "ClaseAsesino=" & Val(ObjData(Object).ClaseAsesino) & vbCrLf)
        If ObjData(Object).RazaHumana > 0 Then Call FP_Append(fObj, "Razahumana=" & Val(ObjData(Object).RazaHumana) & vbCrLf)
        If ObjData(Object).RazaHobbit > 0 Then Call FP_Append(fObj, "Razahobbit=" & Val(ObjData(Object).RazaHobbit) & vbCrLf)
        If ObjData(Object).Expe > 0 Then Call FP_Append(fObj, "Expe=" & Val(ObjData(Object).Expe) & vbCrLf)
        If ObjData(Object).Valor > 0 Then Call FP_Append(fObj, "Valor=" & Val(ObjData(Object).Valor) & vbCrLf)
        If ObjData(Object).Skill > 0 Then Call FP_Append(fObj, "Skill=" & Val(ObjData(Object).Skill) & vbCrLf)
        
        If ObjData(Object).Caos > 0 Then Call FP_Append(fObj, "objetoespecial=" & Val(ObjData(Object).objetoespecial) & vbCrLf)
        If ObjData(Object).Caos > 0 Then Call FP_Append(fObj, "Crucial=" & Val(ObjData(Object).Crucial) & vbCrLf)
        
        'empieza las puertas
        
        If ObjData(Object).ObjType = eOBJType.otPuertas Then
        
            Call FP_Append(fObj, "Abierta=" & Val(ObjData(Object).Cerrada) & vbCrLf)
            
            Call FP_Append(fObj, "Llave=" & Val(ObjData(Object).Llave) & vbCrLf)
            
            If ObjData(Object).Cerrada = 1 Then
                Call FP_Append(fObj, "Llave=" & Val(ObjData(Object).Llave) & vbCrLf)
                Call FP_Append(fObj, "Clave=" & Val(ObjData(Object).CLAVE) & vbCrLf)

            End If
        
        End If

        If ObjData(Object).ObjType = eOBJType.otPuertas Or ObjData(Object).ObjType = eOBJType.otBotellaVacia Or ObjData(Object).ObjType = eOBJType.otBotellaLlena Then   'OBJTYPE_PUERTAS Or ObjData(Object).ObjType = eOBJType.otBotellaVacia Or ObjData(Object).ObjType = eOBJType.otBotellaLlena Then
            Call FP_Append(fObj, "IndexAbierta=" & Val(ObjData(Object).IndexAbierta) & vbCrLf)
            Call FP_Append(fObj, "IndexCerrada=" & Val(ObjData(Object).IndexCerrada) & vbCrLf)
            Call FP_Append(fObj, "IndexCerradaLlave=" & Val(ObjData(Object).IndexCerradaLlave) & vbCrLf)

        End If
        
        If ObjData(Object).CLAVE > 0 Then Call FP_Append(fObj, "Clave=" & Val(ObjData(Object).CLAVE) & vbCrLf)
        
        If ObjData(Object).texto > "" Then Call FP_Append(fObj, "Texto=" & ObjData(Object).texto & vbCrLf)
        If ObjData(Object).GrhSecundario > 0 Then Call FP_Append(fObj, "VGrande=" & Val(ObjData(Object).GrhSecundario) & vbCrLf)
        
        If ObjData(Object).Agarrable > 0 Then Call FP_Append(fObj, "Agarrable=" & Val(ObjData(Object).Agarrable) & vbCrLf)
        If ObjData(Object).ForoID > "" Then Call FP_Append(fObj, "ID=" & ObjData(Object).ForoID)
        
        If ObjData(Object).Acuchilla > 0 Then Call FP_Append(fObj, "Acuchilla=" & Val(ObjData(Object).Acuchilla) & vbCrLf)
        If ObjData(Object).Guante > 0 Then Call FP_Append(fObj, "Guante=" & Val(ObjData(Object).Guante) & vbCrLf)
        
        If ObjData(Object).NoSubasta > 0 Then Call FP_Append(fObj, "NoSubasta=" & Val(ObjData(Object).NoSubasta) & vbCrLf)
        
        If ObjData(Object).NoRegalo > 0 Then Call FP_Append(fObj, "Regalo=" & Val(ObjData(Object).NoRegalo) & vbCrLf)
        
        '<<<<<CLASES PROHIBIDA REHACERLA>>>>>

        Dim i As Integer, Clase As String
        
        For i = 1 To NUMCLASES

            Clase = ObjData(Object).ClasesProhibidas(i)

            If Len(Clase) > 0 Then

                Call FP_Append(fObj, "CP" & i & "=" & UCase$(Clase) & vbCrLf)

            End If

        Next

        'Debug.Assert ObjData(Object).ClasesProhibidas >= 0
        
        If ObjData(Object).Resistencia > 0 Then Call FP_Append(fObj, "Resistencia=" & Val(ObjData(Object).Resistencia) & vbCrLf)

        'Pociones
        If ObjData(Object).ObjType = 11 Then
            If ObjData(Object).TipoPocion > 0 Then Call FP_Append(fObj, "TipoPocion=" & Val(ObjData(Object).TipoPocion) & vbCrLf)
            If ObjData(Object).MaxModificador > 0 Then Call FP_Append(fObj, "MaxModificador=" & Val(ObjData(Object).MaxModificador) & vbCrLf)
            If ObjData(Object).MinModificador > 0 Then Call FP_Append(fObj, "MinModificador=" & Val(ObjData(Object).MinModificador) & vbCrLf)
            If ObjData(Object).DuracionEfecto > 0 Then Call FP_Append(fObj, "DuracionEfecto=" & Val(ObjData(Object).DuracionEfecto) & vbCrLf)

        End If

        If ObjData(Object).ObjType = 51 Then
           
            If ObjData(Object).TipoRegalo > 0 Then Call FP_Append(fObj, "TipoRegalo=" & Val(ObjData(Object).TipoRegalo) & vbCrLf)
            
            For i = 1 To 10
                
                If ObjData(Object).Objetos(i) > 0 Then Call FP_Append(fObj, "Objetos" & i & "=" & Val(ObjData(Object).Objetos(i)) & vbCrLf)
                If ObjData(Object).Cantidad(i) > 0 Then Call FP_Append(fObj, "Cantidad" & i & "=" & Val(ObjData(Object).Cantidad(i)) & vbCrLf)
               
            Next
           
        End If
        
        If ObjData(Object).SkSastreria > 0 Then Call FP_Append(fObj, "SkSastreria=" & Val(ObjData(Object).SkSastreria) & vbCrLf)
        If ObjData(Object).SkCarpinteria > 0 Then Call FP_Append(fObj, "SkCarpinteria=" & Val(ObjData(Object).SkCarpinteria) & vbCrLf)
        If ObjData(Object).SkHechizeria > 0 Then Call FP_Append(fObj, "SkHechizeria=" & Val(ObjData(Object).SkHechizeria) & vbCrLf)

        If ObjData(Object).SkCarpinteria > 0 Then If ObjData(Object).Madera > 0 Then Call FP_Append(fObj, "Madera=" & Val(ObjData(Object).Madera) & vbCrLf)
        If ObjData(Object).MaderaElfica > 0 Then Call FP_Append(fObj, "MaderaElfica=" & Val(ObjData(Object).MaderaElfica) & vbCrLf)

        If ObjData(Object).SkSastreria > 0 Then
        
            If ObjData(Object).Lana > 0 Then Call FP_Append(fObj, "Lana=" & Val(ObjData(Object).Lana) & vbCrLf)
            If ObjData(Object).Lobos > 0 Then Call FP_Append(fObj, "Lobos=" & Val(ObjData(Object).Lobos) & vbCrLf)
            If ObjData(Object).Osos > 0 Then Call FP_Append(fObj, "Osos=" & Val(ObjData(Object).Osos) & vbCrLf)
            If ObjData(Object).Tigre > 0 Then Call FP_Append(fObj, "Tigre=" & Val(ObjData(Object).Tigre) & vbCrLf)
            If ObjData(Object).OsoPolar > 0 Then Call FP_Append(fObj, "OsoPolar=" & Val(ObjData(Object).OsoPolar) & vbCrLf)
            If ObjData(Object).Vaca > 0 Then Call FP_Append(fObj, "Vaca=" & Val(ObjData(Object).Vaca) & vbCrLf)
            If ObjData(Object).Jabali > 0 Then Call FP_Append(fObj, "Jabali=" & Val(ObjData(Object).Jabali) & vbCrLf)
           
        End If

        If ObjData(Object).SkHechizeria > 0 Then If ObjData(Object).ObjHierba > 0 Then Call FP_Append(fObj, "ObjHierba=" & Val(ObjData(Object).ObjHierba) & vbCrLf)
            
        If ObjData(Object).ObjType = eOBJType.otBarcos Then
            If ObjData(Object).MaxHit > 0 Then Call FP_Append(fObj, "MaxHit=" & Val(ObjData(Object).MaxHit) & vbCrLf)
            If ObjData(Object).MinHit > 0 Then Call FP_Append(fObj, "MinHit=" & Val(ObjData(Object).MinHit) & vbCrLf)
            If ObjData(Object).Velocidad > 0 Then Call FP_Append(fObj, "Velocidad=" & Val(ObjData(Object).Velocidad) & vbCrLf)

        End If

        If ObjData(Object).ObjType = eOBJType.otBarcosArmada Then

            If ObjData(Object).MaxHit > 0 Then Call FP_Append(fObj, "MaxHit=" & Val(ObjData(Object).MaxHit) & vbCrLf)
            If ObjData(Object).MinHit > 0 Then Call FP_Append(fObj, "MinHit=" & Val(ObjData(Object).MinHit) & vbCrLf)
            If ObjData(Object).Caos > 0 Then Call FP_Append(fObj, "Caos=" & Val(ObjData(Object).Caos) & vbCrLf)
            If ObjData(Object).Real > 0 Then Call FP_Append(fObj, "Real=" & Val(ObjData(Object).Real) & vbCrLf)
            If ObjData(Object).templ > 0 Then Call FP_Append(fObj, "Templ=" & Val(ObjData(Object).templ) & vbCrLf)
            If ObjData(Object).Nemes > 0 Then Call FP_Append(fObj, "Nemes=" & Val(ObjData(Object).Nemes) & vbCrLf)
            If ObjData(Object).Facc > 0 Then Call FP_Append(fObj, "Facc=" & Val(ObjData(Object).Facc) & vbCrLf)
            If ObjData(Object).Velocidad > 0 Then Call FP_Append(fObj, "Velocidad=" & Val(ObjData(Object).Velocidad) & vbCrLf)

        End If
       
        If ObjData(Object).ObjType = eOBJType.otPases Then
            If ObjData(Object).Zona > 0 Then Call FP_Append(fObj, "Zona=" & Val(ObjData(Object).Zona) & vbCrLf)

        End If

        If ObjData(Object).ObjType = eOBJType.otFlechas Then
            If ObjData(Object).MaxHit > 0 Then Call FP_Append(fObj, "MaxHit=" & Val(ObjData(Object).MaxHit) & vbCrLf)
            If ObjData(Object).MinHit > 0 Then Call FP_Append(fObj, "MinHit=" & Val(ObjData(Object).MinHit) & vbCrLf)
            If ObjData(Object).LingH > 0 Then Call FP_Append(fObj, "LingH=" & Val(ObjData(Object).LingH) & vbCrLf)
            If ObjData(Object).LingP > 0 Then Call FP_Append(fObj, "LingP=" & Val(ObjData(Object).LingP) & vbCrLf)
            If ObjData(Object).LingO > 0 Then Call FP_Append(fObj, "LingO=" & Val(ObjData(Object).LingO) & vbCrLf)
            If ObjData(Object).LingM > 0 Then Call FP_Append(fObj, "LingM=" & Val(ObjData(Object).LingM) & vbCrLf)
            If ObjData(Object).Paraliza > 0 Then Call FP_Append(fObj, "Paraliza=" & Val(ObjData(Object).Paraliza) & vbCrLf)
            If ObjData(Object).TipoMunicion > 0 Then Call FP_Append(fObj, "TipoMunicion=" & Val(ObjData(Object).TipoMunicion) & vbCrLf)
            If ObjData(Object).SkHerreria > 0 Then Call FP_Append(fObj, "SkHerreria=" & Val(ObjData(Object).SkHerreria) & vbCrLf)

        End If

        'Bebidas
        If ObjData(Object).MinSta > 0 Then Call FP_Append(fObj, "MinST=" & Val(ObjData(Object).MinSta) & vbCrLf)
        
        'Sistema de subir nivel a clanes.
        If ObjData(Object).Clan > 0 Then Call FP_Append(fObj, "Clan=" & Val(ObjData(Object).Clan) & vbCrLf)
        
        If ObjData(Object).Name <> "" Then Call FP_Append(fObj, vbCrLf)
        'FrmMain.Bgrabar.value = Object
    
        FrmMain.ucBgrabar.value = Object
        DoEvents
    Next Object
    
    FrmMain.Lbltexto.Visible = True
    FrmMain.Lbltexto.Caption = "Objetos guardados: " & FrmMain.ListView.ListItems.Count & " de: " & Object & "."
    FrmMain.Timer.Enabled = True
    FrmMain.Timer2.Enabled = True
    FrmMain.Timer3.Enabled = True
       
    Exit Sub

ErrHandler:
    MsgBox "error grabando objetos en objeto " & Object

End Sub

Public Sub CabezeraBlockObj()
     
    On Error GoTo ErrHandler

    Dim fObj As String
    
    fObj = App.path & "\" & "Obj.dat"
    
    Call FP_Append(fObj, "'            TIPOS DE OBJETOS" & vbCrLf)
    Call FP_Append(fObj, "' 1: Comida" & vbCrLf)
    Call FP_Append(fObj, "' 2: Armas" & vbCrLf)
    Call FP_Append(fObj, "' 3: Armaduras" & vbCrLf)
    Call FP_Append(fObj, "' 4: Arboles" & vbCrLf)
    Call FP_Append(fObj, "' 5: Dinero" & vbCrLf)
    Call FP_Append(fObj, "' 6: Puertas" & vbCrLf)
    Call FP_Append(fObj, "' 7: Objetos contenedores (por ejemplo bolsas y cofres)" & vbCrLf)
    Call FP_Append(fObj, "' 8: Carteles" & vbCrLf)
    Call FP_Append(fObj, "' 9: Llaves" & vbCrLf)
    Call FP_Append(fObj, "' 10: Foros" & vbCrLf)
    Call FP_Append(fObj, "' 11: Pociones" & vbCrLf)
    Call FP_Append(fObj, "' 12: Libros" & vbCrLf)
    Call FP_Append(fObj, "' 13: Bebida" & vbCrLf)
    Call FP_Append(fObj, "' 14: Leña" & vbCrLf)
    Call FP_Append(fObj, "' 15: Fogata" & vbCrLf)
    Call FP_Append(fObj, "' 16: Guantes" & vbCrLf)
    Call FP_Append(fObj, "' 17: Anillo" & vbCrLf)
    Call FP_Append(fObj, "' 18: Herramientas" & vbCrLf)
    Call FP_Append(fObj, "' 19: telep" & vbCrLf)
    Call FP_Append(fObj, "' 20: Muebles" & vbCrLf)
    Call FP_Append(fObj, "' 21: joyas" & vbCrLf)
    Call FP_Append(fObj, "' 22: yacimiento" & vbCrLf)
    Call FP_Append(fObj, "' 23: metales" & vbCrLf)
    Call FP_Append(fObj, "' 24: pergaminos" & vbCrLf)
    Call FP_Append(fObj, "' 25: aura" & vbCrLf)
    Call FP_Append(fObj, "' 26: Instrumentos Musicales" & vbCrLf)
    Call FP_Append(fObj, "' 27: Yunque" & vbCrLf)
    Call FP_Append(fObj, "' 28: Fragua" & vbCrLf)
    Call FP_Append(fObj, "' 29: Gemas" & vbCrLf)
    Call FP_Append(fObj, "' 30: Flores" & vbCrLf)
    Call FP_Append(fObj, "' 31: barcos Galeon, Fragata" & vbCrLf)
    Call FP_Append(fObj, "' 32: flechas" & vbCrLf)
    Call FP_Append(fObj, "' 33: botellas vacias" & vbCrLf)
    Call FP_Append(fObj, "' 34: botellas llenas" & vbCrLf)
    Call FP_Append(fObj, "' 35: manchas" & vbCrLf)
    Call FP_Append(fObj, "' 36: Pocis que Resucitan" & vbCrLf)
    Call FP_Append(fObj, "' 37: HEAD" & vbCrLf)
    Call FP_Append(fObj, "' 38: Hierbas" & vbCrLf)
    Call FP_Append(fObj, "' 39: Lana" & vbCrLf)
    Call FP_Append(fObj, "' 40: Oveja" & vbCrLf)
    Call FP_Append(fObj, "' 41: Pases" & vbCrLf)
    Call FP_Append(fObj, "' 42: Arbol Elfico" & vbCrLf)
    Call FP_Append(fObj, "' 43: BarcoArmada" & vbCrLf)
    Call FP_Append(fObj, "' 45: Amuleto Defensa y Anillos" & vbCrLf)
    Call FP_Append(fObj, "' 47: Amuleto Teletransportador" & vbCrLf)
    Call FP_Append(fObj, "' 48: Huevo de la gloria eterna" & vbCrLf)
    Call FP_Append(fObj, "' 49: Talisman de la gloria eterna" & vbCrLf)
    Call FP_Append(fObj, "' 50: Vale por 500k de experiencia" & vbCrLf)
    Call FP_Append(fObj, "' 51: Pack Premium" & vbCrLf)
    Call FP_Append(fObj, "'<--------------------Sub Tipos--------------------------->" & vbCrLf)
    Call FP_Append(fObj, "' 0: Armaduras" & vbCrLf)
    Call FP_Append(fObj, "' 1: Cascos" & vbCrLf)
    Call FP_Append(fObj, "' 2: Escudos" & vbCrLf)
    Call FP_Append(fObj, "'<--------------------Sub Tipos--------------------------->" & vbCrLf)
    Call FP_Append(fObj, "'<--------------------Tipos De Pieles--------------------------->" & vbCrLf)
    Call FP_Append(fObj, "'Lobo=" & vbCrLf)
    Call FP_Append(fObj, "'Osos=" & vbCrLf)
    Call FP_Append(fObj, "'Lana=" & vbCrLf)
    Call FP_Append(fObj, "'Tigre=" & vbCrLf)
    Call FP_Append(fObj, "'Jabali=" & vbCrLf)
    Call FP_Append(fObj, "'LoboPolar=" & vbCrLf)
    Call FP_Append(fObj, "'OsoPolar=" & vbCrLf)
    Call FP_Append(fObj, "'Vaca=" & vbCrLf)
    Call FP_Append(fObj, "'<--------------------Tipos De Pieles--------------------------->" & vbCrLf)
    Call FP_Append(fObj, "'<--------------------Tipo de posicones------------------->" & vbCrLf)
    Call FP_Append(fObj, "'1 Modifica la Agilidad" & vbCrLf)
    Call FP_Append(fObj, "'2 Modifica la Fuerza" & vbCrLf)
    Call FP_Append(fObj, "'3 Repone HP" & vbCrLf)
    Call FP_Append(fObj, "'4 Repone Mana" & vbCrLf)
    Call FP_Append(fObj, "'5 Cura Envenenamiento" & vbCrLf)
    Call FP_Append(fObj, "'6 Desparalizante" & vbCrLf)
    Call FP_Append(fObj, "'7 Invisible" & vbCrLf)
    Call FP_Append(fObj, "'8 Telepatia" & vbCrLf)
    Call FP_Append(fObj, "'9 Teletransporte" & vbCrLf)
    Call FP_Append(fObj, "'10 Energia" & vbCrLf)
    Call FP_Append(fObj, "'11 Anti Cegera" & vbCrLf)
    Call FP_Append(fObj, "'12 Anti estupide" & vbCrLf)
    Call FP_Append(fObj, "'<--------------------Tipo de posicones------------------->" & vbCrLf)
    Call FP_Append(fObj, "'<--------------------RAZAS---------------------->" & vbCrLf)
    Call FP_Append(fObj, "'Humano" & vbCrLf)
    Call FP_Append(fObj, "'Elfo" & vbCrLf)
    Call FP_Append(fObj, "'Elfo Oscuro" & vbCrLf)
    Call FP_Append(fObj, "'Gnomo" & vbCrLf)
    Call FP_Append(fObj, "'Enano" & vbCrLf)
    Call FP_Append(fObj, "'Orco" & vbCrLf)
    Call FP_Append(fObj, "'Hobbit" & vbCrLf)
    Call FP_Append(fObj, "'Ciclope" & vbCrLf)
    Call FP_Append(fObj, "'Vampiros" & vbCrLf)
    Call FP_Append(fObj, "'Licantropo" & vbCrLf)
    Call FP_Append(fObj, "'<--------------------RAZAS---------------------->" & vbCrLf)
    Call FP_Append(fObj, "'<--------------------Clases--------------------->" & vbCrLf)
    Call FP_Append(fObj, "'CP1=MAGO" & vbCrLf)
    Call FP_Append(fObj, "'CP2=CLERIGO" & vbCrLf)
    Call FP_Append(fObj, "'CP3=GUERRERO" & vbCrLf)
    Call FP_Append(fObj, "'CP4=ASESINO" & vbCrLf)
    Call FP_Append(fObj, "'CP5=LADRON" & vbCrLf)
    Call FP_Append(fObj, "'CP6=BARDO" & vbCrLf)
    Call FP_Append(fObj, "'CP7=DRUIDA" & vbCrLf)
    Call FP_Append(fObj, "'CP8=PALADIN" & vbCrLf)
    Call FP_Append(fObj, "'CP9=TRABAJADOR" & vbCrLf)
    Call FP_Append(fObj, "'CP10=BRUJO" & vbCrLf)
    Call FP_Append(fObj, "'CP11=ARQUERO" & vbCrLf)
    Call FP_Append(fObj, "'CP12=GLADIADOR MAGICO" & vbCrLf)
    Call FP_Append(fObj, "'CP13=DIOS" & vbCrLf)
    Call FP_Append(fObj, "'CP14=BANDIDO" & vbCrLf)
    Call FP_Append(fObj, "'CP15=DOMADOR" & vbCrLf)
    Call FP_Append(fObj, "<---------------------Clases----------------------->" & vbCrLf)
    Call FP_Append(fObj, "'<----------------------Armas con efectos----------------->" & vbCrLf)
    Call FP_Append(fObj, "'Paraliza= probabvilidad 1 a 100 100 paraliza siempre" & vbCrLf)
    Call FP_Append(fObj, "'<----------------------Armas con efectos----------------->" & vbCrLf)
    Call FP_Append(fObj, "'<----------------------Armaduras con efectos----------------->" & vbCrLf)
    Call FP_Append(fObj, "'Resistencia= probabvilidad 1 a 100 100 Resiste siempre" & vbCrLf)
    Call FP_Append(fObj, "'<----------------------Armaduras con efectos----------------->" & vbCrLf & vbCrLf)
    
    Exit Sub

ErrHandler:
    
End Sub

