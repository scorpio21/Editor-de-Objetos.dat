Attribute VB_Name = "Dat"
Option Explicit

Sub LoadGraficosIni()
        
    Dim Grh(1 To NumGrh) As Integer

    'abrimos el fichero
    Call INI_Open(InitPath & "\", "Graficos.ini")
        
    Call INI_GetString("Graficos.ini", "Graphics", "Grh1")
    
    For i = 1 To NumGrh
        ' recorremos todos los Graficos del fichero.
        Grh(i) = Val(ReadField(2, INI_GetString("Graficos.ini", "Graphics", "Grh" & i), 45))
     
        FrmMain.Picture.Picture = GrafPath & LoadPicture(Grh(i)) & ".bmp"

    Next i
    
End Sub

Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String

    '*****************************************************************
    'Gets a field from a delimited string
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
    '*****************************************************************
    Dim i          As Long

    Dim lastPos    As Long

    Dim CurrentPos As Long

    Dim delimiter  As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)

    End If

End Function

Function FileExist(file As String, FileType As VbFileAttribute) As Boolean

    If Dir(file, FileType) = "" Then
        FileExist = False
    Else
        FileExist = True

    End If

End Function

Public Sub CargarListviewNpc(Optional ByVal NPCtype As Integer = 0, _
                             Optional ByVal Nombre As String = vbNullString)
        
    If FrmMain.Listnpcs.ListItems.Count > 0 Then
        FrmMain.Listnpcs.ListItems.Clear

    End If
    
    If NPCtype > 0 Then
        FormaBusca = 1
    ElseIf Nombre > vbNullString Then
        FormaBusca = 2
   
    End If

    Dim ln As String

    Dim X  As Integer

    FrmMain.ucProgress.Visible = True
 
    For NpcNumber = 1 To NumNpcs
    
        If (NpcDats(NpcNumber).NPCtype = NPCtype And FormaBusca = 1) Or (InStr(UCase(NpcDats(NpcNumber).Name), UCase(Nombre)) > 0 And FormaBusca = 2) Then
            Set subelemento = FrmMain.Listnpcs.ListItems.Add(, , NpcNumber)
            
            subelemento.SubItems(1) = NpcDats(NpcNumber).Name
            subelemento.SubItems(2) = NpcDats(NpcNumber).Desc
            
            subelemento.SubItems(3) = NpcDats(NpcNumber).char.Head
            subelemento.SubItems(4) = NpcDats(NpcNumber).char.Body
            subelemento.SubItems(5) = NpcDats(NpcNumber).char.CascoAnim
            
            subelemento.SubItems(6) = NpcDats(NpcNumber).char.WeaponAnim
            subelemento.SubItems(7) = NpcDats(NpcNumber).char.ShieldAnim
            
            subelemento.SubItems(8) = NpcDats(NpcNumber).char.heading
            
            subelemento.SubItems(9) = NpcDats(NpcNumber).flags.OldMovement
            subelemento.SubItems(10) = NpcDats(NpcNumber).Comercia
            subelemento.SubItems(11) = NpcDats(NpcNumber).Stats.Alineacion
            subelemento.SubItems(12) = NpcDats(NpcNumber).flags.Respawn
            subelemento.SubItems(13) = NpcDats(NpcNumber).Invent.NroItems
            
            'Dim loopc As Integer, obj As Integer, Prob As Integer
            'Obj cantidad +25
            Call ObjetosRegalo
            'Pro cantidad +16
            subelemento.SubItems(39) = NpcDats(NpcNumber).flags.AfectaParalisis
            subelemento.SubItems(56) = NpcDats(NpcNumber).NPCtype
            subelemento.SubItems(57) = NpcDats(NpcNumber).Attackable
            subelemento.SubItems(58) = NpcDats(NpcNumber).Hostile
            subelemento.SubItems(59) = NpcDats(NpcNumber).GiveEXP
            subelemento.SubItems(60) = NpcDats(NpcNumber).GiveGLD
            subelemento.SubItems(61) = NpcDats(NpcNumber).Stats.MinHP
            subelemento.SubItems(62) = NpcDats(NpcNumber).Stats.MaxHP
            subelemento.SubItems(63) = NpcDats(NpcNumber).Stats.MaxHit
            subelemento.SubItems(64) = NpcDats(NpcNumber).Stats.MinHit
            subelemento.SubItems(65) = NpcDats(NpcNumber).Stats.DEF
            subelemento.SubItems(66) = NpcDats(NpcNumber).PoderAtaque
            subelemento.SubItems(67) = NpcDats(NpcNumber).PoderEvasion
            subelemento.SubItems(68) = NpcDats(NpcNumber).flags.Domable
            subelemento.SubItems(69) = NpcDats(NpcNumber).TipoItems
            subelemento.SubItems(70) = NpcDats(NpcNumber).Inflacion
            subelemento.SubItems(71) = NpcDats(NpcNumber).flags.BackUp
            subelemento.SubItems(72) = NpcDats(NpcNumber).Hostile
            
            X = X + 1

        End If
 
        FrmMain.ucProgress.value = Object
        
    Next NpcNumber

    FrmMain.Caption = "Editor de Objetos By Helios 2022. " & "               Nº De Npcs: " & X
    FrmMain.Lbltexto.Visible = True
    FrmMain.Lbltexto.Caption = "Encontrado " & X & " Npcs."
    FrmMain.Timer.Enabled = True
    FrmMain.Timer2.Enabled = True
    FrmMain.Timer3.Enabled = True
    Call ForeColorColumn(vbCyan, 0, FrmMain.Listnpcs)
    FrmMain.txtnombre.Text = vbNullString
    FrmMain.txtObjType.Text = vbNullString
    FrmMain.LstLista.ListIndex = -1
    Call redimenciona
    
End Sub

Public Sub CargarListView(Optional ByVal ObjType As Integer = 0, _
                          Optional ByVal Nombre As String = vbNullString)
    
    If ObjType > 0 Then
        FormaBusca = 1
    ElseIf Nombre > vbNullString Then
        FormaBusca = 2
   
    End If

    'compruebo la barra DoEvents progreso
  
    FrmMain.ucProgress.Visible = True
   
    Dim X As Integer
    
    For Object = 1 To NumObjDatas
        FrmMain.ucProgress.max = NumObjDatas

        If (ObjData(Object).ObjType = ObjType And FormaBusca = 1) Or (InStr(UCase(ObjData(Object).Name), UCase(Nombre)) > 0 And FormaBusca = 2) Then
            ' If (ObjData(Object).ObjType = ObjType And FormaBusca = 1) Or (ObjData(Object).Name = Nombre And FormaBusca = 2) Then
                
            Set subelemento = FrmMain.ListView.ListItems.Add(, , Object)
            subelemento.SubItems(1) = ObjData(Object).Name
            subelemento.SubItems(2) = ObjData(Object).ObjType
            subelemento.SubItems(3) = ObjData(Object).SubTipo
            subelemento.SubItems(4) = ObjData(Object).MaxHit
            subelemento.SubItems(5) = ObjData(Object).MinHit
            subelemento.SubItems(6) = ObjData(Object).MinDef
            subelemento.SubItems(7) = ObjData(Object).MaxDef
            subelemento.SubItems(8) = ObjData(Object).MinHam
            subelemento.SubItems(9) = ObjData(Object).Valor
            subelemento.SubItems(10) = ObjData(Object).Resistencia
            subelemento.SubItems(11) = ObjData(Object).Vendible
            subelemento.SubItems(12) = ObjData(Object).SkSastreria
            subelemento.SubItems(13) = ObjData(Object).SkHechizeria
            subelemento.SubItems(14) = ObjData(Object).ObjHierba
            subelemento.SubItems(15) = ObjData(Object).MinSkill
            subelemento.SubItems(16) = ObjData(Object).Velocidad
            subelemento.SubItems(17) = ObjData(Object).Facc
            subelemento.SubItems(18) = ObjData(Object).objetoespecial
            subelemento.SubItems(19) = ObjData(Object).MinSta
            subelemento.SubItems(20) = ObjData(Object).Caos
            subelemento.SubItems(21) = ObjData(Object).Real
            subelemento.SubItems(22) = ObjData(Object).templ
            subelemento.SubItems(23) = ObjData(Object).Nemes
            subelemento.SubItems(24) = ObjData(Object).Cae
            subelemento.SubItems(25) = ObjData(Object).TiRaRObj
            subelemento.SubItems(26) = ObjData(Object).Robable
            subelemento.SubItems(27) = ObjData(Object).bonifica
            subelemento.SubItems(28) = ObjData(Object).tipobonifica
            subelemento.SubItems(29) = ObjData(Object).Donacion
            subelemento.SubItems(30) = ObjData(Object).Newbie
            subelemento.SubItems(31) = ObjData(Object).Destruir
            subelemento.SubItems(32) = ObjData(Object).ShieldAnim
            subelemento.SubItems(33) = ObjData(Object).LingH
            subelemento.SubItems(34) = ObjData(Object).LingP
            subelemento.SubItems(35) = ObjData(Object).LingO
            subelemento.SubItems(36) = ObjData(Object).LingM
            subelemento.SubItems(37) = ObjData(Object).Madera
            subelemento.SubItems(38) = ObjData(Object).MaderaElfica
            subelemento.SubItems(39) = ObjData(Object).SkHerreria
            subelemento.SubItems(40) = ObjData(Object).Drena
            subelemento.SubItems(41) = ObjData(Object).SkHerreriaMagica
            subelemento.SubItems(42) = ObjData(Object).Gemas
            subelemento.SubItems(43) = ObjData(Object).Diamantes
            subelemento.SubItems(44) = ObjData(Object).CascoAnim
            subelemento.SubItems(45) = ObjData(Object).DefensaMagicaMax
            subelemento.SubItems(46) = ObjData(Object).DefensaMagicaMin
            subelemento.SubItems(47) = ObjData(Object).Botas
            subelemento.SubItems(48) = ObjData(Object).Alas
            subelemento.SubItems(49) = ObjData(Object).Ropaje
            subelemento.SubItems(50) = ObjData(Object).HechizoIndex
            subelemento.SubItems(51) = ObjData(Object).Clase
            subelemento.SubItems(52) = ObjData(Object).WeaponAnim
            subelemento.SubItems(53) = ObjData(Object).Apuñala
            subelemento.SubItems(54) = ObjData(Object).Paraliza
            subelemento.SubItems(55) = ObjData(Object).Ceguera
            subelemento.SubItems(56) = ObjData(Object).Estupidez
            subelemento.SubItems(57) = ObjData(Object).Vida
            subelemento.SubItems(58) = ObjData(Object).Mana
            subelemento.SubItems(59) = ObjData(Object).Envenena
            subelemento.SubItems(60) = ObjData(Object).LvlMin
            subelemento.SubItems(61) = ObjData(Object).LvlMax
            subelemento.SubItems(62) = ObjData(Object).Heroe
            subelemento.SubItems(63) = ObjData(Object).proyectil
            subelemento.SubItems(64) = ObjData(Object).Municion
            subelemento.SubItems(65) = ObjData(Object).TipoProyectil
            subelemento.SubItems(66) = ObjData(Object).DosManos
            subelemento.SubItems(67) = ObjData(Object).Nivel
            subelemento.SubItems(68) = ObjData(Object).GameMaster
            subelemento.SubItems(69) = ObjData(Object).UsoNpc
            subelemento.SubItems(70) = ObjData(Object).Snd1
            subelemento.SubItems(71) = ObjData(Object).Snd2
            subelemento.SubItems(72) = ObjData(Object).Snd3
            subelemento.SubItems(73) = ObjData(Object).MinInt
            subelemento.SubItems(74) = ObjData(Object).LingoteIndex
            subelemento.SubItems(75) = ObjData(Object).sagrado
            subelemento.SubItems(76) = ObjData(Object).MineralIndex
            subelemento.SubItems(77) = ObjData(Object).MaxHP
            subelemento.SubItems(78) = ObjData(Object).MinHP
            subelemento.SubItems(79) = ObjData(Object).Mujer
            subelemento.SubItems(80) = ObjData(Object).Hombre
            subelemento.SubItems(81) = ObjData(Object).MinSed
            subelemento.SubItems(82) = ObjData(Object).Respawn
            subelemento.SubItems(83) = ObjData(Object).RazaEnana
            subelemento.SubItems(84) = ObjData(Object).RazaElfa
            subelemento.SubItems(85) = ObjData(Object).RazaVampiro
            subelemento.SubItems(86) = ObjData(Object).RazaOrco
            subelemento.SubItems(87) = ObjData(Object).ClaseAsesino
            subelemento.SubItems(88) = ObjData(Object).RazaHumana
            subelemento.SubItems(89) = ObjData(Object).RazaHobbit
            subelemento.SubItems(90) = ObjData(Object).Expe
            subelemento.SubItems(91) = ObjData(Object).Skill
            subelemento.SubItems(92) = ObjData(Object).Crucial
            subelemento.SubItems(93) = ObjData(Object).Cerrada
            subelemento.SubItems(94) = ObjData(Object).Llave
            subelemento.SubItems(95) = ObjData(Object).CLAVE
            subelemento.SubItems(96) = ObjData(Object).IndexAbierta
            subelemento.SubItems(97) = ObjData(Object).IndexCerrada
            subelemento.SubItems(98) = ObjData(Object).IndexCerradaLlave
            subelemento.SubItems(99) = ObjData(Object).texto
            subelemento.SubItems(100) = ObjData(Object).GrhSecundario
            subelemento.SubItems(101) = ObjData(Object).Agarrable
            subelemento.SubItems(102) = ObjData(Object).ForoID
            subelemento.SubItems(103) = ObjData(Object).Acuchilla
            subelemento.SubItems(104) = ObjData(Object).Guante
            subelemento.SubItems(105) = ObjData(Object).NoSubasta
            subelemento.SubItems(106) = ObjData(Object).NoRegalo
            subelemento.SubItems(107) = ObjData(Object).TipoPocion
            subelemento.SubItems(108) = ObjData(Object).MaxModificador
            subelemento.SubItems(109) = ObjData(Object).MinModificador
            subelemento.SubItems(110) = ObjData(Object).DuracionEfecto
            subelemento.SubItems(111) = ObjData(Object).TipoRegalo
            
            subelemento.SubItems(114) = ObjData(Object).SkCarpinteria
            subelemento.SubItems(115) = ObjData(Object).Lana
            subelemento.SubItems(116) = ObjData(Object).Lobos
            subelemento.SubItems(117) = ObjData(Object).Osos
            subelemento.SubItems(118) = ObjData(Object).Tigre
            subelemento.SubItems(119) = ObjData(Object).OsoPolar
            subelemento.SubItems(120) = ObjData(Object).Vaca
            subelemento.SubItems(121) = ObjData(Object).Jabali
            subelemento.SubItems(122) = ObjData(Object).Zona
            subelemento.SubItems(123) = ObjData(Object).TipoMunicion
            
            Call ClasesPro '137 hasta 147
            Call Regalo '137 hasta 156
            subelemento.SubItems(157) = ObjData(Object).TiempoPocion
            subelemento.SubItems(158) = ObjData(Object).TipoAura
            subelemento.SubItems(159) = ObjData(Object).TipoLibro
            subelemento.SubItems(160) = ObjData(Object).GrhIndex
            
            X = X + 1

        End If

        FrmMain.ucProgress.value = Object
        
    Next Object

    'FrmMain.lblLabel1.Caption = FrmMain.ListView.ListItems.Count
    'FrmMain.lblLabel1.Caption = "Nº De Objetos: " & x
    FrmMain.Caption = "Editor de Objetos By Helios 2022. " & "               Nº De Objetos: " & X
    FrmMain.Lbltexto.Visible = True
    FrmMain.Lbltexto.Caption = "Encontrado " & X & " Objetos."
    FrmMain.Timer.Enabled = True
    FrmMain.Timer2.Enabled = True
    FrmMain.Timer3.Enabled = True
    Call ForeColorColumn(vbCyan, 0, FrmMain.ListView)
    FrmMain.txtnombre.Text = vbNullString
    FrmMain.txtObjType.Text = vbNullString
    FrmMain.LstLista.ListIndex = -1
    Call redimencionar
    
End Sub

Public Sub ObjetosRegalo()

    Dim R As Integer

    Dim X As Integer

    Dim i As Integer

    R = 13

    If NpcDats(NpcNumber).Invent.NroItems > 0 Then

        For i = 1 To NpcDats(NpcNumber).Invent.NroItems
            R = R + 1
            X = X + 1
                    
            subelemento.SubItems(R) = NpcDats(NpcNumber).Invent.Object(X).ObjIndex & "-" & NpcDats(NpcNumber).Invent.Object(X).Amount
                    
        Next i

        X = 0
        R = 39

        For i = 1 To NpcDats(NpcNumber).Invent.NroItems
            R = R + 1
           
            X = X + 1
            subelemento.SubItems(R) = NpcDats(NpcNumber).Invent.Object(X).ProbDrop
        
        Next
        X = 0
        
    End If

End Sub

Public Sub Regalo()

    Dim R As Integer

    Dim X As Integer

    Dim i As Integer

    R = 136

    For i = 1 To 10
        R = R + 1
           
        X = X + 1
        subelemento.SubItems(R) = INI_GetString("Obj.dat", "OBJ" & Object, "Objetos" & X)
        
    Next
    X = 0
    R = 146

    For i = 1 To 10
        R = R + 1
           
        X = X + 1
                            
        subelemento.SubItems(R) = INI_GetString("Obj.dat", "OBJ" & Object, "Cantidad" & X)
    Next
    X = 0
   
End Sub

Public Sub ClasesPro()

    Dim R As Integer

    Dim X As Integer

    Dim i As Integer

    R = 123

    For i = 1 To NUMCLASES
        R = R + 1
           
        X = X + 1
        subelemento.SubItems(R) = INI_GetString("Obj.dat", "OBJ" & Object, "cp" & i)
        
    Next
  
End Sub

Sub LoadOBJData()

    On Error GoTo ErrHandler

    Dim cTimer  As Currency

    Dim cTimer2 As Currency
    
    'FrmMain.Barra.Visible = True
    
    FrmMain.ucProgress.Visible = True

    Dim Object As Integer

    Dim Clase  As String
    
    'obtiene el numero de obj
    Call INI_Open(DatPath, "obj.dat", False)

    If cargaind = 1 Then
        Fichero = "objAom.dat"
    Else
        Fichero = "Obj.Dat"

    End If

    NumObjDatas = Val(INI_GetString("Obj.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    'FrmMain.Barra.max = NumObjDatas
    FrmMain.ucProgress.max = NumObjDatas

    'Llena la lista
    For Object = 1 To NumObjDatas
        
        ObjData(Object).Name = INI_GetString("Obj.Dat", "OBJ" & Object, "Name")

        ObjData(Object).Caos = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Caos"))
        ObjData(Object).Real = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Real"))
        ObjData(Object).templ = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Templ"))
        ObjData(Object).Nemes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Nemes"))
        ObjData(Object).Facc = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Facc"))
        ObjData(Object).Cae = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Cae"))
        ObjData(Object).TiRaRObj = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Tirar"))
       
        ObjData(Object).Robable = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "NoRobable"))
        ObjData(Object).bonifica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "bonifica"))
        ObjData(Object).tipobonifica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "tipobonifica"))
        ObjData(Object).Donacion = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Donacion"))
    
        ObjData(Object).GrhIndex = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "GrhIndex"))

        ObjData(Object).ObjType = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "ObjType"))
        ObjData(Object).SubTipo = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Subtipo"))

        ObjData(Object).Newbie = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Newbie"))

        ObjData(Object).Destruir = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Destruir"))

        If ObjData(Object).SubTipo = sObjtype.sotEscudo Then
            ObjData(Object).ShieldAnim = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingO"))
            ObjData(Object).LingM = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingM"))
            ObjData(Object).Madera = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Madera"))
            ObjData(Object).MaderaElfica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaderaElfica"))
            ObjData(Object).SkHerreria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Drena = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Drena"))
            
            ObjData(Object).SkHerreriaMagica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Gemas = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Diamantes"))
            
        End If

        If ObjData(Object).SubTipo = sObjtype.sotCasco Then    ' OBJTYPE_CASCO Then
            ObjData(Object).CascoAnim = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingO"))
            ObjData(Object).LingM = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingM"))
            ObjData(Object).Madera = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Madera"))
            ObjData(Object).MaderaElfica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaderaElfica"))
            ObjData(Object).SkHerreria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).DefensaMagicaMax = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "DefensaMagicaMax"))
            ObjData(Object).DefensaMagicaMin = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "DefensaMagicaMin"))
            
            ObjData(Object).SkHerreriaMagica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Gemas = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Diamantes"))
            
        End If
        
        If ObjData(Object).SubTipo = sObjtype.sotBotas Then
            ObjData(Object).Botas = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingO"))
            ObjData(Object).LingM = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingM"))
            ObjData(Object).SkHerreria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))

        End If

        If ObjData(Object).SubTipo = sObjtype.sotAlas Then
            ObjData(Object).Alas = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "AlasAnim"))
            ObjData(Object).LingH = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingO"))
            ObjData(Object).LingM = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingM"))
            ObjData(Object).Real = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Caos"))
            ObjData(Object).templ = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Templ"))
            ObjData(Object).Nemes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Nemes"))
            ObjData(Object).SkHerreria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Gemas = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Diamantes"))
            ObjData(Object).Resistencia = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Resistencia"))

        End If
        
        ObjData(Object).Ropaje = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "NumRopaje"))
        ObjData(Object).HechizoIndex = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "HechizoIndex"))
        ObjData(Object).Clase = INI_GetString("Obj.Dat", "OBJ" & Object, "Clase")
    
        If ObjData(Object).ObjType = eOBJType.otWeapon Then
            ObjData(Object).WeaponAnim = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Apuñala"))
            ObjData(Object).Paraliza = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Paraliza"))
            ObjData(Object).Ceguera = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Ceguera"))
            ObjData(Object).Estupidez = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Estupidez"))
            ObjData(Object).Vida = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Vida"))
            ObjData(Object).Mana = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Mana"))
            ObjData(Object).Envenena = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHit = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHit = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinHIT"))
            ObjData(Object).LvlMin = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LvlMin"))
            ObjData(Object).LvlMax = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LvlMax"))
            ObjData(Object).Heroe = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Heroe"))
            ObjData(Object).LingH = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingO"))
            ObjData(Object).LingM = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingM"))
            ObjData(Object).Real = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Caos"))
            ObjData(Object).templ = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Templ"))
            ObjData(Object).Nemes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Nemes"))
            ObjData(Object).SkHerreria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).proyectil = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Municiones"))
            ObjData(Object).TipoProyectil = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "TipoProyectil"))
            ObjData(Object).DosManos = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "DosManos"))
            ObjData(Object).SkHerreriaMagica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Gemas = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Diamantes"))
            ObjData(Object).Nivel = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Nivel"))
            ObjData(Object).GameMaster = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "GM"))
            ObjData(Object).UsoNpc = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "NPC"))
            ObjData(Object).Madera = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Madera"))
            ObjData(Object).MaderaElfica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaderaElfica"))

        End If
        
        If ObjData(Object).ObjType = eOBJType.otArmadura Then
            ObjData(Object).LingH = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingO"))
            ObjData(Object).LingM = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingM"))
            ObjData(Object).SkHerreria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Caos"))
            ObjData(Object).templ = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Templ"))
            ObjData(Object).Nemes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Nemes"))
            ObjData(Object).GameMaster = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "GM"))
            
            ObjData(Object).MinDef = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MINDEF"))
            ObjData(Object).MaxDef = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MAXDEF"))
            ObjData(Object).Resistencia = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Resistencia"))
            
            ObjData(Object).SkHerreriaMagica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Gemas = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Diamantes"))
            ObjData(Object).Madera = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Madera"))
            ObjData(Object).MaderaElfica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaderaElfica"))
            
        End If

        If ObjData(Object).ObjType = eOBJType.otHerramientas Then    ' OBJTYPE_HERRAMIENTAS Then
            ObjData(Object).LingH = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingO"))
            ObjData(Object).LingM = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingM"))
            ObjData(Object).SkHerreria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            
            ObjData(Object).SkHerreriaMagica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Gemas = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Diamantes"))
            ObjData(Object).MaderaElfica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaderaElfica"))

        End If

        If ObjData(Object).ObjType = eOBJType.otInstrumentos Then
            ObjData(Object).Snd1 = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SND3"))
            ObjData(Object).MinInt = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinInt"))
            '[Helios:6]
            ObjData(Object).SkHerreriaMagica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Gemas = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Diamantes"))
            ObjData(Object).MaderaElfica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaderaElfica"))

            '[\END]
        End If

        ObjData(Object).LingoteIndex = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingoteIndex"))

        If ObjData(Object).ObjType = 31 Or ObjData(Object).ObjType = 23 Then
            ObjData(Object).MinSkill = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinSkill"))

        End If
        
        ObjData(Object).sagrado = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Sagrado"))
        ObjData(Object).MineralIndex = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MineralIndex"))

        ObjData(Object).MaxHP = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaxHP"))
        ObjData(Object).MinHP = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinHP"))

        ObjData(Object).Mujer = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Mujer"))
        ObjData(Object).Hombre = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Hombre"))

        ObjData(Object).MinHam = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinHam"))
        ObjData(Object).MinSed = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinAgu"))

        ObjData(Object).MinDef = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MINDEF"))
        ObjData(Object).MaxDef = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MAXDEF"))

        ObjData(Object).Vendible = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Vendible"))

        ObjData(Object).Respawn = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "ReSpawn"))

        ObjData(Object).RazaEnana = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "RazaEnana"))
        ObjData(Object).RazaElfa = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "RazaElfa"))
        ObjData(Object).RazaVampiro = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Razavampiro"))
        ObjData(Object).RazaOrco = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "RazaOrco"))
        ObjData(Object).ClaseAsesino = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Claseasesino"))
        ObjData(Object).RazaHumana = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Razahumana"))
        ObjData(Object).RazaHobbit = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Razahobbit"))
        ObjData(Object).Expe = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Expe"))
        ObjData(Object).Valor = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Valor"))
        ObjData(Object).Skill = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Skill"))
        '[Helios6]Objetos especiales
        ObjData(Object).objetoespecial = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "objetoespecial"))
        '[/END]
        ObjData(Object).Crucial = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Crucial"))

        ObjData(Object).Cerrada = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "abierta"))

        If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Llave"))
            ObjData(Object).CLAVE = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Clave"))

        End If

        If ObjData(Object).ObjType = eOBJType.otPuertas Or ObjData(Object).ObjType = eOBJType.otBotellaVacia Or ObjData(Object).ObjType = eOBJType.otBotellaLlena Then   'OBJTYPE_PUERTAS Or ObjData(Object).ObjType = eOBJType.otBotellaVacia Or ObjData(Object).ObjType = eOBJType.otBotellaLlena Then
            ObjData(Object).IndexAbierta = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "IndexCerradaLlave"))

        End If

        'Puertas y llaves
        ObjData(Object).CLAVE = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Clave"))

        ObjData(Object).texto = INI_GetString("Obj.Dat", "OBJ" & Object, "Texto")
        ObjData(Object).GrhSecundario = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "VGrande"))

        ObjData(Object).Agarrable = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Agarrable"))
        ObjData(Object).ForoID = INI_GetString("Obj.Dat", "OBJ" & Object, "ID")

        ObjData(Object).Acuchilla = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Acuchilla"))
        ObjData(Object).Guante = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Guante"))

        'Subastas 15-03-10 MIRAR BIN
        ObjData(Object).NoSubasta = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "NoSubasta"))
        'Subastas 15-03-10

        'Regalo 12-03-12
        ObjData(Object).NoRegalo = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Regalo"))
        'Regalo 12-03-12
        
        Dim i As Integer
        
        For i = 1 To NUMCLASES

            Clase = UCase$(INI_GetString("Obj.Dat", "OBJ" & Object, "CP" & i))
            
            If Len(Clase) > 0 Then
            
                ObjData(Object).ClasesProhibidas(i) = Clase
                 
            End If

        Next

        ObjData(Object).Resistencia = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Resistencia"))

        'Pociones
        If ObjData(Object).ObjType = 11 Then
            ObjData(Object).TipoPocion = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "DuracionEfecto"))

        End If

        If ObjData(Object).ObjType = 51 Then
            ObjData(Object).TipoRegalo = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "TipoRegalo"))

            For i = 1 To 10
                ObjData(Object).Objetos(i) = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Objetos" & i))
                ObjData(Object).Cantidad(i) = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Cantidad" & i))
            Next

        End If

        ObjData(Object).SkSastreria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkSastreria"))
        ObjData(Object).SkCarpinteria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkCarpinteria"))
        ObjData(Object).SkHechizeria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHechizeria"))

        If ObjData(Object).SkCarpinteria > 0 Then ObjData(Object).Madera = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Madera"))
        ObjData(Object).MaderaElfica = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaderaElfica"))

        If ObjData(Object).SkSastreria > 0 Then
            ObjData(Object).Lana = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Lana"))
            ObjData(Object).Lobos = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Lobos"))
            ObjData(Object).Osos = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Osos"))
            ObjData(Object).Tigre = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Tigre"))
            ObjData(Object).OsoPolar = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "OsoPolar"))
            ObjData(Object).Vaca = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Vaca"))
            ObjData(Object).Jabali = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Jabali"))

            'ObjData(Object).LoboPolar = val(INI_GetString("Obj.Dat", "OBJ" & Object,"LoboPolar"))
        End If

        If ObjData(Object).SkHechizeria > 0 Then ObjData(Object).ObjHierba = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "ObjHierba"))

        '[Helios]
        If ObjData(Object).ObjType = eOBJType.otBarcos Then
            ObjData(Object).MaxHit = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHit = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinHIT"))
            ObjData(Object).Velocidad = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Velocidad"))

        End If

        If ObjData(Object).ObjType = eOBJType.otBarcosArmada Then

            ObjData(Object).MaxHit = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHit = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinHIT"))
            ObjData(Object).Caos = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Caos"))
            ObjData(Object).Real = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Real"))
            ObjData(Object).templ = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Templ"))
            ObjData(Object).Nemes = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Nemes"))
            ObjData(Object).Facc = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Facc"))
            ObjData(Object).Velocidad = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Velocidad"))

        End If

        '[Helios]
        'If ObjData(Object).ObjType = OBJTYPE_BARCOS Then
        '       ObjData(Object).MaxHIT = val(INI_GetString("Obj.Dat", "OBJ" & Object,"MaxHIT"))
        '       ObjData(Object).MinHIT = val(INI_GetString("Obj.Dat", "OBJ" & Object,"MinHIT"))
        '      ObjData(Object).Real = val(INI_GetString("Obj.Dat", "OBJ" & Object,"Real"))
        '      ObjData(Object).Caos = val(INI_GetString("Obj.Dat", "OBJ" & Object,"Caos"))
        ' End If
        If ObjData(Object).ObjType = eOBJType.otPases Then
            ObjData(Object).Zona = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Zona"))

        End If

        If ObjData(Object).ObjType = eOBJType.otFlechas Then
            ObjData(Object).MaxHit = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHit = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinHIT"))
            ObjData(Object).LingH = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingO"))
            ObjData(Object).LingM = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "LingM"))
            ObjData(Object).Paraliza = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Paraliza"))
            ObjData(Object).TipoMunicion = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "TipoMunicion"))
            ObjData(Object).SkHerreria = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "SkHerreria"))

        End If

        'Bebidas
        ObjData(Object).MinSta = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "MinST"))
        
        'Sistema de subir nivel a clanes.
        ObjData(Object).Clan = Val(INI_GetString("Obj.Dat", "OBJ" & Object, "Clan"))
        'FrmMain.Barra.value = Object
        FrmMain.ucProgress.value = Object
        DoEvents
    Next Object
    
    ObjCargado = True
    Call ForeColorColumn(vbCyan, 0, FrmMain.ListView)
    Exit Sub

ErrHandler:
    MsgBox "error cargando objetos en objeto " & Object

End Sub

Public Sub CargaNpcsDat()
    '############################################
    '## binmode: nueva carga de inis con       ##
    '## fini_*, por indexacion hash            ##
    '############################################

    Dim NpcNumber As Integer, NumNpcsNoH As Integer

    FrmMain.ucProgress.Visible = True
    Call AbreFINI(DatPath & "NPCs-HOSTILES.dat")
    Call SelectFINI("INIT")
    
    NumNpcs = Val(LeeFINI("NumNPCs")) + 500

    ReDim NpcDats(1 To NumNpcs) As npc    'NPCS

    Call AbreFINI(DatPath & "NPCs.dat")
    Call SelectFINI("INIT")
    NumNpcsNoH = Val(LeeFINI("NumNPCs"))
    FrmMain.ucProgress.max = NumNpcs

    For NpcNumber = 1 To NumNpcs

        If (NpcNumber = 500) Then
            Call AbreFINI(DatPath & "NPCs-HOSTILES.dat")

        End If

        If NpcNumber <= NumNpcsNoH Or NpcNumber >= 500 Then
            Call SelectFINI("NPC" & NpcNumber)
            NpcDats(NpcNumber).numero = NpcNumber

            NpcDats(NpcNumber).Name = LeeFINI("Name")
            NpcDats(NpcNumber).Desc = LeeFINI("Desc")

            NpcDats(NpcNumber).Movement = Val(LeeFINI("Movement"))
            NpcDats(NpcNumber).flags.OldMovement = NpcDats(NpcNumber).Movement

            NpcDats(NpcNumber).flags.AguaValida = Val(LeeFINI("AguaValida"))
            NpcDats(NpcNumber).flags.TierraInvalida = Val(LeeFINI("TierraInValida"))
            NpcDats(NpcNumber).flags.Faccion = Val(LeeFINI("Faccion"))
            NpcDats(NpcNumber).flags.Especial = Val(LeeFINI("Especial"))

            NpcDats(NpcNumber).NPCtype = Val(LeeFINI("NpcType"))

            NpcDats(NpcNumber).char.Body = Val(LeeFINI("Body"))
            NpcDats(NpcNumber).char.Head = Val(LeeFINI("Head"))
            NpcDats(NpcNumber).char.heading = Val(LeeFINI("Heading"))
            NpcDats(NpcNumber).Attackable = Val(LeeFINI("Attackable"))
            NpcDats(NpcNumber).Comercia = Val(LeeFINI("Comercia"))
            NpcDats(NpcNumber).RespawnOrig = Val(LeeFINI("PosOrig"))

            If (NpcNumber = 500) Then
                NpcDats(NpcNumber).Hostile = Val(LeeFINI("Hostile"))
                NpcDats(NpcNumber).flags.OldHostil = NpcDats(NpcNumber).Hostile

            End If

            NpcDats(NpcNumber).Hostile = Val(LeeFINI("Hostile"))
            NpcDats(NpcNumber).flags.OldHostil = NpcDats(NpcNumber).Hostile

            'Mithrandir - Respawn NPC's
            Dim loopX As Integer

            NpcDats(NpcNumber).NRespawn = Val(LeeFINI("NRespawn"))
            NpcDats(NpcNumber).MaxRespawn = Val(LeeFINI("MaxRespawn"))

            If NpcDats(NpcNumber).NRespawn > 0 Then
                ReDim NpcDats(NpcNumber).CriaturaR(1 To NpcDats(NpcNumber).NRespawn)

                For loopX = 1 To NpcDats(NpcNumber).NRespawn
                    NpcDats(NpcNumber).CriaturaR(loopX) = LeeFINI("r" & loopX)
                Next loopX

            End If

            NpcDats(NpcNumber).GiveEXP = Val(LeeFINI("GiveEXP"))

            'npcdats(npcnumber).flags.ExpDada = npcdats(npcnumber).GiveEXP
            NpcDats(NpcNumber).flags.ExpCount = NpcDats(NpcNumber).GiveEXP

            NpcDats(NpcNumber).Veneno = Val(LeeFINI("Veneno"))

            NpcDats(NpcNumber).flags.Domable = Val(LeeFINI("Domable"))
            NpcDats(NpcNumber).flags.ItemDoma = Val(LeeFINI("ItemDoma"))

            NpcDats(NpcNumber).subasta = Val(LeeFINI("subasta"))    'Subasta 15-03-10
            NpcDats(NpcNumber).GiveGLD = Val(LeeFINI("GiveGLD"))

            NpcDats(NpcNumber).PoderAtaque = Val(LeeFINI("PoderAtaque"))
            NpcDats(NpcNumber).PoderEvasion = Val(LeeFINI("PoderEvasion"))

            NpcDats(NpcNumber).SkillDomar = Val(LeeFINI("SkillDomar"))    'Skill doma

            NpcDats(NpcNumber).InvReSpawn = Val(LeeFINI("InvReSpawn"))
            NpcDats(NpcNumber).DecirPalabras = Val(LeeFINI("DecirPalabras"))

            NpcDats(NpcNumber).Stats.MaxHP = Val(LeeFINI("MaxHP"))
            NpcDats(NpcNumber).Stats.MinHP = Val(LeeFINI("MinHP"))
            NpcDats(NpcNumber).Stats.MaxHit = Val(LeeFINI("MaxHIT"))
            NpcDats(NpcNumber).Stats.MinHit = Val(LeeFINI("MinHIT"))
            NpcDats(NpcNumber).Stats.DEF = Val(LeeFINI("DEF"))
            NpcDats(NpcNumber).Stats.Alineacion = Val(LeeFINI("Alineacion"))
            NpcDats(NpcNumber).Stats.ImpactRate = Val(LeeFINI("ImpactRate"))

            Dim loopc As Integer

            Dim ln    As String

            NpcDats(NpcNumber).Invent.NroItems = Val(LeeFINI("NROITEMS"))

            For loopc = 1 To NpcDats(NpcNumber).Invent.NroItems
                ln = LeeFINI("Obj" & loopc)
                NpcDats(NpcNumber).Invent.Object(loopc).ObjIndex = Val(ReadField(1, ln, 45))
                NpcDats(NpcNumber).Invent.Object(loopc).Amount = Val(ReadField(2, ln, 45))

                ln = LeeFINI("Prob" & loopc)

                If (Len(ln) = 0) Then ln = "100"
                NpcDats(NpcNumber).Invent.Object(loopc).ProbDrop = CByte(ln)

            Next loopc

            NpcDats(NpcNumber).flags.Mago = Val(LeeFINI("Mago"))
            NpcDats(NpcNumber).flags.LanzaSpells = Val(LeeFINI("LanzaSpells"))

            If NpcDats(NpcNumber).flags.LanzaSpells > 0 Then ReDim NpcDats(NpcNumber).Spells(1 To NpcDats(NpcNumber).flags.LanzaSpells)

            For loopc = 1 To NpcDats(NpcNumber).flags.LanzaSpells
                NpcDats(NpcNumber).Spells(loopc) = Val(LeeFINI("Sp" & loopc))
            Next loopc

            If NpcDats(NpcNumber).NPCtype = eNPCType.Entrenador Then
                NpcDats(NpcNumber).NroCriaturas = Val(LeeFINI("NroCriaturas"))
                ReDim NpcDats(NpcNumber).criaturas(1 To NpcDats(NpcNumber).NroCriaturas) As tCriaturasEntrenador

                For loopc = 1 To NpcDats(NpcNumber).NroCriaturas
                    NpcDats(NpcNumber).criaturas(loopc).NpcIndex = LeeFINI("CI" & loopc)
                    NpcDats(NpcNumber).criaturas(loopc).NpcName = LeeFINI("CN" & loopc)
                Next loopc

            End If
            
            NpcDats(NpcNumber).flags.Respawn = Val(LeeFINI("ReSpawn"))
            
            NpcDats(NpcNumber).Inflacion = Val(LeeFINI("Inflacion"))

            NpcDats(NpcNumber).flags.NPCActive = True
            NpcDats(NpcNumber).flags.UseAINow = False

            NpcDats(NpcNumber).flags.BackUp = Val(LeeFINI("BackUp"))
            NpcDats(NpcNumber).flags.RespawnOrigPos = Val(LeeFINI("OrigPos"))
            NpcDats(NpcNumber).flags.AfectaParalisis = Val(LeeFINI("AfectaParalisis"))
            NpcDats(NpcNumber).flags.ParalisisSagrado = Val(LeeFINI("ParalisisSagrado"))
            NpcDats(NpcNumber).flags.Magiainvisible = Val(LeeFINI("Magiainvisible"))
            NpcDats(NpcNumber).flags.npcSagrado = Val(LeeFINI("npcSagrado"))
            NpcDats(NpcNumber).flags.GolpeExacto = Val(LeeFINI("GolpeExacto"))

            NpcDats(NpcNumber).flags.Snd1 = Val(LeeFINI("Snd1"))
            NpcDats(NpcNumber).flags.Snd2 = Val(LeeFINI("Snd2"))
            NpcDats(NpcNumber).flags.Snd3 = Val(LeeFINI("Snd3"))
            NpcDats(NpcNumber).flags.Snd4 = Val(LeeFINI("Snd4"))
            'NpcDats(NpcNumber).Habla = val(LeeFINI("Habla"))

            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

            Dim aux As String

            aux = LeeFINI("NROEXP")

            If aux = "" Then
                NpcDats(NpcNumber).NroExpresiones = 0
            Else
                NpcDats(NpcNumber).NroExpresiones = Val(aux)
                ReDim NpcDats(NpcNumber).Expresiones(1 To NpcDats(NpcNumber).NroExpresiones) As String

                For loopc = 1 To NpcDats(NpcNumber).NroExpresiones
                    NpcDats(NpcNumber).Expresiones(loopc) = LeeFINI("Exp" & loopc)
                Next loopc

            End If

            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

            'Tipo de items con los que comercia
            NpcDats(NpcNumber).TipoItems = Val(LeeFINI("TipoItems"))

        End If

        'CRAW; estuve por aqui xd
        '  NpcDats(NpcNumber).Name = LeeFINI("Name")
        '    NpcDats(NpcNumber).Anima = val(LeeFINI("Animacion"))
        FrmMain.ucProgress.value = NpcNumber
        DoEvents
    Next NpcNumber

    NpcsCargado = True

End Sub

