Attribute VB_Name = "Declares"
Option Explicit

Public NpcNumber           As Integer, NumNpcsNoH As Integer

Public cargaind            As Integer

Public Const NumGrh        As Long = 37000

Public DatPath             As String

Public NumObjDatas         As Integer

Public Fichero             As String

Public subelemento         As ListItem

Public InitPath            As String

Public GrafPath            As String

Public ClickMouseX         As Long, ClickMouseY As Long

Public WindowMoving        As Boolean

Public itl                 As Integer

Public TaskList(1000)      As String

Public i                   As Integer

Public GrhIndex            As String

Public Ruta                As String

Public Data                As String

Public Datos               As String

Public ObjCargado          As Boolean

Public NumNpcs             As Integer

Public NpcsCargado         As Boolean

Public NpcsCargadoHostiles As Boolean

Public ListadoQueQuiero    As String

Public FormaBusca          As Integer

Public Object              As Integer

Public Leer                As New clsIniManager

Public Const MAX_INVENTORY_SLOTS = 25

Public Const NUMCLASES = 15

'-----------------------NPC--------------------------------
'Npc
Public NpcList() As npc    'NPCS

Public NpcDats() As npc    'NPCS

Public Type tVertice

    X          As Long
    Y          As Long

End Type

Type NpcPathFindingInfo

    path()     As tVertice      ' This array holds the path
    Target     As tVertice      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos     As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath     As Boolean       ' If it is true there is no path to the target location

    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.

End Type

Type tCriaturasEntrenador

    NpcIndex   As Integer
    NpcName    As String
    tmpIndex   As Integer

End Type

'Datos de user o npc
Type char

    CharIndex  As Integer
    Head       As Integer
    HeadBandas As Integer    'Guerra Bandas 15-03-10
    Body       As Integer
    BodyBandas As Integer    'Guerra Bandas 15-03-10
    '[GAU]
    Botas      As Integer
    '[GAU]
    Alas       As Integer    'Helios 13-12-10
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim  As Integer
    'AmuletoAnim As Integer 'Helios
    'AnilloAnim As Integer 'Helios
    FX         As Integer
    loops      As Integer

    heading    As Byte
    
    AuraAnim As Integer

End Type

Public Enum eNPCType

    Comun = 0
    revividor = 1
    Guardiass = 2
    Entrenador = 3
    Banquero = 4
    Armadas = 5
    'El 6 lo usa los dracos (OJO!)
    Cirujia = 7
    Vendecaballos = 8
    CaballoMaron = 9
    Juzgado = 10
    ObjetosSagrados = 11
    Transporte = 12
    Organizadorquest = 13
    Quesat = 14
    VendeUnicornio = 15
    Unicornio = 16
    GuardiaCaos = 17
    Matrimonio = 18
    Resucitamascota = 19
    BarqueroReal = 20
    BarqueroCaos = 21
    torneo = 22
    casino = 23
    RegalaMascota = 24
    Carcelero = 25
    BarqueroNem = 26
    Subastador = 27
    SeñorQuest = 28
    BarqueroTem = 29
    Quesst = 30
    Chismoso = 31
    Caballos = 32
    GuerrMaster = 33
    Propiedades = 34
    AoMCreditos = 35
    AoMCanjes = 36
    TorneoPareja = 37
    BatallaMedusa = 38
    
    nQuest = 40
    misiones = 41
    'Sistema de subir nivel a clanes.
    Banqueroclan = 42
    
    RespawBichos = 50
    Olvidarechizo = 54

    Monturas = 60
    
    NoMejora = 90

End Enum

Type NPCStats

    Alineacion As Integer
    MaxHP      As Long
    MinHP      As Long
    MaxHit     As Integer
    MinHit     As Integer
    DEF        As Integer
    UsuariosMatados As Integer
    ImpactRate As Integer

End Type

Type NpcCounters

    Inmoviliza As Integer
    Paralisis  As Integer
    TiroCriaturas As Integer
    TiempoExistencia As Long

End Type

Type NPCFlags

    AfectaParalisis As Byte
    ParalisisSagrado As Byte
    Magiainvisible As Byte
    npcSagrado As Byte
    GolpeExacto As Byte
    Domable    As Integer
    ItemDoma As Integer
    Respawn    As Byte
    NPCActive  As Boolean    '¿Esta vivo?
    Follow     As Boolean
    Faccion    As Byte
    LanzaSpells As Byte
    Especial   As Byte

    Mago       As Byte

    ExpCount   As Long    '[ALEJO]

    OldMovement As Byte
    OldHostil  As Byte

    AguaValida As Byte
    TierraInvalida As Byte

    UseAINow   As Boolean
    Sound      As Integer
    Attacking  As Integer
    AttackedBy As String
    AttackedByN As Integer

    BackUp     As Byte
    RespawnOrigPos As Byte

    Envenenado As Long
    Paralizado As Byte
    Inmovilizado As Byte
    Invisible  As Byte
    Maldicion  As Byte
    Bendicion  As Byte

    Snd1       As Integer
    Snd2       As Integer
    Snd3       As Integer
    Snd4       As Integer
    QuienParalizo As Integer

End Type

Type WorldPos

    Map        As Integer
    X          As Integer
    Y          As Integer

End Type

Type UserOBJ

    ObjIndex   As Integer
    Amount     As Integer
    Equipped   As Byte
    ProbDrop   As Byte

End Type

Type Inventario

    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    DemoleEqpObjIndex As Integer
    DemoleEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    HerramientaEqpObjIndex As Integer
    HerramientaEqpSlot As Integer
    BarcoObjIndex As Integer
    BarcoSlot  As Byte

    NroItems   As Integer
    AmuletoEqpObjIndex As Integer    'Helios
    AmuletoEqpSlot As Byte    'Helios
    GuantesEqpObjIndex As Integer    'Helios
    GuantesEqpSlot As Byte    'Helios

    AnilloMagoEqpobjIndex As Integer
    AnilloMagoEqpSlot As Byte

    AnilloEqpObjIndex As Integer    'Helios
    AnilloEqpSlot As Byte    'Helios
    '[GAU]
    BotaEqpObjIndex As Integer
    BotaEqpSlot As Byte
    '[GAU]
    '[Helios 13-12-10]
    AlasEqpObjIndex As Integer
    AlasEqpSlot As Byte
    
    CarroEqpObjIndex As Integer
    CarroEqpSlot As Byte

    '[Helios 13-12-10]
    
    BonusFuerza As Integer
    BonusAgilidad As Integer
    
End Type

Type npc

    Name       As String
    char       As char    'Define como se vera
    Desc       As String

    NPCtype    As eNPCType
    'NPCtype As Integer
    numero     As Integer

    level      As Integer

    InvReSpawn As Byte

    subasta    As Integer    'Subastas 15-03-10
    Comercia   As Integer
    Target     As Long
    TargetNpc  As Long
    TipoItems  As Integer

    Veneno     As Byte

    Pos        As WorldPos    'Posicion
    Orig       As WorldPos
    RespawnOrig As Boolean
    SkillDomar As Integer

    Movement   As Integer
    Attackable As Byte
    Hostile    As Byte
    PoderAtaque As Long
    PoderEvasion As Long

    Inflacion  As Long

    GiveEXP    As Long

    GiveGLD    As Long
    TengoFlechas(1 To 6) As Long
    Stats      As NPCStats
    flags      As NPCFlags
    Contadores As NpcCounters

    MurallaEquipo As Byte    'asedio
    MurallaIndex As Byte    'asedio

    Invent     As Inventario
    Prob(1 To MAX_INVENTORY_SLOTS) As Double
    CanAttack  As Byte

    NroExpresiones As Byte
    Expresiones() As String    ' le da vida ;)

    'Mithrandir - Respawn NPC's
    NRespawn   As Byte
    MaxRespawn As Byte
    CriaturaR() As Integer

    NroSpells  As Byte
    Spells()   As Integer  ' le da vida ;)

    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas   As Integer

    '<---------New!! Needed for pathfindig----------->
    PFINFO     As NpcPathFindingInfo
    DoorPos    As WorldPos

    'AI vario
    VuelvoAMiSitio As Boolean

    'Habla As Integer 'npc habla (dat Habla=X "x es el numero del wav")

    DecirPalabras As Byte

    'Anima      As Integer

End Type

'-----------------------NPC--------------------------------
Type Grh

    bmp As String
    X As Integer
    Y As Integer
    ancho As Integer
    alto As Integer
    anim As Integer

End Type

'########DECLARACIONES: OBJ.DAT########

Public Enum sObjtype

    sotArmadura = 0
    sotCasco = 1
    sotEscudo = 2
    sotBotas = 3
    sotAlas = 4
    sotCaña = 138
    sotRed = 543

End Enum

Public Enum eOBJType

    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7    'no se usa
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otAtlas = 12    'no se usa
    otBebidas = 13
    otleña = 14
    otFogata = 15
    otHerramientas = 18
    otGuantes = 16
    oAmago = 17
    'otESCUDO = 16
    'otCASCO = 17
    otTeleport = 19
    otRegalo = 20
    otyacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31

    otBarcosArmada = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otPocionResucitar = 36
    otHead = 37
    otHierba = 38
    otLana = 39
    otOveja = 40
    otPases = 41
    otArbolElfico = 42    'maderaelfica
    otAlcohol = 43
    otAmuletos = 44
    otAnillo = 45
    otSagrado = 46
    otTeleportar = 47
    otClanes = 48
    otGemaClan = 49
    otValeExp = 50
    otRegalos = 51
    otTumbas = 52
    otTesoroA = 53
    otTesoroB = 54
    otTesoroC = 55
    otTiocolgado = 58
    otMontura = 60
    
    'Sistema de subir nivel a clanes.
    otClanNivel = 61

    'otMochilas = 37
    otCualquiera = 1000

End Enum

Public Type ObjData

    ParaCarpin As Byte 'nuevo
    
    ParaHerre  As Byte 'nuevo
    
    Name       As String
    numero     As Integer
    ObjType    As Integer
    SubTipo    As Integer

    GrhIndex   As Integer
    GrhSecundario As Integer

    Respawn    As Byte
    
    MaxItems   As Integer
    Conte      As Inventario
    Apuñala    As Byte
    Acuchilla  As Byte

    Drena      As Byte
    Absorbe    As Byte
    Mana       As Byte
    Vida       As Byte
    Perfora    As Byte
    Paraliza   As Byte
    Ceguera    As Byte
    Estupidez  As Byte
    
    HechizoIndex As Integer
    Clase      As String
    ForoID     As String

    MinHP      As Integer
    MaxHP      As Integer

    MineralIndex As Integer
    LingoteInex As Integer

    DosManos   As Integer
    proyectil  As Integer
    TipoProyectil As Integer
    Municion   As Integer
    TipoMunicion As Integer

    Crucial    As Byte
    Newbie     As Integer
    Destruir   As Byte
    
    MinSta     As Integer
    
    TipoPocion As Byte
    TipoRegalo As Byte
    Objetos(1 To 10) As Integer
    Cantidad(1 To 10) As Integer
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill   As Integer
    LingoteIndex As Integer
    NoTransferible As Integer 'nuevo

    MinHit     As Integer
    MaxHit     As Integer

    LvlMin     As Integer
    LvlMax     As Integer
    Heroe      As Integer

    MinHam     As Integer
    MinSed     As Integer

    Zona       As Integer
    Cae        As Byte
    TiRaRObj   As Byte
    Velocidad  As Byte
    DEF        As Integer
    MinDef     As Integer
    MaxDef     As Integer

    Ropaje     As Integer

    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim  As Integer
    
    Botas      As Integer
   
    Alas       As Integer
    Expe       As Long
    Valor      As Long
    Skill      As Byte
    objetoespecial As Integer
    Nivel      As Byte
    GameMaster As Byte
    Cerrada    As Integer
    Llave      As Byte
    CLAVE      As Long

    UsoNpc     As Byte

    Guante     As Byte

    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer

    RazaEnana  As Byte
    Mujer      As Byte
    Hombre     As Byte
    Envenena   As Byte

    NoSubasta  As Byte    'Subastas 15-03-10
    NoRegalo   As Byte    'Regalo 12-03-12
    Resistencia As Long
    Agarrable  As Byte
    sagrado As Byte

    LingH      As Integer
    LingO      As Integer
    LingP      As Integer
    LingM      As Integer
    Madera     As Integer
    MaderaElfica As Integer
    Lana       As Integer
    Lobos      As Integer
    Osos       As Integer
    
    Gemas      As Integer
    Diamantes  As Integer
    
    Tigre      As Integer
    OsoPolar   As Integer
    Vaca       As Integer
    Jabali     As Integer
    
    ObjHierba  As Integer
    Vendible   As Byte

    SkHerreria As Integer
    SkHerreriaMagica As Integer
    SkSastreria As Integer
    SkCarpinteria As Integer
    SkHechizeria As Integer

    texto      As String

    ClasesProhibidas(1 To NUMCLASES) As String

    Snd1       As Integer
    Snd2       As Integer
    Snd3       As Integer
    MinInt     As Integer

    Nemes      As Byte
    templ      As Byte
    Real       As Byte
    Caos       As Byte
    Facc       As Byte 'nuevo
    Robable    As Byte

    bonifica   As Integer
    tipobonifica As Integer
    Donacion   As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    AtaqueMagicaMax As Integer
    AtaqueMagicaMin As Integer

    RazaElfa   As Integer
    RazaVampiro As Integer
    RazaOrco   As Integer
    RazaHumana As Integer
    ClaseAsesino As Integer
    RazaHobbit As Integer
    
    TiempoPocion As Long 'nuevo
    
    Clan As Byte 'nuevo
    
    TipoAura As Long 'nuevo
    
    TipoLibro As Long 'nuevo

End Type

Public ObjData() As ObjData
