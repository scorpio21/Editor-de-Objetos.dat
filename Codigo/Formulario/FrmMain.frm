VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form FrmMain 
   Caption         =   "Editor de Objetos By Helios"
   ClientHeight    =   5940
   ClientLeft      =   4695
   ClientTop       =   3240
   ClientWidth     =   12180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmMain.frx":08CA
   ScaleHeight     =   5940
   ScaleWidth      =   12180
   Begin Proyecto1.uAOButton UOBuscar 
      Height          =   390
      Left            =   480
      TabIndex        =   17
      Top             =   1260
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   688
      TX              =   "Buscar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "FrmMain.frx":EB7BA
      PICF            =   "FrmMain.frx":EC1E4
      PICH            =   "FrmMain.frx":ECEA6
      PICV            =   "FrmMain.frx":EDE38
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtNpcNombre 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   300
      TabIndex        =   16
      Top             =   330
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtNpcs 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   300
      MaxLength       =   5
      TabIndex        =   15
      Top             =   345
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MSComctlLib.ListView Listnpcs 
      Height          =   5680
      Left            =   1980
      TabIndex        =   13
      Top             =   120
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   10028
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   65535
      BackColor       =   0
      Appearance      =   0
      MousePointer    =   99
      NumItems        =   73
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº Npc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripcion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Head"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Body"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CascoAnim"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "WeaponAnim"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ShieldAnim"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Heading"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Movimiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Comercio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Alineacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "ReSpawn"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "NroItems"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Obj1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Obj2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Obj3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Obj4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Obj5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Obj6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Obj7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Obj8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Obj9"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Obj10"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Obj11"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Obj12"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Obj13"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Obj14"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Obj15"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Obj16"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Obj17"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Obj18"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "Obj19"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Obj20"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Obj21"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Obj22"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "Obj23"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "Obj24"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "Obj25"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "AfectaParalisis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Text            =   "Prob1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "Prob2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "Prob3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "Prob4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "Prob5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "Prob6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "Prob7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "Prob8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "Prob9"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "Prob10"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "Prob11"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "Prob12"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "Prob13"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "Prob14"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "Prob15"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Text            =   "Prob16"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "Npctype"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "Attackable"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "Hostile"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "GiveEXP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Text            =   "GiveGLD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Text            =   "MinHP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Text            =   "MaxHP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   63
         Text            =   "MaxHIT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   64
         Text            =   "MinHIT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   65
         Text            =   "DEF"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   66
         Text            =   "PoderAtaque"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   67
         Text            =   "PoderEvasion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   68
         Text            =   "Domable"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   69
         Text            =   "TipoItems"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   70
         Text            =   "Inflacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(72) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   71
         Text            =   "BackUp"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(73) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   72
         Text            =   "Hostile"
         Object.Width           =   2540
      EndProperty
   End
   Begin Proyecto1.uAOCheckbox uAOCheckbox 
      Height          =   225
      Left            =   1245
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   915
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   397
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "FrmMain.frx":EED3A
   End
   Begin VB.CheckBox chktipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aom"
      ForeColor       =   &H00FF00FF&
      Height          =   210
      Left            =   1275
      TabIndex        =   11
      Top             =   930
      Width           =   630
   End
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   315
      TabIndex        =   7
      Top             =   345
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.ComboBox LstLista 
      Height          =   315
      ItemData        =   "FrmMain.frx":EFC52
      Left            =   180
      List            =   "FrmMain.frx":EFC62
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   900
      Width           =   1050
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3600
      Top             =   5280
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   3195
      Top             =   5280
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   2760
      Top             =   5235
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5680
      Left            =   9735
      TabIndex        =   2
      Top             =   120
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   10028
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   65535
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ObjType"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtObjType 
      BackColor       =   &H000000FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   270
      MaxLength       =   5
      TabIndex        =   0
      ToolTipText     =   "Elige el ObjType de la lista."
      Top             =   360
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   5680
      Left            =   1980
      TabIndex        =   6
      Top             =   120
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   10028
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   65535
      BackColor       =   0
      Appearance      =   0
      MousePointer    =   99
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmMain.frx":EFC8A
      NumItems        =   161
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº Objeto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ObjType"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Subtipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "MaxHit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "MinHit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "MaxDef"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "MinDef"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "MinHam"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Resistencia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Vendible"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "SkSastreria"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "SkHechizeria"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ObjHierba"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "MinSkill"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Velocidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Facc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "objetoespecial"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "MinStat"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Caos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Real"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Templario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Nemesis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Cae"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "TiRaRObj"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Robable"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Bonifica"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "TipoBonifica"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Donacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Newbie"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Destruir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "ShieldAnim"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "LingH"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "LingP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "LingO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "LingM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "Madera"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "MaderaElfica"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "SkHerreria"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Text            =   "Drena"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "SkHerreriaMagica"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "Gemas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "Diamantes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "CascoAnim"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "DefensaMagicaMax"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "DefensaMagicaMin"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "Botas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "Alas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "Ropaje"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "HechizoIndex"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "Clase"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "WeaponAnim"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "Apuñala"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "Paraliza"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Text            =   "Ceguera"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "Estupidez"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "Vida"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "Mana"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "Envenena"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Text            =   "LvlMin"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Text            =   "LvlMax"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Text            =   "Heroe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   63
         Text            =   "proyectil"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   64
         Text            =   "Municion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   65
         Text            =   "TipoProyectil"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   66
         Text            =   "DosManos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   67
         Text            =   "Nivel"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   68
         Text            =   "GameMaster"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   69
         Text            =   "UsoNpc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   70
         Text            =   "Snd1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(72) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   71
         Text            =   "Snd2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(73) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   72
         Text            =   "Snd3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(74) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   73
         Text            =   "MinInt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(75) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   74
         Text            =   "LingoteIndex"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(76) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   75
         Text            =   "sagrado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(77) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   76
         Text            =   "MineralIndex"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(78) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   77
         Text            =   "MaxHP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(79) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   78
         Text            =   "MinHP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(80) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   79
         Text            =   "Mujer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(81) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   80
         Text            =   "Hombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(82) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   81
         Text            =   "MinSed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(83) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   82
         Text            =   "Respawn"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(84) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   83
         Text            =   "RazaEnana"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(85) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   84
         Text            =   "RazaElfa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(86) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   85
         Text            =   "RazaVampiro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(87) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   86
         Text            =   "RazaOrco"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(88) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   87
         Text            =   "ClaseAsesino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(89) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   88
         Text            =   "RazaHumana"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(90) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   89
         Text            =   "RazaHobbit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(91) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   90
         Text            =   "Expe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(92) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   91
         Text            =   "Skill"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(93) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   92
         Text            =   "Crucial"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(94) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   93
         Text            =   "Cerrada"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(95) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   94
         Text            =   "Llave"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(96) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   95
         Text            =   "Clave"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(97) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   96
         Text            =   "IndexAbierta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(98) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   97
         Text            =   "IndexCerrada"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(99) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   98
         Text            =   "IndexCerradaLlave"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(100) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   99
         Text            =   "Texto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(101) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   100
         Text            =   "GrhSecundario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(102) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   101
         Text            =   "Agarrable"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(103) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   102
         Text            =   "ForoID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(104) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   103
         Text            =   "Acuchilla"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(105) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   104
         Text            =   "Guante"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(106) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   105
         Text            =   "NoSubasta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(107) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   106
         Text            =   "NoRegalo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(108) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   107
         Text            =   "TipoPocion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(109) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   108
         Text            =   "MaxModificador"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(110) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   109
         Text            =   "MinModificador"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(111) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   110
         Text            =   "DuracionEfecto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(112) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   111
         Text            =   "TipoRegalo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(113) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   112
         Text            =   "Objeto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(114) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   113
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(115) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   114
         Text            =   "SkCarpinteria"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(116) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   115
         Text            =   "Lana"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(117) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   116
         Text            =   "Lobos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(118) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   117
         Text            =   "Osos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(119) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   118
         Text            =   "Tigre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(120) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   119
         Text            =   "OsoPolar"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(121) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   120
         Text            =   "Vaca"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(122) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   121
         Text            =   "Jabali"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(123) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   122
         Text            =   "Zona"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(124) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   123
         Text            =   "TipoMunicion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(125) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   124
         Text            =   "Cp1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(126) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   125
         Text            =   "Cp2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(127) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   126
         Text            =   "Cp3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(128) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   127
         Text            =   "Cp4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(129) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   128
         Text            =   "Cp5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(130) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   129
         Text            =   "Cp6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(131) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   130
         Text            =   "Cp7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(132) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   131
         Text            =   "Cp8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(133) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   132
         Text            =   "Cp9"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(134) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   133
         Text            =   "Cp10"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(135) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   134
         Text            =   "Cp11"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(136) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   135
         Text            =   "Cp12"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(137) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   136
         Text            =   "Cp13"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(138) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   137
         Text            =   "Objeto1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(139) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   138
         Text            =   "Objeto2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(140) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   139
         Text            =   "Objeto3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(141) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   140
         Text            =   "Objeto4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(142) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   141
         Text            =   "Objeto5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(143) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   142
         Text            =   "Objeto6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(144) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   143
         Text            =   "Objeto7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(145) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   144
         Text            =   "Objeto8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(146) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   145
         Text            =   "Objeto9"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(147) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   146
         Text            =   "Objeto10"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(148) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   147
         Text            =   "Cantidad1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(149) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   148
         Text            =   "Cantidad2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(150) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   149
         Text            =   "Cantidad3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(151) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   150
         Text            =   "Cantidad4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(152) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   151
         Text            =   "Cantidad5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(153) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   152
         Text            =   "Cantidad6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(154) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   153
         Text            =   "Cantidad7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(155) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   154
         Text            =   "Cantidad8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(156) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   155
         Text            =   "Cantidad9"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(157) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   156
         Text            =   "Cantidad10"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(158) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   157
         Text            =   "TiempoPocion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(159) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   158
         Text            =   "TipoAura"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(160) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   159
         Text            =   "TipoLibro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(161) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   160
         Text            =   "GrhIndex"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListtypoNpcs 
      Height          =   5680
      Left            =   9730
      TabIndex        =   14
      Top             =   125
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   10028
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   65535
      BackColor       =   0
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NpcsType"
         Object.Width           =   2540
      EndProperty
   End
   Begin Proyecto1.ucProgressCircular ucBgrabar 
      Height          =   1065
      Left            =   300
      TabIndex        =   10
      Top             =   1605
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1879
      Caption1_ForeColor=   13892210
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      Caption2_ForeColor=   16711680
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StepSpaceSize   =   2
      PF_Width        =   5
      PF_Steps        =   36
      PB_Color1       =   16711680
      PB_Width        =   5
      PB_BorderColor  =   15527149
      Value           =   60
      CenterColor1    =   8388608
      PF_ForeColor    =   255
      AnimationInterval=   100
      PF_ColorsCount  =   2
      PF_Colors       =   "FrmMain.frx":F0964
   End
   Begin Proyecto1.ucProgressCircular ucProgress 
      Height          =   1065
      Left            =   300
      TabIndex        =   9
      Top             =   1605
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1879
      Caption1_ForeColor=   12632064
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PF_Width        =   5
      PF_Steps        =   36
      PB_Color1       =   13369344
      PB_Width        =   5
      Value           =   60
      PF_ForeColor    =   16711680
      AnimationInterval=   100
      PF_ColorsCount  =   3
      PF_Colors       =   "FrmMain.frx":F098D
   End
   Begin VB.Label lblLabel2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Busqueda:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   645
      Width           =   1785
   End
   Begin VB.Label Lbltexto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   4650
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   75
      TabIndex        =   3
      Top             =   2670
      Width           =   45
   End
   Begin VB.Label lblObjeType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº ObjeType"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   1230
   End
   Begin VB.Menu mmuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuGuardar 
         Caption         =   "&Guardar Ini"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objLabelEdit  As LabelEdit

Dim objLabelEdit2 As LabelEdit

Private Sub ListtypoNpcs_Click()

10  On Error GoTo ListtypoNpcs_Click_Err
        
20  GrhIndex = ListtypoNpcs.ListItems(ListtypoNpcs.SelectedItem.Index).Text
30  Datos = Val(ReadField(1, GrhIndex, 58))
40  txtNpcs.Text = Datos
    
50  Exit Sub

ListtypoNpcs_Click_Err:

End Sub

'Private Sub chktipo_Click()

Private Sub uAOCheckbox_Click()

    If uAOCheckbox.Checked = True Then
        cargaind = 1
    Else
        cargaind = 0

    End If

End Sub

Private Sub LstLista_Click()
    ListadoQueQuiero = LstLista.Text
    
    Select Case ListadoQueQuiero

        Case "Nombre"
            lblObjeType.Caption = "Por Nombre:"
            txtObjType.Visible = False
            txtNpcs.Visible = False
            txtObjType.Text = ""
            txtnombre.Visible = True
            ListView1.Visible = False
            ListView.Visible = True
            ListView.ListItems.Clear
            Listnpcs.Visible = False
            ListtypoNpcs.Visible = False
            uAOCheckbox.Visible = True
            chktipo.Visible = True
            ListView.Width = 9945
            TxtNpcNombre.Visible = False 'oculto el textbox de nombre npc

        Case "ObjType"
            ListView1.Enabled = True
            lblObjeType.Caption = "Nº ObjType"
            txtObjType.Visible = True
            txtNpcs.Visible = False
            txtnombre.Visible = False
            uAOCheckbox.Visible = True
            chktipo.Visible = True
            ListView1.Visible = True
            Listnpcs.Visible = False
            ListtypoNpcs.Visible = False
            txtnombre.Text = ""
            ListView.Width = 7635
            ListView.Visible = True
            TxtNpcNombre.Visible = False 'oculto el textbox de nombre npc

        Case "Npcs"
            FrmMain.Listnpcs.ListItems.Clear
            ListView1.Enabled = False
            lblObjeType.Caption = "Nº de Npcs"
            txtObjType.Visible = False
            txtnombre.Visible = False
            ListView1.Visible = False
            ListView.Visible = False
            Listnpcs.Visible = True
            ListtypoNpcs.Visible = True
            If uAOCheckbox.Checked = True Then
                uAOCheckbox.Checked = False
            End If
            uAOCheckbox.Visible = False
            chktipo.Visible = False
            txtNpcs.Visible = True
            Listnpcs.Width = 7635 'redimensiono la lista
            TxtNpcNombre.Visible = False 'oculto el textbox de nombre npc
            mnuGuardar.Enabled = False

        Case "Nombre Npcs"
            FrmMain.Listnpcs.ListItems.Clear 'Limpio la lista DoEvent npcs
            ListView1.Enabled = False 'desactivo la lista de objtype del listview1 del objeto
            ListView1.Visible = False 'Pongo invisible la lista de objtype del listview1 del objeto
            ListView.Visible = False 'Pongo invisible la lista del listview del objeto
            ListView.Enabled = False 'desactivo la lista del listview del objeto
            lblObjeType.Caption = "Nombre de Npcs" 'Pongo el texto al caption
            txtObjType.Visible = False 'Pongo invisible el texto del objtype
            txtnombre.Visible = False 'Pongo invisible el nombre de los objetos
            If uAOCheckbox.Checked = True Then
                uAOCheckbox.Checked = False
            End If
            uAOCheckbox.Visible = False 'oculto el checkbox de aomania
            chktipo.Visible = False 'oculto el checkbox de aomania
            txtNpcs.Visible = False 'oculto el textbox de la busqueda del npctype
            ListtypoNpcs.Visible = False 'oculto la lista del NpcType
            Listnpcs.Width = 9945 'redimensiono la lista
            TxtNpcNombre.Visible = True 'Muestro el textbox de nombre npc
            Listnpcs.Visible = True
            mnuGuardar.Enabled = False

    End Select

End Sub

Private Sub mnuSalir_Click()
    'Stop subclassing
    CloseSubClass

    'Clean up by setting the classes to Nothing
    Set objLabelEdit = Nothing
    End

End Sub

Private Sub Timer_Timer()
    'FrmMain.Barra.Visible = False
    ucProgress.Visible = False
    'FrmMain.Bgrabar.Visible = False
    ucBgrabar.Visible = False
    Lbltexto.ForeColor = vbRed

    'LblInfo.BackColor = &HFFFF&
    'LblInfo.Visible = False
    'FrmMain.Timer.Enabled = False
End Sub

Private Sub Timer2_Timer()
    Lbltexto.ForeColor = vbMagenta

    'LblInfo.BackColor = &HC0C0C0
End Sub

Private Sub Timer3_Timer()
    Lbltexto.Visible = False
    Timer.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    mnuGuardar.Enabled = True
    Call ForeColorColumn(vbCyan, 0, ListView)

    'Call ForeColorColumn(vbRed, 1,  ListView)
End Sub

Private Sub UOBuscar_Click()

    If LstLista.Text = vbNullString Then FrmMain.LstLista.SetFocus: Exit Sub
    
    If verpicture.Visible = True Then Unload verpicture
    
    ListView.Enabled = True
    ListView.ListItems.Clear

    Select Case LstLista.Text
    
        Case "Nombre"

            If txtnombre.Text = "" Then
                txtnombre.Text = "¿Nombre?"
                Exit Sub

            End If
            
            If txtnombre.Text = vbNullString Then Exit Sub
            If Not ObjCargado Then Call LoadOBJData
            ListadoQueQuiero = txtnombre.Text
            Call CargarListView(0, ListadoQueQuiero)

        Case "ObjType"
                         
            If txtObjType.Text = "" Then
                txtObjType.Text = "0"
                Exit Sub

            End If
            
            If txtObjType.Text = vbNullString Then Exit Sub
            If Not ObjCargado Then Call LoadOBJData
            ListadoQueQuiero = txtObjType.Text
            Call CargarListView(ListadoQueQuiero)

        Case "Npcs" 'Npcs
            
            If txtNpcs.Text = "" Then
                txtNpcs.Text = "0"
                Exit Sub

            End If

            If txtNpcs.Text = vbNullString Then Exit Sub
            ListadoQueQuiero = txtNpcs.Text

            If Not NpcsCargado Then Call CargaNpcsDat
            
            '            If Not IsNumeric(ListadoQueQuiero) Then
            '                Call CargarListviewNpc(txtNpcs.Text, ListadoQueQuiero)
            '            Else
            Call CargarListviewNpc(ListadoQueQuiero)

            '            End If
        Case "Nombre Npcs" 'Nombre Npcs

            If TxtNpcNombre.Text = "" Then
                TxtNpcNombre.Text = "¿Nombre?"
                Exit Sub

            End If
            
            If TxtNpcNombre.Text = vbNullString Then Exit Sub
            If Not NpcsCargado Then Call CargaNpcsDat
            ListadoQueQuiero = TxtNpcNombre.Text
            
            Call CargarListviewNpc(0, ListadoQueQuiero)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

    Dim Linea    As String

    Dim LineaNpc As String

    'Start subclassing
    InitSubClass
    
    'Enable label editing for listview2
    Set objLabelEdit = New LabelEdit
    objLabelEdit.Init Me, ListView
    
    Set objLabelEdit2 = New LabelEdit
    objLabelEdit2.Init Me, Listnpcs
    
    Open DatPath & "\TypoObjetos.dat" For Input As #1 'Abrimos el archivo en modo lectura

    Do While Not EOF(1) 'Recorremos todas las líneas hasta el final del archivo
        Line Input #1, Linea 'Leemos la sgte. línea y almacenamos en "Linea"
        'Aquí se puede trabajar con la línea
        Set subelemento = ListView1.ListItems.Add(, , Linea)
        
        'Carga la lista en el listbox
        '        GrhIndex = Linea
        '        Datos = ReadField(2, GrhIndex, 58)
        '        LstLista.AddItem (Datos)
            
    Loop

    Close #1 'Cerramos el archivo
    Open DatPath & "\TypoNpcs.dat" For Input As #2 'Abrimos el archivo en modo lectura

    Do While Not EOF(2) 'Recorremos todas las líneas hasta el final del archivo
        Line Input #2, LineaNpc 'Leemos la sgte. línea y almacenamos en "Linea"
        'Aquí se puede trabajar con la línea
        Set subelemento = ListtypoNpcs.ListItems.Add(, , LineaNpc)
        
        'Carga la lista en el listbox
        '        GrhIndex = Linea
        '        Datos = ReadField(2, GrhIndex, 58)
        '        LstLista.AddItem (Datos)
            
    Loop

    Close #2 'Cerramos el archivo

    If LstLista.Text = "" Then ListView1.Enabled = False

End Sub

Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'MsgBox "El ancho es: " & ColumnHeader.Width

End Sub

'Private Sub ListView_MouseDown(Button As Integer, _
'                                Shift As Integer, _
'                                x As Single, _
'                                y As Single)
'
'    Dim item As ListItem
'
'    Set item = ListView.HitTest(x, y)
'
'    If Not item Is Nothing And Button = vbRightButton Then
'        item.Selected = True
'        Me.PopupMenu mmuMenu
'
'    End If
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Stop subclassing
    CloseSubClass

    'Clean up by setting the classes to Nothing
    Set objLabelEdit = Nothing
    End

End Sub

Private Sub ListView_Click()

    Dim FicheroAom As String

    Dim Fichero    As String

    FicheroAom = "GraficosAom.ini"
    Fichero = "Graficos.ini"

    '    Dim xUltimo As String
    Dim Grh(1 To NumGrh) As Grh
    
    'si checkbox esta marcado se elije el aomania
    If cargaind = 1 Then
        'abrimos el fichero
        Call INI_Open(InitPath & "\", FicheroAom)
        Call INI_GetString(FicheroAom, "Graphics", "Grh1")
 
    Else
        'abrimos el fichero
        Call INI_Open(InitPath & "\", Fichero)
        Call INI_GetString(FicheroAom, "Graphics", "Grh1")

    End If

    'selecciono el GrhIndex del listView
    GrhIndex = ListView.SelectedItem.SubItems(160)

    If GrhIndex = "" Then Exit Sub

    ' recorremos todos los Graficos del fichero.
    If cargaind = 1 Then

        Data = INI_GetString(FicheroAom, "Graphics", "Grh" & GrhIndex)
        Grh(GrhIndex).anim = ReadField(1, Data, 58)
    Else
 
        Data = INI_GetString(Fichero, "Graphics", "Grh" & GrhIndex)
        Grh(GrhIndex).anim = ReadField(1, Data, 45)

    End If

    If Grh(GrhIndex).anim > 1 Then

        'si checkbox esta marcado se elije el aomania
        If cargaind = 1 Then
        
            GrhIndex = ReadField(2, Data, 58)
  
            Data = INI_GetString(FicheroAom, "Graphics", "Grh" & GrhIndex)
        Else
            GrhIndex = ReadField(2, Data, 45)
            Data = INI_GetString(Fichero, "Graphics", "Grh" & GrhIndex)

        End If

    End If

    'Muestro Informacion de Grh en el picture
    verpicture.pic.ToolTipText = ListView.SelectedItem.SubItems(1) & " - " & Data
    
    If Data = "" Then Exit Sub

    'si checkbox esta marcado se elije el aomania
    If cargaind = 1 Then
        Grh(GrhIndex).bmp = ReadField(2, Data, 58)
        Grh(GrhIndex).X = Val(ReadField(3, Data, 58))
        Grh(GrhIndex).Y = Val(ReadField(4, Data, 58))
        Grh(GrhIndex).ancho = Val(ReadField(5, Data, 58))
        Grh(GrhIndex).alto = Val(ReadField(6, Data, 58))
    Else
        Grh(GrhIndex).bmp = ReadField(2, Data, 45)
        Grh(GrhIndex).X = Val(ReadField(3, Data, 45))
        Grh(GrhIndex).Y = Val(ReadField(4, Data, 45))
        Grh(GrhIndex).ancho = Val(ReadField(5, Data, 45))
        Grh(GrhIndex).alto = Val(ReadField(6, Data, 45))

    End If

    Ruta = GrafPath & Grh(GrhIndex).bmp & ".bmp"
    verpicture.Show

    If (FileExist(Ruta, vbArchive)) Then
        
        Dim grafico As Picture
        
        Set grafico = LoadPicture(Ruta)
        
        Dim scaleX, scaleY As Integer
        
        If (Grh(GrhIndex).ancho = 32 And (Grh(GrhIndex).alto = Grh(GrhIndex).ancho)) Then
            scaleX = 32
            scaleY = 32
        ElseIf (Grh(GrhIndex).alto > 128 Or Grh(GrhIndex).ancho > 128) Then
            scaleX = 128
            scaleY = 128
        ElseIf (Grh(GrhIndex).alto = 64 And Grh(GrhIndex).ancho = 43) Then
            scaleX = 43
            scaleY = 64
        ElseIf (Grh(GrhIndex).alto = 32 Or Grh(GrhIndex).ancho = 63) Then
            scaleX = 64
            scaleY = 32
        ElseIf (Grh(GrhIndex).alto = 64 Or Grh(GrhIndex).ancho = 96) Then
            scaleX = 96
            scaleY = 64
        
        Else
            scaleX = Grh(GrhIndex).alto
            scaleY = Grh(GrhIndex).ancho

        End If

        verpicture.pic.Cls
        verpicture.pic.PaintPicture grafico, 0, 0, scaleX, scaleY, Grh(GrhIndex).X, Grh(GrhIndex).Y, Grh(GrhIndex).ancho, Grh(GrhIndex).alto
        verpicture.pic.Refresh

        'saber el ultimo item seleccionado
        '    If ListView1.ListItems.Count = 0 Then
        '
        '    Else
        '
        '        If xUltimo = "" Then
        '            xUltimo = ListView1.SelectedItem.Text
        '        Else
        '            MsgBox xUltimo
        '            xUltimo = ""
        '            xUltimo = ListView1.SelectedItem.Text
        '        End If
        '
        '    End If
    End If

End Sub

'CSEH: ErrHelios
Private Sub ListView1_Click()
    
10  On Error GoTo ListView1_Click_Err
        
20  GrhIndex = ListView1.ListItems(ListView1.SelectedItem.Index).Text
30  Datos = Val(ReadField(1, GrhIndex, 58))
40  txtObjType.Text = Datos
    
50  Exit Sub

ListView1_Click_Err:

    '60       Call LogError("ListView1_Click_Err" & " N: " & Err.Number & " D: " & Err.Description & " Linea del Error: " & Erl)

    '</EhFooter>
End Sub

Private Sub mnuGuardar_Click()
    Guardarobj.SaveOBJData

End Sub

Private Sub txtObjType_KeyPress(KeyAscii As Integer)

    If ListadoQueQuiero = "ObjType" Then
10      If KeyAscii <> 8 Then
20          If KeyAscii < 48 Or KeyAscii > 57 Then
30              KeyAscii = 0
        
40          End If

50      End If

    Else
        
    End If
    
End Sub

Private Sub txtNpcs_KeyPress(KeyAscii As Integer)

    If ListadoQueQuiero = "Npcs" Then
10      If KeyAscii <> 8 Then
20          If KeyAscii < 48 Or KeyAscii > 57 Then
30              KeyAscii = 0
        
40          End If

50      End If

    Else
        
    End If
    
End Sub
