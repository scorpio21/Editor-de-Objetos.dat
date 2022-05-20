VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabelEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private NumObj    As Integer

Private ColumnObj As Integer

Private Type tPos
    
    Button As Integer
    X As Integer
    Y As Integer
            
End Type

Private Pos As tPos

'===============================================================
'ListView LabelEdit
'© 2004 by Michiel Meulendijk
'
'This code enables label editing for SubItems in ListViews.
'By default, when using the LabelEdit property, only the first
'ListItem of a ListView can be edited. With this code all
'ListSubItems can be edited as well.
'
'This code is contained within a class, so multiple instances
'can run at the same time (e.g. more ListViews on one form can
'all support label editing).
'
'This file is provided "as is" with no expressed or implied
'warranty. The author accepts no liability for any damage caused
'to your system because of using this code.
'===============================================================

Private Declare Function GetScrollInfo _
                Lib "user32.dll" (ByVal Hwnd As Long, _
                                  ByVal n As Long, _
                                  lpScrollInfo As SCROLLINFO) As Long

Private Type SCROLLINFO

    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long

End Type

Private Const SB_HORZ = 0

Private Const SB_VERT = 1

Private Const SIF_POS = &H4

Dim WithEvents txtEdit     As TextBox
Attribute txtEdit.VB_VarHelpID = -1

Dim WithEvents ltvListView As ListView
Attribute ltvListView.VB_VarHelpID = -1

Dim objItem                As Object

Public Sub Init(ByRef ctlForm As Form, ByRef ctlListView As ListView)
    'Initiates object. Adds textbox control.
    Set ltvListView = ctlListView
    Set txtEdit = ctlForm.Controls.Add("VB.TextBox", "txtLabelEdit_" & ctlListView.Name)
    Set txtEdit.Container = ctlListView.Container
    Set txtEdit.Font = ctlListView.Font
    txtEdit.Appearance = 0
    txtEdit.Forecolor = vbRed 'vbHighlight
    'Subclass listview
    SubClassWnd ltvListView.Hwnd, Me

End Sub

Private Function GetHorizontalScroll() As Long

    'Returns the position of the horizontal scroll bar
    Dim scrInfo As SCROLLINFO

    scrInfo.cbSize = LenB(scrInfo)
    scrInfo.fMask = SIF_POS
    GetScrollInfo ltvListView.Hwnd, SB_HORZ, scrInfo
    GetHorizontalScroll = scrInfo.nPos

End Function

Private Sub EditText(ByVal X As Integer, ByVal Y As Integer)

    'Handles label editing
    On Error GoTo endSub

    Dim i As Integer, objCol As ColumnHeader, lngScroll As Long

    lngScroll = GetHorizontalScroll * Screen.TwipsPerPixelX
    X = X + lngScroll

    For i = 1 To ltvListView.ColumnHeaders.Count

        If X < ltvListView.ColumnHeaders.item(1).Width Or ltvListView.ColumnHeaders.Count = 1 Then
            Set objCol = ltvListView.ColumnHeaders.item(1)
            Set objItem = ltvListView.SelectedItem
            Exit For
        ElseIf X < ltvListView.ColumnHeaders.item(i).Left Then
            Set objCol = ltvListView.ColumnHeaders.item(i - 1)
            Set objItem = ltvListView.SelectedItem.ListSubItems.item(i - 2)
            NumObj = ltvListView.SelectedItem
            ColumnObj = (i - 2)
            Exit For
        ElseIf i = ltvListView.ColumnHeaders.Count Then
            Set objCol = ltvListView.ColumnHeaders(i)
            Set objItem = ltvListView.SelectedItem.ListSubItems.item(i - 1)
            NumObj = ltvListView.SelectedItem
            ColumnObj = (i - 1)
            Exit For

        End If

    Next i

    txtEdit.BorderStyle = 0
    txtEdit.Left = ltvListView.Left + objCol.Left - lngScroll
    txtEdit.Top = ltvListView.Top + ltvListView.SelectedItem.Top
    txtEdit.Width = objCol.Width
    txtEdit.Height = ltvListView.SelectedItem.Height
    txtEdit.BorderStyle = 1

    txtEdit.Text = objItem.Text
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit)
    txtEdit.Visible = True
    txtEdit.SetFocus

endSub:

End Sub

Public Sub SetText()

    On Error Resume Next

    objItem.Text = txtEdit.Text
    Call GuardarValoresList(NumObj, ColumnObj, objItem.Text)
    txtEdit.Visible = False

End Sub

Public Sub GuardarValoresList(ByVal Object As Integer, _
                              ByVal Columna As Byte, _
                              ByVal Datos As String)
    
    If Object = 0 Then Exit Sub
    Debug.Print Columna

    Select Case Columna

        Case 1 'Nombre
            ObjData(Object).Name = Datos

        Case 2 'GrhIndex
            ObjData(Object).ObjType = Datos

        Case 3 'Valor
            ObjData(Object).SubTipo = Datos

        Case 4 'Subtipo
            ObjData(Object).MaxHit = Datos

        Case 5
            ObjData(Object).MinHit = Datos

        Case 6
            ObjData(Object).MinDef = Datos

        Case 7
            ObjData(Object).MaxDef = Datos

        Case 8
            ObjData(Object).MinHam = Datos

        Case 9
            ObjData(Object).Valor = Datos

        Case 10
            ObjData(Object).Resistencia = Datos

        Case 11
            ObjData(Object).Vendible = Datos

        Case 12
            ObjData(Object).SkSastreria = Datos

        Case 13
            ObjData(Object).SkHechizeria = Datos

        Case 14
            ObjData(Object).ObjHierba = Datos

        Case 15
            ObjData(Object).MinSkill = Datos

        Case 16
            ObjData(Object).Velocidad = Datos

        Case 17
            ObjData(Object).Facc = Datos

        Case 18
            ObjData(Object).objetoespecial = Datos

        Case 19
            ObjData(Object).MinSta = Datos

        Case 20
            ObjData(Object).Caos = Datos

        Case 21
            ObjData(Object).Real = Datos

        Case 22
            ObjData(Object).templ = Datos

        Case 23
            ObjData(Object).Nemes = Datos

        Case 24
            ObjData(Object).Cae = Datos

        Case 25
            ObjData(Object).TiRaRObj = Datos

        Case 26
            ObjData(Object).Robable = Datos

        Case 27
            ObjData(Object).bonifica = Datos

        Case 28
            ObjData(Object).tipobonifica = Datos

        Case 29
            ObjData(Object).Donacion = Datos

        Case 30
            ObjData(Object).Newbie = Datos

        Case 31
            ObjData(Object).Destruir = Datos

        Case 32
            ObjData(Object).ShieldAnim = Datos

        Case 33
            ObjData(Object).LingH = Datos

        Case 34
            ObjData(Object).LingP = Datos

        Case 35
            ObjData(Object).LingO = Datos

        Case 36
            ObjData(Object).LingM = Datos

        Case 37
            ObjData(Object).Madera = Datos

        Case 38
            ObjData(Object).MaderaElfica = Datos

        Case 39
            ObjData(Object).SkHerreria = Datos

        Case 40
            ObjData(Object).Drena = Datos

        Case 41
            ObjData(Object).SkHerreriaMagica = Datos

        Case 42
            ObjData(Object).Gemas = Datos

        Case 43
            ObjData(Object).Diamantes = Datos

        Case 44
            ObjData(Object).CascoAnim = Datos

        Case 45
            ObjData(Object).DefensaMagicaMax = Datos

        Case 46
            ObjData(Object).DefensaMagicaMin = Datos

        Case 47
            ObjData(Object).Botas = Datos

        Case 48
            ObjData(Object).Alas = Datos

        Case 49
            ObjData(Object).Ropaje = Datos

        Case 50
            ObjData(Object).HechizoIndex = Datos

        Case 51
            ObjData(Object).Clase = Datos

        Case 52
            ObjData(Object).WeaponAnim = Datos

        Case 53
            ObjData(Object).Apuñala = Datos

        Case 54
            ObjData(Object).Paraliza = Datos

        Case 55
            ObjData(Object).Ceguera = Datos

        Case 56
            ObjData(Object).Estupidez = Datos

        Case 57
            ObjData(Object).Vida = Datos

        Case 58
            ObjData(Object).Mana = Datos

        Case 59
            ObjData(Object).Envenena = Datos

        Case 60
            ObjData(Object).LvlMin = Datos

        Case 61
            ObjData(Object).LvlMax = Datos

        Case 62
            ObjData(Object).Heroe = Datos

        Case 63
            ObjData(Object).proyectil = Datos

        Case 64
            ObjData(Object).Municion = Datos

        Case 65
            ObjData(Object).TipoProyectil = Datos

        Case 66
            ObjData(Object).DosManos = Datos

        Case 67
            ObjData(Object).Nivel = Datos

        Case 68
            ObjData(Object).GameMaster = Datos

        Case 69
            ObjData(Object).UsoNpc = Datos

        Case 70
            ObjData(Object).Snd1 = Datos

        Case 71
            ObjData(Object).Snd2 = Datos

        Case 72
            ObjData(Object).Snd3 = Datos

        Case 73
            ObjData(Object).MinInt = Datos

        Case 74
            ObjData(Object).LingoteIndex = Datos

        Case 75
            ObjData(Object).sagrado = Datos

        Case 76
            ObjData(Object).MineralIndex = Datos

        Case 77
            ObjData(Object).MaxHP = Datos

        Case 78
            ObjData(Object).MinHP = Datos

        Case 79
            ObjData(Object).Mujer = Datos

        Case 80
            ObjData(Object).Hombre = Datos

        Case 81
            ObjData(Object).MinSed = Datos

        Case 82
            ObjData(Object).Respawn = Datos

        Case 83
            ObjData(Object).RazaEnana = Datos

        Case 84
            ObjData(Object).RazaElfa = Datos

        Case 85
            ObjData(Object).RazaVampiro = Datos

        Case 86
            ObjData(Object).RazaOrco = Datos

        Case 87
            ObjData(Object).ClaseAsesino = Datos

        Case 88
            ObjData(Object).RazaHumana = Datos

        Case 89
            ObjData(Object).RazaHobbit = Datos

        Case 90
            ObjData(Object).Expe = Datos

        Case 91
            ObjData(Object).Skill = Datos

        Case 92
            ObjData(Object).Crucial = Datos

        Case 93
            ObjData(Object).Cerrada = Datos

        Case 94
            ObjData(Object).Llave = Datos

        Case 95
            ObjData(Object).CLAVE = Datos

        Case 96
            ObjData(Object).IndexAbierta = Datos

        Case 97
            ObjData(Object).IndexCerrada = Datos

        Case 98
            ObjData(Object).IndexCerradaLlave = Datos

        Case 99
            ObjData(Object).texto = Datos

        Case 100
            ObjData(Object).GrhSecundario = Datos

        Case 101
            ObjData(Object).Agarrable = Datos

        Case 102
            ObjData(Object).ForoID = Datos

        Case 103
            ObjData(Object).Acuchilla = Datos

        Case 104
            ObjData(Object).Guante = Datos

        Case 105
            ObjData(Object).NoSubasta = Datos

        Case 106
            ObjData(Object).NoRegalo = Datos

        Case 107
            ObjData(Object).TipoPocion = Datos

        Case 108
            ObjData(Object).MaxModificador = Datos

        Case 109
            ObjData(Object).MinModificador = Datos

        Case 110
            ObjData(Object).DuracionEfecto = Datos

        Case 111
            ObjData(Object).TipoRegalo = Datos

        Case 112
            ObjData(Object).SkCarpinteria = Datos

        Case 113
            ObjData(Object).Lana = Datos

        Case 114
            ObjData(Object).Lobos = Datos

        Case 115
            ObjData(Object).Osos = Datos

        Case 116
            ObjData(Object).Tigre = Datos

        Case 117
            ObjData(Object).OsoPolar = Datos

        Case 118
            ObjData(Object).Vaca = Datos

        Case 119
            ObjData(Object).Jabali = Datos

        Case 120
            ObjData(Object).Zona = Datos

        Case 121
            ObjData(Object).TipoMunicion = Datos

    End Select
    
End Sub

Private Sub Class_Terminate()
    UnSubClassWnd ltvListView.Hwnd
    Set txtEdit = Nothing
    Set ltvListView = Nothing

End Sub

Private Sub ltvListView_Click()

    If Pos.Button = vbRightButton Then
        EditText Pos.X, Pos.Y

    End If

End Sub

Private Sub ltvListView_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
     
    Pos.Button = Button
    Pos.X = X
    Pos.Y = Y
    'EditText x, y

End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then SetText

End Sub

Private Sub txtEdit_LostFocus()
    SetText

End Sub
