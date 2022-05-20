Attribute VB_Name = "modobjetos"
Option Explicit

Public Sub DatosPack(nfile As Integer)

    On Error GoTo ErrHandler
    
    Print #nfile, "'-----------------------------------------------------------------------------"
    Print #nfile, "'------------------Pack Premium (Dat by Lugus, Programado por Helios)---------"
    Print #nfile, "'-----------------------------(8€/12€)--(Pago Paypal)-------------------------"

    Exit Sub

ErrHandler:

End Sub

Public Sub DatosfinPack(nfile As Integer)

    On Error GoTo ErrHandler
    
    Print #nfile, "'--------------------------------------------------------------------------------"
    Print #nfile, "'-----------------Fin Pack Premium (Dat by Lugus, Programado por Helios)---------"
    Print #nfile, "'--------------------------------------------------------------------------------"
   
    Exit Sub

ErrHandler:

End Sub

Public Sub DatosObj(nfile As Integer)

    On Error GoTo ErrHandler
    
    Print #nfile, "'            TIPOS DE OBJETOS"
    Print #nfile, "' 1: Comida"
    Print #nfile, "' 2: Armas"
    Print #nfile, "' 3: Armaduras"
    Print #nfile, "' 4: Arboles"
    Print #nfile, "' 5: Dinero"
    Print #nfile, "' 6: Puertas"
    Print #nfile, "' 7: Objetos contenedores (por ejemplo bolsas y cofres)"
    Print #nfile, "' 8: Carteles"
    Print #nfile, "' 9: Llaves"
    Print #nfile, "' 10: Foros"
    Print #nfile, "' 11: Pociones"
    Print #nfile, "' 12: Libros"
    Print #nfile, "' 13: Bebida"
    Print #nfile, "' 14: Leña"
    Print #nfile, "' 15: Fogata"
    Print #nfile, "' 16: escudos"
    Print #nfile, "' 17: cascos"
    Print #nfile, "' 18: Herramientas"
    Print #nfile, "' 19: telep"
    Print #nfile, "' 20: Muebles"
    Print #nfile, "' 21: joyas"
    Print #nfile, "' 22: yacimiento"
    Print #nfile, "' 23: metales"
    Print #nfile, "' 24: pergaminos"
    Print #nfile, "' 25: Cheques 10k"
    Print #nfile, "' 26: Instrumentos Musicales"
    Print #nfile, "' 27: Yunque"
    Print #nfile, "' 28: Fragua"
    Print #nfile, "' 31: barcos"
    Print #nfile, "' 32: flechas"
    Print #nfile, "' 33: botellas vacias"
    Print #nfile, "' 34: botellas llenas"
    Print #nfile, "' 35: manchas"
    Print #nfile, "' 36: Pocis que Resucitan"
    Print #nfile, "' 37: Alas"
    Print #nfile, "' 41: Pasajes"
    Print #nfile, "' 47: Amuleto Teletransporte"
    Print #nfile, "' 50: Vales de experencia"
    Print #nfile, "' 77: Plata"
    Print #nfile, "' 1000: Cualquiera"
    Print #nfile, ""
    Print #nfile, "'<--------------------Sub Tipos--------------------------->"
    Print #nfile, "' 0: Armaduras"
    Print #nfile, "' 1: Cascos"
    Print #nfile, "' 2: Escudos"
    Print #nfile, "'<--------------------Sub Tipos--------------------------->"
    Print #nfile, "'<--------------------Tipos De Pieles--------------------------->"
    Print #nfile, "'Lobo="
    Print #nfile, "'Osos="
    Print #nfile, "'Lana="
    Print #nfile, "'Tigre="
    Print #nfile, "'Jabali="
    Print #nfile, "'LoboPolar="
    Print #nfile, "'OsoPolar="
    Print #nfile, "'Vaca="
    Print #nfile, "'<--------------------Tipos De Pieles--------------------------->"
    Print #nfile, ""
    Print #nfile, "'<--------------------Tipo de posicones------------------->"
    Print #nfile, "'1 Modifica la Agilidad"
    Print #nfile, "'2 Modifica la Fuerza"
    Print #nfile, "'3 Repone HP"
    Print #nfile, "'4 Repone Mana"
    Print #nfile, "'5 Cura Envenenamiento"
    Print #nfile, "'6 Desparalizante"
    Print #nfile, "'7 Invisible (Impedimiento)"
    Print #nfile, "'8 Telepatia (Impedimiento)"
    Print #nfile, "'9 Teletransporte a Nix (Impedimiento)"
    Print #nfile, "'10 Energia"
    Print #nfile, "'11 Anti Cegera (Impedimiento)"
    Print #nfile, "'12 Anti estupidez (Impedimiento)"
    Print #nfile, "'13 Teletransporte a Ulla (Impedimiento)"
    Print #nfile, "'14 Teletransporte a Bander (Impedimiento)"
    Print #nfile, "'15 Gran Pocion Azul"
    Print #nfile, "'16 Cambio de Dia"
    Print #nfile, "'Nota: Las numeraciones que tienen (Impedimiento) al lado son lo que puse sus funcionamiento, se deben agregar MinSkill= (0 al 60) y Cae=1."
    Print #nfile, "'<--------------------Tipo de posicones------------------->"
    Print #nfile, "'<--------------------Objetos Especiales------------------>"
    Print #nfile, "' 1 33% Ahorro flechas"
    Print #nfile, "' 2 +5 Fuerza"
    Print #nfile, "' 3 +2 Fuerza"
    Print #nfile, "' 4 +3 Fuerza"
    Print #nfile, "' 5 +5 Agilidad"
    Print #nfile, "' 6 +2 Agilidad"
    Print #nfile, "' 7 +3 Agilidad"
    Print #nfile, "' 8 +100 Mana"
    Print #nfile, "' 9 +200 Mana"
    Print #nfile, "' 10 +300 Mana"
    Print #nfile, "' 50 (2-8) Recuperacion Mana"
    Print #nfile, "' 51 (8-15) Recuperacion Vida"
    Print #nfile, "' 53 50% Ahorro Flechas"
    Print #nfile, "' 54 75% Ahorro Flechas"
    Print #nfile, "' 55 (5-7) Agilidad y Fuerza"
    Print #nfile, "'<--------------------Objetos Especiales------------------>"
    Print #nfile, "'<--------------------RAZAS---------------------->"
    Print #nfile, "'Humano"
    Print #nfile, "'Elfo"
    Print #nfile, "'Elfo Oscuro"
    Print #nfile, "'Gnomo"
    Print #nfile, "'Enano"
    Print #nfile, "'Orco"
    Print #nfile, "'Hobbit"
    Print #nfile, "'Ciclope"
    Print #nfile, "'Vampiros"
    Print #nfile, "'Licantropo"
    Print #nfile, "'<--------------------RAZAS---------------------->"
    Print #nfile, "'<--------------------Clases--------------------->"
    Print #nfile, "'CP1=MAGO"
    Print #nfile, "'CP2=CLERIGO"
    Print #nfile, "'CP3=GUERRERO"
    Print #nfile, "'CP4=ASESINO"
    Print #nfile, "'CP5=LADRON"
    Print #nfile, "'CP6=BARDO"
    Print #nfile, "'CP7=DRUIDA"
    Print #nfile, "'CP8=PALADIN"
    Print #nfile, "'CP9=PIRATA"
    Print #nfile, "'CP10=TRABAJADOR"
    Print #nfile, "'CP11=BRUJO"
    Print #nfile, "'CP12=ARQUERO"
    Print #nfile, "'CP13=BANDIDO"
    Print #nfile, "'CP14=DOMADOR"
    Print #nfile, "' <---------------------Clases----------------------->"
    Print #nfile, "'<----------------------Armas con efectos----------------->"
    Print #nfile, "'Paraliza= probabvilidad 1 a 100 100 paraliza siempre"
    Print #nfile, "'<----------------------Armas con efectos----------------->"
    Print #nfile, ""
    Print #nfile, "'<----------------------Armaduras con efectos----------------->"
    Print #nfile, "'Resistencia= probabvilidad 1 a 100 100 Resiste siempre"
    Print #nfile, "'<----------------------Armaduras con efectos----------------->"
    Print #nfile, ""

    Exit Sub

ErrHandler:

End Sub

Public Sub Espacio(nfile As Integer)
   
    On Error GoTo ErrHandler
    
    Print #nfile, " "

    Exit Sub

ErrHandler:

End Sub

