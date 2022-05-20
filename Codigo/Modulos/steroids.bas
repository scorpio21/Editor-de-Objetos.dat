Attribute VB_Name = "steroids"
Option Explicit

Public Declare Sub AbreFINI _
               Lib "aomania.dll" (ByVal Archivo As String, _
                                  Optional ByVal Crear As Byte = 0)

Public Declare Function LeeFINI Lib "aomania.dll" (ByVal CLAVE As String) As String

Public Declare Sub SelectFINI Lib "aomania.dll" (ByVal Seccion As String)

Public Declare Function GetVar _
               Lib "aomania.dll" (ByVal Archivo As String, _
                                  ByVal Main As String, _
                                  ByVal Var As String) As String

Public Declare Sub INI_Free Lib "AOSteroids.dll" (ByVal file As String)

Public Declare Sub INI_Open _
               Lib "AOSteroids.dll" (ByVal path As String, _
                                     ByVal file As String, _
                                     Optional ordered As Boolean = False)

Public Declare Function INI_GetString _
               Lib "AOSteroids.dll" (ByVal file As String, _
                                     ByVal key As String, _
                                     ByVal subkey As String) As String

Public Declare Function INI_GetInteger _
               Lib "AOSteroids.dll" (ByVal file As String, _
                                     ByVal key As String, _
                                     ByVal subkey As String) As Integer
                                     
Public Declare Sub FP_Append _
               Lib "aomania.dll" (ByVal Archivo As String, _
                                  ByVal Text As String)

Public Declare Sub INI_SetValue _
               Lib "AOSteroids.dll" (ByVal file As String, _
                                     ByVal key As String, _
                                     ByVal subkey As String, _
                                     ByVal value As String)

Sub Main()

    Call ChDir(App.path)
    Call ChDrive(App.path & "\libs\")

    DatPath = App.path & "\Dats\"
    InitPath = App.path & "\Init"
    GrafPath = App.path & "\Graficos\"

    'verpicture.Show
    FrmMain.Show

End Sub
