VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' Options de sécurité de la clé de la base de registres.
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Types RACINE de la clé de la base de registes.
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Chaîne se terminant par un caractère nul Unicode.
Const REG_DWORD = 4                      ' Nombre 32 bits.

Private Sub Form_Activate()

    Const HKEY_LOCAL_MACHINE = &H80000002
    Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Train3D.exe"
    Const gREGVALSYSINFO = ""
    Dim SysInfoPath$
    Dim Retour#

    'Retour# = Shell(".\Setup.prg", vbNormalFocus)
    'AppActivate Retour#, True
    'AppActivate ".\Setup.prg", True
    
    '
    ' Attente de fin d'installation
    '
    Dim t1!
    Do
        t1! = Timer
        Do
            DoEvents
        Loop While Timer - t1! < 1
    Loop While GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) = False
    'Debug.Print SysInfoPath$
    
    'If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath$) = False Then
    '    MsgBox "Key not found", vbExclamation + vbOKOnly, "Train3D Install"
    'Else
        Dim i%, f%
        For i% = 1 To Len(SysInfoPath$)
            If Mid$(SysInfoPath$, i%, 1) = "\" Then f% = i%
        Next i%
        'Debug.Print Left$(SysInfoPath$, f%)
        
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        fs.CopyFolder ".\Petit-train 3D\*", Left$(SysInfoPath$, f%)
        Set fs = Nothing
        
        MsgBox "Files copied", vbInformation + vbOKOnly, "Train3D Install"
    'End If
    End
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
  Dim i As Long          ' Compteur de boucle.
  Dim rc As Long         ' Code de retour.
  Dim hKey As Long       ' Gestion vers la clé de la base de registres ouverte.
  Dim hDepth As Long
  Dim KeyValType As Long ' Type de données d'une clé de la base de registres.
  Dim tmpVal As String   ' Stockage temporaire pour une valeur de la clé de la base de registres.
  Dim KeyValSize As Long ' Taille de la variable de la clé de la base de registres.

  '------------------------------------------------------------
  ' Ouvre la clé de la base de registres sous la racine clé {HKEY_LOCAL_MACHINE...}
  '------------------------------------------------------------
  rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Ouvre la clé de la base de registres.
  
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gère l'erreur...
  
  tmpVal = String$(1024, 0)                               ' Alloue de l'espace à la variable.
  KeyValSize = 1024                                       ' Marque la taille de la variable.
  
  '------------------------------------------------------------
  ' Extrait la valeur de la clé de la base de registres...
  '------------------------------------------------------------
  rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                       KeyValType, tmpVal, KeyValSize)    ' Obtient/Crée la valeur de la clé.
                      
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gère les erreurs.
  
  If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then ' Win95 ajoute une chaîne se terminant par un caractère null...
      tmpVal = Left(tmpVal, KeyValSize - 1) ' Null a été trouvé, l'extrait de la chaîne
  Else ' WinNT n'ajoute PAS de chaîne se terminant par un carctère null...
      tmpVal = Left(tmpVal, KeyValSize) ' Null introuvable, extrait uniquement la chaîne
  End If
  '------------------------------------------------------------
  ' Détermine le type de valeur de la clé pour la conversion...
  '------------------------------------------------------------
  Select Case KeyValType   ' Recherche les types de données...
  Case REG_SZ              ' Type de données de la clé de la base de registres de la chaîne
      KeyVal = tmpVal      ' Copie la valeur de la chaîne
  Case REG_DWORD           ' Type de données de la clé de la base de registres de mots doubles
      For i = Len(tmpVal) To 1 Step -1 ' Convertit chaque bit
          KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Construit la valeur caractère par caractère
      Next
      KeyVal = Format$("&h" + KeyVal) ' Convertit le mot double en chaîne
  End Select
  
  GetKeyValue = True       ' Renvoie une opération réussie
  rc = RegCloseKey(hKey)   ' Ferme la clé de la base de registres
  Exit Function            ' Quitte la fonction
  
GetKeyError:               ' Efface après qu'une erreur se soit produite...
  KeyVal = ""              ' Définit la valeur de retour avec une chaîne vide
  GetKeyValue = False      ' Renvoie un échec
  rc = RegCloseKey(hKey)   ' Ferme la clé de la base de registres
End Function

