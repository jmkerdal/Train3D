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

' Options de s�curit� de la cl� de la base de registres.
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

' Types RACINE de la cl� de la base de registes.
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cha�ne se terminant par un caract�re nul Unicode.
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
  Dim hKey As Long       ' Gestion vers la cl� de la base de registres ouverte.
  Dim hDepth As Long
  Dim KeyValType As Long ' Type de donn�es d'une cl� de la base de registres.
  Dim tmpVal As String   ' Stockage temporaire pour une valeur de la cl� de la base de registres.
  Dim KeyValSize As Long ' Taille de la variable de la cl� de la base de registres.

  '------------------------------------------------------------
  ' Ouvre la cl� de la base de registres sous la racine cl� {HKEY_LOCAL_MACHINE...}
  '------------------------------------------------------------
  rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Ouvre la cl� de la base de registres.
  
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' G�re l'erreur...
  
  tmpVal = String$(1024, 0)                               ' Alloue de l'espace � la variable.
  KeyValSize = 1024                                       ' Marque la taille de la variable.
  
  '------------------------------------------------------------
  ' Extrait la valeur de la cl� de la base de registres...
  '------------------------------------------------------------
  rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                       KeyValType, tmpVal, KeyValSize)    ' Obtient/Cr�e la valeur de la cl�.
                      
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' G�re les erreurs.
  
  If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then ' Win95 ajoute une cha�ne se terminant par un caract�re null...
      tmpVal = Left(tmpVal, KeyValSize - 1) ' Null a �t� trouv�, l'extrait de la cha�ne
  Else ' WinNT n'ajoute PAS de cha�ne se terminant par un carct�re null...
      tmpVal = Left(tmpVal, KeyValSize) ' Null introuvable, extrait uniquement la cha�ne
  End If
  '------------------------------------------------------------
  ' D�termine le type de valeur de la cl� pour la conversion...
  '------------------------------------------------------------
  Select Case KeyValType   ' Recherche les types de donn�es...
  Case REG_SZ              ' Type de donn�es de la cl� de la base de registres de la cha�ne
      KeyVal = tmpVal      ' Copie la valeur de la cha�ne
  Case REG_DWORD           ' Type de donn�es de la cl� de la base de registres de mots doubles
      For i = Len(tmpVal) To 1 Step -1 ' Convertit chaque bit
          KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Construit la valeur caract�re par caract�re
      Next
      KeyVal = Format$("&h" + KeyVal) ' Convertit le mot double en cha�ne
  End Select
  
  GetKeyValue = True       ' Renvoie une op�ration r�ussie
  rc = RegCloseKey(hKey)   ' Ferme la cl� de la base de registres
  Exit Function            ' Quitte la fonction
  
GetKeyError:               ' Efface apr�s qu'une erreur se soit produite...
  KeyVal = ""              ' D�finit la valeur de retour avec une cha�ne vide
  GetKeyValue = False      ' Renvoie un �chec
  rc = RegCloseKey(hKey)   ' Ferme la cl� de la base de registres
End Function

