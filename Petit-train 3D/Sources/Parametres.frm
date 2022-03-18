VERSION 5.00
Begin VB.Form Parametres 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parametres"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3735
   Icon            =   "Parametres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CCiel 
      Caption         =   "Ciel"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.CheckBox CSens 
      Caption         =   "Inverse sens de départ"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.VScrollBar SElevation 
      Height          =   1455
      Left            =   1440
      Max             =   9
      Min             =   1
      TabIndex        =   3
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin VB.CheckBox CCatenaire 
      Caption         =   "Catenaire"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox CElevation 
      Caption         =   "Elevation"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   1800
      Y1              =   1560
      Y2              =   0
   End
   Begin VB.Label LElevation 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "Parametres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' *************************
' Paramêtres divers
' utilisés pour l'affichage
' *************************
'
Option Explicit

'
' ************************
' Affichage des caténaires
' ************************
'
Private Sub CCatenaire_Click()
    If CCatenaire.Value = vbChecked Then
        ParamCatenaire = True
    Else
        ParamCatenaire = False
    End If
End Sub

'
' *****************
' Affichage du ciel
' *****************
'
Private Sub CCiel_Click()
    If CCiel.Value = vbChecked Then
        ParamCiel = True
    Else
        ParamCiel = False
    End If
End Sub

'
' **********************************
' Construit une élevation de terrain
' **********************************
'
Private Sub CElevation_Click()
    If CElevation.Value = vbChecked Then
        ParamElevation = True
        SElevation.Enabled = True
        LElevation.Enabled = True
    Else
        ParamElevation = False
        SElevation.Enabled = False
        LElevation.Enabled = False
    End If
End Sub

'
' **************************************
' Change la direction de départ du train
' **************************************
'
Private Sub CSens_Click()
    If CSens.Value = vbChecked Then
        ParamSens = True
    Else
        ParamSens = False
    End If
End Sub

'
' ******************************
' Affecte les valeurs par défaut
' ******************************
'
Private Sub Form_Load()
    Me.Caption = Localisation$(22)
    Call Me.Move(Screen.Width / 2, 0)
    CElevation.Caption = Localisation$(CleSAISIETRAIN% + 1)
    CCatenaire.Caption = Localisation$(CleSAISIETRAIN% + 2)
    CSens.Caption = Localisation$(CleSAISIETRAIN% + 3)
    CCiel.Caption = Localisation$(CleSAISIETRAIN% + 4)
    If ParamElevation = True Then
        CElevation.Value = vbChecked
        SElevation.Enabled = True
        LElevation.Enabled = True
    Else
        CElevation.Value = vbUnchecked
        SElevation.Enabled = False
        LElevation.Enabled = False
    End If
    
    SElevation.Value = 10 - ParamHauteur%
    LElevation.Caption = Str$(ParamHauteur% * 10) + "%"
    
    If ParamCatenaire = True Then
        CCatenaire.Value = vbChecked
    Else
        CCatenaire.Value = vbUnchecked
    End If
    If ParamSens = True Then
        CSens.Value = vbChecked
    Else
        CSens.Value = vbUnchecked
    End If
    If ParamCiel = True Then
        CCiel.Value = vbChecked
    Else
        CCiel.Value = vbUnchecked
    End If
    If Principale.ModeSaisie(0).Value = True Then
        Me.CCatenaire.Enabled = True
        Me.CElevation.Enabled = True
        Me.CSens.Enabled = True
        Me.SElevation.Enabled = True
    Else
        Me.CCatenaire.Enabled = False
        Me.CElevation.Enabled = False
        Me.CSens.Enabled = False
        Me.SElevation.Enabled = False
    End If
End Sub

'
' ************************************
' Modifie le % de variation du terrain
' ************************************
'
Private Sub SElevation_Change()
    ParamHauteur% = 10 - SElevation.Value
    LElevation.Caption = Str$(ParamHauteur% * 10) + "%"
End Sub

