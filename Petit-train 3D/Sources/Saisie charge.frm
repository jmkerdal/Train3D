VERSION 5.00
Begin VB.Form SaisieCharge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saisie des éléments chargés"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "Saisie charge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox LWagon 
      Height          =   1410
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1680
      Width           =   4455
   End
   Begin VB.ListBox LDecor 
      Height          =   1410
      ItemData        =   "Saisie charge.frx":014A
      Left            =   120
      List            =   "Saisie charge.frx":0151
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "SaisieCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' ******************************************
' Affiche la liste des éléments selectionnés
' ******************************************
'
Private Sub Form_Load()
    Dim i%
    Me.Caption = Localisation$(CleEDITION% + 21)
    LDecor.Clear
    For i% = 1 To UBound(ListeDecor())
        Call LDecor.AddItem(ListeDecor(i%).Nom)
        If Not (ListeDecor(i%).dxDecor Is Nothing) Then
            LDecor.Selected(i% - 1) = True
        End If
    Next i%
    LDecor.ListIndex = 0
    '
    LWagon.Clear
    For i% = 1 To UBound(ListeWagon())
        Call LWagon.AddItem(ListeWagon(i%).Nom)
        If Not (ListeWagon(i%).dxWagon Is Nothing) Then
            LWagon.Selected(i% - 1) = True
        End If
    Next i%
    LWagon.ListIndex = 0
End Sub

'
' ***************************
' Charge les éléments en plus
' ***************************
'
Private Sub Form_Unload(Cancel As Integer)
    Dim i%
    For i% = 1 To UBound(ListeDecor())
        If LDecor.Selected(i% - 1) = True Then
            If ListeDecor(i%).dxDecor Is Nothing Then
                Call Initialisation.Décor_Charge_Mesh(i%)
            End If
        End If
    Next i%
    '
    For i% = 1 To UBound(ListeWagon())
        If LWagon.Selected(i% - 1) = True Then
            If ListeWagon(i%).dxWagon Is Nothing Then
                Call Initialisation.Wagon_Charger_Mesh(i%)
            End If
        End If
    Next i%
End Sub

