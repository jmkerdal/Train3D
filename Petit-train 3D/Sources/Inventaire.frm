VERSION 5.00
Begin VB.Form Inventaire 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventaire des éléments du réseau"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "Inventaire.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Imprime 
      Caption         =   "&Imprimer"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ListBox ListeInventaire 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Inventaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' ***********************
' Mise à jour de la liste
' ***********************
'
Private Sub Form_Load()
    Dim i%, t As Boolean
    Me.Caption = Localisation$(CleINVENTAIRE%)
    Me.Imprime.Caption = Localisation$(CleINVENTAIRE% + 1)
    ListeInventaire.Clear
    '
    ' ***** Les voies utilisés
    '
    For i% = 1 To UBound(Voie())
        Voie(i%).Inventaire% = 0
    Next i%
    For i% = 1 To UBound(Reseau())
        If Reseau(i%).NoVoie% <> 0 Then
            Voie(Reseau(i%).NoVoie%).Inventaire% = Voie(Reseau(i%).NoVoie%).Inventaire% + 1
        End If
    Next i%
    For i% = 1 To UBound(Voie())
        If Voie(i%).Inventaire% <> 0 Then
            If t = False Then
                Call ListeInventaire.AddItem(Localisation$(CleEDITION%) + ":")
                t = True
                Imprime.Enabled = True
            End If
            Call ListeInventaire.AddItem("  " + Voie(i%).Ref + " :" + Str$(Voie(i%).Inventaire%) + "  (" + Voie(i%).Libelle$(0) + ")")
        End If
    Next i%
    '
    ' ***** Les décors utilisés
    '
    For i% = 1 To UBound(ListeDecor())
        ListeDecor(i%).Inventaire% = 0
    Next i%
    For i% = 1 To UBound(ElementDecor())
        If ElementDecor(i%).NoDecor% <> 0 Then
            ListeDecor(ElementDecor(i%).NoDecor%).Inventaire% = ListeDecor(ElementDecor(i%).NoDecor%).Inventaire% + 1
        End If
    Next i%
    t = False
    For i% = 1 To UBound(ListeDecor())
        If ListeDecor(i%).Inventaire% <> 0 Then
            If t = False Then
                Call ListeInventaire.AddItem(Localisation$(CleEDITION% + 1) + ":")
                t = True
                Imprime.Enabled = True
            End If
            Call ListeInventaire.AddItem("  " + ListeDecor(i%).Ref + " :" + Str$(ListeDecor(i%).Inventaire%) + "  (" + ListeDecor(i%).Nom$ + ")")
        End If
    Next i%
    '
    ' Taille de la planche
    '
    Call ListeInventaire.AddItem(Localisation$(CleINVENTAIRE% + 2) + ":" + Format$(xPlateauMax% - xPlateauMin% + 2 * BORD%) + "x" + Format$(zPlateauMax% - zPlateauMin% + 2 * BORD%) + " mm")
End Sub

'
' ***********************************
' Impression de la liste des éléments
' ***********************************
'
Private Sub Imprime_Click()
    Dim i%
    On Error GoTo Fin
    Printer.FontSize = 24
    Printer.FontBold = True
    Printer.Print "BASE: " + Localisation$(CleINVENTAIRE% + 3)
    Printer.FontBold = False
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 16
    Printer.Print Localisation$(CleINVENTAIRE% + 4) + ": " + CheminReseau$
    Printer.FontSize = 8
    Printer.Print
    Printer.Print Localisation$(CleINVENTAIRE% + 5) + ":"
    Printer.Print
    For i% = 0 To ListeInventaire.ListCount - 1
        Printer.Print ListeInventaire.List(i%)
    Next i%
    Printer.EndDoc
    Exit Sub
Fin:
    Call MsgBox("N°" + Format$(Err.Number) + vbCr + Err.Description, vbCritical + vbOKOnly, Localisation$(CleINVENTAIRE% + 6))
End Sub

