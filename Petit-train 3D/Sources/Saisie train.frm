VERSION 5.00
Begin VB.Form SaisieTrain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Composition du train"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "Saisie train.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   631
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Plus 
      Caption         =   "+"
      Enabled         =   0   'False
      Height          =   735
      Left            =   9000
      TabIndex        =   6
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Moins 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   735
      Left            =   9000
      TabIndex        =   5
      Top             =   240
      Width           =   375
   End
   Begin VB.Timer Rafraichir 
      Interval        =   100
      Left            =   4200
      Top             =   120
   End
   Begin VB.PictureBox VueWagon 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   3240
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.ListBox Selection 
      Height          =   2595
      Left            =   5760
      TabIndex        =   3
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton Retire 
      Caption         =   "<<"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Ajoute 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox Disponible 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "SaisieTrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelectionWagon%
Dim dxSaisieTrain As New ClassWagon

'
' *****************************
' Ajoute un wagon dans le train
' *****************************
'
Private Sub Ajoute_Click()
    Dim n%
    If Disponible.ListIndex < 0 Then Exit Sub
    n% = UBound(ListeTrain()) + 1
    ReDim Preserve ListeTrain(n%) As TypeTrain
    ListeTrain(n%).NoWagon% = Disponible.ListIndex + 1
    Call Affiche_Selection
End Sub

Private Sub Disponible_Click()
    If Disponible.ListIndex < 0 Then Exit Sub
    SelectionWagon% = Disponible.ListIndex + 1
    If ListeWagon(SelectionWagon%).dxWagon Is Nothing Then
        Ajoute.Enabled = False
    Else
        Ajoute.Enabled = True
    End If
    Moins.Enabled = False
    Plus.Enabled = False
End Sub

'
' **************************
' Charge la liste des wagons
' **************************
'
Private Sub Form_Load()
    Dim i%
    Me.Caption = Localisation$(CleSAISIETRAIN%)
    For i% = 1 To UBound(ListeWagon())
        Call Disponible.AddItem(ListeWagon(i%).Nom$)
    Next i%
    Call Affiche_Selection
End Sub

'
' ******************************************
' Charge la zone DX
' et decharge SaisieWagon si elle est active
' ******************************************
'
Private Sub Form_Resize()
    Call dxSaisieTrain.Charge(Me.VueWagon)
End Sub

'
' ********************
' Efface les objets DX
' ********************
'
Private Sub Form_Unload(Cancel As Integer)
    Set dxSaisieTrain = Nothing
End Sub

'
' ******************************
' Remonte un train dans la liste
' ******************************
'
Private Sub Moins_Click()
    Dim Temp As TypeTrain
    Dim TempTexte$
    Temp = ListeTrain(Selection.ListIndex + 1)
    ListeTrain(Selection.ListIndex + 1) = ListeTrain(Selection.ListIndex)
    ListeTrain(Selection.ListIndex) = Temp
    TempTexte$ = Selection.List(Selection.ListIndex)
    Selection.List(Selection.ListIndex) = Selection.List(Selection.ListIndex - 1)
    Selection.List(Selection.ListIndex - 1) = TempTexte$
    Selection.ListIndex = Selection.ListIndex - 1
End Sub

'
' ******************************
' Descend un train dans la liste
' ******************************
'
Private Sub Plus_Click()
    Dim Temp As TypeTrain
    Dim TempTexte$
    Temp = ListeTrain(Selection.ListIndex + 1)
    ListeTrain(Selection.ListIndex + 1) = ListeTrain(Selection.ListIndex + 2)
    ListeTrain(Selection.ListIndex + 2) = Temp
    TempTexte$ = Selection.List(Selection.ListIndex)
    Selection.List(Selection.ListIndex) = Selection.List(Selection.ListIndex + 1)
    Selection.List(Selection.ListIndex + 1) = TempTexte$
    Selection.ListIndex = Selection.ListIndex + 1
End Sub

'
' ****************************
' Affiche le wagon séléctionné
' ****************************
'
Private Sub Rafraichir_Timer()
    Call dxSaisieTrain.Rafraichir(SelectionWagon%)
End Sub

'
' *******************************
' Supprime un wagon dans le train
' *******************************
'
Private Sub Retire_Click()
    Dim i%, n%
    If Selection.ListIndex < 0 Then Exit Sub
    n% = UBound(ListeTrain()) - 1
    For i% = Selection.ListIndex + 1 To n%
        ListeTrain(i%) = ListeTrain(i% + 1)
    Next i%
    ReDim Preserve ListeTrain(n%) As TypeTrain
    Call Affiche_Selection
End Sub

'
' ***********************************************
' Affiche la liste des wagon constituant le train
' ***********************************************
'
Public Sub Affiche_Selection()
    Dim i%
    Call Selection.Clear
    For i% = 1 To UBound(ListeTrain())
        Call Selection.AddItem(ListeWagon(ListeTrain(i%).NoWagon%).Nom$)
    Next i%
End Sub

'
' **********************************
' Affiche le wagon dans la séléction
' **********************************
'
Private Sub Selection_Click()
    If Selection.ListIndex < 0 Then Exit Sub
    SelectionWagon% = ListeTrain(Selection.ListIndex + 1).NoWagon%
    If Selection.ListIndex = 0 Then
        Moins.Enabled = False
    Else
        Moins.Enabled = True
    End If
    If Selection.ListIndex = UBound(ListeTrain()) - 1 Then
        Plus.Enabled = False
    Else
        Plus.Enabled = True
    End If
End Sub

