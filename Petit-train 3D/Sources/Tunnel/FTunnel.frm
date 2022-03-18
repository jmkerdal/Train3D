VERSION 5.00
Begin VB.Form FTunnel 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   Caption         =   "Generation d'un tunnel"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MENU_Edition 
      Caption         =   "Edition"
      Visible         =   0   'False
      Begin VB.Menu MENU_Ajoute 
         Caption         =   "Ajoute"
      End
      Begin VB.Menu MENU_Type 
         Caption         =   "Inverse le type"
      End
      Begin VB.Menu MENU_Supprime 
         Caption         =   "Supprime"
      End
      Begin VB.Menu MENU_Moins0 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_Generation 
         Caption         =   "Génération"
      End
   End
End
Attribute VB_Name = "FTunnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' ********************
' Active le menu popup
' ********************
'
Private Sub Form_DblClick()
    If LeTunnel.SegmentPointe% <> -1 Then
        If LeTunnel.Nb_Point% > 3 Then
            MENU_Supprime.Enabled = True
        Else
            MENU_Supprime.Enabled = False
        End If
        PopupMenu MENU_Edition
    End If
End Sub

'
' ****
' Init
' ****
'
Private Sub Form_Load()
    'Call MTunnel.Init
    Call LeTunnel.Point_Ajoute(20, -20)
    Call LeTunnel.Point_Ajoute(60, -20)
    Call LeTunnel.Point_Ajoute(60, -60)
    Call LeTunnel.Point_Ajoute(20, -60)
    '
    Call Aff_Tunnel
End Sub

'
' **************
' Bouge un point
' **************
'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If LeTunnel.PointSelection% = -1 Then
            LeTunnel.PointSelection% = LeTunnel.PointPointe%
        Else
            Call LeTunnel.Point_Deplace(LeTunnel.PointSelection%, X, -Y)
            LeTunnel.PointSelection% = -1
        End If
        Call Aff_Tunnel
    End If
End Sub

'
' *************************************
' Cherche le point et le segment pointé
' *************************************
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LeTunnel.Cherche_Pointe(X, Y)
    Call Aff_Tunnel
End Sub

'
' ***************
' Ajoute un point
' ***************
'
Private Sub MENU_Ajoute_Click()
    Call LeTunnel.Point_Insere
    Call Aff_Tunnel
End Sub

'
' *****************************
' Génération du tunnel a partir
' de la liste des points
' *****************************
'
Private Sub MENU_Generation_Click()
    Generation_Tunnel.Show vbModal
End Sub

'
' ***************
' Ajoute un point
' ***************
'
Private Sub MENU_Supprime_Click()
    Call LeTunnel.Point_Supprime
    Call Aff_Tunnel
End Sub

'
' **********************
' Change le type du coté
' **********************
'
Private Sub MENU_Type_Click()
    Call LeTunnel.Inverse
    Call Aff_Tunnel
End Sub

'
' **********************************
' Trace la forme du tunnel à l'écran
' **********************************
'
Private Sub Aff_Tunnel()
    Dim i%, n%
    Dim c&
    FTunnel.Cls
    n% = LeTunnel.Nb_Point%()
    For i% = 0 To n% - 1
        If LeTunnel.Face(i%) = True Then
            If LeTunnel.SegmentPointe% = i% Then
                c& = RGB(255, 255, 0) ' Jaune
            Else
                c& = RGB(255, 255, 255) ' Blanc
            End If
        Else
            If LeTunnel.SegmentPointe% = i% Then
                c& = RGB(255, 128, 64) ' Orange
            Else
                c& = RGB(64, 64, 64) ' Gris foncé
            End If
        End If
        FTunnel.Line (LeTunnel.PositionX(i%), -LeTunnel.PositionZ(i%))-(LeTunnel.PositionX((i% + 1) Mod n%), -LeTunnel.PositionZ((i% + 1) Mod n%)), c&
    Next i%
    For i% = 0 To n% - 1
        If i% = LeTunnel.PointSelection% Then
            c& = RGB(255, 0, 0) ' Rouge
        Else
            If LeTunnel.PointPointe% = i% Then
                c& = RGB(0, 255, 255)  ' Cyan
            Else
                c& = RGB(0, 0, 255)  ' Bleu
            End If
        End If
        FTunnel.Line (LeTunnel.PositionX(i%) - 3, -LeTunnel.PositionZ(i%) - 3)-(LeTunnel.PositionX(i%) + 3, -LeTunnel.PositionZ(i%) + 3), c&, B
    Next i%
End Sub

