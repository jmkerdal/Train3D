VERSION 5.00
Begin VB.Form Generation_Tunnel 
   AutoRedraw      =   -1  'True
   Caption         =   "Génération de la forme du tunnel"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
End
Attribute VB_Name = "Generation_Tunnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' ***********************************************
' Création et affichage de la liste des triangles
' et trace les bords
' ***********************************************
'
Private Sub Form_Activate()
    Dim i%, n%
    Dim i1%
    Dim c&
    '
    If LeTunnel.Genere = True Then
        n% = LeTunnel.Nb_Triangle%()
        For i% = 0 To n% - 1
            c& = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
            Me.Line (LeTunnel.TriangleX(0, i%), -LeTunnel.TriangleZ(0, i%))-(LeTunnel.TriangleX(1, i%), -LeTunnel.TriangleZ(1, i%)), c&
            Me.Line (LeTunnel.TriangleX(1, i%), -LeTunnel.TriangleZ(1, i%))-(LeTunnel.TriangleX(2, i%), -LeTunnel.TriangleZ(2, i%)), c&
            Me.Line (LeTunnel.TriangleX(2, i%), -LeTunnel.TriangleZ(2, i%))-(LeTunnel.TriangleX(0, i%), -LeTunnel.TriangleZ(0, i%)), c&
            MsgBox "Suivant"
        Next i%
    End If
    '
    n% = LeTunnel.Nb_Point%()
    For i% = 0 To n% - 1
        i1% = (i% + 1) Mod n%
        Me.Line (LeTunnel.PositionX(i1%), -LeTunnel.PositionZ(i1%))-(LeTunnel.PositionX(i1%) + LeTunnel.NormalX(i1%) * 10, -LeTunnel.PositionZ(i1%) - LeTunnel.NormalZ(i1%) * 10), RGB(255, 255, 255)
        If LeTunnel.Face(i%) = True Then
            Me.Line (LeTunnel.PositionX(i%) + LeTunnel.NormalX(i%) * 10, -LeTunnel.PositionZ(i%) - LeTunnel.NormalZ(i%) * 10)- _
                    (LeTunnel.PositionX(i1%) + LeTunnel.NormalX(i1%) * 10, -LeTunnel.PositionZ(i1%) - LeTunnel.NormalZ(i1%) * 10), RGB(0, 0, 0)
        End If
    Next i%
End Sub

