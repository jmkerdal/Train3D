VERSION 5.00
Begin VB.Form FrmContour 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmContour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Show
    Call Module1.Charger("..\..\Fleischmann.rail")
    'Call Test1
    Call Test2
End Sub

Public Sub Test1()
    Dim i%
    Dim p(3) As D3DVECTOR
    Dim P0 As D3DVECTOR
    
    Dim n%, r%, c&
    
    P0.X = 0
    P0.z = 0
    For n% = 1 To UBound(Voie())
    'For n% = 23 To 24
        Cls
        Call FormeVoie.Raz
        Call FormeVoie.Cree_Couche(n%, 45, P0, 0)
        'Call FormeVoie.Cree_Couche(n%, 13)
        'Call FormeVoie.Cree_Couche(16, 13)
        For i% = 1 To FormeVoie.NbLRectangle%
            Call FormeVoie.Liste_Rectangle(i%, p(0), p(1), p(2), p(3))
            'Me.Line (p(0).X + 50, p(0).z + 100)-(p(1).X + 50, p(1).z + 100), RGB(0, 0, 255)
            'Me.Line (p(1).X + 50, p(1).z + 100)-(p(2).X + 50, p(2).z + 100), RGB(0, 0, 255)
            'Me.Line (p(2).X + 50, p(2).z + 100)-(p(3).X + 50, p(3).z + 100), RGB(0, 0, 255)
            'Me.Line (p(3).X + 50, p(3).z + 100)-(p(0).X + 50, p(0).z + 100), RGB(0, 0, 255)
        Next i%
        
        For i% = 1 To FormeVoie.NbLSegment%
            If FormeVoie.Liste_Segment(2, i%) <> -1 Then
'Debug.Print FormeVoie.Liste_Segment(0, i%); FormeVoie.Liste_Segment(1, i%)
                p(0) = FormeVoie.Liste_Point(FormeVoie.Liste_Segment(0, i%))
                p(1) = FormeVoie.Liste_Point(FormeVoie.Liste_Segment(1, i%))
                Me.Line (p(0).X + 50, p(0).z + 100)-(p(1).X + 50, p(1).z + 100), RGB(Rnd * 256, Rnd * 256, Rnd * 256)
                'Me.Line (p(0).X + 50, p(0).z + 100)-(p(1).X + 50, p(1).z + 100), RGB(0, 0, 255)
            End If
        Next i%
'Debug.Print

        If FormeVoie.Cree_Contour = False Then
            MsgBox "Erreur contour non crée"
        End If
        
        For i% = 1 To FormeVoie.NbNSegment%
            p(0) = FormeVoie.Nouveau_Point(FormeVoie.Nouveau_Segment(0, i%))
            p(1) = FormeVoie.Nouveau_Point(FormeVoie.Nouveau_Segment(1, i%))

            r% = (r% + 1) Mod 8
            Select Case r%
            Case 0
                c& = RGB(0, 0, 0)
            Case 1
                c& = RGB(0, 0, 255)
            Case 2
                c& = RGB(0, 255, 0)
            Case 3
                c& = RGB(0, 255, 255)
            Case 4
                c& = RGB(255, 0, 0)
            Case 5
                c& = RGB(255, 0, 255)
            Case 6
                c& = RGB(255, 255, 0)
            Case 7
                c& = RGB(255, 255, 255)
            End Select
            'Me.Line (p(0).X + 50, p(0).z + 250)-(p(1).X + 50, p(1).z + 250), RGB(Rnd * 256, Rnd * 256, Rnd * 256)
            'Me.Line (p(0).X + 50, p(0).z + 250)-(p(1).X + 50, p(1).z + 250), c&
            
            If FormeVoie.Nouveau_Segment(2, i%) <> -1 Then
                If FormeVoie.Nouveau_Segment(3, i%) = 1 Then
                    Me.Line (p(0).X + 50, p(0).z + 250)-(p(1).X + 50, p(1).z + 250), RGB(0, 0, 255)
                Else
                    If FormeVoie.Nouveau_Segment(4, i%) = 0 Then
                        Me.Line (p(0).X + 50, p(0).z + 250)-(p(1).X + 50, p(1).z + 250), RGB(0, 255, 0)
                    Else
                        Me.Line (p(0).X + 50, p(0).z + 250)-(p(1).X + 50, p(1).z + 250), RGB(255, 0, 255)
                    End If
                End If
            Else
                'Me.Line (p(0).X + 50, p(0).z + 250)-(p(1).X + 50, p(1).z + 250), RGB(255, 0, 0)
            End If
        Next i%
        MsgBox ("Suivant" + Str$(n%))
    Next n%
End Sub

Public Sub Test2()
    Dim i%
    Dim p(3) As D3DVECTOR
    Dim P0 As D3DVECTOR
    Dim Dir!
    
    Dim n%, r%, c&
    Dim f%
    Const REDUC% = 5
    Const OffX% = 100
    Const OffZ% = 50
    
    f% = FreeFile()
    Open "reseau.txt" For Input As #f%
    For i% = 1 To 26
        Input #f%, n%
        Input #f%, n%
        Input #f%, P0.X
        Input #f%, P0.z
        Input #f%, Dir!
        If i% <> 1 Then
        'If (i% >= 11 And i% <= 12) Or (i% >= 17 And i% <= 19) Or i% = 23 Then
        'If (i% >= 11 And i% <= 12) Or (i% >= 17 And i% <= 19) Then
            Call FormeVoie.Cree_Couche(n%, 45, P0, Dir!)
            'Call FormeVoie.Cree_Couche(n%, 13, P0, Dir!)
        End If
        
    Next i%
    
Debug.Print FormeVoie.NbLRectangle%
    For i% = 1 To FormeVoie.NbLRectangle%
        Call FormeVoie.Liste_Rectangle(i%, p(0), p(1), p(2), p(3))
        'Me.Line (p(0).X / REDUC% + OffX%, -p(0).z / REDUC% + OffZ%)-(p(1).X / REDUC% + OffX%, -p(1).z / REDUC% + OffZ%), RGB(0, 0, 255)
        'Me.Line (p(1).X / REDUC% + OffX%, -p(1).z / REDUC% + OffZ%)-(p(2).X / REDUC% + OffX%, -p(2).z / REDUC% + OffZ%), RGB(0, 0, 255)
        'Me.Line (p(2).X / REDUC% + OffX%, -p(2).z / REDUC% + OffZ%)-(p(3).X / REDUC% + OffX%, -p(3).z / REDUC% + OffZ%), RGB(0, 0, 255)
        'Me.Line (p(3).X / REDUC% + OffX%, -p(3).z / REDUC% + OffZ%)-(p(0).X / REDUC% + OffX%, -p(0).z / REDUC% + OffZ%), RGB(0, 0, 255)
    Next i%

'Debug.Print FormeVoie.NbLSegment%
    For i% = 1 To FormeVoie.NbLSegment%
        If FormeVoie.Liste_Segment(2, i%) <> -1 Then
'Debug.Print FormeVoie.Liste_Segment(0, i%); FormeVoie.Liste_Segment(1, i%)
            p(0) = FormeVoie.Liste_Point(FormeVoie.Liste_Segment(0, i%))
            p(1) = FormeVoie.Liste_Point(FormeVoie.Liste_Segment(1, i%))
            'Me.Line (p(0).X / REDUC% + OffX%, -p(0).z / REDUC% + OffZ%)-(p(1).X / REDUC% + OffX%, -p(1).z / REDUC% + OffZ%), RGB(Rnd * 256, Rnd * 256, Rnd * 256)
            Me.Line (p(0).X / REDUC% + OffX%, -p(0).z / REDUC% + OffZ%)-(p(1).X / REDUC% + OffX%, -p(1).z / REDUC% + OffZ%), RGB(0, 0, 255)
        End If
    Next i%

    If FormeVoie.Cree_Contour = False Then
        MsgBox "Erreur contour non crée"
        Exit Sub
    End If

'Debug.Print FormeVoie.NbNSegment%
    For i% = 1 To FormeVoie.NbNSegment%
        p(0) = FormeVoie.Nouveau_Point(FormeVoie.Nouveau_Segment(0, i%))
        p(1) = FormeVoie.Nouveau_Point(FormeVoie.Nouveau_Segment(1, i%))

        If FormeVoie.Nouveau_Segment(2, i%) <> -1 Then
            If FormeVoie.Nouveau_Segment(3, i%) = 1 Then
                c& = RGB(0, 0, 255)
            Else
                If FormeVoie.Nouveau_Segment(4, i%) = 0 Then
                    c& = RGB(0, 255, 0)
                Else
                    c& = RGB(255, 0, 255)
                End If
            End If
            Me.Line (p(0).X / REDUC% + OffX%, -p(0).z / REDUC% + 150 + OffZ%)-(p(1).X / REDUC% + OffX%, -p(1).z / REDUC% + 150 + OffZ%), c&
        Else
            c& = RGB(255, 0, 0)
            'Me.Line (p(0).X + 50, p(0).z + 250)-(p(1).X + 50, p(1).z + 250), c&
        End If
        'Me.Line (p(0).X / REDUC% + OffX%, -p(0).z / REDUC% + 150 + OffZ%)-(p(1).X / REDUC% + OffX%, -p(1).z / REDUC% + 150 + OffZ%), c&
    Next i%

End Sub

