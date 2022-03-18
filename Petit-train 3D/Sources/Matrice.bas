Attribute VB_Name = "Matrice"
Option Explicit

Sub TranslateMatrix(m As D3DMATRIX, v As D3DVECTOR)
    Call dxVue.dX7.IdentityMatrix(m)
    m.rc41 = v.x
    m.rc42 = v.y
    m.rc43 = v.z
End Sub

Sub CreateXRotation(ret As D3DMATRIX, rads As Single)
    Dim cosine As Single
    Dim sine As Single
    cosine = Cos(rads)
    sine = Sin(rads)
    
    Call dxVue.dX7.IdentityMatrix(ret) ' Method of the DirectX7 object.
    
    ret.rc22 = cosine
    ret.rc23 = sine
    ret.rc32 = -sine
    ret.rc33 = cosine
End Sub

'
' *******************************
' Test si un OctObjet est visible
' *******************************
'
Public Function Test_Visible(Frame As Direct3DRMFrame3, Mesh As Direct3DRMMeshBuilder3) As Boolean
'Test_Visible = True
'Exit Function
    Dim dxBox As D3DRMBOX ' Boite du mesh
    Dim VueBox As D3DRMBOX ' Coordonnées projetés
    Dim OctObjet As D3DVECTOR
    Dim OctObjet4D As D3DRMVECTOR4D
    Dim Tv As D3DVECTOR
    Dim i%
    '
    ' ***** Transforme un D3DRMBOX en OctObjet
    '
    Call Mesh.GetBox(dxBox)
    VueBox.MAX.x = -1: VueBox.MAX.y = -1
    VueBox.MIN.x = TailleX% + 1: VueBox.MIN.y = TailleY% + 1
    For i% = 0 To 7
        If i% = 0 Then
            OctObjet.x = dxBox.MIN.x: OctObjet.y = dxBox.MAX.y: OctObjet.z = dxBox.MAX.z
        ElseIf i% = 1 Then
            OctObjet.x = dxBox.MAX.x: OctObjet.y = dxBox.MAX.y: OctObjet.z = dxBox.MAX.z
        ElseIf i% = 2 Then
            OctObjet.x = dxBox.MAX.x: OctObjet.y = dxBox.MAX.y: OctObjet.z = dxBox.MIN.z
        ElseIf i% = 3 Then
            OctObjet.x = dxBox.MIN.x: OctObjet.y = dxBox.MAX.y: OctObjet.z = dxBox.MIN.z
        ElseIf i% = 4 Then
            OctObjet.x = dxBox.MIN.x: OctObjet.y = dxBox.MIN.y: OctObjet.z = dxBox.MAX.z
        ElseIf i% = 5 Then
            OctObjet.x = dxBox.MAX.x: OctObjet.y = dxBox.MIN.y: OctObjet.z = dxBox.MAX.z
        ElseIf i% = 6 Then
            OctObjet.x = dxBox.MAX.x: OctObjet.y = dxBox.MIN.y: OctObjet.z = dxBox.MIN.z
        Else
            OctObjet.x = dxBox.MIN.x: OctObjet.y = dxBox.MIN.y: OctObjet.z = dxBox.MIN.z
        End If
        '
        ' ***** Applique les transformations
        '
        Call Frame.Transform(Tv, OctObjet)
        Call dxVue.dxViewport.Transform(OctObjet4D, Tv)
        If OctObjet4D.w > 0 Then
            OctObjet4D.x = OctObjet4D.x / OctObjet4D.w
            OctObjet4D.y = OctObjet4D.y / OctObjet4D.w
            If OctObjet4D.x > VueBox.MAX.x Then VueBox.MAX.x = OctObjet4D.x
            If OctObjet4D.x < VueBox.MIN.x Then VueBox.MIN.x = OctObjet4D.x
            If OctObjet4D.y > VueBox.MAX.y Then VueBox.MAX.y = OctObjet4D.y
            If OctObjet4D.y < VueBox.MIN.y Then VueBox.MIN.y = OctObjet4D.y
        End If
    Next i%
    If VueBox.MAX.x >= 0 Then
        If VueBox.MIN.x <= TailleX% Then
            If VueBox.MAX.y >= 0 Then
                If VueBox.MIN.y <= TailleY% Then
                    Test_Visible = True
                End If
            End If
        End If
    End If
End Function

