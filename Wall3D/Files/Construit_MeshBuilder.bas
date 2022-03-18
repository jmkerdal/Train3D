Attribute VB_Name = "dxConstruitMeshBuilder"
'
' **************************************
' Meshbuilder automatic generation tools
' Kerdal Jean-Michel © 1998-2000
' **************************************
'
Option Explicit
'
' ***** Face definition
'
Enum EnumSensTexture
    SENS_TRIANGLE
    SENS_TRIANGLE2
    SENS_RECTANGLE
End Enum
'
Private Type TypeFace
    Texture As Integer
    SensTexture As EnumSensTexture
    Couleur As Long
End Type
'
Dim NbPoints&
Dim Points() As D3DVECTOR
Dim NbFaces&
Dim Faces&()
Dim NbTextFaces%
Dim TextFaces() As TypeFace
'
Dim FaceArray As Direct3DRMFaceArray
Dim Face As Direct3DRMFace2
Dim aNormals(0) As D3DVECTOR

'
' ***********
' Add a point
' ***********
'
Private Function Add_Point&(p As D3DVECTOR)
    Points(NbPoints&) = p
    Add_Point& = NbPoints&
    NbPoints& = NbPoints& + 1
    ReDim Preserve Points(NbPoints&) As D3DVECTOR
End Function

'
' ***************
' Add a rectangle
' ***************
'
Public Sub Add_Rectangle(P1 As D3DVECTOR, P2 As D3DVECTOR, P3 As D3DVECTOR, P4 As D3DVECTOR, NoText%, NoCouleur&)
    NbFaces& = NbFaces& + 5
    ReDim Preserve Faces&(NbFaces&)
    Faces&(NbFaces& - 5) = 4
    Faces&(NbFaces& - 4) = Add_Point&(P1)
    Faces&(NbFaces& - 3) = Add_Point&(P2)
    Faces&(NbFaces& - 2) = Add_Point&(P3)
    Faces&(NbFaces& - 1) = Add_Point&(P4)
    NbTextFaces% = NbTextFaces% + 1
    ReDim Preserve TextFaces(NbTextFaces%) As TypeFace
    TextFaces(NbTextFaces%).SensTexture = SENS_RECTANGLE
    TextFaces(NbTextFaces%).Texture = NoText%
    TextFaces(NbTextFaces%).Couleur = NoCouleur&
End Sub

'
' **************
' Add a triangle
' **************
'
Public Sub Add_Triangle(P1 As D3DVECTOR, P2 As D3DVECTOR, P3 As D3DVECTOR, NoText%, NoCouleur&)
    NbFaces& = NbFaces& + 4
    ReDim Preserve Faces&(NbFaces&)
    Faces&(NbFaces& - 4) = 3
    Faces&(NbFaces& - 3) = Add_Point&(P1)
    Faces&(NbFaces& - 2) = Add_Point&(P2)
    Faces&(NbFaces& - 1) = Add_Point&(P3)
    NbTextFaces% = NbTextFaces% + 1
    ReDim Preserve TextFaces(NbTextFaces%) As TypeFace
    TextFaces(NbTextFaces%).SensTexture = SENS_TRIANGLE
    TextFaces(NbTextFaces%).Texture = NoText%
    TextFaces(NbTextFaces%).Couleur = NoCouleur&
End Sub

'
' **************************
' Apply the color for a face
' **************************
'
Private Sub Apply_Color(f%, Color&)
    Set Face = FaceArray.GetElement(f%)
    Call Face.SetColor(Color&)
End Sub

'
' *****************************
' Apply a texture on a triangle
' or a rectangle
' *****************************
'
Private Sub Apply_Texture(f%, Sens As EnumSensTexture, TheTexture As Direct3DRMTexture3)
    If TheTexture Is Nothing Then Exit Sub
    Set Face = FaceArray.GetElement(f%)
    Select Case Sens
    Case EnumSensTexture.SENS_TRIANGLE
        Call Face.SetTextureCoordinates(0, 0, 1)
        Call Face.SetTextureCoordinates(1, 0, 0)
        Call Face.SetTextureCoordinates(2, 1, 0)
    Case EnumSensTexture.SENS_TRIANGLE2
        Call Face.SetTextureCoordinates(0, 0, 0)
        Call Face.SetTextureCoordinates(1, 1, 0)
        Call Face.SetTextureCoordinates(2, 1, 1)
    Case EnumSensTexture.SENS_RECTANGLE
        Call Face.SetTextureCoordinates(0, 0, 1)
        Call Face.SetTextureCoordinates(1, 0, 0)
        Call Face.SetTextureCoordinates(2, 1, 0)
        Call Face.SetTextureCoordinates(3, 1, 1)
    End Select
    Call Face.SetTexture(TheTexture)
End Sub

'
' **************************************************
' Build the mesh with the information found in array
' **************************************************
'
Public Sub Build(MeshBuilder As Direct3DRMMeshBuilder3, Textures() As Direct3DRMTexture3)
    Dim i%
    '
    ' ***** Add points and faces
    '
    Set FaceArray = MeshBuilder.AddFaces(NbPoints&, Points(), 0, aNormals(), Faces&())
    '
    ' ***** Perspective-corrected
    '
    Call MeshBuilder.SetPerspective(D_TRUE)
    '
    ' ***** Apply textures
    '
    If FaceArray.GetSize <> 0 Then
        For i% = 0 To NbTextFaces%
            Call Apply_Color(i%, TextFaces(i%).Couleur&)
            If TextFaces(i%).Texture% <> -1 Then
                Call Apply_Texture(i%, TextFaces(i%).SensTexture, Textures(TextFaces(i%).Texture%))
            End If
        Next i%
    End If
    Call MeshBuilder.Optimize
End Sub

'
' **************
' Reset the mesh
' **************
'
Public Sub Init()
    NbPoints& = 0 ' List of points
    ReDim Points(NbPoints&) As D3DVECTOR
    NbFaces& = 0 ' Faces create with points
    ReDim Faces&(NbFaces&)
    NbTextFaces% = -1 ' Faces textures
    ReDim TextFaces(0) As TypeFace
End Sub

'
' *************
' Add a polygon
' *************
'
Public Sub Add_Polygon(Polygone() As D3DVECTOR, NoText%, NoCouleur&)
    Dim n%, i%
    n% = UBound(Polygone())
    Faces&(NbFaces&) = n%
    NbFaces& = NbFaces& + n% + 1
    ReDim Preserve Faces&(NbFaces&)
    For i% = 1 To n%
        Faces&(NbFaces& - (n% - i% + 1)) = Add_Point&(Polygone(i%))
    Next i%
    NbTextFaces% = NbTextFaces% + 1
    ReDim Preserve TextFaces(NbTextFaces%) As TypeFace
    TextFaces(NbTextFaces%).SensTexture = SENS_RECTANGLE
    TextFaces(NbTextFaces%).Texture = NoText%
    TextFaces(NbTextFaces%).Couleur = NoCouleur&
End Sub

