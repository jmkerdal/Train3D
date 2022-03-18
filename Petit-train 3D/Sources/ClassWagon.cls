VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassWagon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Dim dxEdite As New ClassDirectX
Dim dxEditeRoue(3) As Direct3DRMFrame3
Dim dxEditeBogie(1) As Direct3DRMFrame3
Dim dxAtteles(1) As Direct3DRMFrame3
Dim Angle!

'
' *******************
' Création de la zone
' *******************
'
Public Sub Charge(Zone As PictureBox)
    Call dxEdite.Create_3DRM(Zone, Zone.ScaleWidth, Zone.ScaleHeight, Mode3DSurface)
    Call dxEdite.dxScene.SetSceneBackgroundRGB(0, 0, 0.5)
    With dxEdite.dxViewport
        .SetFront 1 / REDUCTION%
        .SetBack 10000! / REDUCTION%
        .SetField 0.5 / REDUCTION%
    End With
End Sub

'
' **********************
' Charge les éléments DX
' **********************
'
Private Sub Class_Initialize()
    Dim i%
    For i% = 0 To 3
        Set dxEditeRoue(i%) = dxEdite.dxD3Drm.CreateFrame(Nothing)
        Call dxEditeRoue(i%).SetOrientation(dxEdite.dxScene, 1, 0, 0, 0, 1, 0)
        Call dxEditeRoue(i%).AddVisual(dxRoue)
    Next i%
    For i% = 0 To 1
        Set dxAtteles(i%) = dxEdite.dxD3Drm.CreateFrame(Nothing)
        Call dxAtteles(i%).AddVisual(dxAttele)
        Set dxEditeBogie(i%) = dxEdite.dxD3Drm.CreateFrame(Nothing)
    Next i%
    Call dxAtteles(0).SetOrientation(dxEdite.dxScene, 1, 0, 0, 0, 1, 0)
    Call dxAtteles(1).SetOrientation(dxEdite.dxScene, -1, 0, 0, 0, 1, 0)
    For i% = 0 To 1
        Set dxEditeBogie(i%) = dxEdite.dxD3Drm.CreateFrame(Nothing)
    Next i%
End Sub

'
' *************************
' Destruction des objets DX
' *************************
'
Private Sub Class_Terminate()
    Dim i%
    For i% = 0 To 3
        Call dxEditeRoue(i%).DeleteVisual(dxRoue)
        Set dxEditeRoue(i%) = Nothing
    Next i%
    For i% = 0 To 1
        Call dxAtteles(i%).DeleteVisual(dxAttele)
        Set dxAtteles(i%) = Nothing
    Next i%
    For i% = 0 To 1
        Set dxEditeBogie(i%) = Nothing
    Next i%
    Set dxEdite = Nothing
End Sub

'
' ***************************
' Affiche le wagon qui tourne
' ***************************
'
Public Sub Rafraichir(n%)
    Dim box As D3DRMBOX
    Dim Recul!, dX!, dY!, dz!
    Dim P1!
    Dim t1&, t2&
    t1& = timeGetTime
    If n% > 0 Then
        If Not (ListeWagon(n%).dxWagon Is Nothing) Then
            Call dxEdite.dxScene.AddVisual(ListeWagon(n%).dxWagon)
            Call ListeWagon(n%).dxWagon.GetBox(box)
            dX! = box.Max.X - box.Min.X
            dY! = box.Max.Y - box.Min.Y
            dz! = box.Max.z - box.Min.z
            Recul! = dX!
            If dY! > Recul! Then Recul! = dY!
            If dz! > Recul! Then Recul! = dz!
            Recul! = Recul! * 1.25
            '
            ' ***** Pose les roues
            '
            Call dxEdite.dxScene.AddChild(dxEditeRoue(0))
            Call dxEditeRoue(0).SetPosition(dxEdite.dxScene, 0, -20 / REDUCTION%, ListeWagon(n%).EccartBogie!(0) / 2 / REDUCTION%)
            If ListeWagon(n%).EccartBogie(0) <> 0 Then
                Call dxEdite.dxScene.AddChild(dxEditeRoue(1))
                Call dxEditeRoue(1).SetPosition(dxEdite.dxScene, 0, -20 / REDUCTION%, -ListeWagon(n%).EccartBogie!(0) / 2 / REDUCTION%)
            End If
            Call dxEdite.dxScene.AddChild(dxEditeRoue(2))
            Call dxEditeRoue(2).SetPosition(dxEdite.dxScene, 0, -20 / REDUCTION%, (-ListeWagon(n%).EccartEssieu! - ListeWagon(n%).EccartBogie!(1) / 2) / REDUCTION%)
            If ListeWagon(n%).EccartBogie(1) <> 0 Then
                Call dxEdite.dxScene.AddChild(dxEditeRoue(3))
                Call dxEditeRoue(3).SetPosition(dxEdite.dxScene, 0, -20 / REDUCTION%, (-ListeWagon(n%).EccartEssieu! + ListeWagon(n%).EccartBogie!(1) / 2) / REDUCTION%)
            End If
            '
            ' ***** Pose les atteles
            '
            P1! = ListeWagon(n%).Longueur! - ListeWagon(n%).EccartEssieu! - (ListeWagon(n%).EccartBogie!(0) + ListeWagon(n%).EccartBogie!(1)) / 2
            Call dxEdite.dxScene.AddChild(dxAtteles(0))
            Call dxAtteles(0).SetPosition(dxEditeRoue(0), -P1! / 2 / REDUCTION%, 10 / REDUCTION%, 0)
            Call dxEdite.dxScene.AddChild(dxAtteles(1))
            Call dxAtteles(1).SetPosition(dxEditeRoue(2), P1! / 2 / REDUCTION%, 10 / REDUCTION%, 0)
            '
            ' ***** Pose les bloc de bogies
            '
            If ListeWagon(n%).FichierBogie$ <> "" Then
                Call dxEditeBogie(0).AddVisual(ListeWagon(n%).dxBogie)
                Call dxEdite.dxScene.AddChild(dxEditeBogie(0))
                Call dxEditeBogie(0).SetPosition(dxEditeRoue(0), ListeWagon(n%).EccartBogie!(0) / 2 / REDUCTION%, 10 / REDUCTION%, 0)
                Call dxEditeBogie(0).SetOrientation(dxEditeRoue(0), 0, 0, 1, 0, 1, 0)
                '
                Call dxEditeBogie(1).AddVisual(ListeWagon(n%).dxBogie)
                Call dxEdite.dxScene.AddChild(dxEditeBogie(1))
                Call dxEditeBogie(1).SetPosition(dxEditeRoue(3), ListeWagon(n%).EccartBogie!(1) / 2 / REDUCTION%, 10 / REDUCTION%, 0)
                Call dxEditeBogie(1).SetOrientation(dxEditeRoue(3), 0, 0, 1, 0, 1, 0)
            End If
        Else
            Recul! = 100 / REDUCTION%
        End If
    Else
        Recul! = 100 / REDUCTION%
    End If
    Call dxEdite.dxCamera.SetPosition(dxEdite.dxScene, dX! / 2 + box.Min.X, dY! / 2 + box.Min.Y, dz! / 2 + box.Min.z)
    Call dxEdite.dxCamera.SetOrientation(dxEdite.dxScene, 0, 0, 1, 0, 1, 0)
    Call dxEdite.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 0, 1, 0, Angle!)
    Call dxEdite.dxCamera.AddTranslation(D3DRMCOMBINE_BEFORE, 0, 0, -Recul!)
    '
    Call dxEdite.Render(False)
    Call dxEdite.Render(True)
    '
    If n% > 0 Then
        If Not (ListeWagon(n%).dxWagon Is Nothing) Then
            Call dxEdite.dxScene.DeleteChild(dxEditeRoue(0))
            If ListeWagon(n%).EccartBogie(0) <> 0 Then
                Call dxEdite.dxScene.DeleteChild(dxEditeRoue(1))
            End If
            Call dxEdite.dxScene.DeleteChild(dxEditeRoue(2))
            If ListeWagon(n%).EccartBogie(1) <> 0 Then
                Call dxEdite.dxScene.DeleteChild(dxEditeRoue(3))
            End If
            Call dxEdite.dxScene.DeleteVisual(ListeWagon(n%).dxWagon)
            Call dxEdite.dxScene.DeleteChild(dxAtteles(0))
            Call dxEdite.dxScene.DeleteChild(dxAtteles(1))
            If ListeWagon(n%).FichierBogie$ <> "" Then
                Call dxEdite.dxScene.DeleteChild(dxEditeBogie(0))
                Call dxEditeBogie(0).DeleteVisual(ListeWagon(n%).dxBogie)
                Call dxEdite.dxScene.DeleteChild(dxEditeBogie(1))
                Call dxEditeBogie(1).DeleteVisual(ListeWagon(n%).dxBogie)
            End If
        End If
    End If
    '
    t2& = timeGetTime
    'Angle! = Angle! + PI! / 16
    Angle! = Angle! + PI! * (t2& - t1&) / 1000
    If Angle! > 2 * PI! Then
        Angle! = Angle! - 2 * PI!
    End If
End Sub

