Attribute VB_Name = "Cryptage"
'
' **************************************
' Cryptage module for saved data on disk
' © Kerdal Jean-Michel 1998-2000
' **************************************
'
Option Explicit
Dim Key% ' Mask key for saved data
Global CheckSum& ' Checksum for saved data

'
' ************************************
' Return the hexa value with n numbers
' ************************************
'
Private Function Hexa$(Number As Variant, Format%)
    Hexa$ = Hex$(Number)
    If Len(Hexa$) < Format% Then
        If Number < 0 Then
            Hexa$ = String$(Format% - Len(Hexa$), "F") & Hexa$
        Else
            Hexa$ = String$(Format% - Len(Hexa$), "0") & Hexa$
        End If
    End If
    If Len(Hexa$) > Format% Then Hexa$ = Right$(Hexa$, Format%)
End Function

'
' *********************
' Calculate the new key
' *********************
'
Private Sub Key_Change()
    Key% = (1 + 5 * Key%) Mod 256
End Sub

'
' *********************
' Load a chain caracter
' *********************
'
Function Load_Chain$(canal%)
    Dim n%, i%
    n% = Load_Integer%(canal%)
    Load_Chain$ = ""
    For i% = 1 To n%
        Load_Chain$ = Load_Chain$ + Chr$(Load_Byte(canal%))
    Next i%
End Function

'
' ************************
' Load integer (two bytes)
' ************************
'
Function Load_Integer%(canal%)
    Dim a$
    a$ = Hexa$(Load_Byte(canal%), 2)
    a$ = a$ + Hexa$(Load_Byte(canal%), 2)
    Load_Integer% = Val("&h" & a$)
End Function

'
' **********************
' Load long (four bytes)
' **********************
'
Function Load_Long&(canal%)
    Dim a$
    a$ = Hexa$(Load_Byte(canal%), 2)
    a$ = a$ + Hexa$(Load_Byte(canal%), 2)
    a$ = a$ + Hexa$(Load_Byte(canal%), 2)
    a$ = a$ + Hexa$(Load_Byte(canal%), 2)
    Load_Long& = Val("&h" & a$)
End Function

'
' *********
' Load byte
' *********
'
Function Load_Byte(canal%) As Byte
    Load_Byte = Asc(Input$(1, #canal%))
    If Key% <> -1 Then
        Load_Byte = Load_Byte Xor Key%
        Call Key_Change
    End If
    CheckSum& = (CheckSum& + Load_Byte) Mod 2 ^ 16
End Function

'
' *********************
' Init the new mask key
' *********************
'
Sub Key_Init(n%)
    Key% = n%
    CheckSum& = 0
End Sub

'
' ************************
' Save integer (two bytes)
' ************************
'
Sub Save_Integer(canal%, ByVal n%)
    Dim a$
    a$ = Hexa$(n%, 4)
    Call Save_Byte(canal%, Val("&h" & Mid$(a$, 1, 2)))
    Call Save_Byte(canal%, Val("&h" & Mid$(a$, 3, 2)))
End Sub

'
' **********************
' Save long (four bytes)
' **********************
'
Sub Save_Long(canal%, n&)
    Dim a$
    a$ = Hexa$(n&, 8)
    Call Save_Byte(canal%, Val("&h" & Mid$(a$, 1, 2)))
    Call Save_Byte(canal%, Val("&h" & Mid$(a$, 3, 2)))
    Call Save_Byte(canal%, Val("&h" & Mid$(a$, 5, 2)))
    Call Save_Byte(canal%, Val("&h" & Mid$(a$, 7, 2)))
End Sub

'
' ********************
' Save caractere chain
' ********************
'
Sub Save_Chain(canal%, c$)
    Dim n%
    Call Save_Integer(canal%, Len(c$))
    For n% = 1 To Len(c$)
        Call Save_Byte(canal%, Asc(Mid$(c$, n%, 1)))
    Next n%
End Sub

'
' *********
' Save byte
' *********
'
Sub Save_Byte(canal%, ByVal n As Byte)
    CheckSum& = (CheckSum& + n) Mod 2 ^ 16
    If Key% <> -1 Then
        n = n Xor Key%
        Call Key_Change
    End If
    Print #canal%, Chr$(n);
End Sub

'
' *********************
' Save a number to file
' *********************
'
Sub Save_Number(canal%, n As Variant, Retour As Boolean)
    If Retour = True Then
        Print #canal%, " " + Format$(n)
    Else
        Print #canal%, " " + Format$(n);
    End If
End Sub

'
' ************************
' Test validity of the key
' ************************
'
Public Function Key_Test(canal%) As Boolean
    Dim ChekSum_Value&
    Dim Key_Value&
    ChekSum_Value& = CheckSum& ' Memorise because it will change when we load the real key
    Key_Value& = Load_Long&(canal%)
    If Key_Value& < 0 Then Key_Value& = Key_Value& + 2 ^ 16
    If ChekSum_Value& = Key_Value& Then
        Key_Test = True
    Else
        Key_Test = False
    End If
End Function

