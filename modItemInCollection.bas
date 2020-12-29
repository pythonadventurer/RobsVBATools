Attribute VB_Name = "modItemInCollection"
Option Compare Database
Option Explicit

Function ItemInCollection(varItem As Variant, varColl As Collection) As Boolean

Dim varCheckItem As Variant
Dim Found As Boolean

Found = False

For Each varCheckItem In varColl
    If varCheckItem = varItem Then
        Found = True
    End If

Next varCheckItem

ItemInCollection = Found

End Function

