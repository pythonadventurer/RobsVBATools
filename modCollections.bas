Attribute VB_Name = "modCollections"
Option Compare Database
Option Explicit

'Sorts the given collection using the Arrays.MergeSort algorithm.
' O(n log(n)) time
' O(n) space
Public Sub sort(col As Collection, Optional ByRef c As IVariantComparator)
    Dim a() As Variant
    Dim b() As Variant
    a = Collections.ToArray(col)
    Arrays.sort a(), c
    Set col = Collections.FromArray(a())
End Sub

'Returns an array which exactly matches this collection.
' Note: This function is not safe for concurrent modification.
Public Function ToArray(col As Collection) As Variant
    Dim a() As Variant
    ReDim a(0 To col.Count)
    Dim i As Long
    For i = 0 To col.Count - 1
        a(i) = col(i + 1)
    Next i
    ToArray = a()
End Function

'Returns a Collection which exactly matches the given Array
' Note: This function is not safe for concurrent modification.
Public Function FromArray(a() As Variant) As Collection
    Dim col As Collection
    Set col = New Collection
    Dim element As Variant
    For Each element In a
        col.Add element
    Next element
    Set FromArray = col
End Function
Sub testSort()

Dim myCollection As New Collection
Dim colSorted As New Collection
Dim colItem As Variant

myCollection.Add ("Tables")
myCollection.Add ("Queries")
myCollection.Add ("Forms")
myCollection.Add ("Reports")
myCollection.Add ("Modules")


Collections.sort myCollection

End Sub
