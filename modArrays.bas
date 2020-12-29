Attribute VB_Name = "modArrays"
Option Compare Text
Option Explicit
Option Base 0

Private Const INSERTIONSORT_THRESHOLD As Long = 7

'Sorts the array using the MergeSort algorithm (follows the Java legacyMergesort algorithm
'O(n*log(n)) time; O(n) space
Public Sub sort(ByRef a() As Variant, Optional ByRef c As IVariantComparator)

    If c Is Nothing Then
        MergeSort copyOf(a), a, 0, length(a), 0, Factory.newNumericComparator
    Else
        MergeSort copyOf(a), a, 0, length(a), 0, c
    End If
End Sub

Private Sub MergeSort(ByRef src() As Variant, ByRef dest() As Variant, low As Long, high As Long, off As Long, ByRef c As IVariantComparator)
    Dim length As Long
    Dim destLow As Long
    Dim destHigh As Long
    Dim mid As Long
    Dim i As Long
    Dim p As Long
    Dim q As Long

    length = high - low

    ' insertion sort on small arrays
    If length < INSERTIONSORT_THRESHOLD Then
        i = low
        Dim j As Long
        Do While i < high
            j = i
            Do While True
                If (j <= low) Then
                    Exit Do
                End If
                If (c.compare(dest(j - 1), dest(j)) <= 0) Then
                    Exit Do
                End If
                swap dest, j, j - 1
                j = j - 1 'decrement j
            Loop
            i = i + 1 'increment i
        Loop
        Exit Sub
    End If

    'recursively sort halves of dest into src
    destLow = low
    destHigh = high
    low = low + off
    high = high + off
    mid = (low + high) / 2
    MergeSort dest, src, low, mid, -off, c
    MergeSort dest, src, mid, high, -off, c

    'if list is already sorted, we're done
    If c.compare(src(mid - 1), src(mid)) <= 0 Then
        Copy src, low, dest, destLow, length - 1
        Exit Sub
    End If

    'merge sorted halves into dest
    i = destLow
    p = low
    q = mid
    Do While i < destHigh
        If (q >= high) Then
           dest(i) = src(p)
           p = p + 1
        Else
            'Otherwise, check if p<mid AND src(p) preceeds scr(q)
            'See description of following idom at: https://stackoverflow.com/a/3245183/3795219
            Select Case True
               Case p >= mid, c.compare(src(p), src(q)) > 0
                   dest(i) = src(q)
                   q = q + 1
               Case Else
                   dest(i) = src(p)
                   p = p + 1
            End Select
        End If

        i = i + 1
    Loop

End Sub
