Attribute VB_Name = "modPCase"
Option Compare Database

'The PCase() function accepts any string, and returns
'a string with words converted to initial caps (proper case).
'See:
'https://sourcedaddy.com/ms-access/proper-case-function.html

Public Function ProperCase(AnyText As String) As String
    'Create a string variable, then store AnyText in that variable already
    'converted to proper case using the built-in StrConv() function
    Dim FixedText As String
    FixedText = StrConv(AnyText, vbProperCase)

    'Now, take care of StrConv() shortcomings

    'If first two letters are "Mc", cap third letter.
    If Left(FixedText, 2) = "Mc" Then
    FixedText = Left(FixedText, 2) & _
        UCase(mid(FixedText, 3, 1)) & mid(FixedText, 4)
    End If

    'If first three letters are "Mac", cap fourth letter.
    If Left(FixedText, 3) = "Mac" Then
        FixedText = Left(FixedText, 3) & UCase(mid(FixedText, 4, 1)) & mid(FixedText, 5)
    End If

    'If starts with O and apostrophe, cap the letter after the apostrophe.
    If Left(FixedText, 2) = "O'" Then
        FixedText = Left(FixedText, 2) & UCase(mid(FixedText, 3, 1)) & mid(FixedText, 4)
    
    End If
    

    'Now return the modified string.
    ProperCase = FixedText
    
End Function

