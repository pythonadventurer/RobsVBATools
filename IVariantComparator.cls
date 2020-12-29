VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVariantComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'The IVariantComparator provides a method, compare, that imposes a total ordering over a collection _
of variants. A class that implements IVariantComparator, called a Comparator, can be passed to the _
Arrays.sort and Collections.sort methods to precisely control the sort order of the elements.

'Compares two variants for their sort order. Returns -1 if v1 should be sorted ahead of v2; +1 if _
v2 should be sorted ahead of v1; and 0 if the two objects are of equal precedence. This function _
should exhibit several necessary behaviors: _
  1.) compare(x,y)=-(compare(y,x) for all x,y _
  2.) compare(x,y)>= 0 for all x,y _
  3.) compare(x,y)>=0 and compare(y,z)>=0 implies compare(x,z)>0 for all x,y,z
Public Function compare(ByRef v1 As Variant, ByRef v2 As Variant) As Long
End Function

