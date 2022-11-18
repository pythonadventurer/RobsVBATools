Function funcSortCollection(ItemsList As Collection) As Collection
'Quick and easy way to sort a VBA Collection, using the sortable ListView object
'items list. Requires the component MSCOMNCTL.OCX in the directory C:\Windows\SysWOW64,
'a/k/a Microsoft Windows Common Controls 6.0 (SP6).  Add this component to your project
'using the References option in the VBA code editor within Microsoft Access.
  
'See:
'https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
'Answer by user: "ameisenmann"

Dim lv As ListView
Set lv = New ListView
Dim GivenItem As Variant
Dim SortedList As New Collection
For Each GivenItem In ItemsList
    lv.ListItems.Add Text:=GivenItem
Next GivenItem

lv.SortKey = 0            ' sort based on each item's Text
lv.SortOrder = lvwAscending
lv.Sorted = True

For Each GivenItem In lv.ListItems
    SortedList.Add GivenItem
Next GivenItem

Set funcSortCollection = SortedList

End Function
