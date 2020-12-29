Attribute VB_Name = "modColDataControls"
Option Compare Database
Option Explicit

Function colDataControls() As Object

Set colDataControls = New Collection

colDataControls.Add "ListBox"
colDataControls.Add "CheckBox"
colDataControls.Add "ComboBox"
colDataControls.Add "OptionButton"
colDataControls.Add "OptionGroup"
colDataControls.Add "TextBox"
colDataControls.Add "Togglebutton"

End Function

