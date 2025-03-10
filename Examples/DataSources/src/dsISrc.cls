VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsISrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Get the name of this type of fiber
'@remark - Used to populate data validation in column C "Type"
Public Function getName() As String
  
End Function

'Links a fiber
'@param template - The template fiber to link to.
'@param destPath - The destination of the CSV.
'@param data - Data compliant with `getDataType` function.
Public Function linkFiber(ByVal template As stdFiber, ByVal destPath As String, ByVal Data As stdJSON) As stdFiber
  
End Function

'Returns a required data type
'@returns Object<Key: Descriptor>
'@example ```
'{
'  "File": {
'    "DisplayText": "Source File",
'    "Tooltip": "The location of the Excel file to get the data from.",
'    "Type": "File",
'    "File":{
'      "Extensions": "*.xlsx,*.xlsm,*.xlsb",
'      "Multiselect": false
'    }
'  },
'  "Sheet": {
'    "DisplayText": "Source Sheet",
'    "Tooltip": "The sheet from which to get the data from.",
'    "Type": "Dropdown",
'    "Dropdown":{
'      "Options": stdCallback.CreateFromObjectMethod(Me, "protGetSheetNamesCollection")
'    }
'  },
'  "Range": {
'    "DisplayText": "Range address",
'    "Tooltip": "The range from which to get the data from.",
'    "Type": "Text"
'  }
'}
'==>
'{"File":"", "Sheet":"", "Range":""}
'```
'@remark - Used to generate UI for selecting data, which generates data for column D, "Data"
Public Function getDataType() As stdJSON
  
End Function
