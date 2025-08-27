VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdUIAnchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum UIAnchorTypes
  UIAnchorTypeAll
  UIAnchorTypeRight
  UIAnchorTypeLeft
  UIAnchorTypeTop
  UIAnchorTypeBottom
End Enum

Private Type TPercentiles
  PCLeft As Double   'Percentage of parent width where left side starts
  PCRight As Double  'Percentage of parent width where right side starts
  PCTop As Double    'Percentage of parent height where top side starts
  PCBottom As Double 'Percentage of parent height where bottom side starts
End Type
Private Type TThis
  Element As MSForms.Control
  Parent As Object
  percentiles As TPercentiles
End Type

Public Function Create()
  
End Function

Public Sub protInit()

End Sub


Public Sub Resize()
  
End Sub
