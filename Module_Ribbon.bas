Attribute VB_Name = "Module_Ribbon"
Sub GetVisible(control As IRibbonControl, ByRef MakeVisible)
'PURPOSE: Show/Hide buttons based on how many you need (False = Hide/True = Show)

Select Case control.ID
  Case "GroupA": MakeVisible = True
  Case "aButton01": MakeVisible = True
  Case "aButton02": MakeVisible = True
  Case "aButton03": MakeVisible = True
  Case "aButton04": MakeVisible = True
  Case "aButton05": MakeVisible = True
  Case "aButton06": MakeVisible = True
  Case "aButton07": MakeVisible = True
  
End Select

End Sub

Sub GetLabel(ByVal control As IRibbonControl, ByRef Labeling)
'PURPOSE: Determine the text to go along with your Tab, Groups, and Buttons

Select Case control.ID

  Case "CustomTab": Labeling = "COSMOCALC"

  Case "GroupA": Labeling = "CosmoCalc"
  Case "aButton01": Labeling = "Scaling"
  Case "aButton02": Labeling = "Shielding"
  Case "aButton03": Labeling = "Exposure"
  Case "aButton04": Labeling = "Banana"
  Case "aButton05": Labeling = "Converters"
  Case "aButton06": Labeling = "Settings"
  Case "aButton07": Labeling = "About"

End Select

End Sub

Sub GetImage(control As IRibbonControl, ByRef RibbonImage)
'PURPOSE: Tell each button which image to load from the Microsoft Icon Library
'TIPS: Image names are case sensitive, if image does not appear in ribbon after re-starting Excel, the image name is incorrect

Select Case control.ID
  
  Case "aButton01": RibbonImage = "ShowTimeZones"
  Case "aButton02": RibbonImage = "ObjectPictureFill"
  Case "aButton03": RibbonImage = "PictureBrightnessGallery"
  Case "aButton04": RibbonImage = "HappyFace"
  Case "aButton05": RibbonImage = "ListSynchronize"
  Case "aButton06": RibbonImage = "TableSharePointListsModifyColumnsAndSettings"
  Case "aButton07": RibbonImage = "TentativeAcceptInvitation"
  
End Select

End Sub

Sub GetSize(control As IRibbonControl, ByRef Size)
'PURPOSE: Determine if the button size is large or small

Const Large As Integer = 1
Const Small As Integer = 0

Select Case control.ID
    
  Case "aButton01": Size = Large
  Case "aButton02": Size = Large
  Case "aButton03": Size = Large
  Case "aButton04": Size = Large
  Case "aButton05": Size = Large
  Case "aButton06": Size = Large
  Case "aButton07": Size = Large
  
End Select

End Sub

Sub RunMacro(control As IRibbonControl)
'PURPOSE: Tell each button which macro subroutine to run when clicked

Select Case control.ID
  
  Case "aButton01": Application.Run "ScalingModels"
  Case "aButton02": Application.Run "Shielding"
  Case "aButton03": Application.Run "Age"
  Case "aButton04": Application.Run "Banana"
  Case "aButton05": Application.Run "Converters"
  Case "aButton06": Application.Run "Settings"
  Case "aButton07": Application.Run "About"

 End Select
    
End Sub

Sub GetScreentip(control As IRibbonControl, ByRef Screentip)
'PURPOSE: Display a specific macro description when the mouse hovers over a button

Select Case control.ID
  
  Case "aButton01": Screentip = "Calculate scaling factors"
  Case "aButton02": Screentip = "Topographic, snow, and self-shielding corrections"
  Case "aButton03": Screentip = "Calculate exposure/burial ages and erosion rates"
  Case "aButton04": Screentip = "Generate Al-Be and Ne-Be two-nuclide plots"
  Case "aButton05": Screentip = "Convert latitude and elevation units"
  Case "aButton06": Screentip = "Set calibration sites and production mechanisms"
  Case "aButton07": Screentip = "Further details about CosmoCalc"
  
End Select

End Sub
