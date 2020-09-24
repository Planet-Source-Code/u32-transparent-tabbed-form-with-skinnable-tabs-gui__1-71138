Attribute VB_Name = "mdlSkinning"
Option Explicit

' Made by: u32
' Aug, 08

' Skins a form and skinable controls
' Skins stored in the apps install dir. (own folder)

Sub SkinForm(ByVal frm As Object, ByVal sSkin As String)
    
    Dim SkinPath As String
    
    SkinPath = App.Path & "\" & sSkin & "\"
    
    With frm
      .Picture = LoadPicture(SkinPath & "back.gif")
      .Go.Picture = LoadPicture(SkinPath & "close.gif")
      .closeMeb.Picture = LoadPicture(SkinPath & "closeover.gif")
      .closeMea.Picture = LoadPicture(SkinPath & "close.gif")
      .oka.Picture = LoadPicture(SkinPath & "rightb.gif")
      .okb.Picture = LoadPicture(SkinPath & "rightbOver.gif")
      .cancela.Picture = LoadPicture(SkinPath & "middleb.gif")
      .cancelb.Picture = LoadPicture(SkinPath & "middlebOver.gif")
      .applya.Picture = LoadPicture(SkinPath & "leftb.gif")
      .applyb.Picture = LoadPicture(SkinPath & "leftbOver.gif")
      .tabLow.Picture = LoadPicture(SkinPath & "backTab.gif")
      .tabHigh.Picture = LoadPicture(SkinPath & "currentTab.gif")
      .cancel.Picture = .cancela.Picture
      .ok.Picture = .oka.Picture
      .apply.Picture = .applya.Picture
    End With
    
End Sub

