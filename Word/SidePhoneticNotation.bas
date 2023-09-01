Attribute VB_Name = "SidePhoneticNotation"

Sub SidePhoneticNotation()
Attribute SidePhoneticNotation.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.SidePhoneticNotation"
 If Selection.Type <> wdSelectionIP Then
 
  Dim charCount As Long
  charCount = Selection.Characters.Count
    
  SendKeys "{enter}", True
  Application.Run MacroName:="FormatPhoneticGuide"

  Selection.MoveRight Unit:=wdCharacter, Count:=charCount, Extend:=wdExtend
  Selection.Copy
  Selection.PasteAndFormat (wdFormatPlainText)
 End If
End Sub
