Attribute VB_Name = "vw_page_cfg"

Public Sub Request(p as Page, Optional override as Integer = 0)
  Dim resp as Integer
  Dim o_msg as String

  If p.PageSheet.CellExists(S_PAGE_CFG_FULL, visExistsLocally) <> False Then Exit Sub

  Select Case override
   Case vbYes: o_msg = vbNewLine & vbNewLine & "Override is set to Yes"
   Case vbNo: o_msg = vbNewLine & vbNewLine & "Override is set to No"
   Case vbCancel: o_msg = vbNewLine & vbNewLine & "Override is set to Cancel"
   Case Else: o_msg = ""
  End Select

  resp = MsgBox(Title:="Page Config Request", Buttons:=vbQuestion + vbYesNoCancel,
           Prompt:="Allow default shape data to be read from the page?" & vbNewLine & _
           "Click Yes to allow, will add Shape Data to this page" & vbNewLine & _
           "Click No and we will not ask about this again" & vbNewLine & _
           "Click Cancel and we'll ask when a new signal is dropped on the page" & o_msg
  If o_msg <> "" Then resp = override

  If resp = vbNo Then
    p.PageSheet.AddNamedRow visSectionUser, S_PAGE_CFG, visTagDefault
    p.Cells(S_PAGE_CFG_FULL).Formula = False
  ElseIf resp = vbYes Then
    p.PageSheet.AddNamedRow visSectionUser, S_PAGE_CFG, visTagDefault
    p.Cells(S_PAGE_CFG_FULL).Formula = True
    ' add more data
    'show_dimensions.PageAdd
    
  End If
End Sub