Attribute VB_Name = "vw_test_pageconfig"

Private Function PageHasCell(s as String) as Boolean
  s = vw_strings.LegalName(s)
  If ActivePage.PageSheet.CellExists("Prop." & s, visExistsLocally) <> False Then
    PageHasCell = True
  ElseIf ActivePage.PageSheet.CellExists("User." & s, visExistsLocally) <> False Then
    PageHasCell = True
  Else
    PageHasCell = False
    MsgBox Title:="Page Config Test", Buttons:=vbCritical + vbOkayOnly, _
      Prompt:="Page does not have expected cell: " & s
  End If
End Function

Public Sub TestPageConfig()
  Dim rsp as Integer
  Dim opts as Collection

  If PageHasConfig(ActivePage) = True Then
    MsgBox Title:="Page Config Test", Buttons:=vbCritical + vbOkayOnly, _
      Prompt:="Page already has config, call DeleteCfg and Resume after break"
    Stop
    DeleteCfg ActivePage.PageSheet
  End If

  ' cancel means we will ask again
  vw_page_cfg.Request ActivePage, vbCancel
  If PageHasConfig(ActivePage) = True Then
    MsgBox Title:="Page Config Test", Buttons:=vbCritical + vbOkayOnly, _
      Prompt:="Page should not have config after Cancel"
  End If

  ' no means we will not ask again
  For i = 1 to 2
    vw_page_cfg.Request ActivePage, vbNo
    rsp = -1
    If PageHasConfig(ActivePage) = False Then
      rsp = MsgBox(Title:="Page Config Test", Buttons:=vbCritical + vbOkayOnly, _
              Prompt:="Page should have config after No")
    ElseIf ActivePage.PageSheet.Cells(S_PAGE_CFG_FULL).Result("") <> False Then _
      rsp = MsgBox(Title:="Page Config Test", Buttons:=vbCritical + vbOkayOnly, _
              Prompt:="Page Config expected False after No, read True")
    End If
    If i = 2 And rsp > 0 Then
      MsgBox Title:="Page Config Test", Buttons:=vbCritical + vbOkayOnly, _
        Prompt:="Page Config set should not expect another inquiry"
    End If
  Next
  DeleteCfg ActivePage.PageSheet

  ' yes means will not ask again and values
  For i = 1 to 2
    vw_page_cfg.Request ActivePage, vbYes
    rsp = -1
    If PageHasConfig(ActivePage) = False Then
      rsp = MsgBox(Title:="Page Config Test", Buttons:=vbCritical + vbOkayOnly, _
              Prompt:="Page should have config after Yes")
    ElseIf ActivePage.PageSheet.Cells(S_PAGE_CFG_FULL).Result("") = False Then _
      rsp = MsgBox(Title:="Page Config Test", Buttons:=vbCritical + vbOkayOnly, _
              Prompt:="Page Config expected True after Yes, read False")
    Else
      If PageHasCell(S_SHOW_DIMENSIONS) = False Then Stop
      If PageHasCell(S_CHILDOFFSET) = False Then Stop
      If PageHasCell(S_SKEWWIDTH) = False Then Stop
      If PageHasCell(S_ACTIVELOW) = False Then Stop
      If PageHasCell(S_PERIOD) = False Then Stop
      If PageHasCell(S_SKEW) = False Then Stop
    End If
    If i = 2 And rsp > 0 Then
      MsgBox Title:="Page Config Test", Buttons:=vbCritical + vbOkayOnly, _
        Prompt:="Page Config set should not expect another inquiry"
    End If
  Next

  Stop ' review before resuming
  DeleteCfg ActivePage.PageSheet
End Sub

Public Sub DeleteCfg(p as Shape)
  p.DeleteSection visSectionUser
  p.DeleteSection visSectionProp
End Sub
