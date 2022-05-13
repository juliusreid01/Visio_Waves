Attribute VB_Name = "vw_test_pageconfig"

Public Function PageHasConfig(Optional p as Page = Nothing) as Boolean
  If p is Nothing Then Set p = ActivePage
  If p.PageSheet.CellExists(vw_strings.S_PAGE_CFG_FULL, visExistsLocally) <> False Then
    PageHasConfig = True
  Else
    PageHasConfig = False
  End If
End Function

Public Sub TestPageConfig()
  If PageHasConfig(ActivePage) = True Then
    MsgBox Title:="Page Config Test", Buttons:=vbCritical + vbOkayOnly, _
      Prompt:="Page already has config"
  End If

  vw_page_cfg.Request ActivePage, vbCancel
End Sub
