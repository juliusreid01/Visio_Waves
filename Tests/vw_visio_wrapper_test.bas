Attribute VB_Name = "vw_visio_waves_test"

Private Sub FatalMsg(SubTitle as String, Message as String)
  MsgBox Title:="Visio Waves Test: " & SubTitle, _
    Buttons:=vbCritical + vbOkayOnly, _
    Prompt:=Message
End Sub

Public Sub TestAll(s as Shape)
  TestUserCell s
  TestDataCell s
  TestCells s
End Sub

Public Sub TestUserCell(s as Shape)
  Dim wShape as visio_shape_wrapper_c
  Set wShape = New visio_shape_wrapper_c
  Set wShape.vsoShape = s
  wShape.AddUserCell "HelloWorld", """Testing User Cell"""
  If s.CellExists("User.HelloWorld", visExistsLocally) = False Then
    FatalMsg "UserCell", "Failed to add User Cell ""HelloWorld"""
    Stop
    End
  End If
  wShape.SetUserCell "HelloWorld", "=2+2"
  If s.Cells("User.HelloWorld").Result("") <> 4 Then
    MsgBox Title:="Visio Waves Test: UserCell", Buttons:=vbCritical + vbOkayOnly, _
      Prompt:="Failed to set User Cell ""HelloWorld"""
    Stop
    End
  End If
End Sub

Public Sub TestDataCell(s as Shape)
  Dim wShape as visio_shape_wrapper_c
  Set wShape = New visio_shape_wrapper_c
  Set wShape.vsoShape = s
  wShape.AddDataCell "HelloWorld", """Testing Data Cell"""
  If s.CellExists("Prop.HelloWorld", visExistsLocally) = False Then
    FatalMsg "DataCell", "Failed to add Data Cell ""HelloWorld"""
    Stop
    End
  End If
  wShape.SetDataCell "HelloWorld", "=2+2", "Value"
  If s.Cells("Prop.HelloWorld").Result("") <> 4 Then
    FatalMsg "DataCell", "Failed to set Data Cell ""HelloWorld"""
    Stop
    End
  End If
End Sub

Public Sub TestCells(s as Shape)
  Dim wShape as visio_shape_wrapper_c
  Set wShape = New visio_shape_wrapper_c
  Set wShape.vsoShape = s
  wShape.AddUserCell "HelloWorld_u", """Testing Cell"""
  wShape.AddDataCell "HelloWorld_d", """Testing Cell"""
  wShape.SetCell "HelloWorld_u", "=2+1"
  If s.Cells("User.HelloWorld_u").Result("") <> 3 Then
    FatalMsg "Cell", "Failed to set cell ""HelloWorld_u"""
    Stop
    End
  End If
  wShape.SetCell "HelloWorld_d", "=3+1"
  If s.Cells("Prop.HelloWorld_d").Result("") <> 4 Then
    FatalMsg "Cell", "Failed to set cell ""HelloWorld_d"""
    Stop
    End
  End If
  wShape.SetCell "BeginX", "=1.8"
  If s.Cells("BeginX").Result("") <> 1.8 Then
    FatalMsg "Cell", "Failed to set cell ""BeginX"""
    Stop
    End
  End If
End Sub