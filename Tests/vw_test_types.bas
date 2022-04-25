Attribute VB_Name = "vw_test_types"

' this function may require removing a shape so we stop and then delete the shape
Private Sub Compare(var as Variant, exp as Variant, prompt as String)
  If var <> exp Then
    If MsgBox(Title:="Test Type", Buttons:=vbYesNo, _
            Prompt:=prompt & " != " & exp & ", Read: " & var & _
            vbNewLine & "Continue?") = vbNo Then Stop
  End If
End Sub

' check the values
Public Sub ValueTest()
  ' Shape Type Tests
  Compare vw_types.ShapeType_t.Void, 0, "Shape Type Void"
  Compare vw_types.ShapeType_t.Clock, 4, "Shape Type Clock"
  Compare vw_types.ShapeType_t.Bit, 1, "Shape Type Bit"
  Compare vw_types.ShapeType_t.Bus, 2, "Shape Type Bus"
  Compare vw_types.ShapeType_t.Signal, 7, "Shape Type Signal"
  Compare vw_types.ShapeType_t.Data, 3, "Shape Type NotClock"
  Compare vw_types.ShapeType_t.Label, 8, "Shape Type Label"
  Compare vw_types.ShapeType_t.Node, 16, "Shape Type Node"
  Compare vw_types.ShapeType_t.Gate, 32, "Shape Type Gate"
  Compare vw_types.ShapeType_t.Gap, 64, "Shape Type Gap/Spacer"
  Compare vw_types.ShapeType_t.Child, 120, "Shape Type Child"
  ' Event Trigger Tests
  Compare vw_types.EventTrigger_t.Absolute, 4, "Event Trigger Absolute"
  Compare vw_types.EventTrigger_t.Relative, 12, "Event Trigger Relative"
  Compare vw_types.EventTrigger_t.Edge, 3, "Event Trigger Edge"
  Compare vw_types.EventTrigger_t.Posedge, 1, "Event Trigger Posedge"
  Compare vw_types.EventTrigger_t.Negedge, 2, "Event Trigger Negedge"
End Sub
