Attribute VB_Name = "vw_strings"

' this file contains strings for the custom cells and string functions we need for Visio Waves
Option Explicit

' The Page PageSheet cell must have this cell to be used as a configuration object
Public Const S_PAGE_CFG as String = "vw_cfg"
Public Const S_PAGE_CFG_FULL as String = "User." & S_PAGE_CFG

' minimum distance between children
Public Const S_CHILDOFFSET as String = "Child Offset"
' clocks: period * 0.5 * skew
' other: ignored
Public Const S_SKEWWIDTH as String = "Skew Width"
' number of edges counted
Public Const S_EDGES as String = "Edges"
' clocks: period * duty cycle
' other: ignored
Public Const S_ACTIVEWIDTH as String = "Active Width"
' number of pulses calculated
Public Const S_PULSES as String = "Pulses"
' test cell for calculations
Public Const S_TEST as String = "Test"
' name of the parent shape, which if deleted will delete the shape
Public Const S_PARENT as String = "Parent"

' the type of shape Clock, Signal, Label, etc.
Public Const S_TYPE as String = "Type"
'//TODO remove this. Let Visio handle the names and use the text
' displays the name of the shape for linking
Public Const S_NAME as String = "Name"
' sets a reference clock
Public Const S_CLOCK as String = "Clock"
' sets a reference signal
Public Const S_SIGNAL as String = "Signal"
Public Const S_ACTIVELOW as String = "Active Low"
Public Const S_PERIOD as String = "Period"
Public Const S_SKEW as String = "Skew %"
' clocks: initial delay
' other: delay before a transition
Public Const S_DELAY as String = "Delay"
Public Const S_DUTYCYCLE as String = "Duty Cycle %"
' clocks: additional skew to apply to signals referencing this clock
Public Const S_SIGNALSKEW as String = "Signal Skew"
' bus: width of the bus can use instead of changing the text
Public Const S_BUSWIDTH as String = "Bus Width"
Public Const S_EVENTTYPE as String = "Event Type"
Public Const S_EVENTTRIGGER as String = "Event Trigger"
Public Const S_EVENTPOSITION as String = "Event Position"
Public Const S_LABELEDGES as String = "Label Edges"
Public Const S_LABELSIZE as String = "Label Size"
Public Const S_LABELFONT as String = "Label Font Size"
Public Const S_NODEFONT as String = "Node Font Size"
Public Const S_NODESIZEMULT as String = "Node Size Multiplier"

' event type strings
Public Const S_EVENT_NODE as String = "Node"
Public Const S_EVENT_GAP as String = "Gap"
Public Const S_EVENT_EDGE as String = "Transition"
Public Const S_EVENT_DRIVEX as String = "DriveX"
Public Const S_EVENT_DRIVEZ as String = "DriveZ"
Public Const S_EVENT_DRIVE0 as String = "Drive0"
Public Const S_EVENT_DRIVE1 as String = "Drive1"
Public Const S_EVENT_DELETE as String = "Delete"

' event trigger strings
Public Const S_TRIGGER_EDGE as String = "Any Edge"
Public Const S_TRIGGER_POSEDGE as String = "Posedge"
Public Const S_TRIGGER_NEGEDGE as String = "Negedge"
Public Const S_TRIGGER_ABSOLUTE as String = "Absolute"
Public Const S_TRIGGER_RELATIVE as String = "Relative"

' label shape strings
Public Const S_LBL_SHAPE as String = "Label Shape"
Public Const S_LBL_RECTANGLE as String = "Rectangle"
Public Const S_LBL_SQUARE as String = "Square"
Public Const S_LBL_DIAMOND as String = "Diamond"
Public Const S_LBL_RND_RECTANGLE as String = "Rounded Rectangle"
Public Const S_LBL_RND_SQUARE as String = "Rounded Square"
Public Const S_LBL_RND_DIAMOND as String = "Rounded Diamond"
Public Const S_LBL_OVAL as String = "Oval"
Public Const S_LBL_CIRCLE as String = "Circle"

Public Function GenList(ParamArray items() as Variant)
  GenList = items(LBound(items))
  For i = LBound(items) + 1 to UBound(items)
    GenList = GenList & ";" & items(i)
  Next i
End Function

' cell names cannot have spaces or special characters
Public Function LegalName(str as String) as String
  LegalName = Replace(Replace(str, " ", ""), "%", "")
End Function