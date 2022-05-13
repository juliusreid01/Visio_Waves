Attribute VB_Name = "vw_types"

Option Explicit

Public Const VW_EVENT_SECTION as Integer = visSectionScratch
Public Const VW_EVENT_SECTION_COL as Integer = visScratchA

Public Enum ShapeType_t
  Void = 0
  Clock = 4
  Bit = 1
  Bus = 2
  Data = Bit Or Bus
  Signal = Clock Or Data
  Label = 8
  Node = 16
  ' this is specifically GateX which is drawn with a different shape
  '//TODO is this also GateZ, is that drawn with a different shape?
  Gate = 32
  Gap = 64
  Child = Label Or Node Or Gate Or Gap
End Enum

Public Enum EventTrigger_t
  Absolute = 4
  Relative = 12
  Edge = 3
  Posedge = 1
  Negedge = 2
End Enum

Private Enum EventFields_t
  ' indicates bit 1 is set
  F_BIT_1 = 2^30
  ' indicates bit 0 is set
  F_BIT_0 = 2^29
  ' if neither bit is set, consider Z
  ' if both bits are set, consider X
  ' indicate a gap/spacer shape exists for this event
  F_GAP = 2^20
  ' indicates a node exists for this event
  F_NODE = 2^19
  ' indicates a label exists for this event
  F_LABEL = 2^18
  ' indicates the event is a 1: gate, 0: edge
  F_GATE = 2^17
  ' indicates the event is a valid
  F_EVENT = 2^16
End Enum

Private Enum EventMasks_t
  ROW_MASK = F_EVENT - 1
  EVENT_MASK = F_EVENT Or F_GATE
  CHILD_MASK = F_LABEL Or F_NODE Or F_GAP
  BIT_MASK = F_BIT_0 Or F_BIT_1
End Enum

