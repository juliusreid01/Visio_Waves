VERSION 1.0 Class
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "option_labelfont_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements option_c

Public Sub option_c_AddTo(obj as Shape)
  Dim Name as String
  Dim DisplayName as String
  Dim CellName as String
  Dim Prompt as String
  Dim Section as Integer
  Dim Row as Integer

  ' this is the first unique part
  DisplayName = S_LABELFONT
  Prompt = """Controls the font size of labels"""
  Section = visSectionProp

  Name = vw_strings.LegalName(DisplayName)
  CellName = "Prop." & Name
  If obj.CellExists(Name, visExistsLocally) = False Then
    Row = obj.AddNamedRow(Section, Name, visTagDefault)
  Else
    Row = obj.Cells(CellName).Row
  End If

  ' additional uniqueness may exist here
  If Section = visSectionProp Then
    obj.CellsSRC(Section, Row, visCustPropsPrompt).Formula = Prompt
    obj.CellsSRC(Section, Row, visCustPropsLabel).Formula = """" & DisplayName & """"
    obj.CellsSRC(Section, Row, visCustPropsType).Formula = visPropTypeListNumber
    obj.CellsSRC(Section, Row, visCustPropsFormat).Formula = 0 pt
  ElseIf Section = visSectionUser Then
    obj.CellsSRC(Section, Row, visUserPrompt).Formula = Prompt
  End If
End Sub