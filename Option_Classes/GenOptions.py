def Header(d):
  h  = "VERSION 1.0 Class\n"
  h += "BEGIN\n"
  h += "  MultiUse = -1  'True\n"
  h += "END\n"
  h += "Attribute VB_Name = \"option_" + d['name'].lower() + "_c\"\n"
  h += "Attribute VB_GlobalNameSpace = False\n"
  h += "Attribute VB_Creatable = False\n"
  h += "Attribute VB_PredeclaredId = False\n"
  h += "Attribute VB_Exposed = False"

  h += "\nOption Explicit\nImplements option_c\n"
  h += "\nPublic Sub option_c_AddTo(obj as Shape)\n"
  h += "  Dim Name as String\n"
  h += "  Dim DisplayName as String\n"
  h += "  Dim CellName as String\n"
  h += "  Dim Prompt as String\n"
  h += "  Dim Section as Integer\n"
  h += "  Dim Row as Integer\n\n"

  h += "  ' this is the first unique part\n"
  h += "  DisplayName = " + d['var'] + "\n"
  h += "  Prompt = \"\"\"" + d['desc'] + "\"\"\"\n"
  h += "  Section = " + d['sect'] + "\n\n"

  h += "  Name = vw_strings.LegalName(DisplayName)\n"
  h += "  CellName = \"Prop.\" & Name\n"
  h += "  If obj.CellExists(Name, visExistsLocally) = False Then\n"
  h += "    Row = obj.AddNamedRow(Section, Name, visTagDefault)\n"
  h += "  Else\n"
  h += "    Row = obj.Cells(CellName).Row\n"
  h += "  End If\n\n"
  h += "  If Section = visSectionProp Then _\n"
  h += "    obj.CellsSRC(Section, Row, visCustPropsLabel).Formula = \"\"\"\" & DisplayName & \"\"\"\"\n\n"

  h += "  If Section = visSectionProp Then\n"
  h += "    obj.CellsSRC(Section, Row, visCustPropsPrompt).Formula = Prompt\n"
  h += "    obj.CellsSRC(Section, Row, visCustPropsType).Formula = " + d['type'] + "\n"
  h += "  ElseIf Section = visSectionUser Then\n"
  h += "    obj.CellsSRC(Section, Row, visUserPrompt).Formula = Prompt\n"
  h += "  End If\n"
  return h

dimensions = {'name' : 'ShowDimensions', 'var' : 'S_SHOW_DIMENSIONS',
              'sect' : 'visSectionProp', 'type' : 'visPropTypeBool',
              'desc' : "Set to show dimensions when adding an absolute or relative trigger type"}
childoffset = {'name' : 'ChildOffset', 'var' : 'S_CHILDOFFSET',
              'sect' : 'visSectionProp', 'type' : 'visPropTypeNumber',
              'desc' : "Minimum distance between other signals using this signal as a reference"}
skewwidth = {'name' : 'SkewWidth', 'var' : 'S_SKEWWIDTH',
              'sect' : 'visSectionUser', 'type' : 'visPropTypeNumber',
              'desc' : ""}

options = [dimensions]
path = "D:\Visio Projects\Visio Waves\Option_Classes"

for opt in options:
  #print(Header(show_dimensions))
  file = path + "\option_" + opt['name'].lower() + "_c.cls"
  fout = open(file, 'w')
  fout.write(Header(opt))
  fout.write("End Sub")
  fout.close()
