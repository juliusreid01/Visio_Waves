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

  h += "  ' additional uniqueness may exist here\n"
  h += "  If Section = visSectionProp Then\n"
  h += "    obj.CellsSRC(Section, Row, visCustPropsPrompt).Formula = Prompt\n"
  h += "    obj.CellsSRC(Section, Row, visCustPropsLabel).Formula = \"\"\"\" & DisplayName & \"\"\"\"\n"
  for k in d:
    if k[:3] == 'vis':
      h += "    obj.CellsSRC(Section, Row, " + k + ").Formula = " + d[k] + "\n"
  h += "  ElseIf Section = visSectionUser Then\n"
  h += "    obj.CellsSRC(Section, Row, visUserPrompt).Formula = Prompt\n"
  h += "  End If\n"
  return h

dimensions = {'name' : 'ShowDimensions', 'var' : 'S_SHOW_DIMENSIONS',
              'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeBool',
              'desc' : "Set to show dimensions when adding an absolute or relative trigger type"}
childoffset = {'name' : 'ChildOffset', 'var' : 'S_CHILDOFFSET',
              'sect' : 'visSectionProp', 'type' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0.00 u"""',
              'desc' : "Minimum distance between other signals using this signal as a reference"}
skewwidth = {'name' : 'SkewWidth', 'var' : 'S_SKEWWIDTH',
              'sect' : 'visSectionUser', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0.00 u"""',
              'desc' : ""}
activelow = {'name' : 'activelow', 'var' : 'S_ACTIVELOW',
             'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeBool',
             'desc' : ""}
period = {'name' : 'period', 'var' : 'S_PERIOD',
          'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0.00 u"""',
          'desc' : ""}
skew = {'name' : 'skew', 'var' : 'S_SKEW',
        'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0.0 %"""',
        'desc' : ""}
i_delay = {'name' : 'i_delay', 'var' : '"Clock " & ' + 'S_DELAY',
           'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0.00 u"""',
           'desc' : "Initial delay for a clock signal to start"}
s_delay = {'name' : 's_delay', 'var' : '"Signal " & ' + 'S_DELAY',
           'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0.00 u"""',
           'desc' : "Delay before a data signal will transition when referencing a clock"}
duty = {'name' : 'dutycycle', 'var' : 'S_DUTYCYCLE',
        'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0 %"""',
        'desc' : ""}
sigskew = {'name' : 'signalskew', 'var' : 'S_SIGNALSKEW',
           'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0.00 u"""',
           'desc' : "Additional skew to apply to signals on top of the clock skew"}
busw = {'name' : 'buswidth', 'var' : 'S_BUSWIDTH',
        'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber',
        'desc' : "Controls text for Bus Signal Types"}
lbledges = {'name' : 'labeledges', 'var' : 'S_LABELEDGES',
            'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeListVar',
            'desc' : "Controls which labels are shown about the Signal transitions"}
lblsize = {'name' : 'labelsize', 'var' : 'S_LABELSIZE',
           'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0.00 u"""',
           'desc' : "Controls the size of labels"}
lblfont = {'name' : 'labelfont', 'var' : 'S_LABELFONT',
           'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0 pt"""',
           'desc' : "Controls the font size of labels"}
nodefont = {'name' : 'nodefont', 'var' : 'S_NODEFONT',
           'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0 pt"""',
           'desc' : "Controls the font size of nodes"}
nodesize = {'name' : 'nodesizemult', 'var' : 'S_NODESIZEMULT',
           'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeNumber', 'visCustPropsFormat' : '"""0.00 u"""',
           'desc' : "Controls the size of nodes"}
lblshape = {'name' : 'labelshape', 'var' : 'S_LBL_SHAPE',
            'sect' : 'visSectionProp', 'visCustPropsType' : 'visPropTypeListFIX',
            'visCustPropsFormat' : '"""" & vw_strings.GenList(S_LBL_RECTANGLE, S_LBL_SQUARE, S_LBL_DIAMOND, S_LBL_RND_RECTANGLE, S_LBL_RND_SQUARE, S_LBL_RND_DIAMOND, S_LBL_OVAL, S_LBL_CIRCLE) & """"',
            'desc' : "Controls the shape of labels"}

options = [dimensions, childoffset, skewwidth, activelow, period, skew, i_delay, i_delay, s_delay, duty, sigskew, busw, lbledges, lblsize, lblfont, nodefont, nodesize, lblshape]

path = "D:\Visio Projects\Visio Waves\Option_Classes"

for opt in options:
  #print(Header(show_dimensions))
  file = path + "\option_" + opt['name'].lower() + "_c.cls"
  fout = open(file, 'w')
  fout.write(Header(opt))
  fout.write("End Sub")
  fout.close()
