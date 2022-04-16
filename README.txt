Current Visio Waves project directory

Code is developed using test driven development
Red, Green, Refactor

After enabling Macro Settings -> Allow Programmatic Access to VBProject
Run GenVBAReset.py to update the files being imported and import VBAReset.bas to have control

Directory Structure
.
├── GenVBAReset.py
├── README.txt
├── VBAReset.bas
├── Classes
│   └── vw_*_c.cls
├── Documents
├── Forms
│   └── vw_*_f.frm
├── Modules
│   └── vw_*.bas
├── Tests
│   └── vw_test_*.bas
└── Visio_Shape_Wrapper
    ├── Test.bas
    └── visio_shape_wrapper_c.cls