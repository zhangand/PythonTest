from win32com import client
import os


class AnsysDesigner:
    def __init__(self):
        self.oAnsoftApp = client.Dispatch('Ansoft.ElectronicsDesktop')
        self.oDesktop = self.oAnsoftApp.GetAppDesktop()
        #self.oDesktop.RestoreWindow()
        self.oProject = self.oDesktop.NewProject()
        self.oProject.InsertDesign("Circuit Design", "Circuit1", "", "")
        self.oDefinitionManager = self.oProject.GetDefinitionManager()
        self.oModelManager = self.oDefinitionManager.GetManager("Model")
        self.oDesign = self.oProject.SetActiveDesign("Circuit1")
        self.oEditor = self.oDesign.SetActiveEditor("SchematicEditor")
        self.oComponentManager = self.oDefinitionManager.GetManager("Component")
        self.oSimSetup = self.oDesign.GetModule("SimSetup")
        self.oReportSetup = self.oDesign.GetModule("ReportSetup")

    def import_model(self, _file_path, _comp_name, ):
        self.oModelManager.Add(
            [
                "NAME:"+_comp_name,
                "Name:="	, _comp_name,
                "ModTime:="		, 1551594058,
                "Library:="		, "",
                "LibLocation:="		, "Project",
                "ModelType:="		, "nport",
                "Description:="		, "",
                "ImageFile:="		, "",
                "SymbolPinConfiguration:=", 0,
                [
                    "NAME:PortInfoBlk"
                ],
                [
                    "NAME:PortOrderBlk"
                ],
                "filename:="		, _file_path,     # "./VDD_GPU.s4p",
                "numberofports:="	, 4,
                "sssfilename:="		, "",
                "sssmodel:="		, False,
                "PortNames:="		, ["DUT0_DUT0_VDD_GPU_0_Group_DUT0_GND_Group" ,"DUT2_DUT2_VDD_GPU_2_Group_DUT2_GND_Group"
                                ,"UFP_UVS64_100_8_100_UFP_UVS64_100_8_100_VDD_GPU_2_Group_UFP_UVS64_100_8_100_GND_Group"
                                ,"UFP_UVS64_100_9_100_UFP_UVS64_100_9_100_VDD_GPU_0_Group_UFP_UVS64_100_9_100_GND_Group"],
                "domain:="		, "frequency",
                "datamode:="		, "Link",
                "devicename:="		, "",
                "SolutionName:="	, "",
                "displayformat:="	, "MagnitudePhase",
                "datatype:="		, "SMatrix",
                [
                    "NAME:DesignerCustomization",
                    "DCOption:="		, 0,
                    "InterpOption:="	, 0,
                    "ExtrapOption:="	, 1,
                    "Convolution:="		, 0,
                    "Passivity:="		, 0,
                    "Reciprocal:="		, False,
                    "ModelOption:="		, "",
                    "DataType:="		, 1
                ],
                [
                    "NAME:NexximCustomization",
                    "DCOption:="		, 3,
                    "InterpOption:="	, 1,
                    "ExtrapOption:="	, 3,
                    "Convolution:="		, 0,
                    "Passivity:="		, 0,
                    "Reciprocal:="		, False,
                    "ModelOption:="		, "",
                    "DataType:="		, 2
                ],
                [
                    "NAME:HSpiceCustomization",
                    "DCOption:="		, 1,
                    "InterpOption:="	, 2,
                    "ExtrapOption:="	, 3,
                    "Convolution:="		, 0,
                    "Passivity:="		, 0,
                    "Reciprocal:="		, False,
                    "ModelOption:="		, "",
                    "DataType:="		, 3
                ],
                "NoiseModelOption:="	, "External"
            ])

    def add_comp(self, _comp_name, ):
        self.oComponentManager.Add(
            [
                "NAME:"+_comp_name,
                "Refbase:="		, "S",
                "NumParts:="		, 1,
                "ModSinceLib:="		, False,
                "CircuitEnv:="		, 0,
                "CompExtID:="		, 5,
                [
                    "NAME:Parameters",
                    "MenuProp:="		, ["CoSimulator" ,"SD" ,"" ,"Default" ,0],
                    "ButtonProp:="		, ["CosimDefinition" ,"SD" ,"" ,"Edit" ,"Edit" ,40501,
                    "ButtonPropClientData:=", []]
                ],
                [
                    "NAME:CosimDefinitions",
                    [
                        "NAME:CosimDefinition",
                        "CosimulatorType:="	, 102,
                        "CosimDefName:="	, "Default",
                        "IsDefinition:="	, True,
                        "Connect:="		, True,
                        "ModelDefinitionName:="	, _comp_name,
                        "ShowRefPin2:="		, 2,
                        "LenPropName:="		, ""
                    ],
                    "DefaultCosim:="	, "Default"
                ]
            ])

    def create_comp(self, _comp_name,):
        self.oEditor.CreateComponent(
            [
                "NAME:ComponentProps",
                "Name:="	, _comp_name,
                "Id:="			, "3"
            ],
            [
                "NAME:Attributes",
                "Page:="		, 1,
                "X:="			, 0,
                "Y:="			, 0,
                "Angle:="		, 0,
                "Flip:="		, False
            ])

    def assign_port(self, *_signal_pins, **_gnd_pins):
        self.oEditor.AddPinIPorts(
            [
                "NAME:Selections",
                "Selections:="		, ["CompInst@VDD_GPU;3;1"]
            ])
        self.oEditor.Delete(
            [
                "NAME:Selections",
                "Selections:="		, ["IPort@UFP_UVS64_100_8_100_UFP_UVS64_100_8_100_VDD_GPU_2_Group_UFP_UVS64_100_8_100_GND_Group;12"]
            ])
        self.oEditor.Delete(
            [
                "NAME:Selections",
                "Selections:="	, ["IPort@UFP_UVS64_100_9_100_UFP_UVS64_100_9_100_VDD_GPU_0_Group_UFP_UVS64_100_9_100_GND_Group;16"]
            ])
        self.oEditor.AddPinGrounds(
            [
                "NAME:Selections",
                "Selections:="		, ["CompInst@VDD_GPU;3;1"]
            ])

    def insert_analysis_setup(self, *_freq):
        self.oSimSetup.AddLinearNetworkAnalysis(
            [
                "NAME:SimSetup",
                "DataBlockID:="	, 16,
                "SimSetupID:="		, 8,
                "OptionName:="		, "(Default Options)",
                "AdditionalOptions:="	, "",
                "AlterBlockName:="	, "",
                "FilterText:="		, "",
                "AnalysisEnabled:="	, 1,
                [
                    "NAME:OutputQuantities"
                ],
                [
                    "NAME:NoiseOutputQuantities"
                ],
                "Name:="		, "LinearFrequency",
                "LinearFrequencyData:="	, [False ,0.1 ,False ,"" ,False],
                [
                    "NAME:SweepDefinition",
                    "Variable:="		, "F",
                    "Data:="		, "DEC 1kHz 1GHz 301",
                    "OffsetF1:="		, False,
                    "Synchronize:="		, 0
                ]
            ])

    def run_analyze(self):
        self.oDesign.Analyze("LinearFrequency")

    def create_reports(self):
        self.oReportSetup.CreateReport("VDD_GPU", "Standard", "Rectangular Plot", "LinearFrequency",
                             [
                                 "NAME:Context",
                                 "SimValueContext:=", [3, 0, 2, 0, False, False, -1, 1, 0, 1, 1, "", 0, 0]
                             ],
                             [
                                 "F:="		, ["All"]
                             ],
                             [
                                 "X Component:="		, "F",
                                 "Y Component:="		, ["mag(Z(DUT0_DUT0_VDD_GPU_0_Group_DUT0_GND_Group,DUT0_DUT0_VDD_GPU_0_Group_DUT0_GND_Group)) ","mag(Z(DUT2_DUT2_VDD_GPU_2_Group_DUT2_GND_Group,DUT2_DUT2_VDD_GPU_2_Group_DUT2_GND_Group))"]
                             ], [])

    def adjust_reports(self):
        self.oReportSetup.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Scaling",
                    [
                        "NAME:PropServers",
                        "VDD_GPU:AxisX"
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Axis Scaling",
                            "Value:="		, "Log"
                        ],
                        [
                            "NAME:Auto Units",
                            "Value:="		, False
                        ],
                        [
                            "NAME:Units",
                            "Value:="		, "MHz"
                        ]
                    ]
                ],
                [
                    "NAME:Scaling",
                    [
                        "NAME:PropServers",
                        "VDD_GPU:AxisY1"
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Axis Scaling",
                            "Value:="		, "Log"
                        ],
                        [
                            "NAME:Max",
                            "Value:="		, "1000mOhm"
                        ],
                        [
                            "NAME:Auto Units",
                            "Value:="		, False
                        ]
                    ]
                ]
            ])

    def save_prj(self) -> object:
        _base_path = os.getcwd()
        _prj_num = 1
        while True:
            _path = os.path.join(_base_path, 'Prj{}.aedt'.format(_prj_num))
            if os.path.exists(_path):
                _prj_num += 1
            else:
                break
        self.oProject.SaveAs(_path, True)


if __name__ == '__main__':
    h = AnsysDesigner()
    currentPath=os.getcwd()+'./VDD_GPU.s4p'
    h.import_model(currentPath, 'VDD_GPU')
    h.add_comp('VDD_GPU')
    h.create_comp('VDD_GPU')
    h.oEditor.ZoomToFit()
    h.assign_port()
    h.insert_analysis_setup()
    h.run_analyze()
    h.create_reports()
    h.adjust_reports()
    h.save_prj()
