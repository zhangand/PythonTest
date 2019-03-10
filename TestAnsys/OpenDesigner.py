from Gen_ppt import gen_ppt
from win32com import client
import os
import fnmatch
import sys


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
        self.oCreateVar = self.oDesign.GetModule("OutputVariable")

    def import_model(self, _file_path, _comp_name, _pin_list):
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
                "PortNames:="		, _pin_list,
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

    def create_comp(self, _comp_name, _in_comp_id, _in_page_num, _comp_x, _comp_y):
        inCompId=self.oEditor.CreateComponent(
            [
                "NAME:ComponentProps",
                "Name:="	, _comp_name,
                "Id:="			, _in_comp_id
            ],
            [
                "NAME:Attributes",
                "Page:="		, _in_page_num,
                "X:="			, _comp_x,
                "Y:="			, _comp_y,
                "Angle:="		, 0,
                "Flip:="		, False
            ])
        return inCompId

    def get_comp_pininfo(self, in_comp_id):
        comp_pins = self.oEditor.GetComponentPins(in_comp_id)
        for pin in comp_pins :
            comp_pin_info = self.oEditor.GetComponentPinInfo(in_comp_id, pin)
            x_temp = str(comp_pin_info[0]).split('=')
            x = x_temp[1]
            y_temp = str(comp_pin_info[1]).split('=')
            y = y_temp[1]
            angle_temp = str(comp_pin_info[2]).split('=')
            angle = float(angle_temp[1])
            flip_temp = str(comp_pin_info[3]).split('=')
            flip = flip_temp[1]
            
#oEditor.AddPinPageconnectors(
#	[
#		"NAME:Selections",
#		"Selections:="		, ["CompInst@vdd_gpu;1;1"]
#	])
    
    def add_page_conn(self, in_comp_id):
        self.oEditor.AddPinPageconnectors(
            [
                "NAME:Selections",
                "Selections:="		, [in_comp_id]
            ])

#    oEditor.CreatePagePort(
#        [
#            "NAME:PagePortProps",
#            "Name:="	, "pageport_0",
#            "Id:="			, 20
#        ],
#        [
#            "NAME:Attributes",
#            "Page:="		, 2,
#            "X:="			, 0.0508,
#            "Y:="			, 0.06858,
#            "Angle:="		, 0,
#            "Flip:="		, False
#        ])

    def create_page_port(self, _pageport_name, _pageport_num, _sch_page_num, _pageport_x, _pageport_y, ):
        _page_port_id = self.oEditor.CreatePagePort(
            [
                "NAME:PagePortProps",
                "Name:="	, _pageport_name,
                "Id:="			, _pageport_num
            ],
            [
                "NAME:Attributes",
                "Page:="		, _sch_page_num,
                "X:="			, _pageport_x,
                "Y:="			, _pageport_y,
                "Angle:="		, 0,
                "Flip:="		, False
            ])
        self.oEditor.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:PassedParameterTab",
                    [
                        "NAME:PropServers",
                        _page_port_id
                    ],
                    [
                        "NAME:PortName",
                        "ExtraText:="	, _pageport_name,
                        "Name:="		, _pageport_name,
                        "SplitWires:="		, False
                    ]
                ]
            ])

# oEditor.CreateIPort(
#	[
#		"NAME:IPortProps",
#		"Name:="		, "Port1",
#		"Id:="			, 26
#	],
#	[
#		"NAME:Attributes",
#		"Page:="		, 2,
#		"X:="			, 0.0508,
#		"Y:="			, 0.06858,
#		"Angle:="		, 0,
#		"Flip:="		, False
#	])

    def create_iport(self, i_port_name, port_id, sch_page_num, iport_x, iport_y):
        self.oEditor.CreateIPort(
            [
                "NAME:IPortProps",
                "Name:="		, i_port_name,
                "Id:="			, port_id
            ],
            [
                "NAME:Attributes",
                "Page:="		, sch_page_num,
                "X:="			, iport_x,
                "Y:="			, iport_y,
                "Angle:="		, 0,
                "Flip:="		, False
            ])

    def create_resistor(self, _res_id, _sch_page_num, _res_x, _res_y):
        _full_res_id = self.oEditor.CreateComponent(
            [
                "NAME:ComponentProps",
                "Name:="		, "Nexxim Circuit Elements\\Resistors:RES_",
                "Id:="			, _res_id
            ],
            [
                "NAME:Attributes",
                "Page:="		, _sch_page_num,
                "X:="			, _res_x,
                "Y:="			, _res_y,
                "Degrees:="		, 0,
                "Flip:="		, False
            ])

        self.oEditor.ChangeProperty(
            [
                "NAME:AllTabs",
                    [
                        "NAME:PassedParameterTab",
                            [
                                "NAME:PropServers",
                                _full_res_id
                            ],
                            [
                                "NAME:ChangedProps",
                                [
                                    "NAME:R",
                                    "OverridingDef:="	, True,
                                    "Value:="		, "0.0001"
                                ]
                            ]
                    ]
            ])

    def create_gnd(self, _gnd_id, _sch_page_num, _res_x, _res_y):
        self.oEditor.CreateGround(
            [
                "NAME:GroundProps",
                "Id:="			,  _gnd_id
            ],
            [
                "NAME:Attributes",
                "Page:="		, _sch_page_num,
                "X:="			, _res_x,
                "Y:="			, _res_y,
                "Degrees:="		, 0,
                "Flip:="		, False
            ])

# oEditor.CreateWire(
#	[
#		"NAME:WireData",
#		"Name:="		, "",
#		"Id:="			, 10041,
#		"Points:="		, ["(0.045720, -0.050800)","(0.000000, -0.050800)"]
#	],
#	[
#		"NAME:Attributes",
#		"Page:="		, 2
#	])

    def create_wire(self, _wire_id, _sch_page_num, _res_x1, _res_y1, _res_x2, _res_y2,  ):
        _location1 = "(%(_x1)f, %(_y1)f)" % {'_x1': _res_x1, '_y1': _res_y1}
        _location2 = "(%(_x2)f, %(_y2)f)" % {'_x2': _res_x2, '_y2': _res_y2}
        self.oEditor.CreateWire(
            [
                "NAME:WireData",
                "Name:="	, "",
                "Id:="			,  _wire_id,
                "Points:="		, [_location1, _location2]
            ],
            [
                "NAME:Attributes",
                "Page:="		, _sch_page_num,
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
                    "Data:="		, freq_range + table_freq_point_str,
                    "OffsetF1:="		, False,
                    "Synchronize:="		, 0
                ]
            ])

    def run_analyze(self):
        self.oDesign.Analyze("LinearFrequency")

    def create_pdn_plot(self,_report_name, _ports_list):
        self.oReportSetup.CreateReport(_report_name, "Standard", "Rectangular Plot", "LinearFrequency",
             [
                 "NAME:Context",
                 "SimValueContext:=", [3, 0, 2, 0, False, False, -1, 1, 0, 1, 1, "", 0, 0]
             ],
             [
                 "F:="		, ["All"]
             ],
             [
                 "X Component:="		, "F",
                 "Y Component:="		, _ports_list
             ], [])
        for i in range(0, len(_ports_list) , 1):
            self.oReportSetup.ChangeProperty(
                [
                    "NAME:AllTabs",
                    [
                        "NAME:Trace",
                        [
                            "NAME:PropServers",
                            _report_name+ ":"+_ports_list[i]
                        ],
                        [
                            "NAME:ChangedProps",
                            [
                                "NAME:Name",
                                "Value:="		, short_port_name_list[i]
                            ]
                        ]
                    ]
                ])

    def create_var(self,_report_name, _ports_list):
        Z_min = _ports_list[0]
        Z_max = _ports_list[0]
        Z_avg = '(' + _ports_list[0]
        for i in range(0, len(_ports_list) - 1, 1):
            Z_min = 'min(' + Z_min + ',' + _ports_list[i + 1] + ')'
            Z_max = 'max(' + Z_max + ',' + _ports_list[i + 1] + ')'
            Z_avg = Z_avg + '+' + _ports_list[i + 1]

        Z_avg = Z_avg + ')/' + str(len(_ports_list))

        self.oCreateVar.CreateOutputVariable("Avg_"+str(_report_name), str(Z_avg), "LinearFrequency", "Standard",
            [
                "NAME:Context",
                "SimValueContext:="	, [3,0,2,0,False,False,-1,1,0,1,1,"",0,0]
             ])
        self.oCreateVar.CreateOutputVariable("Min_"+str(_report_name), str(Z_min), "LinearFrequency", "Standard",
            [
                "NAME:Context",
                "SimValueContext:="	, [3,0,2,0,False,False,-1,1,0,1,1,"",0,0]
            ])
        self.oCreateVar.CreateOutputVariable("Max_"+str(_report_name), str(Z_max), "LinearFrequency", "Standard",
            [
                "NAME:Context",
                "SimValueContext:="	, [3,0,2,0,False,False,-1,1,0,1,1,"",0,0]
            ])
        self.oCreateVar.CreateOutputVariable("Dev_"+str(_report_name),
                                        "(Max_"+str(_report_name) + '-' + "Min_"+str(_report_name) + ')/' 
                                        "Min_"+str(_report_name) + '*100', "LinearFrequency", "Standard",
            [
                "NAME:Context",
                "SimValueContext:="	, [3,0,2,0,False,False,-1,1,0,1,1,"",0,0]
            ])

    def create_pdn_table(self, _report_name, _ports_list, _table_freq_point):
        z_table_port_list = []
        z_table_port_list.extend(_ports_list)
        z_table_port_list.append("Avg_" + str(_report_name))
        z_table_port_list.append("Dev_" + str(_report_name))
        self.oReportSetup.CreateReport(_report_name + " Z11 table", "Standard", "Data Table", "LinearFrequency",
            [
                "NAME:Context",
                "SimValueContext:="	, [3,0,2,0,False,False,-1,1,0,1,1,"",0,0]
            ],
            [
                "F:="			, _table_freq_point
            ],
            [
                "X Component:="		, "F",
                "Y Component:="		, z_table_port_list
            ], [])
        self.oReportSetup.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Data Table",
                    [
                        "NAME:PropServers",
                        _report_name + " Z11 table:DataTableDisplayTypeProperty"
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Transpose",
                            "Value:="		, True
                        ]
                    ]
                ]
            ])
        for i in range(0, len(z_table_port_list) - 1, 1):
            self.oReportSetup.ChangeProperty(
                [
                    "NAME:AllTabs",
                    [
                        "NAME:Data Filter",
                        [
                            "NAME:PropServers",
                            str(_report_name) + " Z11 table:"+ str(z_table_port_list[i]) + ":Curve1"
                        ],
                        [
                            "NAME:ChangedProps",
                            [
                                "NAME:Units",
                                "Value:="	, "mOhm"
                            ]
                        ]
                    ]
                ])
            self.oReportSetup.ChangeProperty(
                [
                    "NAME:AllTabs",
                    [
                        "NAME:Data Filter",
                        [
                            "NAME:PropServers",
                            str(_report_name) + " Z11 table:"+ str(z_table_port_list[i]) + ":Curve1"
                        ],
                        [
                            "NAME:ChangedProps",
                            [
                                "NAME:Field Precision",
                                "Value:="	, "2"
                            ]
                        ]
                    ]
                ])
        self.oReportSetup.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Data Filter",
                    [
                        "NAME:PropServers",
                        str(_report_name) + " Z11 table:PrimarySweepDrawing",
                        str(_report_name) + " Z11 table:Dev_" + str(_report_name) + ":Curve1"
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Field Precision",
                            "Value:="	, "2"
                        ]
                    ]
                ]
            ])
        for i in range(0, len(z_table_port_list)-2, 1):
            self.oReportSetup.ChangeProperty(
                [
                    "NAME:AllTabs",
                    [
                        "NAME:Trace",
                        [
                            "NAME:PropServers",
                            _report_name + " Z11 table"+":"+z_table_port_list[i]
                        ],
                        [
                            "NAME:ChangedProps",
                            [
                                "NAME:Name",
                                "Value:="		, short_port_name_list[i]
                            ]
                        ]
                    ]
                ])

        self.oReportSetup.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Trace",
                    [
                        "NAME:PropServers",
                        str(_report_name) + " Z11 table:Avg_" + str(_report_name)
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Name",
                            "Value:=", "Avg"
                        ]
                    ]
                ]
            ])

        self.oReportSetup.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Trace",
                    [
                        "NAME:PropServers",
                        str(_report_name) + " Z11 table:Dev_" + str(_report_name)
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Name",
                            "Value:="	, "Dev %"
                        ]
                    ]
                ]
            ])

    def adjust_reports(self,_report_name):
        self.oReportSetup.ChangeProperty(
        [
            "NAME:AllTabs",
            [
                "NAME:Scaling",
                [
                    "NAME:PropServers",
                    _report_name+":AxisX"
                ],
                [
                    "NAME:ChangedProps",
                    [
                        "NAME:Axis Scaling",
                        "Value:="	, "Log"
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
            ]
        ])
        self.oReportSetup.ChangeProperty(
        [
            "NAME:AllTabs",
            [
                "NAME:Scaling",
                [
                    "NAME:PropServers",
                    _report_name + ":AxisY1"
                ],
                [
                    "NAME:ChangedProps",
                    [
                        "NAME:Axis Scaling",
                        "Value:="		, "Log"
                    ],
                    [
                        "NAME:Min",
                        "Value:="		, "0.1mOhm"
                    ],
                    [
                        "NAME:Auto Units",
                        "Value:="		, False
                    ]
                ]
            ]
        ])
        self.oReportSetup.ChangeProperty(
        [
            "NAME:AllTabs",
            [
                "NAME:Scaling",
                [
                    "NAME:PropServers",
                    _report_name + ":AxisY1"
                ],
                [
                    "NAME:ChangedProps",
                    [
                        "NAME:Max",
                        "Value:="	, "1000mOhm"
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


class ReadFile:
    # Read port information in SNP
    def read_port_info(self, file_path, snp_ext_name):
        global port_count
        global comp_name_list
        j = 0
        for full_filename in self.find_files(file_path, snp_ext_name):
            file_list[j] = full_filename
            temp = os.path.split(full_filename)
            temp1 = temp[-1].split('.')
            comp_name_list.append(str.upper(temp1[0]))

            snp_file = open(full_filename, 'r')
            flag_start_search = 0
            i = 0
            for line in snp_file:
                if line.find('Port[') != -1:
                    temp_line_list = line.split('=')
                    temp_word = temp_line_list[-1].replace(" ", "")
                    temp_word = temp_word.replace("\n", "")
                    port_name_list[j].append(temp_word)
                    i = i + 1
                    flag_start_search = 1
                elif (line.find('Port[') == -1) and (flag_start_search == 1):
                    flag_start_search = 0
                    port_count = port_count + i
                    break
            snp_file.close()
            j = j + 1

    def find_files(self, directory, pattern) -> object:
        for file in os.listdir(directory):
                if os.path.isfile(file) and fnmatch.fnmatch(file, pattern):
                    full_filename = os.path.join(directory, file)
                    yield full_filename


if __name__ == '__main__':
    #  Define static variable
    table_freq_point_str = " 0.01GHz 0.05GHz 0.1GHz 0.2GHz 0.5GHz"
    freq_range = 'DEC 10KHZ 100MHZ 301'
    ext_name = '*.s*p'

    file_list = {}
    comp_name_list = []
    port_name_list = []
    short_port_name_list = []
    report_ports_list = []
    #z_table_port_list = []
    port_name_list.append([])
    port_name_list.append([])

    pageport_x = 0
    pageport_y = 0

    comp_x = 0
    comp_y = 0

    iport_x = 0
    iport_y = 0

    port_count = 0
    comp_id = 1000
    page_port_id = 5000
    iport_id = 10000

    page_num = 1

    currentPath=os.getcwd()
    #  End Define static variable
    rf=ReadFile()
    rf.read_port_info(currentPath, ext_name)

    h = AnsysDesigner()
    k = 0
    for comp_name in comp_name_list:
        h.import_model(file_list[k], comp_name, port_name_list[k])
        h.add_comp(comp_name)
        full_comp_id = h.create_comp(comp_name, comp_id, page_num, comp_x, comp_y)
        h.add_page_conn(full_comp_id)
        comp_y = comp_x - 0.0254
        k = k + 1
        comp_id = comp_id + 1

    page_num = page_num + 1
    h.oEditor.CreatePage("<Page Title>")
    m = 0
    for comp_name in comp_name_list:
        for port_name in port_name_list[m]:
            if port_name.startswith("DUT"):
                h.create_page_port(port_name, page_port_id, page_num, pageport_x, pageport_y)
                h.create_iport(port_name, iport_id, page_num, iport_x, iport_y)
                report_ports_list.append('Mag(Z(' + port_name + ',' + port_name + '))')
                short_port_name_list.append(port_name.split('_')[0])
                page_port_id = page_port_id +1
                iport_id = iport_id + 1
                pageport_y = pageport_y - 0.0254
                iport_y = iport_y - 0.0254
            else:
                h.create_page_port(port_name, page_port_id, page_num, pageport_x, pageport_y)
                h.create_resistor(comp_id, page_num, pageport_x+0.0254*5, pageport_y)
                comp_id = comp_id + 1
                h.create_gnd(comp_id, page_num, pageport_x+0.0254*5+0.00508, pageport_y-0.00254)
                comp_id = comp_id + 1
                h.create_wire(comp_id, page_num, pageport_x, pageport_y, pageport_x+0.0254*5-0.00508, pageport_y)
                comp_id = comp_id + 1

                page_port_id = page_port_id +1
                iport_id = iport_id + 1
                pageport_y = pageport_y - 0.0254
                iport_y = iport_y - 0.0254
        h.oEditor.ZoomToFit()

        h.insert_analysis_setup()
        h.run_analyze()
        h.create_pdn_plot(comp_name, report_ports_list)
        h.adjust_reports(comp_name)
        h.oDesktop.RestoreWindow()
        h.oReportSetup.ExportImageToFile(str(comp_name), str(currentPath) + '\\' + str(comp_name) + ".jpg", 1000, 400)
        h.create_var(comp_name, report_ports_list)
        h.create_pdn_table(comp_name, report_ports_list, ["0.01GHz","0.05GHz","0.1GHz","0.2GHz","0.5GHz"])
        h.oReportSetup.ExportToFile(str(comp_name) + " Z11 table", currentPath + '\\' + str(comp_name) + "_Z11_table.csv")

ppt = gen_ppt()
ppt.do_ppt()
    #h.get_comp_pininfo(compid)
'''
    
    h.save_prj()
'''