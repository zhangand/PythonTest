# ----------------------------------------------
# Script Recorded by ANSYS Electronics Desktop Version 2017.2.0
# 14:19:56  Mar 03, 2019
# ----------------------------------------------
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oProject = oDesktop.SetActiveProject("Project1")
oProject.InsertDesign("Circuit Design", "Circuit1", "", "")
oDefinitionManager = oProject.GetDefinitionManager()
oModelManager = oDefinitionManager.GetManager("Model")
oModelManager.Add(
	[
		"NAME:VDD_GPU",
		"Name:="		, "VDD_GPU",
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
		"filename:="		, "E:/Tina Chen/664-828-00/SIWAVE/0301_S-para/VDD_GPU.s4p",
		"numberofports:="	, 4,
		"sssfilename:="		, "",
		"sssmodel:="		, False,
		"PortNames:="		, ["DUT0_DUT0_VDD_GPU_0_Group_DUT0_GND_Group","DUT2_DUT2_VDD_GPU_2_Group_DUT2_GND_Group","UFP_UVS64_100_8_100_UFP_UVS64_100_8_100_VDD_GPU_2_Group_UFP_UVS64_100_8_100_GND_Group","UFP_UVS64_100_9_100_UFP_UVS64_100_9_100_VDD_GPU_0_Group_UFP_UVS64_100_9_100_GND_Group"],
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
oComponentManager = oDefinitionManager.GetManager("Component")
oComponentManager.Add(
	[
		"NAME:VDD_GPU",
		"Info:="		, [			"Type:="		, 10,			"NumTerminals:="	, 4,			"DataSource:="		, "",			"ModifiedOn:="		, 1551594059,			"Manufacturer:="	, "",			"Symbol:="		, "",			"ModelNames:="		, "",			"Footprint:="		, "",			"Description:="		, "",			"InfoTopic:="		, "",			"InfoHelpFile:="	, "",			"IconFile:="		, "nport.bmp",			"Library:="		, "",			"OriginalLocation:="	, "Project",			"IEEE:="		, "",			"Author:="		, "",			"OriginalAuthor:="	, "",			"CreationDate:="	, 1551594059,			"ExampleFile:="		, "",			"HiddenComponent:="	, 0],
		"Refbase:="		, "S",
		"NumParts:="		, 1,
		"ModSinceLib:="		, False,
		"CircuitEnv:="		, 0,
		"Terminal:="		, ["DUT0_DUT0_VDD_GPU_0_Group_DUT0_GND_Group","DUT0_DUT0_VDD_GPU_0_Group_DUT0_GND_Group","A",False,0,1,"","Electrical","0"],
		"Terminal:="		, ["DUT2_DUT2_VDD_GPU_2_Group_DUT2_GND_Group","DUT2_DUT2_VDD_GPU_2_Group_DUT2_GND_Group","A",False,1,1,"","Electrical","0"],
		"Terminal:="		, ["UFP_UVS64_100_8_100_UFP_UVS64_100_8_100_VDD_GPU_2_Group_UFP_UVS64_100_8_100_GND_Group","UFP_UVS64_100_8_100_UFP_UVS64_100_8_100_VDD_GPU_2_Group_UFP_UVS64_100_8_100_GND_Group","A",False,2,1,"","Electrical","0"],
		"Terminal:="		, ["UFP_UVS64_100_9_100_UFP_UVS64_100_9_100_VDD_GPU_0_Group_UFP_UVS64_100_9_100_GND_Group","UFP_UVS64_100_9_100_UFP_UVS64_100_9_100_VDD_GPU_0_Group_UFP_UVS64_100_9_100_GND_Group","A",False,3,1,"","Electrical","0"],
		"CompExtID:="		, 5,
		[
			"NAME:Parameters",
			"MenuProp:="		, ["CoSimulator","SD","","Default",0],
			"ButtonProp:="		, ["CosimDefinition","SD","","Edit","Edit",40501,				"ButtonPropClientData:=", []]
		],
		[
			"NAME:CosimDefinitions",
			[
				"NAME:CosimDefinition",
				"CosimulatorType:="	, 102,
				"CosimDefName:="	, "Default",
				"IsDefinition:="	, True,
				"Connect:="		, True,
				"ModelDefinitionName:="	, "VDD_GPU",
				"ShowRefPin2:="		, 2,
				"LenPropName:="		, ""
			],
			"DefaultCosim:="	, "Default"
		]
	])
oDesign = oProject.SetActiveDesign("Circuit1")
oEditor = oDesign.SetActiveEditor("SchematicEditor")
oEditor.CreateComponent(
	[
		"NAME:ComponentProps",
		"Name:="		, "VDD_GPU",
		"Id:="			, "1"
	], 
	[
		"NAME:Attributes",
		"Page:="		, 1,
		"X:="			, 0.0889,
		"Y:="			, 0.05588,
		"Angle:="		, 0,
		"Flip:="		, False
	])
oEditor.CreateIPort(
	[
		"NAME:IPortProps",
		"Name:="		, "Port1",
		"Id:="			, 3
	], 
	[
		"NAME:Attributes",
		"Page:="		, 1,
		"X:="			, 0.05334,
		"Y:="			, 0.06096,
		"Angle:="		, 0,
		"Flip:="		, False
	])
oEditor.CreateIPort(
	[
		"NAME:IPortProps",
		"Name:="		, "Port2",
		"Id:="			, 7
	], 
	[
		"NAME:Attributes",
		"Page:="		, 1,
		"X:="			, 0.0381,
		"Y:="			, 0.06096,
		"Angle:="		, 0,
		"Flip:="		, False
	])
oEditor.CreateGround(
	[
		"NAME:GroundProps",
		"Id:="			, 11
	], 
	[
		"NAME:Attributes",
		"Page:="		, 1,
		"X:="			, 0.1524,
		"Y:="			, 0.04318,
		"Angle:="		, 0,
		"Flip:="		, False
	])
oEditor.CreateWire(
	[
		"NAME:WireData",
		"Name:="		, "",
		"Id:="			, 15,
		"Points:="		, ["(0.104140, 0.053340)","(0.152400, 0.053340)","(0.152400, 0.045720)"]
	], 
	[
		"NAME:Attributes",
		"Page:="		, 1
	])
oEditor.CreateWire(
	[
		"NAME:WireData",
		"Name:="		, "",
		"Id:="			, 19,
		"Points:="		, ["(0.104140, 0.055880)","(0.152400, 0.055880)","(0.152400, 0.053340)"]
	], 
	[
		"NAME:Attributes",
		"Page:="		, 1
	])
oEditor.CreateWire(
	[
		"NAME:WireData",
		"Name:="		, "",
		"Id:="			, 23,
		"Points:="		, ["(0.073660, 0.055880)","(0.053340, 0.055880)","(0.053340, 0.060960)"]
	], 
	[
		"NAME:Attributes",
		"Page:="		, 1
	])
oEditor.CreateWire(
	[
		"NAME:WireData",
		"Name:="		, "",
		"Id:="			, 27,
		"Points:="		, ["(0.073660, 0.053340)","(0.038100, 0.053340)","(0.038100, 0.060960)"]
	], 
	[
		"NAME:Attributes",
		"Page:="		, 1
	])
oModule = oDesign.GetModule("SimSetup")
oModule.AddLinearNetworkAnalysis(
	[
		"NAME:SimSetup",
		"DataBlockID:="		, 16,
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
		"LinearFrequencyData:="	, [False,0.1,False,"",False],
		[
			"NAME:SweepDefinition",
			"Variable:="		, "F",
			"Data:="		, "DEC 1kHz 1GHz 301",
			"OffsetF1:="		, False,
			"Synchronize:="		, 0
		]
	])
oDesign.Analyze("LinearFrequency")
oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport("Z Parameter Plot 1", "Standard", "Rectangular Plot", "LinearFrequency", 
	[
		"NAME:Context",
		"SimValueContext:="	, [3,0,2,0,False,False,-1,1,0,1,1,"",0,0]
	], 
	[
		"F:="			, ["All"]
	], 
	[
		"X Component:="		, "F",
		"Y Component:="		, ["mag(Z(Port1,Port1))","mag(Z(Port2,Port2))"]
	], [])
oModule.RenameReport("Z Parameter Plot 1", "VDD_GPU")
oModule.ChangeProperty(
	[
		"NAME:AllTabs",
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
				]
			]
		],
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
					"NAME:Auto Units",
					"Value:="		, False
				],
				[
					"NAME:Max",
					"Value:="		, "1000mOhm"
				]
			]
		]
	])
oModule.ExportToFile("VDD_GPU", "E:/Tina Chen/664-828-00/SIWAVE/0301_S-para/VDD_GPU.csv")
oProject.SaveAs("E:\\Tina Chen\\test\\Project1.aedt", True, "", 
	[
		"NAME:OverrideActions",
		[
			"NAME:ef_overwrite",
			[
				"NAME:Files", 
				"$PROJECTDIR\\project1.aedtresults\\circuit1\\temp\\dv9_s7_v11.cir"
			]
		]
	])
oModule.ExportToFile("VDD_GPU", "E:/Tina Chen/test/VDD_GPU.csv")
oModule.ExportImageToFile("VDD_GPU", "E:/Tina Chen/test/VDD_GPU.bmp", 0, 0)
oProject.SaveProjectArchive("E:\\Tina Chen\\test\\Project1.aedtz", True, False, [], "")
