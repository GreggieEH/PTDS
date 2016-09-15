' writeHeader.vbs
' test write header

dim ptds
dim hdr
set ptds = CreateObject("Sciencetech.PTDS.1")
ptds.SetCurrentTime
ptds.TimeConstant = 10.0
ptds.Comment = "Gregs stupid test"
ptds.UserName = "Greggie"
ptds.ChopperFrequency = 1.0
ptds.DwellTime = 50.0
ptds.SlitWidth = 5.0
ptds.LampInfo = "Lamp Power 150.0 Watts"
ptds.LockinInfo = "Stanford SR810 42227"
ptds.StepSize = 10.0
ptds.ScanStarted "C:\Users\Greg\Documents\Visual Studio 2015\Projects\PTDS\TestWriteHeader"
ptds.DisplayGrating "Grating 2    1200L/500"
ptds.DisplayFilter "Filter WG-280"
ptds.AddDataValue    300.0,   1.75259e-004
ptds.AddDataValue    310.0,   2.55373e-004
ptds.AddDataValue    320.0,   3.67412e-004
ptds.AddDataValue    330.0,   5.11378e-004
ptds.AddDataValue    340.0,   6.65516e-004
ptds.AddDataValue    350.0,   7.74705e-004
ptds.AddDataValue    360.0,   8.34134e-004
ptds.AddDataValue    370.0,   8.76824e-004
ptds.AddDataValue    380.0,   9.15393e-004
ptds.AddDataValue    390.0,   9.89717e-004
ptds.AddDataValue    400.0,   1.05050e-003
ptds.AddDataValue    410.0,   1.09418e-003
ptds.AddDataValue    420.0,   1.13711e-003
ptds.AddDataValue    430.0,   1.15848e-003
ptds.AddDataValue    440.0,   1.20853e-003
ptds.AddDataValue    450.0,   1.31219e-003
ptds.AddDataValue    460.0,   1.50658e-003
ptds.AddDataValue    470.0,   1.57875e-003
ptds.AddDataValue    480.0,   1.30581e-003

ptds.ScanEnded "Total Run Time = 15 Minutes, 55 Seconds"
msgBox hdr