' testReadFile.vbs
' test readFile command

dim ptds
dim filePath
set ptds = CreateObject("Sciencetech.PTDS.1")
filePath = "C:\Users\Greg\Documents\Visual Studio 2015\Projects\PTDS\TestWriteHeader\G8V32725.ptds"
ptds.ReadDataFile filePath