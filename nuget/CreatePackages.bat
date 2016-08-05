@echo off

@mkdir packages

nuget pack ..\src\ClassToExcel.OpenXml\ClassToExcel.OpenXml.csproj -prop Configuration=Release -Build  -OutputDirectory .\packages
