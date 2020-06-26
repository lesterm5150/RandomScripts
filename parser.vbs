'Put your files that need searched in a folder named Files in the same directory'
' as this script. '

Option Explicit

Dim objFSO, objFolder, strFolder, objFile
Dim objReadFile, strLine, objExcel, objSheet, objWrkBk
Dim intRow, intCol, strExcelPath

Const ForReading = 1

Dim percentComplete
Dim numberoffiles
Dim count , index

'!!!The folder that you have your files in'
strFolder = ".\Files"
Dim CurPath: CurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
strExcelPath = CurPath & "\Data.xlsx"
Set objExcel = CreateObject("Excel.Application")
set objWrkBk = objExcel.Workbooks.Add
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
intRow = 1


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolder)
numberoffiles = 0
count = 0
index = 1
For Each objFile In objFolder.Files

' If there are more jornals than sheets, this creates a new sheet, and focuses on it
if index > 3 then Set objSheet = objWrkBk.Sheets.Add(, objWrkBk.Sheets(objWrkBk.Sheets.Count))

' As long as the journal # is less than or equal to the numbers of sheets, then don't create more
if index <= 3 then set objSheet = objExcel.ActiveWorkbook.Worksheets(index)

	numberoffiles = objFolder.Files.Count
	
	'!!!Example of how to put values in excel '
objSheet.Cells(1,1).Value = "Journal Name"
objSheet.Cells(1,1).Font.Bold = True
objSheet.Cells(2,1).Value = "Start Time"



' Autofit the cells 
objSheet.Columns.AutoFit
objSheet.Rows.AutoFit


' Name the Sheet of the WorkBook
objExcel.ActiveWorkbook.Sheets(index).Name = objFile.Name
index = index + 1
dim counter : counter = 0

  intRow = intRow + 1
  count = count + 1
  
  ' Run through each file and pull out the specific strings'
  Set objReadFile = objFSO.OpenTextFile(objFile.Path, ForReading)
  Do Until objReadFile.AtEndOfStream
    strLine = objReadFile.Readline
	
	'!!!you would replace "second" with the phrase you are trying to find'
	If (InStr(strLine, "second") > 0) Then
	  objSheet.Cells(1, 2).Value = mid(strLine,1)
	End If
  Loop
Next
' Align all columns to left
objExcel.Columns.HorizontalAlignment = -4131

' Save and close workbook : Will not ask to save over previous file
objExcel.DisplayAlerts = False
objExcel.ActiveWorkbook.SaveAs strExcelPath
objExcel.DisplayAlerts = True
objExcel.ActiveWorkbook.Close
objExcel.Quit
wscript.quit
