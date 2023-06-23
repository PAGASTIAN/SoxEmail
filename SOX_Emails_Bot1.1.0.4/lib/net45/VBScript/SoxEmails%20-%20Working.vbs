Dim Dict
'Excel Object Initiate
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Application.DisplayAlerts = False
objExcel.Application.EnableEvents=False
objExcel.Application.AskToUpdateLinks = False
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'File Path's
Dim FolderPath
FolderPath = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If(fso.FolderExists(FolderPath))Then
	Set NFolder = fso.GetFolder(FolderPath)
		Set colFiles = NFolder.Files
		For Each FileItem in colFiles
			If(Instr(FileItem.Name,"MRI PA Project Status Report Template")>0)Then
				TemplateFile  = FileItem.Name
				TemplateFile = FolderPath & "\" & TemplateFile
			End If
			If(Instr(FileItem.Name,"MRI PA Project Status Report_Project Status Report")>0)Then
				OracleReport  = FileItem.Name
				OracleReport = FolderPath & "\" & OracleReport
			End If
		Next
		Set colFiles = Nothing
	
Else
	MsgBox("File Not Found" & FolderPath)
End If

'Opening Oracle Report
Set objWBOR = objExcel.Workbooks.Open(OracleReport)
Set objWSOR = objWBOR.Worksheets(2)
objWSOR.Activate

'Opening Previous Month File
'Set objWBPMF = objExcel.Workbooks.Open(PreviousMonthFile)
'Set objWSPMF = objWBPMF.Worksheets(1)

'Unwarpping & Unmerging cells in Downloaded Report
Set objRange = objWSOR.UsedRange
objRange.WrapText = False
objRange.MergeCells = False

'Inserting Columns
objWSOR.Columns("U:U").Insert xlToLeft
objWSOR.Cells(6,21).Value = "Actual Cost MM.YYYY"
objWSOR.Columns("V:V").Insert xlToLeft
objWSOR.Cells(6,22).Value = "No Activity 60 Days"

'Openings Template File
Set objWBTF = objExcel.Workbooks.Open(TemplateFile)
Set objWSTF = objWBTF.Worksheets(1)

'Copying pasting current month data to Template File
LastRowOR = objWSOR.Range("A1048576").end(-4162).row
objWSOR.Range("A7:T" & LastRowOR).specialcells(12).Copy
objWSTF.Range("A2").PasteSpecial 12
objWSOR.Range("W7:Y" & LastRowOR).specialcells(12).Copy
objWSTF.Range("W2").PasteSpecial 12
objWBOR.Close

'Applying Formulae from Z to AD Columns
LastRowTF = objWSTF.Range("A1048576").end(-4162).row
objWSTF.Cells(2,22).Formula = "=U2=T2"
Set range = objWSTF.Range("V2:V"& LastRowTF)
objWSTF.Range("V2:V"& LastRowTF).SpecialCells(12).Formula = "=U2=T2"

'Applying Filter
Dim StatusArray
StatusArray=Split(Replace(WScript.Arguments(9),"_"," "),"|")
objWSTF.Range("A1").AutoFilter 1, StatusArray, 7

MRDD_Filter = Trim(Replace(WScript.Arguments(10),"_", " "))
objWSTF.Range("A1").AutoFilter 7, MRDD_Filter, 7

Dim PropertyArray
PropertyArray=Split(Replace(WScript.Arguments(8),"_"," "),"|")
objWSTF.Range("A1").AutoFilter 6, PropertyArray, 7

objWSTF.Range("A1").AutoFilter 22, "True"
'objWSTF.Range("A1").AutoFilter 22, "False"
LastRowTF = objWSTF.Range("A1048576").end(-4162).row
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)
objWSTF.Range("A1:AD"& LastRowTF).SpecialCells(12).Copy
objWorksheet.Range("A1").PasteSpecial
objWorkbook.SaveAs WScript.Arguments(1) & "\MRDD SOX Control Review - " & WScript.Arguments(11)
objWorkbook.Close
objWSTF.AutoFilterMode = False
objWSTF.Range("A1").AutoFilter 1, StatusArray, 7
objWSTF.Range("A1").AutoFilter 7, MRDD_Filter, 7
VisibleRow = objWSTF.AutoFilter.Range.Offset(1).SpecialCells(12).Row
LastRowTF = objWSTF.Range("A1048576").end(-4162).row
Set range = objWSTF.Range("A" & VisibleRow & ":AD"& LastRowTF).Specialcells(12)
range.EntireRow.Delete
objWSTF.AutoFilterMode = False

StatusArray=Split(Replace(WScript.Arguments(9),"_"," "),"|")
PropertyArray=Split(Replace(WScript.Arguments(8),"_"," "),"|")

objWSTF.Range("A1").AutoFilter 1, StatusArray, 7
objWSTF.Range("A1").AutoFilter 6, PropertyArray, 7

objWSTF.Range("A1").AutoFilter 22, "True"
'objWSTF.Range("A1").AutoFilter 22, "False"

Set Dict = CreateObject("Scripting.Dictionary")
LastRowTF = objWSTF.Range("A1048576").end(-4162).row
VisibleRow = objWSTF.AutoFilter.Range.Offset(1).SpecialCells(12).Row
uniqarr = objWSTF.Range("H" & VisibleRow & ":H" & LastRowTF).Value

For each val in uniqarr
 Dict(val) = val
Next
newarray = Dict.Items

For i = LBound(newarray) to UBound(newarray)
    objWSTF.Range("A1").AutoFilter 8, newarray(i)
	VisibleRow = objWSTF.AutoFilter.Range.Offset(1).SpecialCells(12).Row
	LastRowTF = objWSTF.Range("A1048576").end(-4162).row
	If VisibleRow >= 2 Then
	If(Len(objWSTF.Cells(VisibleRow,1).Value))>0 Then
	Set objWorkbook = objExcel.Workbooks.Add()
	Set objWorksheet = objWorkbook.Worksheets(1)
	objWSTF.Range("A1:AD"& LastRowTF).Copy
        objWorksheet.Range("A1").PasteSpecial
	Value = Left(newarray(i),4) & " SOX Control Review"
	objWorkbook.SaveAs WScript.Arguments(1) & "\" & Value & " - " & WScript.Arguments(11)
	objWorkbook.Close
	End If
	End If
Next

objWBTF.Close

'Dim Filtername as string=Filtername_Array
newarr = Split(WScript.Arguments(6),"|")

'Dim Mappingname as string=Mappingname_Array
newarr1 = Split(WScript.Arguments(7),"|")

Set fso = CreateObject("Scripting.FileSystemObject")
FolderPath = WScript.Arguments(1)

If(fso.FolderExists(FolderPath))Then
	Set NFolder = fso.GetFolder(FolderPath)
		Set colFiles = NFolder.Files
		For Each FileItem in colFiles
			If(Instr(FileItem.Name,"2001")>0)Then
				File2001  = FileItem.Name
				File2001 = FolderPath & "\" & File2001
				Exit For
			End If
		Next
		Set colFiles = Nothing
	
Else
	MsgBox("File Not Found" & FolderPath)
End If

Set objWB2001 = objExcel.Workbooks.Open(File2001)
Set objWS2001 = objWB2001.Worksheets(1)

	
For i = LBound(newarr) to UBound(newarr)
    j = i
	
    objWS2001.Range("A1").AutoFilter 7, Replace(newarr(i),"_"," ")
	VisibleRow = objWS2001.AutoFilter.Range.Offset(1).SpecialCells(12).Row
	LastRow2001 = objWS2001.Range("A1048576").end(-4162).row
	If VisibleRow >= 2 Then
	If(Len(objWS2001.Cells(VisibleRow,1).Value))>0 Then
	Set objWorkbook = objExcel.Workbooks.Add()
	objWorkbook.SaveAs WScript.Arguments(1) & "\2001 " & Replace(newarr1(j),"_"," ") & " SOX Control Review - " & WScript.Arguments(11)
	Set objWorksheet = objWorkbook.Worksheets(1)
	objWS2001.Range("A1:AD"& LastRow2001).Copy
    objWorksheet.Range("A1").PasteSpecial
	objWorkbook.Save
	objWorkbook.Close
	End If
	End If
Next
objWB2001.Close


'Dim FolderPath
Set fso = CreateObject("Scripting.FileSystemObject")
FolderPath = WScript.Arguments(1)
If(fso.FolderExists(FolderPath))Then
	Set NFolder = fso.GetFolder(FolderPath)
		Set colFiles = NFolder.Files
		For Each FileItem in colFiles
			If(Instr(FileItem.Name,"1000")>0)Then
				File1  = FileItem.Name
				File1 = FolderPath & "\" & File1
			End If
		    If(Instr(FileItem.Name,"1001")>0)Then
				File2  = FileItem.Name
				File2 = FolderPath & "\" & File2
			End If	
			If(Instr(FileItem.Name,"1003")>0)Then
				File3  = FileItem.Name
				File3 = FolderPath & "\" & File3
			End If
			If(Instr(FileItem.Name,"2061")>0)Then
				File4  = FileItem.Name
				File4 = FolderPath & "\" & File4
			End If
		Next
		Set colFiles = Nothing
	
Else
	MsgBox("File Not Found" & FolderPath)
End If

'Merging 4 Properties
Set objWBF1 = objExcel.Workbooks.Open(File1)
Set objWSF1 = objWBF1.Worksheets(1)

LastRowF1 = objWSF1.Range("A1048576").end(-4162).row
Set objWBF2 = objExcel.Workbooks.Open(File2)
Set objWSF2 = objWBF2.Worksheets(1)
LastRowF2 = objWSF2.Range("A1048576").end(-4162).row
objWSF2.Range("A2:AD" & LastRowF2).Copy
objWSF1.Range("A" & LastRowF1 + 1).PasteSpecial 
objWBF2.Close
objFSO.DeleteFile File2

LastRowF1 = objWSF1.Range("A1048576").end(-4162).row
Set objWBF3 = objExcel.Workbooks.Open(File3)
Set objWSF3 = objWBF3.Worksheets(1)
LastRowF3 = objWSF3.Range("A1048576").end(-4162).row
objWSF3.Range("A2:AD" & LastRowF3).Copy
objWSF1.Range("A" & LastRowF1 + 1).PasteSpecial
objWBF3.Close
objFSO.DeleteFile File3

LastRowF1 = objWSF1.Range("A1048576").end(-4162).row
Set objWBF4 = objExcel.Workbooks.Open(File4)
Set objWSF4 = objWBF4.Worksheets(1)
LastRowF4 = objWSF4.Range("A1048576").end(-4162).row
objWSF4.Range("A2:AD" & LastRowF4).Copy
objWSF1.Range("A" & LastRowF1 + 1).PasteSpecial
objWBF4.Close
objFSO.DeleteFile File4
objWBF1.SaveAs Replace(WScript.Arguments(2),",_"," ") & " - " & WScript.Arguments(11)
objWBF1.Close
objFSO.DeleteFile File1

'Merging 3 Properties
If(fso.FolderExists(FolderPath))Then
	Set NFolder = fso.GetFolder(FolderPath)
		Set colFiles = NFolder.Files
		For Each FileItem in colFiles
			If(Instr(FileItem.Name,"1010")>0)Then
				File5  = FileItem.Name
				File5 = FolderPath & "\" & File5
				Set objWBF5 = objExcel.Workbooks.Open(File5)
                                Set objWSF5 = objWBF5.Worksheets(1)
			End If
		    If(Instr(FileItem.Name,"1012")>0)Then
				File6  = FileItem.Name
				File6 = FolderPath & "\" & File6
				LastRowF5 = objWSF5.Range("A1048576").end(-4162).row
                Set objWBF6 = objExcel.Workbooks.Open(File6)
                Set objWSF6 = objWBF6.Worksheets(1)
                LastRowF6 = objWSF6.Range("A1048576").end(-4162).row
                objWSF6.Range("A2:AD" & LastRowF6).Copy
                objWSF5.Range("A" & LastRowF5 + 1).PasteSpecial
                objWBF6.Close
                objFSO.DeleteFile File6
                 Else
                 objWSF5.Close
		End If
		Next
		Set colFiles = Nothing
	
Else
	MsgBox("File Not Found" & FolderPath)
End If

If(fso.FolderExists(FolderPath))Then
	Set NFolder = fso.GetFolder(FolderPath)
		Set colFiles = NFolder.Files
		For Each FileItem in colFiles
		    If(Instr(FileItem.Name,"1013")>0)Then
				File7  = FileItem.Name
				File7 = FolderPath & "\" & File7
				LastRowF5 = objWSF5.Range("A1048576").end(-4162).row
                Set objWBF7 = objExcel.Workbooks.Open(File7)
                Set objWSF7 = objWBF7.Worksheets(1)
                LastRowF7 = objWSF7.Range("A1048576").end(-4162).row
                objWSF7.Range("A2:AD" & LastRowF7).Copy
                objWSF5.Range("A" & LastRowF5 + 1).PasteSpecial
                objWBF7.Close
                objFSO.DeleteFile File7
		objWBF5.SaveAs Replace(WScript.Arguments(3),",_"," ") & " - " & WScript.Arguments(11)
                objWBF5.Close
                objFSO.DeleteFile File5
                'objExcel.Quit
                Else
                objWBF6.Close
		End If	
		Next
		Set colFiles = Nothing
	
Else
	MsgBox("File Not Found" & FolderPath)
End If

'Merging 2 Files
If(fso.FolderExists(FolderPath))Then
	Set NFolder = fso.GetFolder(FolderPath)
		Set colFiles = NFolder.Files
		For Each FileItem in colFiles
			If(Instr(FileItem.Name,"1040")>0)Then
				File8  = FileItem.Name
				File8 = FolderPath & "\" & File8
				Set objWBF8 = objExcel.Workbooks.Open(File8)
                Set objWSF8 = objWBF8.Worksheets(1)
			End If	
			If(Instr(FileItem.Name,"1041")>0)Then
				File9  = FileItem.Name
				File9 = FolderPath & "\" & File9
				LastRowF8 = objWSF8.Range("A1048576").end(-4162).row
                Set objWBF9 = objExcel.Workbooks.Open(File9)
                Set objWSF9 = objWBF9.Worksheets(1)
                LastRowF9 = objWSF9.Range("A1048576").end(-4162).row
                objWSF9.Range("A2:AD" & LastRowF9).Copy
                objWSF8.Range("A" & LastRowF8 + 1).PasteSpecial
                objWBF9.Close
                objFSO.DeleteFile File9
				objWBF8.SaveAs Replace(WScript.Arguments(4),",_"," ") & " - " & WScript.Arguments(11)
                objWBF8.Close
                objFSO.DeleteFile File8
			End If
		Next
		Set colFiles = Nothing
	
Else
	MsgBox("File Not Found" & FolderPath)
End If

'Merging 2 Files
If(fso.FolderExists(FolderPath))Then
	Set NFolder = fso.GetFolder(FolderPath)
		Set colFiles = NFolder.Files
		For Each FileItem in colFiles
			If(Instr(FileItem.Name,"1090")>0)Then
				File10  = FileItem.Name
				File10 = FolderPath & "\" & File10
				Set objWBF10 = objExcel.Workbooks.Open(File10)
                Set objWSF10 = objWBF10.Worksheets(1)
			End If	
			If(Instr(FileItem.Name,"1095")>0)Then
				File11  = FileItem.Name
				File11 = FolderPath & "\" & File11
				LastRowF10 = objWSF10.Range("A1048576").end(-4162).row
                Set objWBF11 = objExcel.Workbooks.Open(File11)
                Set objWSF11 = objWBF11.Worksheets(1)
                LastRowF11 = objWSF11.Range("A1048576").end(-4162).row
                objWSF11.Range("A2:AD" & LastRowF11).Copy
                objWSF10.Range("A" & LastRowF10 + 1).PasteSpecial
                objWBF11.Close
                objFSO.DeleteFile File11
				objWBF10.SaveAs Replace(WScript.Arguments(5),",_"," ") & " - " & WScript.Arguments(11)
                objWBF10.Close
                objFSO.DeleteFile File10
                objExcel.Quit
			End If
		Next
		Set colFiles = Nothing
	
Else
	MsgBox("File Not Found" & FolderPath)
End If

objExcel.Quit