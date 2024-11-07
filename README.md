# ExcelMacro
An Excel macro that merges data from multiple Excel files into a single workbook.

Excel File Merge Macro
This Excel macro consolidates data from multiple Excel files (.xlsx format) located in a specified folder into a single workbook. It extracts data from the first sheet of each file and compiles them into a new workbook.

Features
Automatic File Detection: Processes all .xlsx files in the specified folder.
Header Management:
First File: Copies all data including the header row.
Subsequent Files: Copies data from the second row onward to avoid duplicating headers.
New Workbook Output: The consolidated data is output to a new Excel workbook, which is displayed upon completion.
How to Use
Set the Folder Path:

In the Excel workbook where you will run the macro, enter the folder path containing the Excel files into cell A2 on Sheet1.
Set Up the Macro:

Open the VBA editor and paste the provided code into a standard module.
Run the Macro:

Execute the macro. It will automatically process all files in the specified folder.
Review the Results:

Upon completion, a new Excel workbook will open displaying the consolidated data. Save it as needed.
Code
Sub MergeExcelFiles()
    ' Declare variables
    Dim folderPath As String
    Dim outputWB As Workbook
    Dim currentFileName As String
    Dim fileList() As String
    Dim fileCount As Long
    Dim i As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outputWS As Worksheet
    Dim lastRow As Long
    Dim copyRange As Range
    Dim pasteRow As Long
    Dim isFirstFile As Boolean

    ' Get the folder path from Sheet1 cell A2
    folderPath = ThisWorkbook.Sheets("Sheet1").Range("A2").Value
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' Create a new output workbook and make it visible
    Set outputWB = Workbooks.Add
    outputWB.Application.Visible = True
    Set outputWS = outputWB.Sheets(1) ' First sheet of the output workbook

    ' Get all *.xlsx files in the folder
    currentFileName = Dir(folderPath & "*.xlsx")
    fileCount = 0
    Do While currentFileName <> ""
        fileCount = fileCount + 1
        ReDim Preserve fileList(1 To fileCount)
        fileList(fileCount) = currentFileName
        currentFileName = Dir()
    Loop

    ' If no files are found
    If fileCount = 0 Then
        MsgBox "No Excel files were found in the specified folder."
        Exit Sub
    End If

    ' Loop through each file
    isFirstFile = True
    For i = 1 To fileCount
        ' Open the file as read-only
        Set wb = Workbooks.Open(folderPath & fileList(i), ReadOnly:=True)
        Set ws = wb.Sheets(1) ' Get the first sheet

        ' Determine the range to copy
        If isFirstFile Then
            ' For the first file, copy all data including headers
            Set copyRange = ws.UsedRange
            pasteRow = 1
            isFirstFile = False
        Else
            ' For subsequent files, copy from the second row to the last
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow >= 2 Then
                Set copyRange = ws.Range(ws.Rows(2), ws.Rows(lastRow))
                ' Determine the next paste position in the output sheet
                pasteRow = outputWS.Cells(outputWS.Rows.Count, 1).End(xlUp).Row + 1
            Else
                ' If there is no data to copy
                Set copyRange = Nothing
            End If
        End If

        ' Paste the data into the output workbook
        If Not copyRange Is Nothing Then
            copyRange.Copy outputWS.Cells(pasteRow, 1)
        End If

        ' Close the opened file without saving changes
        wb.Close SaveChanges:=False
    Next i

    ' Display a completion message
    MsgBox "All files have been processed successfully."

End Sub
Notes
File Format: Only .xlsx files are processed. To include other file formats, modify the file extension in the code ("*.xlsx").
Sheet Selection: Only the first sheet of each file is processed. To process other sheets, adjust the Set ws = wb.Sheets(1) line.
Data Consistency: Ensure that all files have the same structure (e.g., the same columns) to prevent data misalignment.
Error Handling: The macro includes basic error checking for missing files but may need additional error handling for robustness.
License
This project is licensed under the MIT License.

Contributing
Feel free to submit issues or pull requests for bug fixes or enhancements.
