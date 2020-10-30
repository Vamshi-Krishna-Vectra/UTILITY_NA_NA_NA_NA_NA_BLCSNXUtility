Imports Microsoft.Office.Interop.Excel
Module ExcelWrapper
    Public _objApp As Object
    'Private _objWorkbook As Workbook
    'Private _objSheet As Worksheet
    Private _asColumnData() As String
    Private _asRowData() As String
    Sub KillExcelObject()
        If Not _objApp Is Nothing Then
            _objApp.Quit()
            ReleaseObject(_objApp)
            _objApp = Nothing
        End If
    End Sub
    Private Sub ReleaseObject(ByRef obj As Object)
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
        Try
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    'CODE CHANGED - 4/6/16 - Amitabh - To prevent saving of workbooks (ex: template files)
    Sub SCloseWorkBook(ByRef objWorkbook As Workbook, Optional bSaveChanges As Boolean = True)
        If bSaveChanges Then
            objWorkbook.Save()
            objWorkbook.Close(SaveChanges:=True)
        Else
            objWorkbook.Close(SaveChanges:=False)
        End If
        'ReleaseObject(objWorkbook)
        objWorkbook = Nothing
        KillExcelObject()
    End Sub
    Function FnOpenWorkbook(ByVal sfileName As String) As Workbook
        Dim objWorkbook As Workbook
        If (_objApp Is Nothing) Then
            _objApp = CreateObject("Excel.Application")
            _objApp.Visible = False
            _objApp.DisplayAlerts = False
            'Created programatically
            _objApp.UserControl = False
            'Change the Excel Language
            System.Threading.Thread.CurrentThread.CurrentCulture = _
                            New System.Globalization.CultureInfo("en-US")
        End If
        'If _objWorkbook Is Nothing Then
        Err.Clear()
        objWorkbook = _objApp.Workbooks.Open(sfileName, Password:=EXCEL_OPEN_PASSWORD)
        If (Err.Number <> 0) Then
            'If the workbook does not have a password
            objWorkbook = _objApp.Workbooks.Open(sfileName)
        End If
        'Change the excel calculation mode to manual
        _objApp.Calculation = XlCalculation.xlCalculationManual
        Err.Clear()
        'End If
        'Code added Jun-24-2019
        'OPening the excel in USD was throwing RPC Server disconnected error. 
        '(RPC Disconnection error comes, when you open the excel file and save the file immediately. 
        'Excel takes a bit time to,open the application. 
        'If you try to save the excel file before even it is completely opened, then we will get this error)
        'To handle this error, we are checking if the application is ready before saving the excel file.
        Do
            System.Threading.Thread.Sleep(5000)
        Loop Until (_objApp.Ready)
        FnOpenWorkbook = objWorkbook
    End Function

    Function FnSetWorksheet(ByVal objworkbook As Workbook, ByVal sName As String) As Worksheet
        FnSetWorksheet = objworkbook.Worksheets.Item(sName)
    End Function
    Function FnReadColumnForRow(ByVal objworkbook As Workbook, ByVal sSheetName As String, ByVal iRowNum As Integer) As Array
        Dim iLoopIndex As Integer
        Dim iColumnCount As Integer
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objworkbook, sSheetName)
        iColumnCount = FnGetNumberofColumns(objworkbook, sSheetName, iRowNum)
        ReDim _asColumnData(iColumnCount)
        For iLoopIndex = 1 To iColumnCount
            _asColumnData(iLoopIndex - 1) = objWorksheet.Cells(iRowNum, iLoopIndex).value.ToString().Trim()
        Next
        FnReadColumnForRow = _asColumnData
    End Function
    Function FnReadRowDataForColumn(ByVal objworkbook As Workbook, ByVal sSheetName As String, ByVal iColNum As Integer, ByVal iStartRow As Integer) As Array
        Dim iLoopIndex As Integer
        Dim iRowCount As Integer
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objworkbook, sSheetName)
        iRowCount = FnGetNumberofRows(objworkbook, sSheetName, iColNum, iStartRow)
        ReDim _asRowData(iRowCount)
        For iLoopIndex = 1 To iRowCount
            _asRowData(iLoopIndex - 1) = objWorksheet.Cells(iLoopIndex, iColNum).value.ToString().Trim()
        Next
        FnReadRowDataForColumn = _asRowData
    End Function
    Function FnFindRowNumberByColumnAndValue(ByVal objworkbook As Workbook, ByVal sSheetName As String, ByVal iColNum As Integer, ByVal iStartRow As Integer, _
                                 ByVal sRowText As String) As Integer
        Dim iLoopIndex As Integer
        Dim iRowCount As Integer
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objworkbook, sSheetName)
        iRowCount = FnGetNumberofRows(objworkbook, sSheetName, iColNum, iStartRow)
        For iLoopIndex = iStartRow To iRowCount
            If objWorksheet.Cells(iLoopIndex, iColNum).value.ToString().Trim() = sRowText.ToString() Then
                FnFindRowNumberByColumnAndValue = iLoopIndex
                Exit Function
            End If
        Next
        FnFindRowNumberByColumnAndValue = 0
    End Function
    Sub SWriteValueToCell(ByVal objwrkBook As Workbook, ByVal sSheetName As String, ByVal iRowNum As Integer, ByVal iColNum As Integer, ByVal sValue As String)
        Dim objWorksheet As Worksheet
        Dim objRange As Range = Nothing
        objWorksheet = FnSetWorksheet(objwrkBook, sSheetName)
        Try
            objWorksheet.Cells(iRowNum, iColNum).Value = sValue
        Catch ex As Exception
            objRange = objWorksheet.Range(objWorksheet.Cells(iRowNum, iColNum))
            objRange.Value2 = sValue
        End Try
    End Sub
    Function FnReadSingleRowForColumn(ByVal objworkbook As Workbook, ByVal sSheetName As String, ByVal iColNum As Integer, ByVal iRowNum As Integer) As String
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objworkbook, sSheetName)
        If Not objWorksheet.Cells(iRowNum, iColNum).value Is Nothing Then
            FnReadSingleRowForColumn = objWorksheet.Cells(iRowNum, iColNum).value.ToString().Trim()
        Else
            FnReadSingleRowForColumn = Nothing
        End If
    End Function
    Function FnReadSingleColumnForRow(ByVal objworkbook As Workbook, ByVal sSheetName As String, ByVal iColNum As Integer, ByVal iRowNum As Integer) As String
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objworkbook, sSheetName)
        If Not objWorksheet.Cells(iRowNum, iColNum).value Is Nothing Then
            FnReadSingleColumnForRow = objWorksheet.Cells(iRowNum, iColNum).value.ToString().Trim()
        Else
            FnReadSingleColumnForRow = Nothing
        End If
    End Function
    Function FnGetColumnNumberByName(ByVal objworkbook As Workbook, ByVal sSheetName As String, ByVal sColumnName As String, ByVal iRowNumber As Integer) As Integer
        Dim iColumnCount As Integer
        Dim iLoopIndex As Integer
        Dim bFound As Boolean
        bFound = False
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objworkbook, sSheetName)
        iColumnCount = FnGetNumberofColumns(objworkbook, sSheetName, iRowNumber)
        For iLoopIndex = 1 To iColumnCount
            If (objWorksheet.Cells(iRowNumber, iLoopIndex).value.ToString().Trim = sColumnName.Trim()) Then
                FnGetColumnNumberByName = iLoopIndex
                bFound = True
                Exit For
            End If
        Next
        If Not bFound Then
            FnGetColumnNumberByName = 0
        End If
    End Function
    Function FnGetNumberofColumns(ByVal objworkbook As Workbook, ByVal sSheetName As String, ByVal iRowIndex As Integer) As Integer
        Dim iColumnIndex As Integer
        iColumnIndex = 1
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objworkbook, sSheetName)
        While (Not objWorksheet.Cells(iRowIndex, iColumnIndex) Is Nothing And Not objWorksheet.Cells(iRowIndex, iColumnIndex + 1).value Is Nothing)
            iColumnIndex = iColumnIndex + 1
        End While
        FnGetNumberofColumns = iColumnIndex
    End Function
    Function FnGetNumberofRows(ByVal objworkbook As Workbook, ByVal sSheetName As String, ByVal iColumnIndex As Integer, ByVal iStartRow As Integer) As Integer
        Dim iRowIndex As Integer
        iRowIndex = iStartRow
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objworkbook, sSheetName)
        While (Not objWorksheet.Cells(iRowIndex, iColumnIndex).value Is Nothing And Not objWorksheet.Cells(iRowIndex + 1, iColumnIndex).value Is Nothing)
            iRowIndex = iRowIndex + 1
        End While
        FnGetNumberofRows = iRowIndex
    End Function
    Function FnGetColumnNumber(ByVal objworkbook As Workbook, ByVal sSheetName As String, ByVal sColumnValue As String, ByVal iRowIndex As Integer) As String
        Dim iIndex As Integer
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objworkbook, sSheetName)
        For iIndex = 1 To FnGetNumberofColumns(objworkbook, sSheetName, iRowIndex)
            If objWorksheet.Cells(iRowIndex, iIndex).value.ToString().Trim() = sColumnValue.Trim() Then
                FnGetColumnNumber = iIndex
                Exit For
            End If
        Next
    End Function
    Function FnCreateRange(ByVal objWrkBk As Workbook, ByVal sSheetName As String, ByVal iStartRow As Integer, ByVal iStartCol As Integer, ByVal iEndRow As Integer, ByVal iEndCol As Integer) As Range
        Dim objWorksheet As Worksheet
        Dim objRange As Range = Nothing
        objWorksheet = FnSetWorksheet(objWrkBk, sSheetName)
        FnCreateRange = objWorksheet.Range(objWorksheet.Cells(iStartRow, iStartCol), objWorksheet.Cells(iEndRow, iEndCol))
    End Function
    Public Function FnReadFullExcelSheet(ByVal sSheetName As String, ByVal sfileName As String, ByVal iStartRow As Integer, ByVal iStartColumn As Integer) As String()
        Dim objWorkbook As Workbook
        Dim objWorkSheet As Worksheet
        Dim iLastColumnCount As Integer
        Dim iLastRowCount As Integer
        Dim asExcelData() As String
        Dim iCounter As Integer = 0
        Dim iLoopRowIndex As Integer
        Dim iLoopColumnIndex As Integer
        Dim sDataToAdd As String = ""
        Dim bAlreadyOpen As Boolean = False

        If (_objApp Is Nothing) Then
            _objApp = New Application
            _objApp.Visible = False
        End If

        For Each wrkbk As Workbook In _objApp.Workbooks
            If wrkbk.FullName = sfileName Then
                bAlreadyOpen = True
                objWorkbook = wrkbk
                Exit For
            End If
        Next

        If Not bAlreadyOpen Then
            Err.Clear()
            objWorkbook = _objApp.Workbooks.Open(sfileName, ReadOnly:=True, Password:=EXCEL_OPEN_PASSWORD)
            If (Err.Number <> 0) Then
                'If the workbook does not have a password
                objWorkbook = _objApp.Workbooks.Open(sfileName, ReadOnly:=True)
            End If
            Err.Clear()
        End If

        Try
            objWorkSheet = objWorkbook.Worksheets.Item(sSheetName)
            'Get the Last Row and the Lat column in the sheet (blank cell will be considered as column/ row finish
            iLastColumnCount = FnGetNumberofColumns(objWorkbook, sSheetName, iStartRow)
            iLastRowCount = FnGetNumberofRows(objWorkbook, sSheetName, iStartColumn, iStartRow)

            'Read the Cells and add the values as comma seperated
            For iLoopRowIndex = iStartRow To iLastRowCount
                For iLoopColumnIndex = iStartColumn To iLastColumnCount
                    If sDataToAdd = "" Then
                        sDataToAdd = objWorkSheet.Cells(iLoopRowIndex, iLoopColumnIndex).Value.ToString
                    Else
                        sDataToAdd = sDataToAdd + "," + objWorkSheet.Cells(iLoopRowIndex, iLoopColumnIndex).Value.ToString
                    End If
                Next
                ReDim Preserve asExcelData(iCounter)
                asExcelData(iCounter) = sDataToAdd
                iCounter = iCounter + 1
                sDataToAdd = ""
            Next
        Catch ex As Exception
        End Try

        If Not bAlreadyOpen Then
            objWorkbook.Close()
        End If
        objWorkbook = Nothing
        FnReadFullExcelSheet = asExcelData
    End Function

    Public Sub SChangeToNumberFormat(ByVal objWrkBk As Workbook, ByVal sSheetName As String, ByVal iRow As Integer, ByVal iCol As Integer)
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objWrkBk, sSheetName)
        objWorksheet.Cells(iRow, iCol).NumberFormat = "0.0000"
    End Sub
    Public Function FnCheckIfSheetPresent(ByVal objWrkBk As Workbook, ByVal sSheetName As String) As Boolean
        For Each objSht As Worksheet In objWrkBk.Worksheets
            If objSht.Name.ToUpper = sSheetName.ToUpper Then
                FnCheckIfSheetPresent = True
                Exit Function
            End If
        Next
        FnCheckIfSheetPresent = False
    End Function
    Public Sub sDeleteSheet(ByVal objWrkBk As Workbook, ByVal sSheetName As String)
        For Each objSht As Worksheet In objWrkBk.Worksheets
            If objSht.Name.ToUpper = sSheetName.ToUpper Then
                objSht.Delete()
            End If
        Next
    End Sub
    'To Copy a sheet and place it after the copied sheet
    Public Sub sCopySheet(ByVal objWrkBk As Workbook, ByVal sSheetNameToCopy As String, ByVal sNewSheetName As String)
        Dim objWorksheet As Worksheet
        objWorksheet = FnSetWorksheet(objWrkBk, sSheetNameToCopy)
        objWorksheet.Copy(After:=objWorksheet)
        'Change the name of the copied sheet
        objWrkBk.ActiveSheet.Name = sNewSheetName
    End Sub

    Public Function FnAddWorkSheet(ByVal objWrkBk As Workbook, sNewSheetName As String, Optional ByRef sRefSheetName As String = "") As Worksheet
        Dim objNewWorkSheet As Worksheet = Nothing
        Dim objRefSheet As Worksheet = Nothing

        If sRefSheetName <> "" Then
            objRefSheet = FnSetWorksheet(objWrkBk, sRefSheetName)
        End If

        If Not objRefSheet Is Nothing Then
            objNewWorkSheet = objWrkBk.Worksheets.Add(, After:=objRefSheet)
        Else
            objNewWorkSheet = objWrkBk.Worksheets.Add()
        End If
        objNewWorkSheet.Name = sNewSheetName

        FnAddWorkSheet = objNewWorkSheet

    End Function

    'Copy one specific row from a source worksheet to a destination worksheet 
    Sub sCopyRows(objWrkBk As Workbook, objDestinationWrkSheet As Worksheet, iRowToCopyStart As Integer, iRowToCopyEnd As Integer, _
                 sSourceSheetName As String, Optional iLastColumnToCopy As Integer = 0, Optional iStartColumnCopy As Integer = 1)

        Dim objSourceSheet As Worksheet = Nothing
        Dim iNumOfColumns As Integer = -1
        Dim objRange As Range = Nothing
        Dim objDestinationRange As Range = Nothing

        objSourceSheet = FnSetWorksheet(objWrkBk, sSourceSheetName)
        If Not objSourceSheet Is Nothing Then
            If iLastColumnToCopy <> 0 Then
                iNumOfColumns = iLastColumnToCopy
            Else
                iNumOfColumns = FnGetNumberofColumns(objWrkBk, sSourceSheetName, iRowToCopyStart)
            End If
            objRange = FnCreateRange(objWrkBk, sSourceSheetName, iRowToCopyStart, iStartColumnCopy, iRowToCopyEnd, iNumOfColumns)
            objRange.Copy()
            objDestinationRange = FnCreateRange(objWrkBk, objDestinationWrkSheet.Name, iRowToCopyStart, iStartColumnCopy, iRowToCopyEnd, iNumOfColumns)
            objDestinationRange.PasteSpecial(XlPasteType.xlPasteAll)
        End If
    End Sub
End Module
