Option Strict Off
Imports System.Math
Imports NXOpen
Imports NXOpen.Features
Imports NXOpen.UF
Imports NXOpen.Drawings
Imports NXOpen.Assemblies
Imports NXOpen.Utilities
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Net.Mail
Imports System.Xml
Imports System.Xml.Linq
Module Main
    Private _objWorkBk As Workbook = Nothing
    Private iColFeatType As Integer = 0
    Private iColFeatName As Integer = 0
    Private iColFaceName As Integer = 0
    Private iColFaceType As Integer = 0
    Private iColEdgeName As Integer = 0
    Private iColEdgeDia As Integer = 0
    Private iColEdgeCenterX As Integer = 0
    Private iColEdgeCenterY As Integer = 0
    Private iColEdgeCenterZ As Integer = 0
    Private iColVertex1X As Integer = 0
    Private iColVertex1Y As Integer = 0
    Private iColVertex1Z As Integer = 0
    Private iColVertex2X As Integer = 0
    Private iColVertex2Y As Integer = 0
    Private iColVertex2Z As Integer = 0
    Private iColEdgeLength As Integer = 0
    Private iColStartAngle As Integer = 0
    Private iColEndAngle As Integer = 0
    Private iColDCXx As Integer = 0
    Private iColDCXy As Integer = 0
    Private iColDCXz As Integer = 0

    Private iColDCYx As Integer = 0
    Private iColDCYy As Integer = 0
    Private iColDCYz As Integer = 0

    Private iColHoleX As Integer = 0
    Private iColHoleY As Integer = 0
    Private iColHoleZ As Integer = 0
    Private iColHoleDia As Integer = 0
    Private iColHoleParent As Integer = 0
    Private iColThreadParent1 As Integer = 0
    Private iColThreadParent2 As Integer = 0
    Private iColCallout As Integer = 0
    Private iColEdgeCurvature As Integer = 0

    'Private iColDepth As Integer = 0
    'Private iColBlendRad As Integer = 0
    Private iColEdgeType As Integer = 0
    Private iColFaceNameFaceVec As Integer = 0
    Private iColVectorX As Integer = 0
    Private iColVectorY As Integer = 0
    Private iColVectorZ As Integer = 0
    Private iColFaceArea As Integer = 0
    Private iColHoleSize As Integer = 0
    Private iColPreFab As Integer = 0
    Private iColFlameCutFace As Integer = 0

    Private iColFaceCenterX As Integer = 0
    Private iColFaceCenterY As Integer = 0
    Private iColFaceCenterZ As Integer = 0
    Private iColFaceRadius As Integer = 0
    Private iColFaceDirection As Integer = 0

    Private iColFeatNameAttr As Integer = 0
    Private iColHoleName As Integer = 0
    Private iColHoleVecX As Integer = 0
    Private iColHoleVecY As Integer = 0
    Private iColHoleVecZ As Integer = 0

    Private iColBodyName As Integer = 0
    Private iColBodyShape As Integer = 0
    Private iColBodyDetailNos As Integer = 0
    Private iColBodyToolClass As Integer = 0
    Private iColStockSize As Integer = 0
    Private iColComponentName As Integer = 0
    Private iColBodyLayer As Integer = 0
    Private iColCompDBPartName As Integer = 0
    Private iColMinPointX As Integer = 0
    Private iColMinPointY As Integer = 0
    Private iColMinPointZ As Integer = 0
    Private iColVectorXX As Integer = 0
    Private iColVectorXY As Integer = 0
    Private iColVectorXZ As Integer = 0
    Private iColVectorYX As Integer = 0
    Private iColVectorYY As Integer = 0
    Private iColVectorYZ As Integer = 0
    Private iColVectorZX As Integer = 0
    Private iColVectorZY As Integer = 0
    Private iColVectorZZ As Integer = 0
    Private iColMagX As Integer = 0
    Private iColMagY As Integer = 0
    Private iColMagZ As Integer = 0
    Private iColNCPartContactFace As Integer = 0
    Private iColPMat As Integer = 0
    Private iColFloorMountFace As Integer = 0

    'Private dictStockQuantity As Dictionary(Of String, Integer) = Nothing
    Private _colListOfPartsGenerated As String() = Nothing
    Private _iMaxSubDetNum As Integer = 0
    'Store the Tool folder path selected by the user
    Private _sToolFolderPath As String = ""
    'Store the Sweep data output folder path
    Private _sSweepDataOutputFolderPath As String = ""
    Private _sToolFolderName As String = ""

    'To store all the various layers in which solid bodies are present
    'Private _asBodyLayersList() As Integer = Nothing
    Private _aoTextExtrudeFeatureFaces() As Face = Nothing
    'Store all the solid bodies in the part
    Private _aoSolidBody() As Body = Nothing

    'Exception components where the mating body tolerance needs to be relaxed
    Private _asExceptionCompsForMatingBodyToleranceRelaxation() As String = {"SQUEEZE ACTION CLAMP"}
    'DB_PART_NAME for probable NC Blocks
    Private _asNCBLOCKDbPartNames() As String = {"NC LOCATOR", "LOCATOR", "NC BACKUP", "BACKUP", "N.C. PRESSURE FOOT", "N.C. LOCATOR",
                                                "NC PRESSURE FOOT", "PRESSURE FOOT", "NC FINGER", "FINGER", "REST", "NC REST"}

    Public _dictDowelHoles As Dictionary(Of String, Dictionary(Of String, String())) = Nothing
    'Get the long path name
    Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" _
        (ByVal lpszShortPath As String, ByVal lpszLongPath As String,
        ByVal cchBuffer As Integer) _
        As Integer
    Private iTrailOrentationCount As Integer = 1
    'To get the long path from a short path
    Private Function LongPath(ByVal ShortPath As String) As String

        Dim sLongPath As String = Space(255)
        GetLongPathName(ShortPath & Chr(0), sLongPath, sLongPath.Length)
        Return sLongPath.Substring(0, InStr(sLongPath, Chr(0)) - 1)

    End Function
    'Structure to store all orientation info and then figure out the optimal rotation matrix based on a defined logic
    Public Structure structOrientationInfo
        Public xx As Double
        Public xy As Double
        Public xz As Double
        Public yx As Double
        Public yy As Double
        Public yz As Double
        Public zx As Double
        Public zy As Double
        Public zz As Double
        Public iCountXAlignedPeripheralFaces As Integer
        Public iCountYAlignedPeripheralFaces As Integer
        Public iCountZAlignedPeripheralFaces As Integer
        Public dBoundingBoxVolume As Double
        Public iRank As Integer
        Public iCountMisAlignedMachinedFaces As Integer
        Public iCountAlignedFace As Integer
        Public iCountAlignedFaceWithHoles As Integer
    End Structure

    Dim _aoStructBodyOrientationInfo() As structOrientationInfo = Nothing
    Dim _iOrientationIndex As Integer = 0
    Dim _aoStructPartOrientationInfo() As structOrientationInfo = Nothing
    Dim _iPartOrientationIndex As Integer = 0

    'Structure to store Minimum and Maximum X Y Z values for all solid body in a weldment
    Public Structure structMinMaxXYZValues
        Public sBodyName As String
        Public dMinX As Double
        Public dMaxX As Double
        Public dMinY As Double
        Public dMaxY As Double
        Public dMinZ As Double
        Public dMaxZ As Double
    End Structure
    Dim _aoStructMinMaxXYZ() As structMinMaxXYZValues
    Dim _iSolidBodyIndex As Integer = 0
    Public _bCreatePartLCSOrientation As Boolean = True
    Public _sToolNamefromConfigFile As String = ""
    Public _sFeatureGroupName As String = ""
    'Code added on May-14-2018
    'This attribute will act as a flag to determine if the edge and face names to be assigned at occurrence or prototype level
    Public _bIsComponent As Boolean = False
    'Public _dictBodyOptimalRotMat As Dictionary(Of String, Double()) = Nothing
    Public _FINISHTOLERANCE_ATTR_NAME As String = ""
    Public _SHAPE_ATTR_NAME As String = ""
    Public _PART_NAME As String = ""
    Public _SUB_DETAIL_NUMBER As String = ""
    Public _P_MASS As String = ""
    Public _PURCH_OPTION As String = ""
    Public _QTY As String = ""
    Public _P_MAT As String = ""
    Public _CLIENT_STOCK_SIZE_ATTR As String = ""
    Public _STOCK_SIZE_METRIC As String = ""
    Public _STOCK_SIZE As String = ""
    Public _RELIEF_CUT_FACE_TOLERANCE_VALUE As String = ""
    Public _sOemName As String = ""
    Public _sDivision As String = ""
    Public _sSupplierName As String = ""
    Public _FINISH_TOL_VALUE2 As String = ""
    Public _CLIENT_PART_NAME As String = ""
    Public _BOM_ATTR As String = ""
    Public _TOOL_CLASS As String = ""
    Public _TOOL_ID As String = ""
    Public _ALTPURCH As String = ""
    Public _bProcessPSDAgainForThisPart As Boolean = False
    Public _asErrorNamesToProcessFileAgain() As String = Nothing
    'Public _sStartTime As Date
    'Public _sEndTime As Date
    'Public _lRunTime As Long = 0
    Public Declare Auto Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As IntPtr,
        ByRef lpdwProcessId As Integer) As Integer

    Sub Main()

        Dim objTemplateWrkbk As Workbook = Nothing
        Dim sDesSource As String = ""
        Dim bFilePresentInConfigFolder As Boolean = False
        'Get the path to the tool config file
        Dim sToolConfigFilePath As String = ""
        'Get the path to the output folder path config file
        Dim sOutputConfigFilePath As String = ""
        'Get all the tool paths
        Dim asToolPaths() As String = Nothing
        'To store the path provided by the user
        Dim asOutputFolderPath() As String = Nothing
        'Get the template file path
        Dim sTemplateFilePath As String = ""
        'To store the parent path where this part belongs
        Dim sPartToolParentPath As String = ""
        Dim asFilteredToolPath() As String = Nothing
        Dim sFolderPath As String = ""
        Dim sAttributeXMLConfigFilePath As String = ""
        Dim sErrorNamesFilePath As String = ""
        Dim objPart As Part = FnGetNxSession.Parts.Work

        'Read the config file for the tool folder paths
        'There can be more than one tool folder paths
        sToolConfigFilePath = Path.Combine(FnGetExecutionFolderPath(), CONFIG_FOLDER_NAME, OEM_SUPPLIER_TEXT_FILE_NAME)
        'Read the full file
        asToolPaths = FnReadFullFile(sToolConfigFilePath)
        If asToolPaths Is Nothing Then
            sWriteToLogFile("Tool folder path not specified")
            Exit Sub
        Else
            _sOemName = asToolPaths(0)
            _sDivision = asToolPaths(1)
            _sSupplierName = asToolPaths(2)
        End If

        'Code added Jun-28-2019
        'Check if the Attribtues XML config file is present in the Config folder
        sAttributeXMLConfigFilePath = Path.Combine(FnGetExecutionFolderPath(), CONFIG_FOLDER_NAME, ATTRIBUTE_XML_CONFIG_FILE_NAME)
        If Not FnCheckFileExists(sAttributeXMLConfigFilePath) Then
            sWriteToLogFile(ATTRIBUTE_XML_CONFIG_FILE_NAME & " file is not present inside the Config folder")

            Exit Sub
        Else
            sWriteToLogFile(ATTRIBUTE_XML_CONFIG_FILE_NAME & " file is present inside the Config folder")
        End If

        'Code added May-28-2018
        sWriteToLogFile("B_LCS Version is : " & MODULE_VERSION)
        sWriteToLogFile("OEM is : " & _sOemName)
        sWriteToLogFile("Division is : " & _sDivision)
        sWriteToLogFile("Supplier is : " & _sSupplierName)
        sWriteToLogFile("-----------------------------------------")

        'Code added Feb-27-2018
        If _sOemName = DAIMLER_OEM_NAME Then
            _sFeatureGroupName = FnGetFeatureGroupName(_sDivision)
        End If
        'Code modified on Jun-28-2019
        'All the attributes are read from XML Config file
        'Code added Jun-04-2018
        'Get the attribute names based on different OEM
        sGetAttributeNamesBasedOnOEMandSupplier(sAttributeXMLConfigFilePath, _sOemName, _sSupplierName, _sDivision)

        Try
            sWriteLCSInfo(objPart)
            objPart.ModelingViews.WorkView.Fit()

        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

    End Sub
    Sub sProcessSweepDataForpart(sPartFilePath As String, sToolConfigFilePath As String)

        Dim objPart As Part = Nothing
        Dim sSweepDataExcelFilePath As String = ""
        Dim sSweepDataFolder As String = ""
        Dim sErrorFolderPath As String = ""
        Dim aoAllCompInSession() As Component = Nothing

        ' _sStartTime = DateTime.Now
        'Open the NX Part in Session
        objPart = FnOpenNxPart(sPartFilePath)
        'sCalculateTimeForAPI("Open Part File ")

        'Code added On Jun-13-2017
        sWriteToLogFile("")
        sWriteToLogFile("File loaded from : " & sPartFilePath)
        If Not objPart Is Nothing Then
            FnLoadPartFully(objPart)
            'Code added Jun-29-2017
            'Check if the part is to be detailed or not
            If FnCheckIfPartIsToBeDetailed(objPart) Then
                If FnCheckIfPartIsValidMakeDetail(objPart) Then
                    'Write the Sweep Data file for the given part
                    Try
                        _bIsComponent = False
                        FnGetUFSession.Undo.DeleteAllMarks()
                        Write(objPart, sToolConfigFilePath)
                        FnSavePart(objPart)
                        sWriteToLogFile("Part file saved")
                        sWriteToLogFile("Sweep data created for " & objPart.Leaf.ToUpper)
                        sWriteToLogFile("****************************************************************************")

                    Catch ex As Exception
                        'Code Added - Amitabh - 9/21/16 - report error message
                        sWriteToLogFile(ex.Message)
                        'Code modified on May-05-2020
                        If FnCheckIfErrorNameIsInList(ex.Message, _asErrorNamesToProcessFileAgain) Then
                            _bProcessPSDAgainForThisPart = True
                        End If
                        'Code commented on May-05-2020
                        'Code added July-30-2019
                        'If (ex.Message.ToUpper.Contains("RPC_E_DISCONNECTED")) Or (ex.Message.ToUpper.Contains("RPC_E_CALL_REJECTED")) Or
                        '    (ex.Message.ToUpper.Contains("RPC SERVER IS UNAVAILABLE")) Then
                        '    _bProcessPSDAgainForThisPart = True
                        'End If
                        'Create the file with exceptions report
                        sCreateErrorLogFile(objPart, ex)
                        If Not _objWorkBk Is Nothing Then
                            'CODE ADDED - 5/6/16 - Amitabh - Store sweep data information
                            sSweepDataExcelFilePath = _objWorkBk.FullName
                            sSweepDataFolder = _objWorkBk.Path
                            SCloseWorkBook(_objWorkBk, bSaveChanges:=True)
                        End If
                        'Code added May-14-2018
                        sClearMemory()
                        'Code modified on May-05-2020
                        'If (ex.Message.ToUpper.Contains("RPC_E_DISCONNECTED")) Or (ex.Message.ToUpper.Contains("RPC_E_CALL_REJECTED")) Or
                        '    (ex.Message.ToUpper.Contains("RPC SERVER IS UNAVAILABLE")) Then
                        If FnCheckIfErrorNameIsInList(ex.Message, _asErrorNamesToProcessFileAgain) Then
                            _bProcessPSDAgainForThisPart = True
                            'Delete the file in the sweep data folder
                            If FnCheckFileExists(sSweepDataExcelFilePath) Then
                                SDeleteFile(sSweepDataExcelFilePath)
                            End If
                            sWriteToLogFile("Sweep data not created for " & objPart.Leaf.ToUpper)
                            sWriteToLogFile("This might be due to Excel error, Remote Procedure Call")
                            sWriteToLogFile("PSD will be initiated again to rerun for this part")
                            sWriteToLogFile("****************************************************************************")
                        Else
                            'CODE ADDED - 5/6/16 - Amitabh
                            'Move this sweep data file to the unprocessed folder
                            sErrorFolderPath = Path.Combine(_sSweepDataOutputFolderPath, UNPROCESSED_PARTS, _sToolFolderName)
                            If Not FnCheckFolderExists(sErrorFolderPath) Then
                                SCreateDirectory(sErrorFolderPath)
                            End If
                            SCopyFiles(sErrorFolderPath, sSweepDataFolder, sSweepDataExcelFilePath)
                            'Delete the file in the sweep data folder
                            If FnCheckFileExists(sSweepDataExcelFilePath) Then
                                SDeleteFile(sSweepDataExcelFilePath)
                            End If
                            sWriteToLogFile("Sweep data not created for " & objPart.Leaf.ToUpper)
                            sWriteToLogFile("File moved to " & UNPROCESSED_PARTS & " folder" & objPart.Leaf.ToUpper)
                            sWriteToLogFile("****************************************************************************")
                        End If

                    End Try
                Else
                    sWriteToLogFile("Part is not a Valid Make Detail")
                    sWriteToLogFile("Part is ALT PURCH Component")
                    sWriteToLogFile("This part has the value in " & _ALTPURCH & " Component")
                    sWriteToLogFile("Sweep data not created for this component")
                    sWriteToLogFile("****************************************************************************")
                End If
            Else
                'Code added July-12-2018
                'In some case custom components may have face names and edge names even if they have B_Detail =N attribute.
                'So do a complete clean up to erase face names and edge names and save the part.
                aoAllCompInSession = FnGetAllComponentsInSession()
                'Delete all the existing face names and edge names
                sDeleteOldSweepDataInformation(objPart, aoAllCompInSession)
                sWriteToLogFile("Old Face names and Edge Names are removed from the part")
                FnSavePart(objPart)
                sWriteToLogFile("Part file saved")
                sWriteToLogFile("****************************************************************************")
            End If
            SClosePart(objPart.Leaf.ToString)
        Else
            sWriteToLogFile("Part file was empty, Didnot open the part")
        End If

    End Sub

    Sub Write(ByVal objPart As Part, sConfigFolderPath As String)
        'Dim theSession As Session = Session.GetSession()
        'Dim workPart As Part = theSession.Parts.Work
        Dim iCount As Integer = 1
        Dim iFaceCount As Integer = 1
        Dim ufs As UFSession = UFSession.GetUFSession()
        Dim iLastFilledRow As Integer = 0
        Dim iLastFilledColumn As Integer = 0
        Dim sPartName As String = objPart.Leaf.ToString
        Dim iRowStart As Integer = START_ROW_WRITE
        Dim iRowFaceVecStart As Integer = 0
        Dim sFolderName As String = ""
        Dim sWeldmentSubFolderName As String = ""
        Dim sComponentSubFolderName As String = ""
        Dim sSweepDataFilePath As String = ""
        Dim aoAllCompInSession() As Component = Nothing

        Dim sBodyName As String = ""
        Dim objRefPoint1 As Point3d = Nothing
        Dim objRefPoint2 As Point3d = Nothing
        Dim sValue As String
        Dim iRowFinishStart As Integer = 0
        Dim iRowStartDataStart As Integer = 0
        Dim aoAllValidSolidBody() As Body = Nothing

        Dim objPrimaryView2BodyNameSheet As Worksheet = Nothing
        Dim iNosOfFilledRows As Integer = 0
        Dim objChildPart As Part = Nothing
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim objFeatureGroup As Features.FeatureGroup = Nothing
        Dim aoAllMembers() As Features.Feature = Nothing
        Dim bIsValidSolidBody As Boolean = False
        'Dim sCompInCarDivision As Boolean = False
        Dim objBodyToAnalyse As Body = Nothing

        'Set WCS to Absolute
        Dim origin1 As Point3d = New Point3d(0.0, 0.0, 0.0)
        Dim matrix1 As Matrix3x3
        matrix1.Xx = 1.0
        matrix1.Xy = 0.0
        matrix1.Xz = 0.0
        matrix1.Yx = 0.0
        matrix1.Yy = 1.0
        matrix1.Yz = 0.0
        matrix1.Zx = 0.0
        matrix1.Zy = 0.0
        matrix1.Zz = 1.0
        objPart.WCS.SetOriginAndMatrix(origin1, matrix1)

        'Coordinate system Data
        Dim wcs As NXOpen.Tag = NXOpen.Tag.Null
        Dim wcs_mx As NXOpen.Tag = NXOpen.Tag.Null

        Dim origin(2) As Double
        Dim wcs_mx_vals(8) As Double

        ufs.Csys.AskWcs(wcs)
        ufs.Csys.AskCsysInfo(wcs, wcs_mx, origin)
        ufs.Csys.AskMatrixValues(wcs_mx, wcs_mx_vals)

        'Copy the file to the output directory and then rename it
        If Not FnCheckFolderExists(_sSweepDataOutputFolderPath) Then
            SCreateDirectory(_sSweepDataOutputFolderPath)
        End If

        'Code modified on Feb-28-2018
        'Poulate all config sheet informations
        sPopulateConfigSheet(objPart, _sOemName)
        'Creating both weldment and component folder, pushing the Sweep data file based on the file type.
        'Create sub folders for weldments and components within the main tool folder
        'Code modified on Aug-26-2017
        'Code modified to create both Weldment and component folder even if there is no weldment or no component present in the tool
        sWeldmentSubFolderName = "W_" & _sToolFolderName
        sComponentSubFolderName = "C_" & _sToolFolderName

        'Configured to work for all OEM
        'In GM and Chrysler, we cannot identify the weldemnt based on the Naming convention.
        'In Daimler, we can identify the weldemnt based on the nameing convention.
        'In Fiat, we can identify the weldment based on the attribute (B_PART_TYPE = WELDED_ASS)
        If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
            If FnChkPartisWeldment(objPart) Then
                'Get the combined path
                sFolderName = Path.Combine(UNIT_SWEEP_DATA, _sToolFolderName, sWeldmentSubFolderName)
            Else
                'Get the combined path
                sFolderName = Path.Combine(UNIT_SWEEP_DATA, _sToolFolderName, sComponentSubFolderName)
            End If
        ElseIf _sOemName = DAIMLER_OEM_NAME Then
            If FnCheckIfThisIsAWeldment(sPartName) Then
                'Get the combined path
                sFolderName = Path.Combine(UNIT_SWEEP_DATA, _sToolFolderName, sWeldmentSubFolderName)
                _bIsComponent = False
            Else
                'Get the combined path
                sFolderName = Path.Combine(UNIT_SWEEP_DATA, _sToolFolderName, sComponentSubFolderName)
                _bIsComponent = True
            End If
            'Code added Nov-07-2018
            'Identify the FIAT component type
        ElseIf _sOemName = FIAT_OEM_NAME Then
            If FnChkPartisWeldmentBasedOnAttr(objPart) Then
                'Get the combined path
                sFolderName = Path.Combine(UNIT_SWEEP_DATA, _sToolFolderName, sWeldmentSubFolderName)
            Else
                'Get the combined path
                sFolderName = Path.Combine(UNIT_SWEEP_DATA, _sToolFolderName, sComponentSubFolderName)
            End If
        End If


        SCreateDirectory(Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, sComponentSubFolderName))
        SCreateDirectory(Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, sWeldmentSubFolderName))

        'Delete the existing file if present
        If FnCheckFileExists(Path.Combine(_sSweepDataOutputFolderPath, sFolderName, WRORKBOOK_NAME & "_" & sPartName & XLSM)) Then
            SDeleteFile(Path.Combine(_sSweepDataOutputFolderPath, sFolderName, WRORKBOOK_NAME & "_" & sPartName & XLSM))
        End If

        'Copy file to the output folder
        SCopyFiles(_sSweepDataOutputFolderPath & "\" & sFolderName, Path.Combine(FnGetExecutionFolderPath(), TEMPLATE_FOLDER),
                   Path.Combine(FnGetExecutionFolderPath(), TEMPLATE_FOLDER, WRORKBOOK_NAME + XLSM))
        SRenameFile(Path.Combine(_sSweepDataOutputFolderPath, sFolderName, WRORKBOOK_NAME & XLSM), WRORKBOOK_NAME & "_" & sPartName & XLSM)

        'Open the excel workbook for writing
        sSweepDataFilePath = Path.Combine(_sSweepDataOutputFolderPath, sFolderName, WRORKBOOK_NAME & "_" & sPartName & XLSM)
        _objWorkBk = FnOpenWorkbook(sSweepDataFilePath)
        _objWorkBk.Save()
        'Code added Oct-22-2018
        If Not _objApp Is Nothing Then
            'get the window handle
            Dim xlHWND As Integer = _objApp.Hwnd
            'this will have the process ID after call to GetWindowThreadProcessId
            Dim ProcIdXL As Integer = 0
            'Get the process ID
            GetWindowThreadProcessId(xlHWND, ProcIdXL)
            sWriteProcessID(ProcIdXL)
            sWriteToLogFile("Excel Process initiated with Process ID : " & ProcIdXL)
        End If
        'Code modified on Feb-28-2018
        'Values are hard coded
        'Get the actual column numbers
        sAssignPSDColumnHeaderNums()

        aoAllCompInSession = FnGetAllComponentsInSession()
        'Delete all the existing face names and edge names
        sDeleteOldSweepDataInformation(objPart, aoAllCompInSession)
        '********************************************************************************************************************************

        iRowFinishStart = MISC_INFO_START_ROW_WRITE
        iRowStartDataStart = MISC_INFO_START_ROW_WRITE

        iCount = 1
        iFaceCount = 1
        iRowFaceVecStart = 2

        'Collect all the TEXT EXTRUDE feature face (if any)
        _aoTextExtrudeFeatureFaces = FnGetTextExtrudeFeatureFaces(objPart, aoAllCompInSession)

        If Not aoAllCompInSession Is Nothing Then
            For Each objChildComp As Component In aoAllCompInSession
                'sCompInCarDivision = False
                objChildPart = FnGetPartFromComponent(objChildComp)
                If Not objChildPart Is Nothing Then
                    'Load the part fully
                    FnLoadPartFully(objChildPart)
                    'Code added Jun-01-2018
                    'Collect the solid body from the part based on OEM
                    aoAllValidSolidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)

                    If Not aoAllValidSolidBody Is Nothing Then
                        For Each objBody As Body In aoAllValidSolidBody
                            If Not objBody Is Nothing Then
                                If _sOemName = GM_OEM_NAME Or _sOemName = CHRYSLER_OEM_NAME Or _sOemName = GESTAMP_OEM_NAME Then
                                    If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                        objBodyToAnalyse = objBody
                                        sBodyName = objBodyToAnalyse.JournalIdentifier
                                    Else
                                        objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                        sBodyName = objBodyToAnalyse.JournalIdentifier & " " & objChildComp.JournalIdentifier
                                    End If
                                ElseIf _sOemName = DAIMLER_OEM_NAME Then
                                    If _sDivision = TRUCK_DIVISION Then
                                        'Check if it is the root component or it is a sub assembly compoenent
                                        If (objChildComp Is objPart.ComponentAssembly.RootComponent) Then
                                            'Component in truck. Get Prototype body
                                            objBodyToAnalyse = objBody
                                            sBodyName = objBodyToAnalyse.JournalIdentifier
                                        Else
                                            'Weldment in truck. Get Occurrence body
                                            objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                            If Not objBodyToAnalyse Is Nothing Then
                                                sBodyName = objBodyToAnalyse.JournalIdentifier & " " & objChildComp.JournalIdentifier
                                            End If
                                        End If
                                    ElseIf _sDivision = CAR_DIVISION Then
                                        'Check if the component is a child component in weldment
                                        If FnCheckIfThisIsAChildCompInWeldment(objChildComp, _sOemName) Then
                                            'Weldment in Car. Get the Occurrence body
                                            objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                            If Not objBodyToAnalyse Is Nothing Then
                                                'when populating body name, give Body journal identifier and component container journalidentifier
                                                sBodyName = objBodyToAnalyse.JournalIdentifier & " " & objChildComp.Parent.JournalIdentifier
                                            End If
                                        Else
                                            'Component in Car. Get the Prototype Body
                                            'objBodyToAnalyse = objBody
                                            'Code modified on May-14-2018
                                            'Get the occurrence body. There was some mismatch between the geomety at container level and GEo level
                                            objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                            If Not objBodyToAnalyse Is Nothing Then
                                                'sBodyName = objBodyToAnalyse.JournalIdentifier
                                                sBodyName = objBody.JournalIdentifier
                                            End If
                                        End If
                                    End If
                                    'Code added Nov-07-2018
                                    'Added Fiat OEM
                                ElseIf _sOemName = FIAT_OEM_NAME Then
                                    'Check if the component is a child component in weldment
                                    If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                        'Fiat component
                                        objBodyToAnalyse = objBody
                                        sBodyName = objBodyToAnalyse.JournalIdentifier
                                    Else
                                        'This is a Fiat Weldment child component
                                        objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                        If Not objBodyToAnalyse Is Nothing Then
                                            'when populating body name, give Body journal identifier and Child component journalidentifier
                                            sBodyName = objBodyToAnalyse.JournalIdentifier & " " & objChildComp.JournalIdentifier
                                        End If
                                    End If
                                    ''If FnCheckIfThisIsAChildCompInWeldment(objChildComp, _sOemName) Then
                                    ''    'This is a Fiat Weldment child component
                                    ''    objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                    ''    If Not objBodyToAnalyse Is Nothing Then
                                    ''        'when populating body name, give Body journal identifier and Child component journalidentifier
                                    ''        sBodyName = objBodyToAnalyse.JournalIdentifier & " " & objChildComp.JournalIdentifier
                                    ''    End If
                                    ''Else
                                    ''    'Fiat component
                                    ''    objBodyToAnalyse = objBody
                                    ''    sBodyName = objBodyToAnalyse.JournalIdentifier
                                    ''End If
                                End If

                                If Not objBodyToAnalyse Is Nothing Then

                                    sSetStatus("Collecting data for " & sBodyName.ToUpper)
                                    'Check whether the body is a solid body and the solid body should be on a range of layers
                                    If objBodyToAnalyse.IsSolidBody Then 'And (body.Layer = 1 Or (body.Layer >= 92 And body.Layer <= 184)) Then
                                        sStoreSolidBody(objBodyToAnalyse)
                                        'Check whether it is some body which is not suppose to be generated in sweep data
                                        'CODE MODIFIED - 6/13/16 - Amitabh - Ignore WIRE MESH bodies with respect to the SHAPE attribute
                                        If (FnGetStringUserAttribute(objBodyToAnalyse, _SHAPE_ATTR_NAME) <> WIRE_MESH_SHAPE) And
                                            (Not FnChkIfBodyIsMesh(objPart, objBodyToAnalyse)) Then

                                            For Each objFaceToAnalyse As Face In objBodyToAnalyse.GetFaces()
                                                'Validation added on NOV-10-2016
                                                'TEXT EXTRUDE feature face should be ignored 
                                                If Not FnCheckIfFaceToBeIgnoredBasedOnTextExtrudeFeature(objFaceToAnalyse) Then
                                                    If _bIsComponent Then
                                                        objFaceToAnalyse.Prototype.SetName("Face " & iFaceCount.ToString())
                                                    Else
                                                        objFaceToAnalyse.SetName("Face " & iFaceCount.ToString())
                                                    End If

                                                    iFaceCount = iFaceCount + 1
                                                    'Populate Face information in Face Vec sheet
                                                    sPopulateFaceInfoInFaceVecTab(objPart, objChildComp, objBodyToAnalyse, objFaceToAnalyse, iRowFaceVecStart, SHEETFACEVECTORDETAILS)
                                                    iRowFaceVecStart = iRowFaceVecStart + 1

                                                    'Populate Face attributes in MiscInfo Sheet
                                                    sPopulateFinishTolInfoOfFaceInMiscInfoTab(objFaceToAnalyse, iRowFinishStart)

                                                    For Each objEdge As Edge In objFaceToAnalyse.GetEdges()
                                                        SWriteValueToCell(_objWorkBk, SHEETNAMEWRITE, iRowStart, iColFeatName, sBodyName)
                                                        If _bIsComponent Then
                                                            If Not objEdge.Prototype.Name Is Nothing Then
                                                                If Not objEdge.Prototype.Name.Contains("BODY EDGE") Then
                                                                    objEdge.Prototype.SetName("BODY EDGE " & iCount.ToString)
                                                                    iCount = iCount + 1
                                                                End If
                                                            Else
                                                                objEdge.Prototype.SetName("BODY EDGE " & iCount.ToString)
                                                                iCount = iCount + 1
                                                            End If

                                                        Else
                                                            If Not objEdge.Name Is Nothing Then
                                                                If Not objEdge.Name.Contains("BODY EDGE") Then
                                                                    objEdge.SetName("BODY EDGE " & iCount.ToString)
                                                                    iCount = iCount + 1
                                                                Else
                                                                    'CODE ADDED - 5/13/16 - Amitabh - To add unique edge names to occurrence edges
                                                                    If objEdge.IsOccurrence Then
                                                                        If objEdge.Name.ToUpper = objEdge.Prototype.Name.ToUpper Then
                                                                            objEdge.SetName("BODY EDGE " & iCount.ToString)
                                                                            iCount = iCount + 1
                                                                        End If
                                                                    End If
                                                                End If
                                                            Else
                                                                objEdge.SetName("BODY EDGE " & iCount.ToString)
                                                                iCount = iCount + 1
                                                            End If

                                                        End If
                                                        'Populate EDge information in F_Data sheet
                                                        sPopulateEdgeInfoInFDataTab(objPart, objChildComp, objBodyToAnalyse, objFaceToAnalyse, objEdge, origin, iRowStart, SHEETNAMEWRITE)
                                                        iRowStart = iRowStart + 1
                                                    Next
                                                    'iCount = iCount + 1
                                                End If
                                            Next
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Else
            '_sStartTime = DateTime.Now
            'Code added Jun-01-2018
            'Collect the solid body from the part based on OEM
            aoAllValidSolidBody = FnGetValidBodyForOEM(objPart, _sOemName)
            'sCalculateTimeForAPI("Identify Valid Body ")
            If Not aoAllValidSolidBody Is Nothing Then
                For Each objbody As Body In aoAllValidSolidBody
                    'Check whether the body is a solid body
                    If objbody.IsSolidBody Then ' And (body.Layer = 1 Or (body.Layer >= 92 And body.Layer <= 184)) Then
                        'If objbody.Layer = 1 Then
                        sStoreSolidBody(objbody)
                        'Check whether it is some body which is not suppose to be generated in sweep data
                        'CODE MODIFIED - 6/13/16 - Amitabh - Ignore WIRE MESH bodies with respect to the SHAPE attribute
                        If (FnGetStringUserAttribute(objbody, _SHAPE_ATTR_NAME) <> WIRE_MESH_SHAPE) And
                                    (Not FnChkIfBodyIsMesh(objPart, objbody)) Then
                            sSetStatus("Collecting data for " & objbody.JournalIdentifier.ToUpper)
                            For Each objFace As Face In objbody.GetFaces()
                                'Validation added on NOV-10-2016
                                'TEXT EXTRUDE feature face should be ignored 
                                If Not FnCheckIfFaceToBeIgnoredBasedOnTextExtrudeFeature(objFace) Then
                                    objFace.SetName("Face " & iFaceCount.ToString())
                                    iFaceCount = iFaceCount + 1
                                    '_sStartTime = DateTime.Now
                                    'Populate Face information in Face Vec sheet
                                    sPopulateFaceInfoInFaceVecTab(objPart, Nothing, objbody, objFace, iRowFaceVecStart, SHEETFACEVECTORDETAILS)
                                    iRowFaceVecStart = iRowFaceVecStart + 1
                                    'sCalculateTimeForAPI("Populate Face Vec Info for a single face ")
                                    '_sStartTime = DateTime.Now
                                    'Populate Face attributes in MiscInfo Sheet
                                    sPopulateFinishTolInfoOfFaceInMiscInfoTab(objFace, iRowFinishStart)
                                    'sCalculateTimeForAPI("Populate Finish Tol For a Single Face ")
                                    For Each objEdge As Edge In objFace.GetEdges()
                                        SWriteValueToCell(_objWorkBk, SHEETNAMEWRITE, iRowStart, iColFeatName, objbody.JournalIdentifier)
                                        If Not objEdge.Name Is Nothing Then
                                            If Not objEdge.Name.Contains("BODY EDGE") Then
                                                objEdge.SetName("BODY EDGE " & iCount.ToString)
                                                iCount = iCount + 1
                                            End If
                                        Else
                                            objEdge.SetName("BODY EDGE " & iCount.ToString)
                                            iCount = iCount + 1
                                        End If
                                        '_sStartTime = DateTime.Now
                                        'Populate EDge information in F_Data sheet
                                        sPopulateEdgeInfoInFDataTab(objPart, Nothing, objbody, objFace, objEdge, origin, iRowStart, SHEETNAMEWRITE)
                                        'sCalculateTimeForAPI("Populate F_Data Info for a EDge ")
                                        iRowStart = iRowStart + 1
                                    Next
                                    'Count = iCount + 1
                                End If
                            Next
                            'End If
                        End If
                    End If
                Next
            End If
        End If

        'Code added Dec-05-2017
        'Populate all the associated machined face to the reference Relief cut face.
        sPopulateAssociatedMachinedFaceToReliefCutFace(objPart)

        ''Code commented on Apr-04-2019
        ''Flaw in Logic and Varma will confirm the logic.
        ''Code added Apr-03-2019
        ''Compute B_Primary2 view orientation
        'Try
        '    sComputePrimary2ViewOrientation(objPart)
        '    sWriteToLogFile("B_PRIMARY2 computation completed")
        'Catch ex As Exception
        '    sWriteToLogFile("Error encountered in B_PRIMARY2 view computation")
        '    sWriteToLogFile(ex.Message)
        'End Try


        'Code added Oct-06-2017
        'Write Comp Data in BodyName Sheet
        'Configured to work for all OEM
        'In GM and Chrysler, we cannot identify the weldemnt based on the Naming convention.
        'In Daimler, we can identify the weldemnt based on the nameing convention.
        If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
            If FnChkPartisWeldment(objPart) Then
                sWriteWeldmentDataInBodyNameSheet(objPart, sConfigFolderPath, aoAllCompInSession)
                'Update the mating bodies information in the sweep data along with mating faces
                sDetermineMatingBodiesandFaces(objPart)
            Else
                sWriteComponentDataInBodyNameSheet(sConfigFolderPath)
            End If
        ElseIf (_sOemName = DAIMLER_OEM_NAME) Then
            If FnCheckIfThisIsAWeldment(sPartName) Then
                sWriteWeldmentDataInBodyNameSheet(objPart, sConfigFolderPath, aoAllCompInSession)
                'Update the mating bodies information in the sweep data along with mating faces
                sDetermineMatingBodiesandFaces(objPart)
                'sWriteWeldmentDataInBodyNameSheetAlt(objPart, sConfigFolderPath, aoAllCompInSession)
            Else
                sWriteComponentDataInBodyNameSheet(sConfigFolderPath)
            End If
            'Code added Nov-08-2018
        ElseIf (_sOemName = FIAT_OEM_NAME) Then
            If FnChkPartisWeldmentBasedOnAttr(objPart) Then
                'FIAT weldment
                sWriteWeldmentDataInBodyNameSheet(objPart, sConfigFolderPath, aoAllCompInSession)
                'Update the mating bodies information in the sweep data along with mating faces
                sDetermineMatingBodiesandFaces(objPart)
            Else
                'FIAT Component
                sWriteComponentDataInBodyNameSheet(sConfigFolderPath)
            End If
        End If


        'Update the mating bodies information in the sweep data along with mating faces
        'If FnChkPartisWeldment(objPart) Then
        ''If FnCheckIfThisIsAWeldment(sPartName) Then
        ''    sDetermineMatingBodiesandFaces(objPart)
        ''End If

        'Write History Data
        'Call SWriteHistoryData(objPart)

        'Populate ADA information to the MISC INFO Tab
        sPopulateADAInfoInMiscInfoTab(objPart)

        '4/16/16 - CODE CHANGED - Amitabh - ONLY IMPLEMENT FOR BASE OR FRAME
        'All the parts in the tool have the same model views and direction cosines.
        'If FnGetPartName(objPart) <> "" Then
        'If FnGetPartName(objPart).ToUpper.Contains("BASE") Or FnGetPartName(objPart).ToUpper.Contains("FRAME") Then
        'Code modified on Aug-20-2018
        'If ADA is executed for this part, then fetch the PRIMARY1 model view cosines from the part.
        If FnGetStringUserAttribute(objPart, "B_ADA") <> "" Then
            sWriteModelingViewDirCosines(objPart)
        End If
        'End If

        'Writing Pre-fab Hole attributes to Excel sheet
        Call sAddPreFabHoleAttributeToExcel(objPart)

        'Writing Bounding Box For B_PRIMARY2
        If FnChkIfModelingViewPresent(objPart, SECOND_PRIMARY_VIEW_NAME) Then
            If FnCheckIfSheetPresent(_objWorkBk, BODYSHEETNAME & "_" & SECOND_PRIMARY_VIEW_NAME) Then
                sDeleteSheet(_objWorkBk, BODYSHEETNAME & "_" & SECOND_PRIMARY_VIEW_NAME)
            End If
            'Code changed on Oct-20-2016
            'Copying the header row only so as not to duplicate the named ranges
            'sCopySheet(_objWorkBk, BODYSHEETNAME, BODYSHEETNAME & "_" & SECOND_PRIMARY_VIEW_NAME)
            'Add a work sheet after a refsheet
            objPrimaryView2BodyNameSheet = FnAddWorkSheet(_objWorkBk, BODYSHEETNAME & "_" & SECOND_PRIMARY_VIEW_NAME, BODYSHEETNAME)
            sCopyRows(_objWorkBk, objPrimaryView2BodyNameSheet, 1, 1, BODYSHEETNAME, iLastColumnToCopy:=27, iStartColumnCopy:=1)
            'Copying existing data from Body_Sheet_Name
            iNosOfFilledRows = FnGetNumberofRows(_objWorkBk, BODYSHEETNAME, 1, 1)
            sCopyRows(_objWorkBk, objPrimaryView2BodyNameSheet, 2, iNosOfFilledRows, BODYSHEETNAME, iLastColumnToCopy:=27, iStartColumnCopy:=1)
            sWriteBoundingBoxofAllBodiesInAModelView(objPart, aoAllCompInSession, SECOND_PRIMARY_VIEW_NAME, BODYSHEETNAME & "_" & SECOND_PRIMARY_VIEW_NAME)
        End If
        'Code modified on Feb-22-2018
        'Bounding Box Information must be calculated based on B_LCS for all other parts and B_PRIMARY1 for frame/Base component
        'Writing Bounding Box for B_PRIMARY1
        'sWriteBoundingBoxofAllBodiesInAModelView(objPart, PRIMARY_VIEW_NAME, BODYSHEETNAME)

        'Code added by Shanmugam on May-10-2016
        'Code Modified on Aug-21-2018
        'Only for FRAME and BASE component, populate Visible edge names. It is not needed for any other weldments, though it has B_ADA attribute.
        If FnGetPartAttribute(objPart, "String", _PART_NAME) = "FRAME" Or FnGetPartAttribute(objPart, "String", _PART_NAME) = "BASE" Then
            ' ''Code modified on Aug-20-2018
            ' ''If ADA was executed for this part, then populate the visible edgenames. Earlier we were doing it only for Frame and Base components.
            ''If FnGetStringUserAttribute(objPart, "B_ADA") <> "" Then
            If FnChkIfModelingViewPresent(objPart, B_PRIMARY1) Then

                Call sPopulateVisibleEdgeNames(objPart)
                'Deleting & Writing Modelling View Direction Cosines
                iLastFilledRow = TITLE_ROW_NUM_VIEW_DIR_COS + 1
                iLastFilledColumn = 1
                iLastFilledRow = FnGetNumberofRows(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, 1, TITLE_ROW_NUM_VIEW_DIR_COS + 1)
                iLastFilledColumn = FnGetNumberofColumns(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, TITLE_ROW_NUM_VIEW_DIR_COS)
                'CODE MODIFIED - 6/7/16 - Amitabh - use .clear instead of .Delete as it affects named ranges
                'FnCreateRange(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, TITLE_ROW_NUM_VIEW_DIR_COS + 1, 1, iLastFilledRow, iLastFilledColumn).Delete()
                FnCreateRange(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, TITLE_ROW_NUM_VIEW_DIR_COS + 1, 1, iLastFilledRow, iLastFilledColumn).Clear()
                sWriteModelingViewDirCosines(objPart)
                'Delete MOdeling View starting with Primary1 or Primary2
                For Each objModelView As ModelingView In objPart.ModelingViews
                    If objModelView.Name.ToUpper.StartsWith("PRIMARY") Then
                        sDeleteModellingView(objPart, objModelView.Name)
                    End If
                Next
            End If
        End If

        'Code added May-10-2017
        'Write LCS information of Body and Part to the excel file
        'SWrite("Analyzing " & objPart.Leaf.ToUpper, ERROR_REPORT_FOLDER_PATH & "\" & "TEST_TIMINGS" & ".txt")
        'FnGetUFSession.UF.BeginTimer(start)
        'sWriteLCSInfo(objPart, aoAllCompInSession)

        'Code added Feb-22-2018
        'Bounding Box Information must be calculated based on B_LCS for all other parts and B_PRIMARY1 for frame/Base component
        'Writing Bounding Box for B_PRIMARY1
        'If FnGetPartName(objPart) <> "" Then
        'If FnGetPartName(objPart).ToUpper.Contains("BASE") Or FnGetPartName(objPart).ToUpper.Contains("FRAME") Then
        'Code modified on Aug-20-2018
        'If ADA was executed for this part, then populate Bounding Box information based on the PRIMARY1 view.
        If FnGetStringUserAttribute(objPart, "B_ADA") <> "" Then
            sWriteBoundingBoxofAllBodiesInAModelView(objPart, aoAllCompInSession, PRIMARY_VIEW_NAME, BODYSHEETNAME)
        Else
            sWriteBoundingBoxofAllBodiesInAModelView(objPart, aoAllCompInSession, PART_LCS_VIEW_NAME, BODYSHEETNAME)
        End If
        'End If


        'Code added Apr-18-2018
        If (_sOemName = DAIMLER_OEM_NAME) Or (_sOemName = FIAT_OEM_NAME) Then
            Try
                sPopulateBurnOutBodyInfo(objPart, _sOemName)
            Catch ex As Exception
                sWriteToLogFile("Error occurred in populating Burnout body Information")
                'Clean memory
                _aoStructPartOrientationInfo = Nothing
                _iPartOrientationIndex = 0
                _aoStructBodyOrientationInfo = Nothing
                _iOrientationIndex = 0
                'In case of any error, make sure to Unsuppress all the feature group.
                If (_sOemName = FIAT_OEM_NAME) Then
                    sAddOrRemoveFeatureGroupTemporarily(objPart, bIsSuppress:=False)
                End If
            End Try
        End If

        'FnGetUFSession().UF.EndTimer(start, tv)
        'SWrite("Real Time: " & tv.real_time.ToString, ERROR_REPORT_FOLDER_PATH & "\" & "TEST_TIMINGS" & ".txt")
        'SWrite("CPU Time: " & tv.cpu_time.ToString, ERROR_REPORT_FOLDER_PATH & "\" & "TEST_TIMINGS" & ".txt")

        'Dim displayPart As Part = theSession.Parts.Display
        'theSession.ListingWindow.SelectDevice(ListingWindow.DeviceType.Window, "")
        Call SCloseWorkBook(_objWorkBk, bSaveChanges:=True)
        sWriteToLogFile("WorkBook closed")
        '********** WRITE the FEEDER TEXT FILE TO BE USER BY CORE ALGORITHM EXE FILE *************
        'SWrite(sSweepDataFilePath & Chr(9) & "N", _
        '       Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, PART_SWEEP_DATA_TEXT_FILE))
        '*****************************************************************************************
        'Clear the memory
        Call sClearMemory()

    End Sub
    Public Sub sClearMemory()
        _aoSolidBody = Nothing
        _aoStructPartOrientationInfo = Nothing
        _iPartOrientationIndex = 0
    End Sub

    'Public Sub sWriteLogData(ByVal objPart As Part)
    '    Dim sFolderName As String = ""
    '    Dim s3DErrDesc As String = ""
    '    Dim sLayerInfo As String = ""
    '    'Populate the 3D exception report in case of missing stock size
    '    sFolderName = Split(Split(objPart.FullPath, NXPART_FILES_INPUT_FOLDER_PATH & "\")(1), "\")(0)
    '    If Not _asBodyLayersList Is Nothing Then
    '        For Each iLayer As Integer In _asBodyLayersList
    '            If sLayerInfo = "" Then
    '                sLayerInfo = iLayer.ToString
    '            Else
    '                sLayerInfo = sLayerInfo & "," & iLayer.ToString
    '            End If
    '        Next
    '        s3DErrDesc = "Bodies are also present in the layers " & sLayerInfo & " in the 3D model part " & objPart.Leaf.ToString
    '        SWrite(s3DErrDesc, OUTPUT_FILE_PATH & "\" & sFolderName & "\" & THREE_DIMENSIONAL_MODEL_EXCEPTION_REPORT_FILE_NAME)
    '    End If

    'End Sub

    Public Function GetUnloadOption(ByVal dummy As String) As Integer

        'Unloads the image when the NX session terminates
        'GetUnloadOption = NXOpen.Session.LibraryUnloadOption.AtTermination

        '----Other unload options-------
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately

        'Unloads the image explicitly, via an unload dialog
        'GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Explicitly
        '-------------------------------

    End Function
    'Public Function findradius(ByVal objectinforeport As String) As Double
    '    Return Math.Round(Double.Parse(objectinforeport.Substring(objectinforeport.IndexOf("=", objectinforeport.IndexOf("Radius")) + 1, 12)), 2)
    'End Function
    'Function FnGetEdgeRadius(ByVal edgeTag As Tag) As NXOpen.UF.UFEval.Arc
    '    Dim theUFSession As UFSession = UFSession.GetUFSession()
    '    Dim arc_evaluator As System.IntPtr
    '    Dim arc_data As NXOpen.UF.UFEval.Arc = Nothing

    '    theUFSession.Eval.Initialize(edgeTag, arc_evaluator)
    '    theUFSession.Eval.AskArc(arc_evaluator, arc_data)
    '    theUFSession.Eval.Free(arc_evaluator)

    '    FnGetEdgeRadius = arc_data
    'End Function
    Function FnGetEdgeData(ByVal edgeTag As Tag) As NXOpen.UF.UFEval.Arc
        Dim theUFSession As UFSession = UFSession.GetUFSession()
        Dim arc_evaluator As System.IntPtr
        Dim arc_data As NXOpen.UF.UFEval.Arc = Nothing

        theUFSession.Eval.Initialize(edgeTag, arc_evaluator)
        theUFSession.Eval.AskArc(arc_evaluator, arc_data)
        theUFSession.Eval.Free(arc_evaluator)

        FnGetEdgeData = arc_data
    End Function
    Function FnGetEdgeCenterForEllipse(ByVal edgeTag As Tag) As NXOpen.UF.UFEval.Ellipse
        Dim theUFSession As UFSession = UFSession.GetUFSession()
        Dim arc_evaluator As System.IntPtr
        Dim arc_data As NXOpen.UF.UFEval.Ellipse = Nothing

        theUFSession.Eval.Initialize(edgeTag, arc_evaluator)
        theUFSession.Eval.AskEllipse(arc_evaluator, arc_data)
        theUFSession.Eval.Free(arc_evaluator)

        FnGetEdgeCenterForEllipse = arc_data
    End Function

    'Public Function FnGetPartAttributeValue(ByVal objPart As Part, ByVal sTitle As String) As String
    '    Try
    '        FnGetPartAttributeValue = objPart.GetStringAttribute(sTitle)
    '    Catch ex As Exception
    '        FnGetPartAttributeValue = ""
    '    End Try
    'End Function


    Public Function FnGetViewByName(ByVal objPart As Part, ByVal sViewName As String) As DraftingView
        For Each objView As DraftingView In objPart.DraftingViews
            If objView.Name = sViewName Then
                objView.ActivateForSketching()
                FnGetViewByName = objView
                Exit For
            End If
        Next
    End Function

    'Function to get the front view name based on the X-deviation after computing the maximum Y from among the views
    Public Function FnGetFrontViewName(ByVal asData() As String) As String
        Dim iLoopIndex As Integer
        Dim asSplitData() As String
        Dim dMaxY As Double = 0.0
        Dim sPlanViewName As String = ""
        Dim sFrontViewName As String = ""
        Dim dMaxX As Double = 0.0
        Dim dXDeviation As Double = 0.0
        If UBound(asData) >= 1 Then
            dMaxY = CDbl(Split(asData(0), ",")(2))
        End If

        'Get the view details with the maximum Y Co-ordinate
        For iLoopIndex = 0 To UBound(asData)
            asSplitData = Split(asData(iLoopIndex), ",")
            If dMaxY < CDbl(asSplitData(2)) Then
                sPlanViewName = asSplitData(0)
                dMaxY = CDbl(asSplitData(2))
                dMaxX = CDbl(asSplitData(1))
            End If
        Next

        'Get the view details with the minimum deviation in the X
        For iLoopIndex = 0 To UBound(asData)
            asSplitData = Split(asData(iLoopIndex), ",")
            If Not dMaxX = CDbl(asSplitData(1)) Then
                dXDeviation = Abs(dMaxX - CDbl(asSplitData(1)))
                Exit For
            End If
        Next
        For iLoopIndex = 0 To UBound(asData)
            asSplitData = Split(asData(iLoopIndex), ",")
            If Not dMaxX = CDbl(asSplitData(1)) Then
                If dXDeviation > Abs(dMaxX - CDbl(asSplitData(1))) Then
                    dXDeviation = Abs(dMaxX - CDbl(asSplitData(1)))
                    sFrontViewName = asSplitData(0)
                End If
            End If
        Next
        FnGetFrontViewName = sFrontViewName
    End Function

    Public Sub SWriteHoleData()
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim ufs As UFSession = UFSession.GetUFSession()
        Dim lw As ListingWindow = theSession.ListingWindow
        Dim iFeatCount As Integer = 0
        Dim SelectedfeatTag As Tag = Nothing
        Dim dirX(2) As Double
        Dim dirY(2) As Double
        Dim ijk(2) As Double
        Dim Loc(2) As Double
        Dim mag As Double = 0.0
        Dim feat_Name As String = ""
        Dim iRowStart As Integer
        iRowStart = HOLE_VECT_INFO_START_ROW_WRITE

        'For Each objComp As Component In FnGetAllComponentsInSession()
        For Each body As Body In workPart.Bodies()
            Try
                Dim Feat_Tag() As Tag = Nothing
                iFeatCount = 0
                ufs.Modl.AskBodyFeats(body.Tag, Feat_Tag)
                If Not Feat_Tag Is Nothing Then
                    ufs.Modl.AskListCount(Feat_Tag, iFeatCount)
                End If
                If iFeatCount > 0 Then
                    For iLoopIndex As Integer = 0 To iFeatCount - 1
                        ufs.Modl.AskListItem(Feat_Tag, iLoopIndex, SelectedfeatTag)

                        'Get the feature name in the format by comparing tags
                        For Each objFeat As Feature In workPart.Features()
                            If objFeat.FeatureType.Contains("HOLE") Then
                                If objFeat.Tag = SelectedfeatTag Then
                                    feat_Name = objFeat.GetFeatureName
                                    Exit For
                                End If
                            End If
                        Next

                        Dim feat_Type As String = ""
                        ufs.Modl.AskFeatType(SelectedfeatTag, feat_Type)
                        If feat_Type.ToUpper.Contains("HOLE") Then
                            'ufs.Modl.AskFeatName(SelectedfeatTag, feat_Name)
                            'ufs.Modl.AskFeatLocation(SelectedfeatTag, Loc)
                            ufs.Modl.AskFeatDirection(SelectedfeatTag, dirX, dirY)
                            ufs.Vec3.Unitize(dirX, 0, mag, ijk)
                            'ufs.Disp.Conehead(UFConstants.UF_DISP_WORK_VIEW_ONLY, Loc, ijk, 0)
                            SWriteValueToCell(_objWorkBk, HOLEVECTORSHEETNAME, iRowStart, iColHoleName, feat_Name)
                            SWriteValueToCell(_objWorkBk, HOLEVECTORSHEETNAME, iRowStart, iColHoleVecX, ijk(0).ToString)
                            SChangeToNumberFormat(_objWorkBk, HOLEVECTORSHEETNAME, iRowStart, iColHoleVecX)
                            SWriteValueToCell(_objWorkBk, HOLEVECTORSHEETNAME, iRowStart, iColHoleVecY, ijk(1).ToString)
                            SChangeToNumberFormat(_objWorkBk, HOLEVECTORSHEETNAME, iRowStart, iColHoleVecY)
                            SWriteValueToCell(_objWorkBk, HOLEVECTORSHEETNAME, iRowStart, iColHoleVecZ, ijk(2).ToString)
                            SChangeToNumberFormat(_objWorkBk, HOLEVECTORSHEETNAME, iRowStart, iColHoleVecZ)
                            iRowStart = iRowStart + 1
                            feat_Name = ""
                        End If
                        'ufs.UF.Free(feat_Type)
                    Next
                End If
            Catch ex As NXOpen.NXException
                lw.Open()
                lw.WriteLine(ex.InnerException.Message)
            End Try
            'Next
        Next
    End Sub
    'Code commented on OCT-05-2017
    'Daimler configuration
    ''Public Sub sWriteComponentBodyData(sConfigFolderPath As String)
    ''    Dim theSession As Session = Session.GetSession()
    ''    Dim workPart As Part = theSession.Parts.Work
    ''    Dim iDeatailNos As String = ""
    ''    Dim sShape As String = ""
    ''    Dim sToolClass As String = ""
    ''    Dim iRowStart As Integer
    ''    Dim sStockSize As String = ""
    ''    Dim sPMat As String = ""
    ''    'For computing the exact view bounds of the bodies
    ''    'Dim min_corner(2) As Double
    ''    'Dim directions(2, 2) As Double
    ''    'Dim distances(2) As Double

    ''    'Dim sFolderName As String = ""
    ''    Dim s3DErrDesc As String = ""
    ''    Dim sBodyName As String = ""

    ''    iRowStart = BODY_INFO_START_ROW_WRITE

    ''    For Each body As Body In workPart.Bodies()
    ''        'Check if it is a solid body
    ''        If body.IsSolidBody Then
    ''            'Add the GM Toolkit Attributes
    ''            sSetGMToolkitAttributes(workPart, body, False, True)
    ''            'FnGetUFSession.Modl.AskBoundingBoxExact(body.Tag, NXOpen.Tag.Null, min_corner, directions, distances)

    ''            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyName, body.JournalIdentifier.ToString)
    ''            sShape = FnGetBodyAttribute(body, "String", SHAPE)
    ''            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyShape, sShape)
    ''            'iDeatailNos = FnGetBodyAttribute(body, "Integer", SUB_DET_NUM)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyDetailNos, iDeatailNos)
    ''            sToolClass = FnGetBodyAttribute(body, "String", TOOL_CLASS)
    ''            'To Output the stock size
    ''            'sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColStockSize, sStockSize)

    ''            If Not sToolClass = "" Then
    ''                SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, sToolClass)
    ''            Else
    ''                SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, NOTAPPLICABLE)
    ''            End If

    ''            'To get the bounding box exact for each body
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMinPointX, min_corner(0).ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMinPointY, min_corner(1).ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMinPointZ, min_corner(2).ToString)

    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorXX, directions(0, 0).ToString.ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorXY, directions(0, 1).ToString.ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorXZ, directions(0, 2).ToString.ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMagX, distances(0).ToString.ToString)

    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorYX, directions(1, 0).ToString.ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorYY, directions(1, 1).ToString.ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorYZ, directions(1, 2).ToString.ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMagY, distances(1).ToString.ToString)

    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorZX, directions(2, 0).ToString.ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorZY, directions(2, 1).ToString.ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorZZ, directions(2, 2).ToString.ToString)
    ''            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMagZ, distances(2).ToString.ToString)

    ''            'If FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("M") Then
    ''            sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE_METRIC)
    ''            'Check if the STOCK_SIZE attribute is present if the above attribute is absent
    ''            If sStockSize = "" Then
    ''                sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE)
    ''            End If
    ''            'Need to look into this attribute in case of ALT STD NC blocks
    ''            If sStockSize = "" Then
    ''                sStockSize = FnGetBodyAttribute(body, "String", TOOL_ID)
    ''            End If
    ''            'ElseIf FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("P") Then
    ''            '    sStockSize = FnGetBodyAttribute(body, "String", TOOL_ID)
    ''            '    If sStockSize = "" Then
    ''            '        sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE)
    ''            '    Else
    ''            '        sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE_METRIC)
    ''            '    End If
    ''            'End If

    ''            If sStockSize = "" Then
    ''                'Populate the 3D exception report in case of missing stock size in the component
    ''                'sFolderName = Split(sConfigFolderPath, "\")(UBound(Split(sConfigFolderPath, "\")))
    ''                sBodyName = body.JournalIdentifier
    ''                s3DErrDesc = "Stock size is missing in the body " & sBodyName & " in the 3D model part " & workPart.Leaf.ToString
    ''                'SWrite(s3DErrDesc, _sSweepDataOutputFolderPath & "\" & sFolderName & "\" & THREE_DIMENSIONAL_MODEL_EXCEPTION_REPORT_FILE_NAME)
    ''                SWrite(s3DErrDesc, Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, THREE_DIMENSIONAL_MODEL_EXCEPTION_REPORT_FILE_NAME))
    ''            End If

    ''            'Update the STOCK SIZE METRIC Information in the data
    ''            SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, STOCK_SIZE_METRIC_INFO_ROW_NOS, _
    ''                                STOCK_SIZE_METRIC_INFO_COLUMN_NOS, sStockSize)
    ''            'Add the stock size in the body name sheet.
    ''            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColStockSize, sStockSize)

    ''            'Getthe P_MAT value
    ''            sPMat = FnGetBodyAttribute(body, "String", P_MAT)

    ''            'Add the P_Mat value to the body name sheet
    ''            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColPMat, sPMat)

    ''            iRowStart = iRowStart + 1
    ''        End If
    ''    Next
    ''    'End If
    ''End Sub

    'Get Body level attribute
    Public Function FnGetBodyAttribute(ByVal objBody As Body, ByVal sType As String, ByVal sAttributeName As String) As String
        If sType = "String" Then
            Try
                FnGetBodyAttribute = FnGetStringUserAttribute(objBody, sAttributeName)
            Catch ex As Exception
                FnGetBodyAttribute = ""
            End Try
        ElseIf sType = "Integer" Then
            Try
                FnGetBodyAttribute = FnGetIntegerUserAttribute(objBody, sAttributeName).ToString
            Catch ex As Exception
                FnGetBodyAttribute = ""
            End Try
        ElseIf sType = "Real" Then
            Try
                FnGetBodyAttribute = FnGetRealUserAttribute(objBody, sAttributeName).ToString
            Catch ex As Exception
                FnGetBodyAttribute = ""
            End Try
        End If
    End Function

    Public Function FnGetFaceAttribute(ByVal objFace As Face, ByVal sType As String, ByVal sAttributeName As String) As String
        If sType = "String" Then
            Try
                FnGetFaceAttribute = FnGetStringUserAttribute(objFace, sAttributeName)
            Catch ex As Exception
                FnGetFaceAttribute = ""
            End Try
        ElseIf sType = "Integer" Then
            Try
                FnGetFaceAttribute = FnGetIntegerUserAttribute(objFace, sAttributeName).ToString
            Catch ex As Exception
                FnGetFaceAttribute = ""
            End Try
        End If
    End Function

    Public Sub SWriteHistoryData(ByVal objPart As Part)
        Dim cre_Ver As Integer = 0
        Dim mod_ver As Integer = 0
        Dim icolFeatName As Integer = 0
        Dim icolCreatedDate As Integer = 0
        Dim iColModifiedDate As Integer = 0
        Dim iRowWrite As Integer = 0

        icolFeatName = FnGetColumnNumberByName(_objWorkBk, HISTORY_SHEET_NAME, FEATURE_NAME, HISTORY_SHEET_TITLE_ROW_NOS)
        icolCreatedDate = FnGetColumnNumberByName(_objWorkBk, HISTORY_SHEET_NAME, CREATED_DATE, HISTORY_SHEET_TITLE_ROW_NOS)
        iColModifiedDate = FnGetColumnNumberByName(_objWorkBk, HISTORY_SHEET_NAME, MODIFIED_DATE, HISTORY_SHEET_TITLE_ROW_NOS)
        iRowWrite = HISTORY_SHEET_TITLE_ROW_NOS + 1
        For Each objFeat As Feature In objPart.Features()
            FnGetUFSession.Obj.AskCreModVersions(objFeat.Tag, cre_Ver, mod_ver)
            SWriteValueToCell(_objWorkBk, HISTORY_SHEET_NAME, iRowWrite, icolFeatName, objFeat.GetFeatureName)
            SWriteValueToCell(_objWorkBk, HISTORY_SHEET_NAME, iRowWrite, icolCreatedDate, FnGetTimeandDate(objPart, cre_Ver))
            SWriteValueToCell(_objWorkBk, HISTORY_SHEET_NAME, iRowWrite, iColModifiedDate, FnGetTimeandDate(objPart, mod_ver))
            iRowWrite = iRowWrite + 1
        Next
    End Sub

    Public Function FnGetTimeandDate(ByVal objPart As Part, ByVal versionNumber As Integer) As String
        Dim num_hist As Integer = 0
        Dim hist_List As IntPtr
        Dim prog As String = ""
        Dim user As String = ""
        Dim mac As String = ""
        Dim ver As Integer = 0
        Dim gmTime As Integer = 0
        Dim startDate As New DateTime(1970, 1, 1)
        Dim targetDate As DateTime

        FnGetUFSession.Part.CreateHistoryList(hist_List)
        FnGetUFSession.Part.AskPartHistory(objPart.Tag, hist_List)
        FnGetUFSession.Part.AskNumHistories(hist_List, num_hist)
        For iLoopHistInd As Integer = 0 To num_hist - 1
            FnGetUFSession.Part.AskNthHistory(hist_List, iLoopHistInd, prog, user, mac, ver, gmTime)
            If ver = versionNumber Then
                targetDate = startDate.AddSeconds(gmTime)
                FnGetTimeandDate = targetDate.ToShortDateString.ToString
                Exit For
            End If
        Next
        FnGetUFSession.Part.ClearHistoryList(hist_List)
    End Function

    'COde commented on OCT-05-2017
    'Daimler configuration
    ''Public Sub sWriteWeldmentBodyData(ByVal objPart As Part, sConfigFolderPath As String)
    ''    Dim sStockSize As String = ""
    ''    Dim iRowStart As Integer = 0
    ''    Dim sShape As String = ""
    ''    Dim sBodyName As String = ""
    ''    Dim bSubComponent As Boolean = False
    ''    Dim s3DErrDesc As String = ""
    ''    'Dim sFolderName As String = ""
    ''    Dim sToolClass As String = ""
    ''    Dim sPMat As String = ""
    ''    'For computing the exact view bounds of the bodies
    ''    'Dim min_corner(2) As Double
    ''    'Dim directions(2, 2) As Double
    ''    'Dim distances(2) As Double

    ''    iRowStart = BODY_INFO_START_ROW_WRITE
    ''    Dim dictStockSizeBodyData As Dictionary(Of String, NXObject()) = Nothing
    ''    dictStockSizeBodyData = New Dictionary(Of String, NXObject())

    ''    If Not FnGetAllComponentsInSession() Is Nothing Then
    ''        For Each objComp As Component In FnGetAllComponentsInSession()
    ''            If Not FnGetPartFromComponent(objComp) Is Nothing Then
    ''                FnLoadPartFully(FnGetPartFromComponent(objComp))
    ''                For Each body As Body In FnGetPartFromComponent(objComp).Bodies()
    ''                    'Check if the body belongs to the root comp or the sub assembly , assign a unique name as reuired by core algo
    ''                    'by joining the body name with the component instance tag
    ''                    'Check whether the body is a solid body
    ''                    'Only pick bodies which are in layer 1 (other side may also be present in the same part which need not be detailed) - 26/2/2014
    ''                    If body.IsSolidBody Then 'And (body.Layer = 1 Or (body.Layer >= 92 And body.Layer <= 184)) Then
    ''                        sSetStatus("Collecting attribute data for " & body.JournalIdentifier.ToUpper)
    ''                        sToolClass = FnGetBodyAttribute(body, "String", TOOL_CLASS)
    ''                        If Not sToolClass = "" Then
    ''                            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, sToolClass)
    ''                        End If
    ''                        'Check whether it is some body which is not suppose to be generated in sweep data
    ''                        'CODE MODIFIED - 6/13/16 - Amitabh - Ignore WIRE MESH bodies with respect to the SHAPE attribute
    ''                        If (FnGetStringUserAttribute(body, SHAPE) <> WIRE_MESH_SHAPE) And _
    ''                            (Not FnChkIfBodyIsMesh(objPart, body)) Then
    ''                            If objComp Is objPart.ComponentAssembly.RootComponent Then
    ''                                sBodyName = body.JournalIdentifier
    ''                                'Add this body to the collection of solid bodies
    ''                                sStoreSolidBody(body)

    ''                                bSubComponent = False
    ''                                'Add the GM Toolkit Attributes
    ''                                sSetGMToolkitAttributes(objPart, body, False, False)
    ''                                'If FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("M") Then
    ''                                'To Output the stock size
    ''                                sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE_METRIC)
    ''                                'Check if the STOCK_SIZE attribute is present if the above attribute is absent
    ''                                If sStockSize = "" Then
    ''                                    sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE)
    ''                                End If
    ''                                'Need to look into this attribute in case of ALT STD NC blocks
    ''                                If sStockSize = "" Then
    ''                                    sStockSize = FnGetBodyAttribute(body, "String", TOOL_ID)
    ''                                End If
    ''                                'ElseIf FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("P") Then
    ''                                '    sStockSize = FnGetBodyAttribute(body, "String", TOOL_ID)
    ''                                '    If sStockSize = "" Then
    ''                                '        sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE)
    ''                                '    End If
    ''                                '    If sStockSize = "" Then
    ''                                '        sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE_METRIC)
    ''                                '    End If
    ''                                'End If


    ''                                'FnGetUFSession.Modl.AskBoundingBoxExact(body.Tag, NXOpen.Tag.Null, min_corner, directions, distances)
    ''                            Else
    ''                                sBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier & " " & objComp.JournalIdentifier
    ''                                'Add this occurence body to the collection of solid bodies
    ''                                sStoreSolidBody(CType(objComp.FindOccurrence(body), Body))
    ''                                bSubComponent = True

    ''                                'Add the DB_PART_NAME info for all the child components in a weldment
    ''                                SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColCompDBPartName, _
    ''                                                                        FnGetCompAttribute(objComp, "String", DB_PART_NAME))

    ''                                'Add the GM Toolkit Attributes
    ''                                sSetGMToolkitAttributes(objPart, objComp, True, False)
    ''                                'If FnGetCompAttribute(objComp, "String", PURCH_OPTION).Contains("M") Then
    ''                                'To Output the stock size
    ''                                sStockSize = FnGetCompAttribute(objComp, "String", STOCK_SIZE_METRIC)
    ''                                'Check if the STOCK_SIZE attribute is present if the above attribute is absent
    ''                                If sStockSize = "" Then
    ''                                    sStockSize = FnGetCompAttribute(objComp, "String", STOCK_SIZE)
    ''                                End If
    ''                                'Need to look into this attribute in case of ALT STD NC blocks
    ''                                If sStockSize = "" Then
    ''                                    sStockSize = FnGetCompAttribute(objComp, "String", TOOL_ID)
    ''                                End If

    ''                                'If the stock size attribute is not present in the component , then pull it from the body level attribute.
    ''                                'If sStockSize = "" Then
    ''                                '    If FnGetBodyAttribute(CType(objComp.FindOccurrence(body), Body), "String", PURCH_OPTION).Contains("M") Then
    ''                                '        'To Output the stock size
    ''                                '        sStockSize = FnGetBodyAttribute(CType(objComp.FindOccurrence(body), Body), "String", STOCK_SIZE_METRIC)
    ''                                '        'Check if the STOCK_SIZE attribute is present if the above attribute is absent
    ''                                '        If sStockSize = "" Then
    ''                                '            sStockSize = FnGetBodyAttribute(CType(objComp.FindOccurrence(body), Body), "String", STOCK_SIZE)
    ''                                '        End If
    ''                                '    End If
    ''                                'End If
    ''                                '    ElseIf FnGetPartAttribute(FnGetPartFromComponent(objComp), "String", PURCH_OPTION).Contains("P") Then
    ''                                '    sStockSize = FnGetCompAttribute(objComp, "String", TOOL_ID)

    ''                                'End If
    ''                                'FnGetUFSession.Modl.AskBoundingBoxExact(CType(objComp.FindOccurrence(body), Body).Tag, _
    ''                                '                                        NXOpen.Tag.Null, min_corner, directions, distances)
    ''                            End If
    ''                            'iQnty = 1

    ''                            If Not dictStockSizeBodyData.ContainsKey(sStockSize) Then
    ''                                If Not bSubComponent Then
    ''                                    dictStockSizeBodyData.Add(sStockSize, {body})
    ''                                Else
    ''                                    dictStockSizeBodyData.Add(sStockSize, {objComp})
    ''                                End If
    ''                            Else
    ''                                If Not bSubComponent Then
    ''                                    ReDim Preserve dictStockSizeBodyData(sStockSize)(UBound(dictStockSizeBodyData(sStockSize)) + 1)
    ''                                    dictStockSizeBodyData(sStockSize)(UBound(dictStockSizeBodyData(sStockSize))) = body
    ''                                Else
    ''                                    ReDim Preserve dictStockSizeBodyData(sStockSize)(UBound(dictStockSizeBodyData(sStockSize)) + 1)
    ''                                    dictStockSizeBodyData(sStockSize)(UBound(dictStockSizeBodyData(sStockSize))) = objComp
    ''                                End If
    ''                            End If

    ''                            'Update Value to the cell
    ''                            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyName, sBodyName)
    ''                            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColComponentName, objComp.DisplayName)
    ''                            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyLayer, body.Layer.ToString)
    ''                            sShape = FnGetBodyAttribute(body, "String", SHAPE)
    ''                            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyShape, sShape)
    ''                            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColStockSize, sStockSize)
    ''                            sPMat = FnGetBodyAttribute(body, "String", P_MAT)
    ''                            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColPMat, sPMat)
    ''                            If sStockSize = "" Then
    ''                                'Populate the 3D exception report in case of missing stock size
    ''                                'sFolderName = Split(sConfigFolderPath, "\")(UBound(Split(sConfigFolderPath, "\")))
    ''                                s3DErrDesc = "Stock size is missing in the body " & sBodyName & " in the 3D model part " & objPart.Leaf.ToString
    ''                                SWrite(s3DErrDesc, Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, THREE_DIMENSIONAL_MODEL_EXCEPTION_REPORT_FILE_NAME))
    ''                            End If
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyDetailNos, dictStockSizeSubDetaildata(sStockSize).ToString())

    ''                            'To get the bounding box exact for each body
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMinPointX, min_corner(0).ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMinPointY, min_corner(1).ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMinPointZ, min_corner(2).ToString)

    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorXX, directions(0, 0).ToString.ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorXY, directions(0, 1).ToString.ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorXZ, directions(0, 2).ToString.ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMagX, distances(0).ToString.ToString)

    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorYX, directions(1, 0).ToString.ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorYY, directions(1, 1).ToString.ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorYZ, directions(1, 2).ToString.ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMagY, distances(1).ToString.ToString)

    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorZX, directions(2, 0).ToString.ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorZY, directions(2, 1).ToString.ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorZZ, directions(2, 2).ToString.ToString)
    ''                            'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMagZ, distances(2).ToString.ToString)

    ''                            iRowStart = iRowStart + 1
    ''                        End If
    ''                        'Else
    ''                        '    'To report the layers apart from layer 1 in which all the solid bodies are present
    ''                        '    If _asBodyLayersList Is Nothing Then
    ''                        '        ReDim Preserve _asBodyLayersList(0)
    ''                        '        _asBodyLayersList(0) = body.Layer
    ''                        '    ElseIf Not _asBodyLayersList.Contains(body.Layer) Then
    ''                        '        ReDim Preserve _asBodyLayersList(UBound(_asBodyLayersList) + 1)
    ''                        '        _asBodyLayersList(UBound(_asBodyLayersList)) = body.Layer
    ''                        '    End If
    ''                    End If
    ''                Next
    ''            End If
    ''        Next
    ''    Else
    ''        For Each body As Body In FnGetNxSession.Parts.Work.Bodies()
    ''            'Check whether the body is a solid body
    ''            'Only pick bodies which are in layer 1 (other side may also be present in the same part which need not be detailed) - 26/2/2014
    ''            If body.IsSolidBody Then 'And (body.Layer = 1 Or (body.Layer >= 92 And body.Layer <= 184)) Then
    ''                sSetStatus("Collecting attribute data for " & body.JournalIdentifier.ToUpper)
    ''                sToolClass = FnGetBodyAttribute(body, "String", TOOL_CLASS)
    ''                If Not sToolClass = "" Then
    ''                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, sToolClass)
    ''                End If
    ''                'Check whether it is some body which is not suppose to be generated in sweep data
    ''                'CODE MODIFIED - 6/13/16 - Amitabh - Ignore WIRE MESH bodies with respect to the SHAPE attribute
    ''                If (FnGetStringUserAttribute(body, SHAPE) <> WIRE_MESH_SHAPE) And _
    ''                            (Not FnChkIfBodyIsMesh(objPart, body)) Then
    ''                    'Add the GM Toolkit Attributes
    ''                    sSetGMToolkitAttributes(objPart, body, False, False)
    ''                    'Add this body to the collection of solid bodies
    ''                    sStoreSolidBody(body)

    ''                    'If FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("M") Then
    ''                    'To Output the stock size
    ''                    sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE_METRIC)
    ''                    'Check if the STOCK_SIZE attribute is present if the above attribute is absent
    ''                    If sStockSize = "" Then
    ''                        sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE)
    ''                    End If
    ''                    If sStockSize = "" Then
    ''                        sStockSize = FnGetBodyAttribute(body, "String", TOOL_ID)
    ''                    End If
    ''                    'ElseIf FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("P") Then
    ''                    '    sStockSize = FnGetBodyAttribute(body, "String", TOOL_ID)
    ''                    '    If sStockSize = "" Then
    ''                    '        sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE)
    ''                    '    Else
    ''                    '        sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE_METRIC)
    ''                    '    End If
    ''                    'End If

    ''                    If Not dictStockSizeBodyData.ContainsKey(sStockSize) Then
    ''                        dictStockSizeBodyData.Add(sStockSize, {body})
    ''                    Else
    ''                        ReDim Preserve dictStockSizeBodyData(sStockSize)(UBound(dictStockSizeBodyData(sStockSize)) + 1)
    ''                        dictStockSizeBodyData(sStockSize)(UBound(dictStockSizeBodyData(sStockSize))) = body
    ''                    End If

    ''                    'Get the exact bounding box of the bodies
    ''                    'FnGetUFSession.Modl.AskBoundingBoxExact(body.Tag, NXOpen.Tag.Null, min_corner, directions, distances)

    ''                    'Update Value to the cell
    ''                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyName, body.JournalIdentifier.ToString)
    ''                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColComponentName, objPart.Leaf.ToString())
    ''                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyLayer, body.Layer.ToString)
    ''                    sShape = FnGetBodyAttribute(body, "String", SHAPE)
    ''                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyShape, sShape)
    ''                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColStockSize, sStockSize)
    ''                    sPMat = FnGetBodyAttribute(body, "String", P_MAT)
    ''                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColPMat, sPMat)
    ''                    If sStockSize = "" Then
    ''                        'Populate the 3D exception report in case of missing stock size
    ''                        'sFolderName = Split(Split(objPart.FullPath, sConfigFolderPath & "\")(1), "\")(0)
    ''                        s3DErrDesc = "Stock size is missing in the body " & body.JournalIdentifier.ToString & " in the 3D model part " & objPart.Leaf.ToString
    ''                        SWrite(s3DErrDesc, Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, THREE_DIMENSIONAL_MODEL_EXCEPTION_REPORT_FILE_NAME))
    ''                    End If

    ''                    'To get the bounding box exact for each body
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMinPointX, min_corner(0).ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMinPointY, min_corner(1).ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMinPointZ, min_corner(2).ToString)

    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorXX, directions(0, 0).ToString.ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorXY, directions(0, 1).ToString.ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorXZ, directions(0, 2).ToString.ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMagX, distances(0).ToString.ToString)

    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorYX, directions(1, 0).ToString.ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorYY, directions(1, 1).ToString.ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorYZ, directions(1, 2).ToString.ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMagY, distances(1).ToString.ToString)

    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorZX, directions(2, 0).ToString.ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorZY, directions(2, 1).ToString.ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColVectorZZ, directions(2, 2).ToString.ToString)
    ''                    'SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColMagZ, distances(2).ToString.ToString)

    ''                    iRowStart = iRowStart + 1
    ''                End If
    ''                'Else
    ''                '    'To report the layers apart from layer 1 in which all the solid bodies are present
    ''                '    If _asBodyLayersList Is Nothing Then
    ''                '        ReDim Preserve _asBodyLayersList(0)
    ''                '        _asBodyLayersList(0) = body.Layer
    ''                '    ElseIf Not _asBodyLayersList.Contains(body.Layer) Then
    ''                '        ReDim Preserve _asBodyLayersList(UBound(_asBodyLayersList) + 1)
    ''                '        _asBodyLayersList(UBound(_asBodyLayersList)) = body.Layer
    ''                '    End If
    ''            End If
    ''        Next
    ''    End If

    ''    'Now Check for any physical differences and then renumber sub details if required
    ''    sSetStatus("Analysing solid bodies for sub-details ")
    ''    sCheckForPhysicalDifferencesInSubDetailBasedOnFaceArea(objPart, dictStockSizeBodyData)
    ''End Sub
    'To get the dictionary of sub detail number based on the face area 
    Public Sub sCheckForPhysicalDifferencesInSubDetailBasedOnFaceArea(ByVal objPart As Part, ByVal dictStockSizeCompData As Dictionary(Of String, NXObject()))
        Dim pair As KeyValuePair(Of String, NXObject())
        Dim _MaxSubDetailNumber As Integer = 0
        Dim dFaceArea As Integer = 0
        Dim dPostFabHoleAreaToAdd As Integer = 0
        'Variable to store the face area after adding the post fab holes area
        Dim dToTalFaceArea As Integer = 0
        Dim adBodyFaceAreas() As Integer = Nothing
        Dim adBodyFaceAreasToCompare() As Integer = Nothing
        Dim iSubDetNum As Integer = SUB_DET_NUM_START_VALUE
        'Dim bSubComponent As Boolean = False
        'To store the Body or (Component Object in case of sub assemblies in weldment) and their corresposnding assigned sub detail number 
        Dim dictSubDetailNumber As Dictionary(Of NXObject, Integer) = Nothing
        'To store the body and its all faces (adjusted area after including the area of the post fab holes)
        Dim dictBodyFacesArea As Dictionary(Of NXObject, Integer()) = Nothing
        dictSubDetailNumber = New Dictionary(Of NXObject, Integer)
        Dim objPartFrmComp As Part = Nothing
        Dim objHoleVert1 As Point3d = Nothing
        Dim objHoleVert2 As Point3d = Nothing
        Dim bSimilarFaceArea As Boolean = False
        Dim objCompToAnalysis As Component = Nothing

        'key - Stock Size
        'Value - List of solid bodies with the same stock size
        For Each pair In dictStockSizeCompData
            dictBodyFacesArea = New Dictionary(Of NXObject, Integer())
            'If it has nore than one bodies with the same stock size
            If Not UBound(pair.Value) = 0 Then
                'Check if the value of the dictionary key is a body or a component
                If pair.Value(0).GetType().ToString().ToUpper = "NXOPEN.BODY" Then
                    'bSubComponent = False
                    For Each objBody As Body In pair.Value
                        For Each objFace In objBody.GetFaces()
                            'Check only the planar faces
                            If objFace.SolidFaceType = Face.FaceType.Planar Then
                                'Compute the actual face area, get the integer part as small modelling errors may be there hence
                                'ignoring the decimal values
                                dFaceArea = CInt(FnCalculateFaceArea(objPart, objFace))
                                'Check if the face has any circular edge
                                For Each objEdge As Edge In objFace.GetEdges
                                    'Can be a hole face , get the hole cylindrical face
                                    If objEdge.SolidEdgeType = Edge.EdgeType.Circular Then
                                        'For a hole the vertex 1 and vertex 2 should be same
                                        objEdge.GetVertices(objHoleVert1, objHoleVert2)
                                        If objHoleVert1.Equals(objHoleVert2) Then
                                            For Each objConnectedFace As Face In objEdge.GetFaces()
                                                'Hole face will be always cylindrical
                                                If objConnectedFace.SolidFaceType = Face.FaceType.Cylindrical Then
                                                    'Check if the hole face is PostFab based on attribute value assigned by the modeler
                                                    If FnGetFaceAttribute(objConnectedFace, "String", PRE_FAB_HOLE_ATTR_TITLE) <> PRE_FAB_HOLE_ATTR_VALUE Then
                                                        'Compute the post fab hole area
                                                        'get the circular edge dia
                                                        If dPostFabHoleAreaToAdd = 0 Then
                                                            dPostFabHoleAreaToAdd = CInt(PI * Pow(FnGetEdgeData(objEdge.Tag).radius, 2))
                                                        Else
                                                            dPostFabHoleAreaToAdd = dPostFabHoleAreaToAdd + CInt(PI * Pow(FnGetEdgeData(objEdge.Tag).radius, 2))
                                                        End If
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                                'Adding the post fab area
                                dToTalFaceArea = dFaceArea + dPostFabHoleAreaToAdd
                                If dictBodyFacesArea.ContainsKey(objBody) Then
                                    ReDim Preserve dictBodyFacesArea(objBody)(UBound(dictBodyFacesArea(objBody)) + 1)
                                    dictBodyFacesArea(objBody)(UBound(dictBodyFacesArea(objBody))) = dToTalFaceArea
                                Else
                                    dictBodyFacesArea.Add(objBody, {dToTalFaceArea})
                                End If
                                'Reset the counter for the new face
                                dPostFabHoleAreaToAdd = 0
                            End If
                        Next
                    Next
                Else
                    'It is a component
                    'check for all the bodies inside the components for physical differences
                    For Each objComp As Component In pair.Value
                        'Code modified on Feb-28-2018
                        'In case of truck division, the objComp will be the exact child component
                        'In case of Car division, the objCOmp will be the container child component within the weldment.
                        'So get the immediate child component in case of car division, which has the solid body
                        If (_sOemName = DAIMLER_OEM_NAME) Then
                            If _sDivision = TRUCK_DIVISION Then
                                objCompToAnalysis = objComp
                            ElseIf _sDivision = CAR_DIVISION Then
                                objCompToAnalysis = objComp.GetChildren(0)
                            End If
                            'Code added Mar-12-2019
                            'GM also will have a child components within the weldment (Example Pivot ear in a Clamp arm and Frame component containing child component)
                        ElseIf (_sOemName = FIAT_OEM_NAME) Or (_sOemName = GM_OEM_NAME) Then
                            objCompToAnalysis = objComp
                        End If

                        objPartFrmComp = FnGetPartFromComponent(objCompToAnalysis)
                        If Not objPartFrmComp Is Nothing Then
                            'Get all the bodies inside each component
                            For Each objBody As Body In objPartFrmComp.Bodies()
                                'Get all the faces inside each body
                                For Each objFace As Face In objBody.GetFaces()
                                    'Check for only planar faces
                                    If objFace.SolidFaceType = Face.FaceType.Planar Then
                                        'Compute the actual face area, get the integer part as small modelling errors may be there hence
                                        'ignoring the decimal values
                                        dFaceArea = CInt(FnCalculateFaceArea(objPart, objFace))
                                        'Check if the face has any circular edge
                                        For Each objEdge As Edge In objFace.GetEdges
                                            'Can be a hole face , get the hole cylindrical face
                                            If objEdge.SolidEdgeType = Edge.EdgeType.Circular Then
                                                'For a hole the vertex 1 and vertex 2 should be same
                                                objEdge.GetVertices(objHoleVert1, objHoleVert2)
                                                If objHoleVert1.Equals(objHoleVert2) Then
                                                    For Each objConnectedFace As Face In objEdge.GetFaces()
                                                        'Hole face will be always cylindrical
                                                        If objConnectedFace.SolidFaceType = Face.FaceType.Cylindrical Then
                                                            'Check if the hole face is PostFab based on attribute value assigned by the modeler
                                                            If FnGetFaceAttribute(objConnectedFace, "String", PRE_FAB_HOLE_ATTR_TITLE) <> PRE_FAB_HOLE_ATTR_VALUE Then
                                                                'Compute the post fab hole area
                                                                'get the circular edge dia
                                                                If dPostFabHoleAreaToAdd = 0 Then
                                                                    dPostFabHoleAreaToAdd = CInt(PI * Pow(FnGetEdgeData(objEdge.Tag).radius, 2))
                                                                Else
                                                                    dPostFabHoleAreaToAdd = dPostFabHoleAreaToAdd +
                                                                                                CInt(PI * Pow(FnGetEdgeData(objEdge.Tag).radius, 2))
                                                                End If
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                        'Rounding off the area for matching purpose to 6 decimal places
                                        dToTalFaceArea = dFaceArea + dPostFabHoleAreaToAdd
                                        If dictBodyFacesArea.ContainsKey(objComp) Then
                                            ReDim Preserve dictBodyFacesArea(objComp)(UBound(dictBodyFacesArea(objComp)) + 1)
                                            dictBodyFacesArea(objComp)(UBound(dictBodyFacesArea(objComp))) = dToTalFaceArea
                                        Else
                                            dictBodyFacesArea.Add(objComp, {dToTalFaceArea})
                                        End If
                                        'Reset the counter for the new face
                                        dPostFabHoleAreaToAdd = 0
                                    End If
                                Next
                            Next
                        End If
                        'If Not dictSubDetailNumber.ContainsKey(objComp) Then
                        '    dictSubDetailNumber.Add(objComp, iSubDetNum)
                        'End If
                        'bSubComponent = True
                    Next
                    'iSubDetNum = iSubDetNum + 1
                End If

                'Check the body faces area to see if all the sub detail bodies with the stock size are same
                For Each skey In dictBodyFacesArea.Keys()
                    If Not dictSubDetailNumber.ContainsKey(skey) Then
                        'Sorting will help in comapring the sequence
                        Array.Sort(dictBodyFacesArea(skey))
                        adBodyFaceAreasToCompare = dictBodyFacesArea(skey)
                        For Each comparingBodyFacePair In dictBodyFacesArea
                            Array.Sort(comparingBodyFacePair.Value)
                            bSimilarFaceArea = False
                            'Check if the number of faces are same by checking the count of the Face Area
                            'If the count is different then the number of faces in the matching body is different 
                            'For sure this means that the 2 bodies in comparison are different
                            If UBound(adBodyFaceAreasToCompare) = UBound(comparingBodyFacePair.Value) Then
                                For iArrayIndex As Integer = 0 To UBound(adBodyFaceAreasToCompare)
                                    'Check for the area within a tolerance of 1 sq.mm
                                    If Abs(adBodyFaceAreasToCompare(iArrayIndex) - comparingBodyFacePair.Value(iArrayIndex)) <= 1 Then
                                        bSimilarFaceArea = True
                                    Else
                                        bSimilarFaceArea = False
                                        'Exit the loop if even one of the area values do not match
                                        Exit For
                                    End If
                                Next
                            End If
                            If bSimilarFaceArea Then
                                dictSubDetailNumber.Add(comparingBodyFacePair.Key, iSubDetNum)
                            End If
                        Next
                        'Increase the sub detail number
                        iSubDetNum = iSubDetNum + 1
                    End If
                Next
                adBodyFaceAreasToCompare = Nothing
                dictBodyFacesArea = Nothing
            Else
                'Only one solid body with the given stock size
                dictSubDetailNumber.Add(pair.Value(0), iSubDetNum)
                'Increase the sub detail number
                iSubDetNum = iSubDetNum + 1
            End If
        Next

        'Call function to populate the sub detail information
        sPopulateWeldmentSubdetailInfo(dictSubDetailNumber)
    End Sub

    'To Compute the sub detail number based on the net mass of the body as usggested by Eric Simmons of Hirotec
    'Public Sub sCheckForPhysicalDifferencesInSubDetailBasedOnNetMass(ByVal objPart As Part, ByVal dictStockSizeBodyData As Dictionary(Of String, NXObject()))
    '    Dim pair As KeyValuePair(Of String, NXObject())
    '    'Variable to store the mass
    '    Dim dBodyMass As Double = 0.0
    '    Dim dComponentMass As Double = 0.0
    '    Dim iSubDetNum As Integer = SUB_DET_NUM_START_VALUE - 1
    '    'To store the Body or (Component Object in case of sub assemblies in weldment) and their corresposnding assigned sub detail number 
    '    Dim dictSubDetailNumber As Dictionary(Of NXObject, Integer) = Nothing
    '    dictSubDetailNumber = New Dictionary(Of NXObject, Integer)
    '    Dim objPartFrmComp As Part = Nothing

    '    'key - Stock Size
    '    'Value - List of solid bodies with the same stock size
    '    For Each pair In dictStockSizeBodyData
    '        dBodyMass = 0.0
    '        'If it has nore than one bodies with the same stock size
    '        If Not UBound(pair.Value) = 0 Then
    '            'Check if the value of the dictionary key is a body or a component
    '            If pair.Value(0).GetType().ToString().ToUpper = "NXOPEN.BODY" Then
    '                For Each objBody As Body In pair.Value
    '                    If dBodyMass = 0.0 Then
    '                        'Compute the mass of each body
    '                        dBodyMass = Round(CDbl(FnCalculateMassProperties(objPart, objBody)), 6)
    '                        iSubDetNum = iSubDetNum + 1
    '                        dictSubDetailNumber.Add(objBody, iSubDetNum)
    '                    Else
    '                        If dBodyMass <> Round(CDbl(FnCalculateMassProperties(objPart, objBody)), 6) Then
    '                            dBodyMass = Round(CDbl(FnCalculateMassProperties(objPart, objBody)), 6)
    '                            iSubDetNum = iSubDetNum + 1
    '                        End If
    '                        dictSubDetailNumber.Add(objBody, iSubDetNum)
    '                    End If
    '                Next
    '            Else
    '                'It is a component
    '                'check for all the bodies inside the components for physical differences
    '                For Each objComp As Component In pair.Value
    '                    dComponentMass = 0.0
    '                    objPartFrmComp = FnGetPartFromComponent(objComp)
    '                    'Get all the bodies inside each component
    '                    For Each objBody As Body In objPartFrmComp.Bodies()
    '                        If dComponentMass = 0.0 Then
    '                            'Compute the total mass of all the bodies inside the component
    '                            dComponentMass = Round(CDbl(FnCalculateMassProperties(objPart, objBody)), 6)
    '                        Else
    '                            'Compute the total mass of all the bodies inside the component
    '                            dComponentMass = dComponentMass + Round(CDbl(FnCalculateMassProperties(objPart, objBody)), 6)
    '                        End If
    '                    Next
    '                    dictSubDetailNumber.Add(objComp, iSubDetNum)
    '                Next
    '                iSubDetNum = iSubDetNum + 1
    '            End If
    '        Else
    '            'Increase the sub detail number
    '            iSubDetNum = iSubDetNum + 1
    '            'Only one solid body with the given stock size
    '            dictSubDetailNumber.Add(pair.Value(0), iSubDetNum)
    '        End If
    '    Next

    '    'Call function to populate the sub detail information
    '    sPopulateWeldmentSubdetailInfo(dictSubDetailNumber)
    'End Sub
    'To populate the subdetail number in the excel file Body Name Tab
    Public Sub sPopulateWeldmentSubdetailInfo(ByVal dictSubDetNum As Dictionary(Of NXObject, Integer))
        Dim iNosOfFilledRows As Integer
        Dim objComp As Component = Nothing
        Dim objCompPart As Part = Nothing
        Dim sBodyName As String = ""
        Dim sCompBodyName As String = ""
        Dim sSubDetNumValue As Integer = 0
        Dim iSubDetQnty As Integer = 0
        Dim dictBodyQuantity As Dictionary(Of NXObject, Integer) = Nothing
        dictBodyQuantity = New Dictionary(Of NXObject, Integer)
        Dim bProcess As Boolean = False

        'Get the number of rows of filled data
        iNosOfFilledRows = FnGetNumberofRows(_objWorkBk, BODYSHEETNAME, 1, BODY_INFO_START_ROW_WRITE)
        For Each pair In dictSubDetNum
            If pair.Key.GetType.ToString.ToUpper = "NXOPEN.BODY" Then
                For iloopIndex = BODY_INFO_START_ROW_WRITE To iNosOfFilledRows
                    sBodyName = FnReadSingleRowForColumn(_objWorkBk, BODYSHEETNAME, 1, iloopIndex)
                    If sBodyName = pair.Key.JournalIdentifier Then
                        sSetIntegerUserAttribute(pair.Key, _SUB_DETAIL_NUMBER, pair.Value, "")
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iloopIndex, iColBodyDetailNos, pair.Value.ToString)
                        Exit For
                    End If
                Next
            Else
                'For sub components
                objComp = CType(pair.Key, Component)
                objCompPart = FnGetPartFromComponent(objComp)
                For Each Body As Body In _aoSolidBody
                    bProcess = False
                    'sCompBodyName = CType(objComp.FindOccurrence(Body), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                    'Code added May-14-2018
                    'Validation added to check if the body is alive.
                    Try
                        If Not Body Is Nothing Then
                            Dim iStatus As Integer
                            iStatus = FnGetUFSession.Obj.AskStatus(Body.Tag)
                            If iStatus = UFConstants.UF_OBJ_ALIVE Then
                                bProcess = True
                            End If
                        End If
                    Catch ex As Exception
                        bProcess = False
                    End Try
                    If bProcess Then
                        sCompBodyName = Body.JournalIdentifier & " " & objComp.JournalIdentifier
                        For iloopIndex = BODY_INFO_START_ROW_WRITE To iNosOfFilledRows
                            sBodyName = FnReadSingleRowForColumn(_objWorkBk, BODYSHEETNAME, 1, iloopIndex)
                            If sBodyName = sCompBodyName Then
                                sSetIntegerUserAttribute(objComp, _SUB_DETAIL_NUMBER, pair.Value, "")
                                'Add the sub detail number at part level for each child component so that this value can be used at Sub Detail Callout Module
                                sSetIntegerUserAttribute(objCompPart, _SUB_DETAIL_NUMBER, pair.Value, "")
                                SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iloopIndex, iColBodyDetailNos, pair.Value.ToString)
                                Exit For
                            End If
                        Next
                    End If
                Next
                'For Each body As Body In FnGetPartFromComponent(objComp).Bodies()
                '    sCompBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                '    For iloopIndex = BODY_INFO_START_ROW_WRITE To iNosOfFilledRows
                '        sBodyName = FnReadSingleRowForColumn(_objWorkBk, BODYSHEETNAME, 1, iloopIndex)
                '        If sBodyName = sCompBodyName Then
                '            sSetIntegerUserAttribute(objComp, SUB_DET_NUM, pair.Value)
                '            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iloopIndex, iColBodyDetailNos, pair.Value.ToString)
                '            Exit For
                '        End If
                '    Next
                'Next
            End If
        Next

        'Add the Quantity Data to the sub details
        For Each pair In dictSubDetNum
            iSubDetQnty = 0
            'Get the sub detail number value
            sSubDetNumValue = pair.Value
            For Each sValue In dictSubDetNum.Values()
                If sSubDetNumValue = sValue Then
                    iSubDetQnty = iSubDetQnty + 1
                End If
            Next
            'Add the quantity value for each sub detail
            dictBodyQuantity.Add(pair.Key, iSubDetQnty)
        Next

        'populate the QTY attribute in the solid body or ( component in case of sub assemblies)
        For Each pair In dictBodyQuantity
            'Code modified on Jan-19-2020
            'Most of the attributes belong to GM and should not be under VECTRA category
            If _sOemName = GM_OEM_NAME Then
                sSetStringUserAttribute(pair.Key, _QTY, pair.Value.ToString, "")
            Else
                sSetStringUserAttribute(pair.Key, _QTY, pair.Value.ToString)
            End If

            sUpdateAttributesInModel()
        Next

        dictBodyQuantity = Nothing
        dictSubDetNum = Nothing

    End Sub
    'Compute the face area 
    Public Function FnCalculateFaceArea(ByVal objPart As Part, ByVal objFace As Face) As Double
        Dim areaUnit As Unit = objPart.UnitCollection.GetBase("Area")
        Dim lengthUnit As Unit = objPart.UnitCollection.GetBase("Length")
        'Get the area for the face
        FnCalculateFaceArea = objPart.MeasureManager.NewFaceProperties(areaUnit, lengthUnit, 0.999, {objFace}).Area
    End Function
    'Write the Qunatity Information
    'Public Sub sWriteWeldmentQuantityData(ByVal objPart As Part)
    '    Dim sStockSize As String = ""
    '    If Not FnGetAllComponentsInSession() Is Nothing Then
    '        For Each objComp As Component In FnGetAllComponentsInSession()
    '            For Each body As Body In FnGetPartFromComponent(objComp).Bodies()
    '                If objComp Is objPart.ComponentAssembly.RootComponent Then
    '                    If FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("M") Then
    '                        'To Output the stock size
    '                        sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE_METRIC)
    '                        'Check if the STOCK_SIZE attribute is present if the above attribute is absent
    '                        If sStockSize = "" Then
    '                            sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE)
    '                        End If
    '                    ElseIf FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("P") Then
    '                        sStockSize = FnGetBodyAttribute(body, "String", TOOL_ID)
    '                    End If
    '                    body.SetAttribute(QTY, dictStockQuantity(sStockSize).ToString)
    '                Else
    '                    If FnGetCompAttribute(objComp, "String", PURCH_OPTION).Contains("M") Then
    '                        'To Output the stock size
    '                        sStockSize = FnGetCompAttribute(objComp, "String", STOCK_SIZE_METRIC)
    '                        'Check if the STOCK_SIZE attribute is present if the above attribute is absent
    '                        If sStockSize = "" Then
    '                            sStockSize = FnGetCompAttribute(objComp, "String", STOCK_SIZE)
    '                        End If
    '                    ElseIf FnGetCompAttribute(objComp, "String", PURCH_OPTION).Contains("P") Then
    '                        sStockSize = FnGetCompAttribute(objComp, "String", TOOL_ID)
    '                    End If
    '                    objComp.SetAttribute(QTY, dictStockQuantity(sStockSize).ToString)
    '                End If
    '            Next
    '        Next
    '    Else
    '        For Each body As Body In FnGetNxSession.Parts.Work.Bodies()
    '            If FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("M") Then
    '                'To Output the stock size
    '                sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE_METRIC)
    '                'Check if the STOCK_SIZE attribute is present if the above attribute is absent
    '                If sStockSize = "" Then
    '                    sStockSize = FnGetBodyAttribute(body, "String", STOCK_SIZE)
    '                End If
    '            ElseIf FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("P") Then
    '                sStockSize = FnGetBodyAttribute(body, "String", TOOL_ID)
    '            End If
    '            body.SetAttribute(QTY, dictStockQuantity(sStockSize).ToString)
    '        Next
    '    End If
    'End Sub
    Public Function FnChkPartisWeldment(ByVal objPart As Part) As Boolean
        Dim objBodyCol() As Body = objPart.Bodies.ToArray()
        Dim iSolidBodyCount As Integer = 0
        If Not objBodyCol Is Nothing Then
            For Each objBody As Body In objBodyCol
                If objBody.IsSolidBody Then
                    iSolidBodyCount = iSolidBodyCount + 1
                End If
            Next
            'Check for sub components as well if present
            If Not objPart.ComponentAssembly.RootComponent Is Nothing Then
                If UBound(objPart.ComponentAssembly.RootComponent.GetChildren().ToArray) >= 0 Then
                    For Each objChildComp As Component In objPart.ComponentAssembly.RootComponent.GetChildren()
                        If Not FnGetPartFromComponent(objChildComp) Is Nothing Then
                            FnLoadPartFully(FnGetPartFromComponent(objChildComp))
                            For Each body As Body In FnGetPartFromComponent(objChildComp).Bodies()
                                If body.IsSolidBody Then
                                    iSolidBodyCount = iSolidBodyCount + 1
                                End If
                            Next
                        End If
                    Next
                End If
            End If
            If iSolidBodyCount > 1 Then
                FnChkPartisWeldment = True
            Else
                FnChkPartisWeldment = False
            End If
        End If
    End Function
    'Code commented on Oct-05-2017
    'Daimler configuration
    ' ''Add the GM Toolkit Attributes
    Public Sub sSetGMToolkitAttributes(ByVal objPart As Part, ByVal objToAddAttribute As NXObject, ByVal bSubComponent As Boolean,
                                       ByVal bIsComponent As Boolean)
        Dim dMass As Double = 0.0

        'Add the attribute to the main part first as the same part will be a sub component in the unit tool assembly
        If FnGetPartAttribute(objPart, "String", _TOOL_CLASS) = COMM Or
                                            FnGetPartAttribute(objPart, "String", _TOOL_CLASS) = STD Then
            'Code modified on Jan-19-2020
            'Most of the attributes belong to GM and should not be under VECTRA category
            If _sOemName = GM_OEM_NAME Then
                sSetStringUserAttribute(objPart, _PURCH_OPTION, PURCHASE, "")
            Else
                sSetStringUserAttribute(objPart, _PURCH_OPTION, PURCHASE)
            End If

            sUpdateAttributesInModel()
        Else
            'Code modified on Jan-19-2020
            'Most of the attributes belong to GM and should not be under VECTRA category
            If _sOemName = GM_OEM_NAME Then
                sSetStringUserAttribute(objPart, _PURCH_OPTION, MAKE_DETAIL, "")
            Else
                sSetStringUserAttribute(objPart, _PURCH_OPTION, MAKE_DETAIL)
            End If

            sUpdateAttributesInModel()
        End If

        For Each objBody As Body In objPart.Bodies()
            If FnCalculateMassProperties(objPart, objBody) <> "" Then
                dMass = dMass + CDbl(FnCalculateMassProperties(objPart, objBody))
            End If
        Next
        If FnGetPartAttribute(objPart, "String", _P_MASS) = "" Then
            'Rounding off to 5 decimal places as checkmate throws error
            'Code modified on Jan-19-2020
            'Most of the attributes belong to GM and should not be under VECTRA category
            If _sOemName = GM_OEM_NAME Then
                sSetStringUserAttribute(objPart, _P_MASS, Round(dMass, 5).ToString, "")
            Else
                sSetStringUserAttribute(objPart, _P_MASS, Round(dMass, 5).ToString)
            End If

        End If

        If bSubComponent Then
            If FnGetCompAttribute(objToAddAttribute, "String", _TOOL_CLASS) = COMM Or
                                            FnGetCompAttribute(objToAddAttribute, "String", _TOOL_CLASS) = STD Then
                'Code modified on Jan-19-2020
                'Most of the attributes belong to GM and should not be under VECTRA category
                If _sOemName = GM_OEM_NAME Then
                    sSetStringUserAttribute(objToAddAttribute, _PURCH_OPTION, PURCHASE, "")
                Else
                    sSetStringUserAttribute(objToAddAttribute, _PURCH_OPTION, PURCHASE)
                End If

            Else
                'Code modified on Jan-19-2020
                'Most of the attributes belong to GM and should not be under VECTRA category
                If _sOemName = GM_OEM_NAME Then
                    sSetStringUserAttribute(objToAddAttribute, _PURCH_OPTION, MAKE_DETAIL, "")
                Else
                    sSetStringUserAttribute(objToAddAttribute, _PURCH_OPTION, MAKE_DETAIL)
                End If

            End If
            For Each objBody As Body In FnGetPartFromComponent(CType(objToAddAttribute, Component)).Bodies()
                If FnCalculateMassProperties(objPart, objBody) <> "" Then
                    dMass = dMass + CDbl(FnCalculateMassProperties(objPart, objBody))
                End If
            Next
            If FnGetCompAttribute(objToAddAttribute, "String", _P_MASS) = "" Then
                'Code modified on Jan-19-2020
                'Most of the attributes belong to GM and should not be under VECTRA category
                If _sOemName = GM_OEM_NAME Then
                    'Rounding off to 5 decimal places as checkmate throws error
                    sSetStringUserAttribute(objToAddAttribute, _P_MASS, Round(dMass, 5).ToString, "")
                Else
                    'Rounding off to 5 decimal places as checkmate throws error
                    sSetStringUserAttribute(objToAddAttribute, _P_MASS, Round(dMass, 5).ToString)
                End If

            End If
        Else
            'Add the attributes to each body as well
            If FnGetBodyAttribute(objToAddAttribute, "String", _TOOL_CLASS) = COMM Or
                                            FnGetBodyAttribute(objToAddAttribute, "String", _TOOL_CLASS) = STD Then
                'Code modified on Jan-19-2020
                'Most of the attributes belong to GM and should not be under VECTRA category
                If _sOemName = GM_OEM_NAME Then
                    sSetStringUserAttribute(objToAddAttribute, _PURCH_OPTION, PURCHASE, "")
                Else
                    sSetStringUserAttribute(objToAddAttribute, _PURCH_OPTION, PURCHASE)
                End If

            Else
                'Code modified on Jan-19-2020
                'Most of the attributes belong to GM and should not be under VECTRA category
                If _sOemName = GM_OEM_NAME Then
                    sSetStringUserAttribute(objToAddAttribute, _PURCH_OPTION, MAKE_DETAIL, "")
                Else
                    sSetStringUserAttribute(objToAddAttribute, _PURCH_OPTION, MAKE_DETAIL)
                End If

            End If

            'Delete the attribute if present
            Try
                sDeleteIntegerUserAttribute(objToAddAttribute, _P_MASS)
                sUpdateAttributesInModel()
            Catch ex As Exception
            End Try
            If FnGetBodyAttribute(objToAddAttribute, "Real", _P_MASS) = "" Then
                If Not FnCalculateMassProperties(objPart, objToAddAttribute) = "" Then
                    'Rounding off to 5 decimal places as checkmate throws error
                    sSetRealUserAttribute(objPart, objToAddAttribute, _P_MASS, Round(CDbl(FnCalculateMassProperties(objPart, objToAddAttribute)), 5))
                    sUpdateAttributesInModel()
                End If
            End If
        End If

        'In case of component set the quantity as 1
        If bIsComponent Then
            'Delete the attribute if present
            Try
                sDeleteIntegerUserAttribute(objToAddAttribute, _QTY)
            Catch ex As Exception
                sDeleteIntegerUserAttribute(objToAddAttribute, _QTY)
            End Try
            sUpdateAttributesInModel()
            'Code modified on Jan-19-2020
            'Most of the attributes belong to GM and should not be under VECTRA category
            If _sOemName = GM_OEM_NAME Then
                sSetStringUserAttribute(objToAddAttribute, _QTY, "1", "")
            Else
                sSetStringUserAttribute(objToAddAttribute, _QTY, "1")
            End If

        End If

        sUpdateAttributesInModel()

    End Sub
    Public Function FnGetCompAttribute(ByVal objcomp As Component, ByVal sType As String, ByVal sAttributeName As String) As String
        If sType = "String" Then
            Try
                FnGetCompAttribute = FnGetStringUserAttribute(objcomp, sAttributeName)
            Catch ex As Exception
                FnGetCompAttribute = ""
            End Try
        ElseIf sType = "Integer" Then
            Try
                FnGetCompAttribute = FnGetIntegerUserAttribute(objcomp, sAttributeName).ToString
            Catch ex As Exception
                FnGetCompAttribute = ""
            End Try
        End If
        sUpdateAttributesInModel()
    End Function
    'To get the part attributes
    Public Function FnGetPartAttribute(ByVal objPart As Part, ByVal sType As String, ByVal sAttributeName As String) As String
        If sType = "String" Then
            Try
                FnGetPartAttribute = FnGetStringUserAttribute(objPart, sAttributeName)
            Catch ex As Exception
                FnGetPartAttribute = ""
            End Try
        ElseIf sType = "Integer" Then
            Try
                FnGetPartAttribute = FnGetIntegerUserAttribute(objPart, sAttributeName).ToString
            Catch ex As Exception
                FnGetPartAttribute = ""
            End Try
        End If
        sUpdateAttributesInModel()
    End Function
    'To Check the Edge Curvature whether Concave or Convex
    Public Function FnIsEdgeConvex(ByVal objPart As Part, ByVal objEdge As Edge) As Boolean
        Dim arcData As UFEval.Arc = Nothing
        Dim arcDataEllipse As UFEval.Ellipse = Nothing
        Dim dirCosX As Double = 0.0
        Dim dirCosY As Double = 0.0
        Dim dirCosZ As Double = 0.0

        Dim diffX As Double = 0.0
        Dim diffY As Double = 0.0
        Dim diffZ As Double = 0.0
        Dim bResult As Boolean = False
        Dim inside As Integer = 0

        'Create a mid point on the edge at 50% of the edge length along the arc
        Dim midPoint As NXOpen.Point = FnCreateEdgeMidpoint(objPart, objEdge)
        'Echo("midPoint : " & midPoint.Coordinates.X.ToString() & "," & midPoint.Coordinates.Y.ToString() & "," & midPoint.Coordinates.Z.ToString())

        If objEdge.SolidEdgeType = Edge.EdgeType.Circular Then
            'Get the circular edge center
            arcData = FnGetEdgeData(objEdge.Tag)
            'Compute direction cosines of vector joining Edge Center to "mid-point"
            diffX = midPoint.Coordinates.X - arcData.center(0)
            diffY = midPoint.Coordinates.Y - arcData.center(1)
            diffZ = midPoint.Coordinates.Z - arcData.center(2)
        ElseIf objEdge.SolidEdgeType = Edge.EdgeType.Elliptical Then
            arcDataEllipse = FnGetEdgeCenterForEllipse(objEdge.Tag)
            'Compute direction cosines of vector joining Edge Center to "mid-point"
            diffX = midPoint.Coordinates.X - arcDataEllipse.center(0)
            diffY = midPoint.Coordinates.Y - arcDataEllipse.center(1)
            diffZ = midPoint.Coordinates.Z - arcDataEllipse.center(2)
        End If

        dirCosX = diffX / Math.Sqrt(Math.Pow(diffX, 2) + Math.Pow(diffY, 2) + Math.Pow(diffZ, 2))
        dirCosY = diffY / Math.Sqrt(Math.Pow(diffX, 2) + Math.Pow(diffY, 2) + Math.Pow(diffZ, 2))
        dirCosZ = diffZ / Math.Sqrt(Math.Pow(diffX, 2) + Math.Pow(diffY, 2) + Math.Pow(diffZ, 2))

        'Create another point P such that
        'Coord (P) = Coord(M) + 0.01*(Direction cosines of CM)
        Dim NewPointP As Point3d = New Point3d(midPoint.Coordinates.X + (0.01 * dirCosX),
                                               midPoint.Coordinates.Y + (0.01 * dirCosY),
                                               midPoint.Coordinates.Z + (0.01 * dirCosZ))
        Dim objCreatedPointP As NXOpen.Point = objPart.Points().CreatePoint(NewPointP)

        'If point P is contained on the Body (as ascertained by "Contained" function used in API)
        'then the given Circular Edge is "Concave"


        FnGetUFSession.Modl.AskPointContainment({objCreatedPointP.Coordinates.X, objCreatedPointP.Coordinates.Y,
                                               objCreatedPointP.Coordinates.Z}, objEdge.GetBody().Tag, inside)

        'Edge is Concave
        If inside = 3 Or inside = 1 Then
            bResult = False
        End If

        'Now, Compute direction cosines of vector joining "mid-point" to Edge Center
        'Direction cosines of MC = (-1)*Direction cosines of CM
        'Create another point Q such that
        'Coord (Q) = Coord(M) + 0.01*(Direction cosines of MC)
        Dim NewPointQ As Point3d = New Point3d(midPoint.Coordinates.X + (0.01 * -dirCosX),
                                               midPoint.Coordinates.Y + (0.01 * -dirCosY),
                                               midPoint.Coordinates.Z + (0.01 * -dirCosZ))

        Dim objCreatedPointQ As NXOpen.Point = objPart.Points().CreatePoint(NewPointQ)

        'If point Q is contained on the Body, then the given Circular Edge is "Convex"
        FnGetUFSession.Modl.AskPointContainment({objCreatedPointQ.Coordinates.X, objCreatedPointQ.Coordinates.Y,
                                               objCreatedPointQ.Coordinates.Z}, objEdge.GetBody().Tag, inside)

        'Edge is Convex
        If inside = 3 Or inside = 1 Then
            bResult = True
        End If

        FnIsEdgeConvex = bResult
    End Function
    'To create a point at 50 % of the arc length
    Public Function FnCreateEdgeMidpoint(ByVal objPart As Part, ByVal theEdge As Edge) As NXOpen.Point

        Dim halfway As Scalar = objPart.Scalars().CreateScalar(50.0, Scalar.DimensionalityType.None, SmartObject.UpdateOption.WithinModeling)
        Dim objCreatedPoint As NXOpen.Point = objPart.Points().CreatePoint(theEdge, halfway, PointCollection.PointOnCurveLocationOption.PercentArcLength,
                                                           SmartObject.UpdateOption.WithinModeling)
        FnCreateEdgeMidpoint = objCreatedPoint
    End Function
    'To Write the matrix data corresponding to all the modelling views in NX Part
    'Public Sub sWriteModelingViewDirCosines()
    '    Dim iColModelViewName As Integer = 0
    '    Dim iColXxc As Integer = 0
    '    Dim iColXyc As Integer = 0
    '    Dim iColXzc As Integer = 0
    '    Dim iColYxc As Integer = 0
    '    Dim iColYyc As Integer = 0
    '    Dim iColYzc As Integer = 0
    '    Dim iColZxc As Integer = 0
    '    Dim iColZyc As Integer = 0
    '    Dim iColZzc As Integer = 0
    '    Dim iRowStart As Integer = 0

    '    iColModelViewName = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, MODEL_VIEW_NAME, TITLE_ROW_NUM_VIEW_DIR_COS)
    '    iColXxc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Xxc, TITLE_ROW_NUM_VIEW_DIR_COS)
    '    iColXyc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Xyc, TITLE_ROW_NUM_VIEW_DIR_COS)
    '    iColXzc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Xzc, TITLE_ROW_NUM_VIEW_DIR_COS)
    '    iColYxc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Yxc, TITLE_ROW_NUM_VIEW_DIR_COS)
    '    iColYyc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Yyc, TITLE_ROW_NUM_VIEW_DIR_COS)
    '    iColYzc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Yzc, TITLE_ROW_NUM_VIEW_DIR_COS)
    '    iColZxc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Zxc, TITLE_ROW_NUM_VIEW_DIR_COS)
    '    iColZyc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Zyc, TITLE_ROW_NUM_VIEW_DIR_COS)
    '    iColZzc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Zzc, TITLE_ROW_NUM_VIEW_DIR_COS)
    '    iRowStart = TITLE_ROW_NUM_VIEW_DIR_COS + 1
    '    For Each objModelView As ModelingView In FnGetNxSession.Parts.Work.ModelingViews
    '        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColModelViewName, objModelView.Name)
    '        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColXxc, Round(objModelView.Matrix.Xx, 4).ToString)
    '        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColXyc, Round(objModelView.Matrix.Xy, 4).ToString)
    '        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColXzc, Round(objModelView.Matrix.Xz, 4).ToString)
    '        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColYxc, Round(objModelView.Matrix.Yx, 4).ToString)
    '        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColYyc, Round(objModelView.Matrix.Yy, 4).ToString)
    '        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColYzc, Round(objModelView.Matrix.Yz, 4).ToString)
    '        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColZxc, Round(objModelView.Matrix.Zx, 4).ToString)
    '        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColZyc, Round(objModelView.Matrix.Zy, 4).ToString)
    '        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColZzc, Round(objModelView.Matrix.Zz, 4).ToString)
    '        iRowStart = iRowStart + 1
    '    Next
    'End Sub

    'Public Sub FnLoadPartFully()
    '    For Each objPart As Part In FnGetPartCollectioninSession()
    '        Dim objPartLoadStatus As PartLoadStatus
    '        objPartLoadStatus = objPart.LoadThisPartFully()
    '        objPartLoadStatus.Dispose()
    '    Next
    'End Sub
    'To Change Face and Edge Names
    Public Sub sChangeObjectNames(ByVal objPart As Part, ByVal obj As NXObject, ByVal sObjectName As String)
        Dim objects2(0) As NXObject
        objects2(0) = obj
        Dim objectGeneralPropertiesBuilder1 As ObjectGeneralPropertiesBuilder = Nothing
        objectGeneralPropertiesBuilder1 = objPart.PropertiesManager.CreateObjectGeneralPropertiesBuilder(objects2)
        objectGeneralPropertiesBuilder1.Name = sObjectName
        Dim nXObject2 As NXObject
        nXObject2 = objectGeneralPropertiesBuilder1.Commit()
        objectGeneralPropertiesBuilder1.Destroy()
    End Sub
    'To check if the body is a welded wire mesh
    Public Function FnChkIfBodyIsMesh(ByVal objPart As Part, ByVal objBody As Body) As Boolean
        Dim sPartName As String = ""
        Dim iCountEdge As Integer = 0

        sPartName = objPart.Leaf.ToString()
        'Only do this check for weldments
        'components will never have mesh bodies
        'Logic added Feb-20-2019
        'For a Mesh Body, there should be atleast two circulat, convex edge which are less than 1.5mm.
        'This change was made since, there was a Body which had one circular edge of 0.7mm (Ref Part Mb749531L in tool Maa76606s.f01.0010 Valiant tool)
        'If FnChkPartisWeldment(objPart) Then
        'If FnCheckIfThisIsAWeldment(sPartName) Then
        If FnCheckIfPartIsWeldmentBasedOnOEM(objPart, _sOemName) Then
            If Not objBody Is Nothing Then
                For Each objFace As Face In objBody.GetFaces()
                    iCountEdge = 0
                    For Each objEdge As Edge In objFace.GetEdges()
                        If objEdge.SolidEdgeType = Edge.EdgeType.Circular Then
                            If (FnGetEdgeData(objEdge.Tag).radius * 2) <= 1.5 Then
                                If FnIsEdgeConvex(objPart, objEdge) Then
                                    iCountEdge = iCountEdge + 1
                                End If
                            End If
                        End If
                    Next
                    If iCountEdge >= 2 Then
                        FnChkIfBodyIsMesh = True
                        Exit Function
                    End If
                Next
            End If
        Else
            FnChkIfBodyIsMesh = False
            Exit Function
        End If
        FnChkIfBodyIsMesh = False
    End Function
    'Create a Error Log file text format
    Public Sub sCreateErrorLogFile(ByVal objPart As Part, ByVal ex As Exception)
        Dim sFileFullPath As String = ""
        Dim sErrorFolderPath As String = ""
        'Output to the designer log
        Dim sDesignerLogFilePath As String = ""
        sErrorFolderPath = Path.Combine(_sSweepDataOutputFolderPath, UNPROCESSED_PARTS, _sToolFolderName)
        'Code modified - Amitabh - 11/14/16 - Create the file inside the log folder
        sDesignerLogFilePath = Path.Combine(_sSweepDataOutputFolderPath, LOG_FOLDER, _sToolFolderName, DESIGNER_LOG_FILE)
        SWrite(objPart.Leaf.ToUpper & " part sweep data file not processed", sDesignerLogFilePath)
        SCreateDirectory(sErrorFolderPath)
        sFileFullPath = Path.Combine(sErrorFolderPath, ERROR_LOG_FILE_PREFIX & objPart.Leaf & ".txt")
        If FnCheckFileExists(sFileFullPath) Then
            Try
                SDeleteFile(sFileFullPath)
                'If the file gets deleted , write the header information freshly
                SWrite(" ", sFileFullPath)
                SWrite("Machine Name : " & Environment.MachineName, sFileFullPath)
                SWrite(DateTime.Now.ToString, sFileFullPath)
                SWrite("Module Name : " & System.Reflection.Assembly.GetExecutingAssembly.GetName().Name, sFileFullPath)
                SWrite("3D Model Part Name : " & objPart.Leaf.ToString, sFileFullPath)
                SWrite(" ", sFileFullPath)
                SWrite(" ", sFileFullPath)
                SWrite(ex.Source, sFileFullPath)
                SWrite(ex.Message, sFileFullPath)
                Dim st As StackTrace = New StackTrace(ex, True)
                SWrite(ex.StackTrace.ToString, sFileFullPath)
            Catch excep As Exception
                'Only write the error information as the header already exists
                SWrite(" ", sFileFullPath)
                SWrite(" ", sFileFullPath)
                SWrite(ex.Source, sFileFullPath)
                SWrite(ex.Message, sFileFullPath)
                Dim st As StackTrace = New StackTrace(ex, True)
                SWrite(ex.StackTrace.ToString, sFileFullPath)
            End Try
        Else
            'Populate the complete set of data for the first time
            SWrite(" ", sFileFullPath)
            SWrite("Machine Name : " & Environment.MachineName, sFileFullPath)
            SWrite(DateTime.Now.ToString, sFileFullPath)
            SWrite("Module Name : " & System.Reflection.Assembly.GetExecutingAssembly.GetName().Name, sFileFullPath)
            SWrite("3D Model Part Name : " & objPart.Leaf.ToString, sFileFullPath)
            SWrite(" ", sFileFullPath)
            SWrite(" ", sFileFullPath)
            SWrite(ex.Source, sFileFullPath)
            SWrite(ex.Message, sFileFullPath)
            Dim st As StackTrace = New StackTrace(ex, True)
            SWrite(ex.StackTrace.ToString, sFileFullPath)
        End If


        'Send the email to the vectra team with the attached error log file
        'sSendEmail(ERROR_LOG_FILE_PREFIX & objPart.Leaf, sFileFullPath)
    End Sub
    'Compose and send the E-mail
    Public Sub sSendEmail(ByVal sEmailSub As String, ByVal sErrorFilePath As String)
        Try
            Dim Smtp_Server As New SmtpClient
            Dim objAttachedFile As Attachment
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("vectra.err.reporting@gmail.com", "vectraglobal")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"
            'Smtp_Server.Host = "smtp.bizmail.yahoo.com"
            e_mail = New MailMessage()
            'Fill the data to send the E-mail
            e_mail.From = New MailAddress("vectra.err.reporting@gmail.com")
            e_mail.To.Add("ram@vectraglobal.com")
            e_mail.CC.Add("amitabh@vectraglobal.com")
            e_mail.Subject = sEmailSub
            e_mail.IsBodyHtml = False
            'Single attachment
            objAttachedFile = New Attachment(sErrorFilePath)
            e_mail.Attachments.Add(objAttachedFile)
            e_mail.Body = "Team," & vbNewLine & "Error report is attached." & vbNewLine & "Regards," & vbNewLine & "VECTRA automated error reporting system"
            Smtp_Server.Send(e_mail)
        Catch error_t As Exception
        End Try

    End Sub
    'Set Program Status
    Public Sub sSetStatus(sMsg As String)
        FnGetUFSession.Ui.SetStatus(sMsg)
    End Sub

    'Read the config file
    'Function FnReadConfigFile() As String()
    '    Dim sFilePath As String = FOLDER_PATH & CONFIG_FILE_NAME
    '    Dim asConfigFileData() As String = Nothing
    '    If FnCheckFileExists(sFilePath) Then
    '        asConfigFileData = FnReadFullFile(sFilePath)
    '    End If
    '    FnReadConfigFile = asConfigFileData
    'End Function

    'Public Sub sSetToolFolderPath(sPath As String)
    '    _sToolFolderPath = sPath
    'End Sub
    'Store all the solid bodies present in the part
    Sub sStoreSolidBody(objBody As Body)
        If _aoSolidBody Is Nothing Then
            ReDim Preserve _aoSolidBody(0)
            _aoSolidBody(0) = objBody
        Else
            ReDim Preserve _aoSolidBody(UBound(_aoSolidBody) + 1)
            _aoSolidBody(UBound(_aoSolidBody)) = objBody
        End If
    End Sub
    'Determine and write the mating solid bodies and the mating faces in the sweep data file
    'Sub sDetermineMatingBodiesandFaces(objWorkPart As Part)
    '    Dim sParentBodyName As String = ""
    '    Dim sChildBodyName As String = ""
    '    Dim iStartRowWrite As Integer = 2
    '    'To keep a count of number of mating faces between the mating solid bodies
    '    Dim iCountOfMatingFacesBetweenBodies As Integer = 1
    '    Dim aoMatingFacesinBodies() As Face = Nothing
    '    Dim objOwningComp As Component = Nothing
    '    Dim dTolerance As Double = 0.0
    '    Dim bToleranceRelaxed As Boolean = False
    '    Dim dDistanceBetweenMatingBody As Double = 0.0

    '    sSetStatus("Collecting mating body data...")
    '    If Not _aoSolidBody Is Nothing Then
    '        dTolerance = BODY_TO_BODY_MATING_TOLERANCE
    '        For Each objParentBody As Body In _aoSolidBody
    '            objOwningComp = objParentBody.OwningComponent
    '            If Not objOwningComp Is Nothing Then
    '                If FnChkIfMatingToleranceNeedsToBeRelaxed(objOwningComp) Then
    '                    dTolerance = BODY_TO_BODY_MATING_TOLERANCE_RELAXED
    '                    bToleranceRelaxed = True
    '                    Exit For
    '                End If
    '            End If
    '        Next

    '        For Each objParentBody As Body In _aoSolidBody
    '            If objParentBody.IsOccurrence Then
    '                sParentBodyName = objParentBody.JournalIdentifier & " " & objParentBody.OwningComponent.JournalIdentifier
    '            Else
    '                sParentBodyName = objParentBody.JournalIdentifier
    '            End If
    '            For Each objChildBody As Body In _aoSolidBody
    '                If Not objChildBody Is objParentBody Then
    '                    If objChildBody.IsOccurrence Then
    '                        sChildBodyName = objChildBody.JournalIdentifier & " " & objChildBody.OwningComponent.JournalIdentifier
    '                    Else
    '                        sChildBodyName = objChildBody.JournalIdentifier
    '                    End If
    '                    'Compute the distance between mating bodies
    '                    dDistanceBetweenMatingBody = Round(FnComputeMinDistance(objWorkPart, objParentBody, objChildBody), 1)
    '                    If dDistanceBetweenMatingBody <= dTolerance Then
    '                        'Determine the mating faces between these bodies
    '                        If bToleranceRelaxed Then
    '                            aoMatingFacesinBodies = FnGetMatingChildFacesBetweenBodies(objWorkPart, objParentBody, objChildBody, dDistanceBetweenMatingBody, _
    '                                                                                        bExceptionComps:=True)
    '                        Else
    '                            aoMatingFacesinBodies = FnGetMatingChildFacesBetweenBodies(objWorkPart, objParentBody, objChildBody, dDistanceBetweenMatingBody, _
    '                                                                                       bExceptionComps:=False)
    '                        End If
    '                        If Not aoMatingFacesinBodies Is Nothing Then
    '                            'CODE CHANGED - 4/13/16 - Amitabh - Write mating bodies information only if mating faces are found for the 2 mating bodies.
    '                            SWriteValueToCell(_objWorkBk, MATING_BODY_SHEET_NAME, iStartRowWrite, 1, sParentBodyName)
    '                            'The solid bodies are mating with each other
    '                            SWriteValueToCell(_objWorkBk, MATING_BODY_SHEET_NAME, iStartRowWrite, 2, sChildBodyName)
    '                            iCountOfMatingFacesBetweenBodies = 1
    '                            For iLoopIndex = 0 To UBound(aoMatingFacesinBodies) Step 2
    '                                If iCountOfMatingFacesBetweenBodies = 1 Then
    '                                    SWriteValueToCell(_objWorkBk, MATING_BODY_FACE_INFO_TAB_1_SHEET_NAME, iStartRowWrite, 1, aoMatingFacesinBodies(iLoopIndex).Name)
    '                                    SWriteValueToCell(_objWorkBk, MATING_BODY_FACE_INFO_TAB_1_SHEET_NAME, iStartRowWrite, 2, aoMatingFacesinBodies(iLoopIndex + 1).Name)
    '                                ElseIf iCountOfMatingFacesBetweenBodies = 2 Then
    '                                    SWriteValueToCell(_objWorkBk, MATING_BODY_FACE_INFO_TAB_2_SHEET_NAME, iStartRowWrite, 1, aoMatingFacesinBodies(iLoopIndex).Name)
    '                                    SWriteValueToCell(_objWorkBk, MATING_BODY_FACE_INFO_TAB_2_SHEET_NAME, iStartRowWrite, 2, aoMatingFacesinBodies(iLoopIndex + 1).Name)
    '                                ElseIf iCountOfMatingFacesBetweenBodies = 3 Then
    '                                    SWriteValueToCell(_objWorkBk, MATING_BODY_FACE_INFO_TAB_3_SHEET_NAME, iStartRowWrite, 1, aoMatingFacesinBodies(iLoopIndex).Name)
    '                                    SWriteValueToCell(_objWorkBk, MATING_BODY_FACE_INFO_TAB_3_SHEET_NAME, iStartRowWrite, 2, aoMatingFacesinBodies(iLoopIndex + 1).Name)
    '                                End If
    '                                iCountOfMatingFacesBetweenBodies = iCountOfMatingFacesBetweenBodies + 1
    '                            Next
    '                            iStartRowWrite = iStartRowWrite + 1
    '                        End If
    '                    End If
    '                End If
    '            Next
    '        Next
    '    End If
    'End Sub

    'Code Added - Amitabh - 9/22/16 - New format of reporting
    'Determine and write the mating solid bodies and the mating faces in the sweep data file
    Sub sDetermineMatingBodiesandFaces(objWorkPart As Part)
        Dim sParentBodyName As String = ""
        Dim sChildBodyName As String = ""
        Dim iStartRowWrite As Integer = 2
        'To keep a count of number of mating faces between the mating solid bodies
        'Dim iCountOfMatingFacesBetweenBodies As Integer = 1
        Dim aoMatingFacesinBodies() As Face = Nothing
        Dim objOwningComp As Component = Nothing
        Dim dTolerance As Double = 0.0
        Dim bToleranceRelaxed As Boolean = False
        Dim dDistanceBetweenMatingBody As Double = 0.0
        'Start column for writing mating faces
        Dim iColumnWriteMatingFaces As Integer = 3
        'Dim aoMatingBodyList() As String = Nothing
        'Dim bAnalyzeMatingFaces As Boolean = False

        sSetStatus("Collecting mating body data...")
        If Not _aoSolidBody Is Nothing Then
            dTolerance = BODY_TO_BODY_MATING_TOLERANCE
            For Each objParentBody As Body In _aoSolidBody
                objOwningComp = objParentBody.OwningComponent
                If Not objOwningComp Is Nothing Then
                    If FnChkIfMatingToleranceNeedsToBeRelaxed(objOwningComp) Then
                        dTolerance = BODY_TO_BODY_MATING_TOLERANCE_RELAXED
                        bToleranceRelaxed = True
                        Exit For
                    End If
                End If
            Next

            For Each objParentBody As Body In _aoSolidBody
                If objParentBody.IsOccurrence Then
                    If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                        sParentBodyName = objParentBody.JournalIdentifier & " " & objParentBody.OwningComponent.JournalIdentifier
                    ElseIf (_sOemName = DAIMLER_OEM_NAME) Then
                        If _sDivision = TRUCK_DIVISION Then
                            sParentBodyName = objParentBody.JournalIdentifier & " " & objParentBody.OwningComponent.JournalIdentifier
                        ElseIf _sDivision = CAR_DIVISION Then
                            sParentBodyName = objParentBody.JournalIdentifier & " " & objParentBody.OwningComponent.Parent.JournalIdentifier
                        End If
                    ElseIf (_sOemName = FIAT_OEM_NAME) Then
                        sParentBodyName = objParentBody.JournalIdentifier & " " & objParentBody.OwningComponent.JournalIdentifier
                    End If
                Else
                    sParentBodyName = objParentBody.JournalIdentifier
                End If

                'bAnalyzeMatingFaces = False
                For Each objChildBody As Body In _aoSolidBody
                    If Not objChildBody Is objParentBody Then
                        If objChildBody.IsOccurrence Then
                            If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                sChildBodyName = objChildBody.JournalIdentifier & " " & objChildBody.OwningComponent.JournalIdentifier
                            ElseIf (_sOemName = DAIMLER_OEM_NAME) Then
                                If _sDivision = TRUCK_DIVISION Then
                                    sChildBodyName = objChildBody.JournalIdentifier & " " & objChildBody.OwningComponent.JournalIdentifier
                                ElseIf _sDivision = CAR_DIVISION Then
                                    sChildBodyName = objChildBody.JournalIdentifier & " " & objChildBody.OwningComponent.Parent.JournalIdentifier
                                End If
                            ElseIf (_sOemName = FIAT_OEM_NAME) Then
                                sChildBodyName = objChildBody.JournalIdentifier & " " & objChildBody.OwningComponent.JournalIdentifier
                            End If
                        Else
                            sChildBodyName = objChildBody.JournalIdentifier
                        End If

                        'Prevent Vice-versa comparison
                        'If aoMatingBodyList Is Nothing Then
                        '    bAnalyzeMatingFaces = True
                        'Else
                        '    If Not aoMatingBodyList.Contains(sParentBodyName.ToUpper & " | " & sChildBodyName.ToUpper) Then
                        '        bAnalyzeMatingFaces = True
                        '    End If
                        'End If

                        'If bAnalyzeMatingFaces Then
                        'Compute the distance between mating bodies
                        dDistanceBetweenMatingBody = Round(FnComputeMinDistance(objWorkPart, objParentBody, objChildBody), 1)
                        If dDistanceBetweenMatingBody <= dTolerance Then
                            'Determine the mating faces between these bodies
                            If bToleranceRelaxed Then
                                aoMatingFacesinBodies = FnGetMatingChildFacesBetweenBodies(objWorkPart, objParentBody, objChildBody, dTolerance,
                                                                                            bExceptionComps:=True)
                            Else
                                aoMatingFacesinBodies = FnGetMatingChildFacesBetweenBodies(objWorkPart, objParentBody, objChildBody, dTolerance,
                                                                                           bExceptionComps:=False)
                            End If

                            If Not aoMatingFacesinBodies Is Nothing Then
                                SWriteValueToCell(_objWorkBk, MATING_BODY_SHEET_NAME, iStartRowWrite, 1, sParentBodyName)
                                'The solid bodies are mating with each other
                                SWriteValueToCell(_objWorkBk, MATING_BODY_SHEET_NAME, iStartRowWrite, 2, sChildBodyName)

                                '******************* CODE COMMENTED ON 11OCT16 **************************************************************
                                'NO LONGER REQUIRED AS IT IMPACTS CORE ALG ANALYSIS
                                'Update the data in the mating body list
                                'This is to prevent vice-versa comparison
                                'If aoMatingBodyList Is Nothing Then
                                '    ReDim Preserve aoMatingBodyList(1)
                                '    aoMatingBodyList(0) = sParentBodyName.ToUpper & " | " & sChildBodyName.ToUpper
                                '    aoMatingBodyList(1) = sChildBodyName.ToUpper & " | " & sParentBodyName.ToUpper
                                'Else
                                '    ReDim Preserve aoMatingBodyList(UBound(aoMatingBodyList) + 2)
                                '    aoMatingBodyList(UBound(aoMatingBodyList) - 1) = sParentBodyName.ToUpper & " | " & sChildBodyName.ToUpper
                                '    aoMatingBodyList(UBound(aoMatingBodyList)) = sChildBodyName.ToUpper & " | " & sParentBodyName.ToUpper
                                'End If
                                '*************************************************************************************************************

                                iColumnWriteMatingFaces = 3
                                For iLoopIndex = 0 To UBound(aoMatingFacesinBodies) Step 2
                                    'Write the Header
                                    SWriteValueToCell(_objWorkBk, MATING_BODY_SHEET_NAME, 1, iColumnWriteMatingFaces, MATING_BODY_FACES_PARENT_HEADER)
                                    SWriteValueToCell(_objWorkBk, MATING_BODY_SHEET_NAME, 1, iColumnWriteMatingFaces + 1, MATING_BODY_FACES_CHILD_HEADER)
                                    'Write the Face Names
                                    SWriteValueToCell(_objWorkBk, MATING_BODY_SHEET_NAME, iStartRowWrite, iColumnWriteMatingFaces, aoMatingFacesinBodies(iLoopIndex).Name)
                                    SWriteValueToCell(_objWorkBk, MATING_BODY_SHEET_NAME, iStartRowWrite, iColumnWriteMatingFaces + 1, aoMatingFacesinBodies(iLoopIndex + 1).Name)
                                    iColumnWriteMatingFaces = iColumnWriteMatingFaces + 2
                                Next
                                iStartRowWrite = iStartRowWrite + 1
                            End If
                        End If
                        'End If
                    End If
                Next
            Next
        End If
    End Sub
    'Collection of combination of parent face and mating child face
    Function FnGetMatingChildFacesBetweenBodies(objWorkPart As Part, objParentBody As Body, objChildBody As Body, dMatingBodyDistance As Double,
                                                Optional bExceptionComps As Boolean = False) As Face()
        Dim aoMatingFaces() As Face = Nothing
        For Each objParentFace As Face In objParentBody.GetFaces()
            For Each objChildFace As Face In objChildBody.GetFaces()
                If Not bExceptionComps Then
                    If objParentFace.SolidFaceType = Face.FaceType.Planar And objChildFace.SolidFaceType = Face.FaceType.Planar Then
                        If FnIsFaceMating(objWorkPart, objParentFace, objChildFace, dMatingBodyDistance, PLANAR_FACE) Then
                            sUpdateParentChildMatingFacesInBodies(objParentFace, objChildFace, aoMatingFaces)
                        End If
                    ElseIf objParentFace.SolidFaceType = Face.FaceType.Cylindrical And objChildFace.SolidFaceType = Face.FaceType.Cylindrical Then
                        If FnIsFaceMating(objWorkPart, objParentFace, objChildFace, dMatingBodyDistance, CYLINDRICAL_FACE) Then
                            sUpdateParentChildMatingFacesInBodies(objParentFace, objChildFace, aoMatingFaces)
                        End If
                        'CODE COMMENTED - 4/13/16 - Amitabh - No longer checking for other types of faces
                        'ElseIf Not (objParentFace.SolidFaceType = Face.FaceType.Planar Or objParentFace.SolidFaceType = Face.FaceType.Cylindrical) And _
                        '       Not (objChildFace.SolidFaceType = Face.FaceType.Planar Or objChildFace.SolidFaceType = Face.FaceType.Cylindrical) Then
                        'Else
                        'If FnIsFaceMating(objWorkPart, objParentFace, objChildFace, dMatingBodyDistance, OTHER_TYPES_FACE) Then
                        '    sUpdateParentChildMatingFacesInBodies(objParentFace, objChildFace, aoMatingFaces)
                        'End If
                    End If
                Else
                    'For exception components the mating condition should be checked between any types of faces
                    If FnIsFaceMating(objWorkPart, objParentFace, objChildFace, dMatingBodyDistance, OTHER_TYPES_FACE) Then
                        sUpdateParentChildMatingFacesInBodies(objParentFace, objChildFace, aoMatingFaces)
                    End If
                End If
            Next
        Next
        FnGetMatingChildFacesBetweenBodies = aoMatingFaces
    End Function
    'Update the array containing the mating parent and child faces
    Sub sUpdateParentChildMatingFacesInBodies(objParFace As Face, objChildFace As Face, ByRef aoMatingFaces() As Face)
        If aoMatingFaces Is Nothing Then
            ReDim Preserve aoMatingFaces(1)
            aoMatingFaces(0) = objParFace
            aoMatingFaces(1) = objChildFace
        Else
            ReDim Preserve aoMatingFaces(UBound(aoMatingFaces) + 2)
            aoMatingFaces(UBound(aoMatingFaces) - 1) = objParFace
            aoMatingFaces(UBound(aoMatingFaces)) = objChildFace
        End If
    End Sub

    Function FnChkIfMatingToleranceNeedsToBeRelaxed(objComp As Component) As Boolean
        Dim sPartName As String = ""
        sPartName = FnGetStringUserAttribute(objComp, _PART_NAME)
        If Not sPartName = "" Then
            For Each sName As String In _asExceptionCompsForMatingBodyToleranceRelaxation
                If sPartName.Contains(sName.ToUpper) Then
                    FnChkIfMatingToleranceNeedsToBeRelaxed = True
                    Exit Function
                End If
            Next
        Else
            FnChkIfMatingToleranceNeedsToBeRelaxed = False
            Exit Function
        End If
    End Function


    'Get objects by attribute
    'Currently only coded for fetching string attributes
    'Passing empty string as value selects all the objects having the attribute
    Public Function FnGetFaceObjectByAttributes(objPart As Part, sAttrTitle As String, sAttrValue As String,
                                            Optional ByVal sAttrType As String = "String") As DisplayableObject()
        Dim arule As String = ""
        If sAttrType = "String" Then
            arule = "mqc_selectEntitiesWithFilters(" &
                "select_by_entity_type, { FACE }, " &
                "select_by_attribute, {{ String, """ & sAttrTitle & """, """ & sAttrValue & """, """ & sAttrValue & """ }}, " &
                "ignore_entity_occurrence?, False)"
        End If
        Dim ruleName As String = "VECTRA_Rule"
        objPart.RuleManager.CreateDynamicRule("root:", ruleName, "Any", arule, "")
        Dim theObj As Object = objPart.RuleManager.Evaluate("root:" & ruleName & ":")
        objPart.RuleManager.DeleteDynamicRule("root:", ruleName)

        If Not theObj Is Nothing Then
            Dim found() As DisplayableObject = ConvertTagListToDisplayableObject(theObj)
            FnGetFaceObjectByAttributes = found
        Else
            FnGetFaceObjectByAttributes = Nothing
        End If
    End Function
    Function ConvertTagListToDisplayableObject(ByRef tags() As Object) As DisplayableObject()

        Dim theList As ArrayList = New ArrayList
        Dim anObj As DisplayableObject

        For Each aTag As Tag In tags
            anObj = NXObjectManager.Get(aTag)
            If Not anObj Is Nothing Then theList.Add(anObj)
        Next
        Return theList.ToArray(GetType(DisplayableObject))
    End Function
    'To Write the matrix data corresponding to all the modelling views in NX Part
    Public Sub sWriteModelingViewDirCosines(ByVal objPart)
        Dim iColModelViewName As Integer = 0
        Dim iColXxc As Integer = 0
        Dim iColXyc As Integer = 0
        Dim iColXzc As Integer = 0
        Dim iColYxc As Integer = 0
        Dim iColYyc As Integer = 0
        Dim iColYzc As Integer = 0
        Dim iColZxc As Integer = 0
        Dim iColZyc As Integer = 0
        Dim iColZzc As Integer = 0
        Dim iRowStart As Integer = 0

        iColModelViewName = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, MODEL_VIEW_NAME, TITLE_ROW_NUM_VIEW_DIR_COS)
        iColXxc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Xxc, TITLE_ROW_NUM_VIEW_DIR_COS)
        iColXyc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Xyc, TITLE_ROW_NUM_VIEW_DIR_COS)
        iColXzc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Xzc, TITLE_ROW_NUM_VIEW_DIR_COS)
        iColYxc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Yxc, TITLE_ROW_NUM_VIEW_DIR_COS)
        iColYyc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Yyc, TITLE_ROW_NUM_VIEW_DIR_COS)
        iColYzc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Yzc, TITLE_ROW_NUM_VIEW_DIR_COS)
        iColZxc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Zxc, TITLE_ROW_NUM_VIEW_DIR_COS)
        iColZyc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Zyc, TITLE_ROW_NUM_VIEW_DIR_COS)
        iColZzc = FnGetColumnNumberByName(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, Zzc, TITLE_ROW_NUM_VIEW_DIR_COS)
        'iRowStart = TITLE_ROW_NUM_VIEW_DIR_COS + 1
        iRowStart = FnGetNumberofRows(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, 1, 1)
        iRowStart = iRowStart + 1
        For Each objModelView As ModelingView In objPart.ModelingViews
            '3/24/16- CODE CHANGED - Do not include views with names like TOP#1 FRONT#1 Right#1
            '4/16/16 - CODE CHANGED - Only output Primary view information
            'If (Not objModelView.Name.ToUpper.Contains("DIMETRIC")) And (Not objModelView.Name.ToUpper.Contains("#")) Then
            If (objModelView.Name.ToUpper.Contains(PRIMARY_VIEW_NAME) Or objModelView.Name.ToUpper.Contains(SECOND_PRIMARY_VIEW_NAME)) Then
                'Re-Orient the view by making it the work view before retrieving the matrix
                sReplaceViewInLayout(objPart, objModelView)
                SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColModelViewName, objModelView.Name)
                SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColXxc, objModelView.Matrix.Xx.ToString)
                SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColXyc, objModelView.Matrix.Xy.ToString)
                SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColXzc, objModelView.Matrix.Xz.ToString)
                SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColYxc, objModelView.Matrix.Yx.ToString)
                SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColYyc, objModelView.Matrix.Yy.ToString)
                SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColYzc, objModelView.Matrix.Yz.ToString)
                SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColZxc, objModelView.Matrix.Zx.ToString)
                SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColZyc, objModelView.Matrix.Zy.ToString)
                SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowStart, iColZzc, objModelView.Matrix.Zz.ToString)
                iRowStart = iRowStart + 1
            End If
        Next
    End Sub
    'Every model view in NX is in a layout
    'Find the layout in which the current work view is in and then replace it with a new view so that the model view is displayed
    Public Sub sReplaceViewInLayout(ByVal objPart As Part, ByVal objModelViewToReplace As ModelingView)
        Dim viewType As Integer = 0
        '1 = modeling view
        '2 = drawing view
        'other = error
        FnGetUFSession.Draw.AskDisplayState(viewType)

        'if drawing sheet shown, change to modeling view
        If viewType = 2 Then
            FnGetUFSession.Draw.SetDisplayState(1)
        End If
        'Set the layout to "L1" by default
        sSetDefaultLayout(objPart)
        Dim sCurrentWorkView As ModelingView = Nothing
        sCurrentWorkView = objPart.ModelingViews.WorkView
        For Each objLayout As Layout In objPart.Layouts
            For Each objView As ModelingView In objLayout.GetViews()
                If objView.Name = sCurrentWorkView.Name Then
                    '3/24/16 - CODE CHANGED - Display the layout by changing the current layout
                    'View needs to be dispayed for it to be changed as a work view
                    objPart.Layouts.ChangeLayout(objLayout)
                    objLayout.ReplaceView(sCurrentWorkView, objModelViewToReplace, False)
                    'Make this view as the work view
                    objModelViewToReplace.MakeWork()
                    objModelViewToReplace.Restore()
                    Exit Sub
                End If
            Next
        Next
    End Sub


    'To write PreFab hole attributes from Model to the Excel sheet
    Sub sAddPreFabHoleAttributeToExcel(ByVal objPart As Part)
        Dim _iLastRowFilledFaceVec As Integer = 0
        Dim sFaceName As String = ""
        Dim objFaceAttr() As DisplayableObject = Nothing

        If _iLastRowFilledFaceVec = 0 Then
            _iLastRowFilledFaceVec = FnGetNumberofRows(_objWorkBk, "FaceVec", 1, 1)
        End If

        objFaceAttr = FnGetFaceObjectByAttributes(objPart, PRE_FAB_HOLE_ATTR_TITLE, PRE_FAB_HOLE_ATTR_VALUE)
        If Not objFaceAttr Is Nothing Then
            For Each aface As DisplayableObject In objFaceAttr.ToArray()
                For iLoopIndex As Integer = 2 To _iLastRowFilledFaceVec
                    sFaceName = FnReadSingleRowForColumn(_objWorkBk, "FaceVec", 1, iLoopIndex)
                    If sFaceName.ToUpper = CType(aface, Face).Name.ToUpper Then
                        SWriteValueToCell(_objWorkBk, "FaceVec", iLoopIndex, 8, PRE_FAB_HOLE_ATTR_VALUE)
                        Exit For
                    End If
                Next
            Next
        End If
    End Sub

    'Function to CHeck ModelingView Exist
    Public Function FnChkIfModelingViewPresent(ByVal objPart As Part, ByVal sViewName As String) As Boolean
        For Each objModelView As ModelingView In objPart.ModelingViews
            If objModelView.Name.ToUpper = sViewName.ToUpper Then
                FnChkIfModelingViewPresent = True
                Exit Function
            End If
        Next
        FnChkIfModelingViewPresent = False
    End Function
    'To get the bounding box of a body aligned to a given model view
    Sub sWriteBoundingBoxofAllBodiesInAModelView(ByVal objPart As Part, aoAllCompInSession() As Component, ByVal sModelViewName As String, ByVal sSheetNameToWrite As String)
        'For computing the exact view bounds of the bodies
        Dim min_corner(2) As Double
        Dim directions(2, 2) As Double
        Dim distances(2) As Double
        Dim adBoundingBox() As Double = Nothing
        Dim objBody As Body = Nothing
        Dim sBodyName As String = ""
        Dim sBodyNameCompare As String = ""
        Dim iRowToWrite As Integer = 2
        'Dim aCompCol() As Component = FnGetAllComponentsInSession()
        Dim iNosOfFilledRows As Integer = 0
        Dim iColBodyName As Integer = 0

        Dim matrixTag As Tag = NXOpen.Tag.Null
        Dim csysTag As Tag = NXOpen.Tag.Null
        Dim objChildPart As Part = Nothing
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim bIsValidSolidBody As Boolean = False
        Dim aoAllValidBody() As Body = Nothing

        For Each objModelView As ModelingView In objPart.ModelingViews
            If objModelView.Name.ToUpper = sModelViewName.ToUpper Then
                'Make this view as the work view
                sReplaceViewInLayout(objPart, objModelView)
                Dim admatrixValues As Double() = {objModelView.Matrix.Xx, objModelView.Matrix.Xy, objModelView.Matrix.Xz,
                                                  objModelView.Matrix.Yx, objModelView.Matrix.Yy, objModelView.Matrix.Yz,
                                                  objModelView.Matrix.Zx, objModelView.Matrix.Zy, objModelView.Matrix.Zz}
                'Create the CSYS matrix
                UFSession.GetUFSession().Csys.CreateMatrix(admatrixValues, matrixTag)
                UFSession.GetUFSession.Csys.CreateCsys({0, 0, 0}, matrixTag, csysTag)
                Exit For
            End If
        Next

        iColBodyName = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, BODY_NAME, 1)
        iNosOfFilledRows = FnGetNumberofRows(_objWorkBk, sSheetNameToWrite, 1, 1)

        iColMinPointX = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, MIN_POINTX, 1)
        iColMinPointY = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, MIN_POINTY, 1)
        iColMinPointZ = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, MIN_POINTZ, 1)
        iColVectorXX = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORXX, 1)
        iColVectorXY = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORXY, 1)
        iColVectorXZ = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORXZ, 1)
        iColVectorYX = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORYX, 1)
        iColVectorYY = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORYY, 1)
        iColVectorYZ = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORYZ, 1)
        iColVectorZX = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORZX, 1)
        iColVectorZY = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORZY, 1)
        iColVectorZZ = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORZZ, 1)
        iColMagX = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, Magnitude_X, 1)
        iColMagY = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, Magnitude_Y, 1)
        iColMagZ = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, Magnitude_Z, 1)

        If Not aoAllCompInSession Is Nothing Then
            For Each objComp As Component In aoAllCompInSession
                objChildPart = FnGetPartFromComponent(objComp)
                If Not objChildPart Is Nothing Then
                    FnLoadPartFully(objChildPart)
                    aoAllValidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)
                    If Not aoAllValidBody Is Nothing Then
                        For Each body As Body In aoAllValidBody

                            If Not body Is Nothing Then
                                If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        sBodyName = body.JournalIdentifier
                                        objBody = body
                                    Else
                                        sBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                        objBody = CType(objComp.FindOccurrence(body), Body)
                                    End If
                                ElseIf (_sOemName = DAIMLER_OEM_NAME) Then
                                    If _sDivision = TRUCK_DIVISION Then
                                        If objComp Is objPart.ComponentAssembly.RootComponent Then
                                            'Component in truck
                                            sBodyName = body.JournalIdentifier
                                            objBody = body
                                        Else
                                            'Weldment in truck
                                            sBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                            'sBodyName = objComp.DisplayName.ToUpper
                                            objBody = CType(objComp.FindOccurrence(body), Body)
                                        End If
                                    ElseIf _sDivision = CAR_DIVISION Then
                                        'Check if the component is a child component in weldment
                                        If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                            'Weldment in car. Get the occurrence body
                                            sBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier & " " & objComp.Parent.JournalIdentifier
                                            'sBodyName = objComp.DisplayName.ToUpper
                                            objBody = CType(objComp.FindOccurrence(body), Body)
                                        Else
                                            'Code modified on may-14-2018
                                            'Component in car. Get the prototype body
                                            If Not CType(objComp.FindOccurrence(body), Body) Is Nothing Then
                                                sBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier
                                                objBody = CType(objComp.FindOccurrence(body), Body)
                                            End If

                                        End If
                                    End If
                                ElseIf (_sOemName = FIAT_OEM_NAME) Then
                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        'Fiat component
                                        objBody = body
                                        sBodyName = objBody.JournalIdentifier
                                    Else
                                        'This is a Fiat Weldment child component
                                        objBody = CType(objComp.FindOccurrence(body), Body)
                                        If Not objBody Is Nothing Then
                                            'when populating body name, give Body journal identifier and Child component journalidentifier
                                            sBodyName = objBody.JournalIdentifier & " " & objComp.JournalIdentifier
                                        End If
                                    End If
                                    ' ''Check if the component is a child component in weldment
                                    ''If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                    ''    'Weldment in FIAT. Get the occurrence body
                                    ''    sBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                    ''    'sBodyName = objComp.DisplayName.ToUpper
                                    ''    objBody = CType(objComp.FindOccurrence(body), Body)
                                    ''Else
                                    ''    'Component in FIAT
                                    ''    sBodyName = body.JournalIdentifier
                                    ''    objBody = body
                                    ''End If
                                End If


                                FnGetUFSession.Modl.AskBoundingBoxExact(objBody.Tag, csysTag, min_corner, directions, distances)
                                'Find the row to update
                                For iloopIndex = 2 To iNosOfFilledRows
                                    'Fetch the Body Name number already assigned
                                    sBodyNameCompare = FnReadSingleRowForColumn(_objWorkBk, sSheetNameToWrite, iColBodyName, iloopIndex)
                                    If sBodyNameCompare.ToUpper = sBodyName.ToUpper Then
                                        iRowToWrite = iloopIndex
                                        Exit For
                                    End If
                                Next

                                'To get the bounding box exact for each body
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointX, min_corner(0).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointY, min_corner(1).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointZ, min_corner(2).ToString)

                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXX, directions(0, 0).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXY, directions(0, 1).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXZ, directions(0, 2).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagX, distances(0).ToString)

                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYX, directions(1, 0).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYY, directions(1, 1).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYZ, directions(1, 2).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagY, distances(1).ToString)

                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZX, directions(2, 0).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZY, directions(2, 1).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZZ, directions(2, 2).ToString)
                                SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagZ, distances(2).ToString)
                            End If

                        Next
                    End If
                End If
            Next
        Else
            aoAllValidBody = FnGetValidBodyForOEM(objPart, _sOemName)
            If Not aoAllValidBody Is Nothing Then
                For Each body As Body In aoAllValidBody
                    If Not body Is Nothing Then
                        objBody = body
                        sBodyName = body.JournalIdentifier
                        FnGetUFSession.Modl.AskBoundingBoxExact(objBody.Tag, csysTag, min_corner, directions, distances)

                        'Find the row to update
                        For iloopIndex = 2 To iNosOfFilledRows
                            'Fetch the Body Name number already assigned
                            sBodyNameCompare = FnReadSingleRowForColumn(_objWorkBk, sSheetNameToWrite, iColBodyName, iloopIndex)
                            If sBodyNameCompare.ToUpper = sBodyName.ToUpper Then
                                iRowToWrite = iloopIndex
                                Exit For
                            End If
                        Next

                        'To get the bounding box exact for each body
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointX, min_corner(0).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointY, min_corner(1).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointZ, min_corner(2).ToString)

                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXX, directions(0, 0).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXY, directions(0, 1).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXZ, directions(0, 2).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagX, distances(0).ToString)

                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYX, directions(1, 0).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYY, directions(1, 1).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYZ, directions(1, 2).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagY, distances(1).ToString)

                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZX, directions(2, 0).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZY, directions(2, 1).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZZ, directions(2, 2).ToString)
                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagZ, distances(2).ToString)
                    End If

                Next
            End If
        End If


        'Delete the created CSYS
        SDeleteObjects({NXObjectManager.Get(csysTag)})

    End Sub
    Public Sub SDeleteObjects(ByVal objToDelete() As NXObject)
        FnGetNxSession().UpdateManager.ClearErrorList()
        Dim markIdDelete As Session.UndoMarkId
        markIdDelete = FnGetNxSession.SetUndoMark(Session.MarkVisibility.Visible, "Delete")
        FnGetNxSession().UpdateManager.AddToDeleteList(objToDelete)
        Dim nErrs2 As Integer
        nErrs2 = FnGetNxSession.UpdateManager.DoUpdate(markIdDelete)
    End Sub

    'Function to CHeck if part is NC Groups
    Function FnChkIfPartIsNC(sPartName As String) As Boolean
        For Each sName As String In _asNCBLOCKDbPartNames
            If sPartName.ToUpper = sName Then
                FnChkIfPartIsNC = True
                Exit Function
            End If
        Next
        FnChkIfPartIsNC = False
    End Function
    'Get the Part Name of the COmponent
    Function FnGetPartName(objPart As Part) As String
        FnGetPartName = FnGetStringUserAttribute(objPart, _PART_NAME)

    End Function

    'To get the execution folder path
    Function FnGetExecutionFolderPath() As String
        Dim sAppFolderPath As String = ""
        Dim sFilePath As String = ""
        'Getting the path of the folder where the executable resides
        Dim myassembly As System.Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim tempfile As FileInfo = New FileInfo(myassembly.Location)
        sAppFolderPath = tempfile.DirectoryName
        FnGetExecutionFolderPath = sAppFolderPath
    End Function

    'Code Added by Shanmugam on May-10-2016
    Sub sPopulateVisibleEdgeNames(ByVal objPart As Part)
        Dim objBodyToCheck As Body = Nothing

        Dim asVisibleEdgeNamesInPrimary1FrontVw() As String = Nothing
        Dim asVisibleEdgeNamesInPrimary1RightVw() As String = Nothing
        Dim asVisibleEdgeNamesInPrimary1TopVw() As String = Nothing
        Dim asVisibleEdgeNamesInPrimary1RearVw() As String = Nothing
        Dim asVisibleEdgeNamesInPrimary1LeftVw() As String = Nothing
        Dim asVisibleEDgeNamesInPrimary1BottomVw() As String = Nothing
        Dim asVisibleEdgeNamesInPrimary2FrontVw() As String = Nothing
        Dim asVisibleEdgeNamesInPrimary2RightVw() As String = Nothing
        Dim asVisibleEdgeNamesInPrimary2TopVw() As String = Nothing
        Dim asVisibleEdgeNamesInPrimary2RearVw() As String = Nothing
        Dim asVisibleEdgeNamesInPrimary2LeftVw() As String = Nothing
        Dim asVisibleEDgeNamesInPrimary2BottomVw() As String = Nothing

        Dim iRowStartHeader As Integer = 1
        Dim iRowStartVisibleEdge As Integer = 2
        Dim iColEdgeNames As Integer = 0
        Dim iColPrimary1Front As Integer = 0
        Dim iColPrimary1Right As Integer = 0
        Dim iColPrimary1Left As Integer = 0
        Dim iColPrimary1Back As Integer = 0
        Dim iColPrimary1Top As Integer = 0
        Dim iColPrimary1Bottom As Integer = 0
        Dim iColPrimary2Front As Integer = 0
        Dim iColPrimary2Right As Integer = 0
        Dim iColPrimary2Left As Integer = 0
        Dim iColPrimary2Back As Integer = 0
        Dim iColPrimary2Top As Integer = 0
        Dim iColPrimary2Bottom As Integer = 0

        Dim bIsPrimaryView2Present As Boolean = False
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim bIsValidSolidBody As Boolean = False
        Dim aoAllValidBody() As Body = Nothing

        If Not objPart Is Nothing Then
            'Delete MOdeling View starting with Primary1 or Primary2
            For Each objModelView As ModelingView In objPart.ModelingViews
                If objModelView.Name.ToUpper.StartsWith("PRIMARY") Then
                    sDeleteModellingView(objPart, objModelView.Name)
                End If
            Next

            If FnChkIfModelingViewPresent(objPart, B_PRIMARY1) Then
                If FnChkIfModelingViewPresent(objPart, B_PRIMARY2) Then
                    'Its a two frame. 
                    sCreateModelingViewByName(objPart, PRIMARY1_FRONTVIEW, B_PRIMARY1)
                    sCreateModelingViewByName(objPart, PRIMARY1_RIGHTVIEW, B_PRIMARY1)
                    sCreateModelingViewByName(objPart, PRIMARY1_TOPVIEW, B_PRIMARY1)
                    sCreateModelingViewByName(objPart, PRIMARY1_REARVIEW, B_PRIMARY1)
                    sCreateModelingViewByName(objPart, PRIMARY1_LEFTVIEW, B_PRIMARY1)
                    sCreateModelingViewByName(objPart, PRIMARY1_BOTTOMVIEW, B_PRIMARY1)

                    sCreateModelingViewByName(objPart, PRIMARY2_FRONTVIEW, B_PRIMARY2)
                    sCreateModelingViewByName(objPart, PRIMARY2_RIGHTVIEW, B_PRIMARY2)
                    sCreateModelingViewByName(objPart, PRIMARY2_TOPVIEW, B_PRIMARY2)
                    sCreateModelingViewByName(objPart, PRIMARY2_REARVIEW, B_PRIMARY2)
                    sCreateModelingViewByName(objPart, PRIMARY2_LEFTVIEW, B_PRIMARY2)
                    sCreateModelingViewByName(objPart, PRIMARY2_BOTTOMVIEW, B_PRIMARY2)

                    asVisibleEdgeNamesInPrimary1FrontVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_FRONTVIEW)
                    asVisibleEdgeNamesInPrimary1RightVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_RIGHTVIEW)
                    asVisibleEdgeNamesInPrimary1TopVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_TOPVIEW)
                    asVisibleEdgeNamesInPrimary1RearVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_REARVIEW)
                    asVisibleEdgeNamesInPrimary1LeftVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_LEFTVIEW)
                    asVisibleEDgeNamesInPrimary1BottomVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_BOTTOMVIEW)

                    asVisibleEdgeNamesInPrimary2FrontVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY2_FRONTVIEW)
                    asVisibleEdgeNamesInPrimary2RightVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY2_RIGHTVIEW)
                    asVisibleEdgeNamesInPrimary2TopVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY2_TOPVIEW)
                    asVisibleEdgeNamesInPrimary2RearVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY2_REARVIEW)
                    asVisibleEdgeNamesInPrimary2LeftVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY2_LEFTVIEW)
                    asVisibleEDgeNamesInPrimary2BottomVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY2_BOTTOMVIEW)

                    bIsPrimaryView2Present = True
                Else
                    'Its a single frame
                    sCreateModelingViewByName(objPart, PRIMARY1_FRONTVIEW, B_PRIMARY1)
                    sCreateModelingViewByName(objPart, PRIMARY1_RIGHTVIEW, B_PRIMARY1)
                    sCreateModelingViewByName(objPart, PRIMARY1_TOPVIEW, B_PRIMARY1)
                    sCreateModelingViewByName(objPart, PRIMARY1_REARVIEW, B_PRIMARY1)
                    sCreateModelingViewByName(objPart, PRIMARY1_LEFTVIEW, B_PRIMARY1)
                    sCreateModelingViewByName(objPart, PRIMARY1_BOTTOMVIEW, B_PRIMARY1)

                    asVisibleEdgeNamesInPrimary1FrontVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_FRONTVIEW)
                    asVisibleEdgeNamesInPrimary1RightVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_RIGHTVIEW)
                    asVisibleEdgeNamesInPrimary1TopVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_TOPVIEW)
                    asVisibleEdgeNamesInPrimary1RearVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_REARVIEW)
                    asVisibleEdgeNamesInPrimary1LeftVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_LEFTVIEW)
                    asVisibleEDgeNamesInPrimary1BottomVw = FnPopulateVisibleEdgeNamesInAView(objPart, PRIMARY1_BOTTOMVIEW)

                End If
            End If

            iColEdgeNames = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, EDGENAMES_HEADER, iRowStartHeader)

            iColPrimary1Front = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY1FRONT_HEADER, iRowStartHeader)
            iColPrimary1Right = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY1RIGHT_HEADER, iRowStartHeader)
            iColPrimary1Left = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY1LEFT_HEADER, iRowStartHeader)
            iColPrimary1Back = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY1REAR_HEADER, iRowStartHeader)
            iColPrimary1Top = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY1TOP_HEADER, iRowStartHeader)
            iColPrimary1Bottom = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY1BOTTOM_HEADER, iRowStartHeader)

            If bIsPrimaryView2Present Then
                iColPrimary2Front = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY2FRONT_HEADER, iRowStartHeader)
                iColPrimary2Right = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY2RIGHT_HEADER, iRowStartHeader)
                iColPrimary2Left = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY2LEFT_HEADER, iRowStartHeader)
                iColPrimary2Back = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY2REAR_HEADER, iRowStartHeader)
                iColPrimary2Top = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY2TOP_HEADER, iRowStartHeader)
                iColPrimary2Bottom = FnGetColumnNumberByName(_objWorkBk, VISIBLEEDGE_SHEETNAME, PRIMARY2BOTTOM_HEADER, iRowStartHeader)
            End If

            sSetStatus("Populating Visible Edge Name")

            If Not FnGetAllComponentsInSession() Is Nothing Then
                For Each objComp As Component In FnGetAllComponentsInSession()
                    If Not FnGetPartFromComponent(objComp) Is Nothing Then
                        FnLoadPartFully(FnGetPartFromComponent(objComp))
                        aoAllValidBody = FnGetValidBodyForOEM(FnGetPartFromComponent(objComp), _sOemName)
                        If Not aoAllValidBody Is Nothing Then
                            For Each objBody As Body In aoAllValidBody

                                If objBody.IsSolidBody Then

                                    If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                        If objComp Is objPart.ComponentAssembly.RootComponent Then
                                            objBodyToCheck = objBody
                                        Else
                                            objBodyToCheck = CType(objComp.FindOccurrence(objBody), Body)
                                        End If
                                    ElseIf (_sOemName = DAIMLER_OEM_NAME) Then
                                        If _sDivision = TRUCK_DIVISION Then
                                            If objComp Is objPart.ComponentAssembly.RootComponent Then
                                                objBodyToCheck = objBody
                                            Else
                                                objBodyToCheck = CType(objComp.FindOccurrence(objBody), Body)
                                            End If
                                        ElseIf _sDivision = CAR_DIVISION Then
                                            If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                                'Weldment in car
                                                objBodyToCheck = CType(objComp.FindOccurrence(objBody), Body)
                                            Else
                                                'Component in car
                                                objBodyToCheck = objBody
                                            End If
                                        End If
                                    ElseIf (_sOemName = FIAT_OEM_NAME) Then
                                        'Check if the component is a child component in weldment
                                        If objComp Is objPart.ComponentAssembly.RootComponent Then
                                            'Fiat component
                                            objBodyToCheck = objBody
                                        Else
                                            'This is a Fiat Weldment child component
                                            objBodyToCheck = CType(objComp.FindOccurrence(objBody), Body)
                                        End If

                                    End If


                                    If Not objBodyToCheck Is Nothing Then
                                        For Each objEdge As Edge In objBodyToCheck.GetEdges()
                                            'Write the EDge Name
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColEdgeNames, objEdge.Name.ToUpper)
                                            'iRowStartVisibleEdge = FnFindRowNumberByColumnAndValue(_objWorkBk, VISIBLEEDGE_SHEETNAME, 4, 3, objEdge.Name.ToUpper)

                                            If Not asVisibleEdgeNamesInPrimary1FrontVw Is Nothing And asVisibleEdgeNamesInPrimary1FrontVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Front, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Front, EDGE_NOT_VISIBLE)
                                            End If

                                            If Not asVisibleEdgeNamesInPrimary1RightVw Is Nothing And asVisibleEdgeNamesInPrimary1RightVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Right, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Right, EDGE_NOT_VISIBLE)
                                            End If

                                            If Not asVisibleEdgeNamesInPrimary1LeftVw Is Nothing And asVisibleEdgeNamesInPrimary1LeftVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Left, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Left, EDGE_NOT_VISIBLE)
                                            End If

                                            If Not asVisibleEdgeNamesInPrimary1RearVw Is Nothing And asVisibleEdgeNamesInPrimary1RearVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Back, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Back, EDGE_NOT_VISIBLE)
                                            End If

                                            If Not asVisibleEdgeNamesInPrimary1TopVw Is Nothing And asVisibleEdgeNamesInPrimary1TopVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Top, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Top, EDGE_NOT_VISIBLE)
                                            End If

                                            If Not asVisibleEDgeNamesInPrimary1BottomVw Is Nothing And asVisibleEDgeNamesInPrimary1BottomVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Bottom, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Bottom, EDGE_NOT_VISIBLE)
                                            End If

                                            If bIsPrimaryView2Present Then
                                                If Not asVisibleEdgeNamesInPrimary2FrontVw Is Nothing And asVisibleEdgeNamesInPrimary2FrontVw.Contains(objEdge.Name) Then
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Front, EDGE_VISIBLE)
                                                Else
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Front, EDGE_NOT_VISIBLE)
                                                End If

                                                If Not asVisibleEdgeNamesInPrimary2RightVw Is Nothing And asVisibleEdgeNamesInPrimary2RightVw.Contains(objEdge.Name) Then
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Right, EDGE_VISIBLE)
                                                Else
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Right, EDGE_NOT_VISIBLE)
                                                End If

                                                If Not asVisibleEdgeNamesInPrimary2LeftVw Is Nothing And asVisibleEdgeNamesInPrimary2LeftVw.Contains(objEdge.Name) Then
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Left, EDGE_VISIBLE)
                                                Else
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Left, EDGE_NOT_VISIBLE)
                                                End If

                                                If Not asVisibleEdgeNamesInPrimary2RearVw Is Nothing And asVisibleEdgeNamesInPrimary2RearVw.Contains(objEdge.Name) Then
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Back, EDGE_VISIBLE)
                                                Else
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Back, EDGE_NOT_VISIBLE)
                                                End If

                                                If Not asVisibleEdgeNamesInPrimary2TopVw Is Nothing And asVisibleEdgeNamesInPrimary2TopVw.Contains(objEdge.Name) Then
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Top, EDGE_VISIBLE)
                                                Else
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Top, EDGE_NOT_VISIBLE)
                                                End If

                                                If Not asVisibleEDgeNamesInPrimary2BottomVw Is Nothing And asVisibleEDgeNamesInPrimary2BottomVw.Contains(objEdge.Name) Then
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Bottom, EDGE_VISIBLE)
                                                Else
                                                    SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Bottom, EDGE_NOT_VISIBLE)
                                                End If
                                            End If
                                            iRowStartVisibleEdge = iRowStartVisibleEdge + 1
                                        Next
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            Else
                If Not objPart Is Nothing Then
                    FnLoadPartFully(objPart)
                    aoAllValidBody = FnGetValidBodyForOEM(objPart, _sOemName)
                    If Not aoAllValidBody Is Nothing Then
                        For Each objBody As Body In aoAllValidBody
                            If objBody.IsSolidBody Then
                                objBodyToCheck = objBody
                                If Not objBodyToCheck Is Nothing Then
                                    For Each objEdge As Edge In objBodyToCheck.GetEdges()
                                        'Write the EDge Name
                                        SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColEdgeNames, objEdge.Name.ToUpper)
                                        'iRowStartVisibleEdge = FnFindRowNumberByColumnAndValue(_objWorkBk, VISIBLEEDGE_SHEETNAME, 4, 3, objEdge.Name.ToUpper)

                                        If Not asVisibleEdgeNamesInPrimary1FrontVw Is Nothing And asVisibleEdgeNamesInPrimary1FrontVw.Contains(objEdge.Name) Then
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Front, EDGE_VISIBLE)
                                        Else
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Front, EDGE_NOT_VISIBLE)
                                        End If

                                        If Not asVisibleEdgeNamesInPrimary1RightVw Is Nothing And asVisibleEdgeNamesInPrimary1RightVw.Contains(objEdge.Name) Then
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Right, EDGE_VISIBLE)
                                        Else
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Right, EDGE_NOT_VISIBLE)
                                        End If

                                        If Not asVisibleEdgeNamesInPrimary1LeftVw Is Nothing And asVisibleEdgeNamesInPrimary1LeftVw.Contains(objEdge.Name) Then
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Left, EDGE_VISIBLE)
                                        Else
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Left, EDGE_NOT_VISIBLE)
                                        End If

                                        If Not asVisibleEdgeNamesInPrimary1RearVw Is Nothing And asVisibleEdgeNamesInPrimary1RearVw.Contains(objEdge.Name) Then
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Back, EDGE_VISIBLE)
                                        Else
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Back, EDGE_NOT_VISIBLE)
                                        End If

                                        If Not asVisibleEdgeNamesInPrimary1TopVw Is Nothing And asVisibleEdgeNamesInPrimary1TopVw.Contains(objEdge.Name) Then
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Top, EDGE_VISIBLE)
                                        Else
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Top, EDGE_NOT_VISIBLE)
                                        End If

                                        If Not asVisibleEDgeNamesInPrimary1BottomVw Is Nothing And asVisibleEDgeNamesInPrimary1BottomVw.Contains(objEdge.Name) Then
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Bottom, EDGE_VISIBLE)
                                        Else
                                            SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary1Bottom, EDGE_NOT_VISIBLE)
                                        End If

                                        If bIsPrimaryView2Present Then
                                            If Not asVisibleEdgeNamesInPrimary2FrontVw Is Nothing And asVisibleEdgeNamesInPrimary2FrontVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Front, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Front, EDGE_NOT_VISIBLE)
                                            End If

                                            If Not asVisibleEdgeNamesInPrimary2RightVw Is Nothing And asVisibleEdgeNamesInPrimary2RightVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Right, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Right, EDGE_NOT_VISIBLE)
                                            End If

                                            If Not asVisibleEdgeNamesInPrimary2LeftVw Is Nothing And asVisibleEdgeNamesInPrimary2LeftVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Left, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Left, EDGE_NOT_VISIBLE)
                                            End If

                                            If Not asVisibleEdgeNamesInPrimary2RearVw Is Nothing And asVisibleEdgeNamesInPrimary2RearVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Back, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Back, EDGE_NOT_VISIBLE)
                                            End If

                                            If Not asVisibleEdgeNamesInPrimary2TopVw Is Nothing And asVisibleEdgeNamesInPrimary2TopVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Top, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Top, EDGE_NOT_VISIBLE)
                                            End If

                                            If Not asVisibleEDgeNamesInPrimary2BottomVw Is Nothing And asVisibleEDgeNamesInPrimary2BottomVw.Contains(objEdge.Name) Then
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Bottom, EDGE_VISIBLE)
                                            Else
                                                SWriteValueToCell(_objWorkBk, VISIBLEEDGE_SHEETNAME, iRowStartVisibleEdge, iColPrimary2Bottom, EDGE_NOT_VISIBLE)
                                            End If
                                        End If
                                        iRowStartVisibleEdge = iRowStartVisibleEdge + 1
                                    Next
                                End If
                            End If

                        Next
                    End If
                End If
            End If
        End If
    End Sub

    Public Function FnPopulateVisibleEdgeNamesInAView(ByVal objPart As Part, ByVal sRefViewName As String) As String()
        Dim currentView As ModelingView = Nothing
        Dim baseView As ModelingView = Nothing
        Dim RefDraftingView As DraftingView = Nothing
        Dim objView As View = Nothing
        Dim sVisibleEdgeNames() As String = Nothing

        'CHange the View type to Modelling
        FnViewType(objPart)

        If Not objPart Is Nothing Then
            currentView = objPart.ModelingViews.WorkView
            baseView = FnGetModellingView(objPart, sRefViewName)
            'Make the current view as sREfViewName which is he base view
            If Not currentView Is baseView Then
                sReplaceViewInLayout(objPart, baseView)
            End If
            FnGetNxSession.Parts.Work.Views.WorkView.Fit()
            'OPen Drafting application
            sOpenDraftingApplication(objPart)
            'Code added May-14-2018
            'Delete existing Drawing sheet if any
            sDeleteDrawingSheet(objPart, sTempSheet)
            FnInsertDrawingSheet(objPart, sTempSheet)
            FnInsertBaseView(objPart, "B_" & sRefViewName, baseView.Name, 0.125, 150, 150, 0)
            'Get the Drafting View
            RefDraftingView = FnGetViewByName(objPart, "B_" & sRefViewName)

            sChangeViewSettingsForIsoVw(RefDraftingView)
            objView = RefDraftingView
            If Not objView Is Nothing Then
                sVisibleEdgeNames = FnGetVisibleEdges(objPart, objView)
                'Code added on Nov-29-2016
                'There was a peak in memory
                FnSavePart(objPart)
            End If
            sDeleteDrawingSheet(objPart, sTempSheet)
            FnViewType(objPart)
        End If
        FnPopulateVisibleEdgeNamesInAView = sVisibleEdgeNames
    End Function
    Public Function FnGetModellingView(ByVal objPart As Part, ByVal sViewName As String) As ModelingView
        For Each objModelView As ModelingView In objPart.ModelingViews
            If objModelView.Name.ToUpper = sViewName.ToUpper Then
                FnGetModellingView = objModelView
                Exit Function
            End If
        Next
        FnGetModellingView = Nothing
    End Function
    'Function to Insert an Empty Drawing Sheet
    Public Function FnInsertDrawingSheet(ByVal objPart As Part, ByVal sSheetName As String)
        Dim sUtil_dir As String = ""
        Dim sPath As String = ""
        Dim ufs As UFSession = UFSession.GetUFSession

        Dim nullDrawings_DrawingSheet As DrawingSheet = Nothing
        Dim objDrawingSheetBuilder As DrawingSheetBuilder = Nothing

        objDrawingSheetBuilder = objPart.DrawingSheets.DrawingSheetBuilder(nullDrawings_DrawingSheet)
        objDrawingSheetBuilder.AutoStartViewCreation = False
        objDrawingSheetBuilder.Option = NXOpen.Drawings.DrawingSheetBuilder.SheetOption.StandardSize
        objDrawingSheetBuilder.StandardMetricScale = NXOpen.Drawings.DrawingSheetBuilder.SheetStandardMetricScale.S11
        objDrawingSheetBuilder.StandardEnglishScale = NXOpen.Drawings.DrawingSheetBuilder.SheetStandardEnglishScale.S11
        ufs.UF.TranslateVariable("UGII_UTIL", sUtil_dir)
        objDrawingSheetBuilder.ProjectionAngle = NXOpen.Drawings.DrawingSheetBuilder.SheetProjectionAngle.Third
        objDrawingSheetBuilder.Name = sSheetName

        Dim nXObject1 As NXObject
        nXObject1 = objDrawingSheetBuilder.Commit()
        CType(nXObject1, DrawingSheet).Open()
        objDrawingSheetBuilder.Destroy()
        FnInsertDrawingSheet = CType(nXObject1, DrawingSheet)
    End Function

    Public Function FnInsertBaseView(ByVal objPart As Part, ByVal sViewName As String, ByVal sModelingViewName As String, ByVal dScale As Double, ByVal dx As Double,
                              ByVal dY As Double, ByVal dZ As Double) As BaseView
        Dim nullDrawings_BaseView As BaseView = Nothing

        Dim baseViewBuilder1 As BaseViewBuilder
        baseViewBuilder1 = objPart.DraftingViews.CreateBaseViewBuilder(nullDrawings_BaseView)
        baseViewBuilder1.SelectModelView.SelectedView = FnGetModellingView(objPart, sModelingViewName)
        baseViewBuilder1.Style.ViewStyleBase.Part = objPart
        baseViewBuilder1.Style.ViewStyleBase.PartName = objPart.FullPath
        baseViewBuilder1.Style.ViewStyleDetail.ViewBoundaryWidth = Preferences.Width.Thin
        Dim nullAssemblies_Arrangement As Assemblies.Arrangement = Nothing
        baseViewBuilder1.Style.ViewStyleBase.Arrangement.SelectedArrangement = nullAssemblies_Arrangement
        baseViewBuilder1.Scale.Denominator = 1.0
        baseViewBuilder1.Scale.Numerator = dScale

        'Place the view
        Dim point1 As Point3d = New Point3d(dx, dY, dZ)
        'point that helps determine the view's position based on the alignment method and alignment point specified. 
        baseViewBuilder1.Placement.Placement.SetValue(Nothing, objPart.Views.WorkView, point1)

        Dim nXObject2 As NXObject
        nXObject2 = baseViewBuilder1.Commit()
        baseViewBuilder1.Destroy()

        CType(nXObject2, NXOpen.Drawings.BaseView).SetName(sViewName)
        'CODE COMMENTED - 4/21/16 - View Label not required
        'If FnChkPartisWeldment() Then
        '    CType(nXObject2, NXOpen.Drawings.BaseView).Style.General.ViewLabel = True
        '    CType(nXObject2, NXOpen.Drawings.BaseView).Style.General.ScaleLabel = True
        '    CType(nXObject2, NXOpen.Drawings.BaseView).Commit()
        'Else
        CType(nXObject2, NXOpen.Drawings.BaseView).Style.General.ViewLabel = False
        CType(nXObject2, NXOpen.Drawings.BaseView).Style.General.ScaleLabel = False
        CType(nXObject2, NXOpen.Drawings.BaseView).Commit()
        'End If
        FnInsertBaseView = CType(nXObject2, NXOpen.Drawings.BaseView)

    End Function

    'Function to get the visible objects in drafting view

    Public Function FnGetVisibleObjectsInAView(ByVal objView As DraftingView) As NXObject()
        Dim iNoOfVisibleObjects As Integer = 0
        Dim tVisibleObjects() As Tag = Nothing
        Dim iNoOfClippedObjects As Integer = 0
        Dim tClippedObjects() As Tag = Nothing
        Dim objVisibleEdges() As Edge = Nothing

        FnGetUFSession.View.AskVisibleObjects(objView.Tag, iNoOfVisibleObjects, tVisibleObjects, iNoOfClippedObjects, tClippedObjects)

        Dim groupTag As Tag = NXOpen.Tag.Null
        Dim agroup As NXOpen.Group = Nothing
        Dim iGroupCount As Integer = 0
        Dim objVisible() As NXObject = Nothing
        If Not tVisibleObjects Is Nothing Then
            For Each objTag As Tag In tVisibleObjects
                ReDim Preserve objVisible(iGroupCount)
                objVisible(iGroupCount) = NXObjectManager.Get(objTag)
                iGroupCount = iGroupCount + 1
            Next
        End If

        FnGetVisibleObjectsInAView = objVisible
    End Function

    'Function to get the visibleEdge Names from the visible objects in a drafting view

    Public Function FnGetVisibleEdges(ByVal objPart As Part, ByVal objView As DraftingView) As String()
        Dim objEdgeTag() As Tag = Nothing
        Dim iTagCount As Integer = 0
        Dim objVisibleEdge As Edge = Nothing
        Dim sVisibleEdgesName() As String = Nothing
        Dim objTag As Tag = NXOpen.Tag.Null
        Dim objVisible() As NXObject = Nothing
        Dim objBodyToCheck As Body = Nothing
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim bIsValidSolidBody As Boolean = False
        Dim aoAllValidBody() As Body = Nothing

        If Not objPart Is Nothing Then
            If Not FnGetAllComponentsInSession() Is Nothing Then
                For Each objComp As Component In FnGetAllComponentsInSession()
                    If Not FnGetPartFromComponent(objComp) Is Nothing Then
                        FnLoadPartFully(FnGetPartFromComponent(objComp))
                        aoAllValidBody = FnGetValidBodyForOEM(FnGetPartFromComponent(objComp), _sOemName)
                        If Not aoAllValidBody Is Nothing Then
                            For Each objBody As Body In aoAllValidBody
                                If objBody.IsSolidBody Then
                                    If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                        If objComp Is objPart.ComponentAssembly.RootComponent Then
                                            objBodyToCheck = objBody
                                        Else
                                            objBodyToCheck = CType(objComp.FindOccurrence(objBody), Body)
                                        End If
                                    ElseIf (_sOemName = DAIMLER_OEM_NAME) Then
                                        If _sDivision = TRUCK_DIVISION Then
                                            If objComp Is objPart.ComponentAssembly.RootComponent Then
                                                objBodyToCheck = objBody
                                            Else
                                                objBodyToCheck = CType(objComp.FindOccurrence(objBody), Body)
                                            End If
                                        ElseIf _sDivision = CAR_DIVISION Then
                                            If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                                'Weldment in car
                                                objBodyToCheck = CType(objComp.FindOccurrence(objBody), Body)
                                            Else
                                                'Component in car
                                                objBodyToCheck = objBody
                                            End If
                                        End If
                                    ElseIf (_sOemName = FIAT_OEM_NAME) Then
                                        If objComp Is objPart.ComponentAssembly.RootComponent Then
                                            objBodyToCheck = objBody
                                        Else
                                            objBodyToCheck = CType(objComp.FindOccurrence(objBody), Body)
                                        End If
                                    End If


                                    If Not objBodyToCheck Is Nothing Then
                                        'For Each objFace As Face In objBody.GetFaces()
                                        For Each objEdge As Edge In objBodyToCheck.GetEdges()
                                            For Each objDraftingBody As DraftingBody In objView.DraftingBodies
                                                For Each objDisplayedObject As DisplayableObject In objDraftingBody.DraftingCurves
                                                    If objDisplayedObject.JournalIdentifier.ToUpper.Contains(objEdge.JournalIdentifier.ToUpper) Then

                                                        If sVisibleEdgesName Is Nothing Then
                                                            ReDim Preserve sVisibleEdgesName(0)
                                                            sVisibleEdgesName(0) = objEdge.Name
                                                        Else
                                                            ReDim Preserve sVisibleEdgesName(UBound(sVisibleEdgesName) + 1)
                                                            sVisibleEdgesName(UBound(sVisibleEdgesName)) = objEdge.Name
                                                        End If

                                                    End If
                                                Next

                                            Next
                                        Next
                                    End If
                                    'Next
                                End If

                            Next
                        End If
                    End If
                Next
            Else
                If Not objPart Is Nothing Then
                    FnLoadPartFully(objPart)
                    aoAllValidBody = FnGetValidBodyForOEM(objPart, _sOemName)
                    If Not aoAllValidBody Is Nothing Then
                        For Each objBody As Body In aoAllValidBody
                            If objBody.IsSolidBody Then

                                objBodyToCheck = objBody
                                If Not objBodyToCheck Is Nothing Then
                                    'For Each objFace As Face In objBody.GetFaces()
                                    For Each objEdge As Edge In objBodyToCheck.GetEdges()
                                        For Each objDraftingBody As DraftingBody In objView.DraftingBodies
                                            For Each objDisplayedObject As DisplayableObject In objDraftingBody.DraftingCurves
                                                If objDisplayedObject.JournalIdentifier.ToUpper.Contains(objEdge.JournalIdentifier.ToUpper) Then

                                                    If sVisibleEdgesName Is Nothing Then
                                                        ReDim Preserve sVisibleEdgesName(0)
                                                        sVisibleEdgesName(0) = objEdge.Name
                                                    Else
                                                        ReDim Preserve sVisibleEdgesName(UBound(sVisibleEdgesName) + 1)
                                                        sVisibleEdgesName(UBound(sVisibleEdgesName)) = objEdge.Name
                                                    End If

                                                End If
                                            Next
                                        Next
                                    Next
                                    'Next
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
        FnGetVisibleEdges = sVisibleEdgesName
    End Function
    'Function to Delete the Drawing Sheet
    Sub sDeleteDrawingSheet(ByVal objPart As Part, ByVal sSheetName As String)
        For Each sheet As DrawingSheet In objPart.DrawingSheets
            If sheet.Name = sSheetName Then
                For Each sview As DraftingView In sheet.GetDraftingViews
                    SDeleteObjects({sview})
                Next
                SDeleteObjects({sheet})
            End If
        Next
    End Sub

    'Function to create modelling view based on the input name given the user to obtain standard right or left or top or bottom or front or back projections

    Public Sub sCreateModelingViewByName(ByVal objPart As Part, ByVal sNewViewName As String, ByVal sDefaultRefViewName As String)

        Dim baseView As ModelingView = Nothing
        Dim currentView As ModelingView = Nothing
        Dim rotMatPrimaryView As Matrix3x3 = Nothing
        Dim origin As Point3d = Nothing
        Dim origin1 As Point3d = Nothing
        Dim vector As Vector3d
        Dim vector1 As Vector3d
        Dim iAngle As Integer = 0
        Dim objView As View = Nothing

        'Get the sdefaultRefView as base view
        If Not sDefaultRefViewName Is Nothing Then
            If FnChkIfModelingViewPresent(objPart, sDefaultRefViewName) Then
                baseView = objPart.ModelingViews.FindObject(sDefaultRefViewName)
            End If
            currentView = objPart.ModelingViews.WorkView
            If Not currentView Is baseView Then
                sReplaceViewInLayout(objPart, baseView)
            End If
            rotMatPrimaryView = FnGetRotationMatrix(objPart)
            origin = objPart.ModelingViews.WorkView.AbsoluteOrigin
            origin1 = New Point3d(origin.X, origin.Y, origin.Z)

            If sNewViewName.Contains("RIGHT") Then
                vector = New Vector3d(rotMatPrimaryView.Yx, rotMatPrimaryView.Yy, rotMatPrimaryView.Yz)
                vector1 = New Vector3d(vector.X, vector.Y, vector.Z)
                iAngle = -90
            ElseIf sNewViewName.Contains("LEFT") Then
                vector = New Vector3d(rotMatPrimaryView.Yx, rotMatPrimaryView.Yy, rotMatPrimaryView.Yz)
                vector1 = New Vector3d(vector.X, vector.Y, vector.Z)
                iAngle = 90
            ElseIf sNewViewName.Contains("TOP") Then
                vector = New Vector3d(rotMatPrimaryView.Xx, rotMatPrimaryView.Xy, rotMatPrimaryView.Xz)
                vector1 = New Vector3d(vector.X, vector.Y, vector.Z)
                iAngle = 90
            ElseIf sNewViewName.Contains("BOTTOM") Then
                vector = New Vector3d(rotMatPrimaryView.Xx, rotMatPrimaryView.Xy, rotMatPrimaryView.Xz)
                vector1 = New Vector3d(vector.X, vector.Y, vector.Z)
                iAngle = -90
            ElseIf sNewViewName.Contains("REAR") Then
                vector = New Vector3d(rotMatPrimaryView.Yx, rotMatPrimaryView.Yy, rotMatPrimaryView.Yz)
                vector1 = New Vector3d(vector.X, vector.Y, vector.Z)
                iAngle = -180
            ElseIf sNewViewName.Contains("FRONT") Then
                vector = New Vector3d(rotMatPrimaryView.Yx, rotMatPrimaryView.Yy, rotMatPrimaryView.Yz)
                vector1 = New Vector3d(vector.X, vector.Y, vector.Z)
                iAngle = 0
            End If
            FnViewType(objPart)

            If objPart.ModelingViews.WorkView Is Nothing Then
                objPart.ModelingViews.WorkView.Rotate(origin1, vector1, iAngle)
                objView = objPart.Views.SaveAs(objPart.ModelingViews.WorkView, sNewViewName, True, False)
            Else
                For Each objModelView As ModelingView In objPart.ModelingViews
                    If objModelView.Name.ToUpper = sDefaultRefViewName Then
                        sReplaceViewInLayout(objPart, objModelView)
                        objModelView.Rotate(origin1, vector1, iAngle)
                        objModelView.Fit()
                        objPart.Views.SaveAs(objModelView, sNewViewName, True, False)
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub

    ' Function to Get the Rotation Matrix
    Public Function FnGetRotationMatrix(ByRef objpart As Part) As Matrix3x3
        FnGetNxSession()
        FnGetDisplayedPart()
        FnViewType(objpart)
        Dim currentview As ModelingView
        currentview = objpart.ModelingViews.WorkView
        Dim matrix As Matrix3x3 = currentview.Matrix
        Dim Xx As Double = matrix.Xx
        Dim Xy As Double = matrix.Xy
        Dim Xz As Double = matrix.Xz
        Dim Yx As Double = matrix.Yx
        Dim Yy As Double = matrix.Yy
        Dim Yz As Double = matrix.Yz
        Dim Zx As Double = matrix.Zx
        Dim Zy As Double = matrix.Zy
        Dim Zz As Double = matrix.Zz
        FnGetRotationMatrix = matrix
    End Function
    'Code added on Nov-10-2016
    'Function to collect the faces from Extrude Feature (which is the child of TEXT feature).
    'Assuming TEXT feature will always has EXTRUDE Feature as its child.
    Function FnGetTextExtrudeFeatureFaces(objToolPart As Part, aoAllCompInSession() As Component) As Face()
        'Dim aoTextExtrudeFeatureFaces() As Face = Nothing
        Dim objExtrude As Extrude = Nothing
        Dim objPart As Part = Nothing
        Dim objOccFace As Face = Nothing
        Dim bIsCustomComp As Boolean = False

        If Not objToolPart Is Nothing Then
            If Not objToolPart.ComponentAssembly.RootComponent Is Nothing Then
                For Each objComp As Component In aoAllCompInSession
                    bIsCustomComp = False
                    If Not objComp Is Nothing Then
                        'If FnChkIfComponentIsCustom(objComp, bIgnoreCarryOverComponents:=False) Then
                        If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = DAIMLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                            If FnCheckIfThisIsAMakeDetailBasedOnName(objComp.DisplayName) Then
                                bIsCustomComp = True
                            End If
                        ElseIf (_sOemName = FIAT_OEM_NAME) Then
                            If FnCheckIfThisIsAMakeDetailBasedOnAttribute(FnGetPartFromComponent(objComp)) Then
                                bIsCustomComp = True
                            End If
                        End If

                        If bIsCustomComp Then
                            objPart = FnGetPartFromComponent(objComp)
                            If Not objPart Is Nothing Then
                                FnLoadPartFully(objPart)
                                For Each objFeature As Features.Feature In objPart.Features
                                    If objFeature.FeatureType.ToString().ToUpper = "TEXT" Then
                                        For Each objChild As Features.Feature In objFeature.GetChildren()
                                            If Not objChild Is Nothing Then
                                                'Currently only handling extrude type features and getting its constituent faces
                                                If objChild.GetType().ToString.ToUpper = "NXOPEN.FEATURES.EXTRUDE" Then
                                                    objExtrude = CType(objChild, Extrude)
                                                    For Each objFace As Face In objExtrude.GetFaces()
                                                        If objComp Is objToolPart.ComponentAssembly.RootComponent Then
                                                            objOccFace = objFace
                                                        Else
                                                            objOccFace = objComp.FindOccurrence(objFace)
                                                        End If
                                                        If Not objOccFace Is Nothing Then
                                                            If _aoTextExtrudeFeatureFaces Is Nothing Then
                                                                ReDim Preserve _aoTextExtrudeFeatureFaces(0)
                                                                _aoTextExtrudeFeatureFaces(0) = objOccFace
                                                            Else
                                                                ReDim Preserve _aoTextExtrudeFeatureFaces(UBound(_aoTextExtrudeFeatureFaces) + 1)
                                                                _aoTextExtrudeFeatureFaces(UBound(_aoTextExtrudeFeatureFaces)) = objOccFace
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
            Else
                For Each objFeature As Features.Feature In objToolPart.Features
                    If objFeature.FeatureType.ToString().ToUpper = "TEXT" Then
                        For Each objChild As Features.Feature In objFeature.GetChildren()
                            If Not objChild Is Nothing Then
                                If objChild.GetType().ToString.ToUpper = "NXOPEN.FEATURES.EXTRUDE" Then
                                    objExtrude = CType(objChild, Extrude)
                                    For Each objFace As Face In objExtrude.GetFaces()
                                        If _aoTextExtrudeFeatureFaces Is Nothing Then
                                            ReDim Preserve _aoTextExtrudeFeatureFaces(0)
                                            _aoTextExtrudeFeatureFaces(0) = objFace
                                        Else
                                            ReDim Preserve _aoTextExtrudeFeatureFaces(UBound(_aoTextExtrudeFeatureFaces) + 1)
                                            _aoTextExtrudeFeatureFaces(UBound(_aoTextExtrudeFeatureFaces)) = objFace
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If
                Next
            End If

        End If
        FnGetTextExtrudeFeatureFaces = _aoTextExtrudeFeatureFaces
    End Function

    'Function to check if the face has to be ignored based on Text Extrude Feature Face
    Function FnCheckIfFaceToBeIgnoredBasedOnTextExtrudeFeature(objFace As Face) As Boolean
        If Not _aoTextExtrudeFeatureFaces Is Nothing Then
            For Each objExtrudeFace As Face In _aoTextExtrudeFeatureFaces
                If objExtrudeFace Is objFace Then
                    FnCheckIfFaceToBeIgnoredBasedOnTextExtrudeFeature = True
                    Exit Function
                End If
            Next
        End If
        FnCheckIfFaceToBeIgnoredBasedOnTextExtrudeFeature = False
    End Function
    'Code commented on May-25-2017
    'B_FLOOR_MOUNT view is not created as part of Part sweep data. It is handled by Pradeep in KeySheet
    ''Code Added on NOV-28-2016
    'Sub sCreateFloorMountView(objPart As Part)

    '    Dim aoFloorMountBody() As Body = Nothing
    '    Dim objFloorMountBody As Body = Nothing
    '    Dim objMatingFaceOnFloorMountBody As Face = Nothing
    '    Dim adMatingFaceNormalInFloorMountBody() As Double = Nothing
    '    'Dim objSmallestEdgeInMatingFaceOnAFloorMountingBody As Edge = Nothing
    '    Dim objOrthogonalFaceOnFloorMountBody As Face = Nothing
    '    Dim adOrthogonalFaceNormalInFloorMountBody() As Double = Nothing
    '    Dim adRotationMatrix(8) As Double

    '    'Get the collection of all floor mount body
    '    aoFloorMountBody = FnCollectAllFloorMountBodys(objPart)
    '    If Not aoFloorMountBody Is Nothing Then
    '        objFloorMountBody = aoFloorMountBody(0)
    '        If Not objFloorMountBody Is Nothing Then
    '            'Finding the planar face for the non dowel hole which is mating with another solid body within the frame / Base component.
    '            objMatingFaceOnFloorMountBody = FnGetMatingFaceInASolidBody(objPart, objFloorMountBody)
    '            If Not objMatingFaceOnFloorMountBody Is Nothing Then
    '                'Get the orthogonal face of objMatingFaceonFloorMountBody
    '                objOrthogonalFaceOnFloorMountBody = FnGetOrthogonalFaceToRefFaceInABody(objFloorMountBody, objMatingFaceOnFloorMountBody)
    '                If Not objOrthogonalFaceOnFloorMountBody Is Nothing Then
    '                    'Find the face normal vector of both the Mating face and Orthogonal face in a Floor Mounting Body
    '                    adMatingFaceNormalInFloorMountBody = FnGetFaceNormalVector(objMatingFaceOnFloorMountBody)
    '                    adOrthogonalFaceNormalInFloorMountBody = FnGetFaceNormalVector(objOrthogonalFaceOnFloorMountBody)
    '                    If Not adMatingFaceNormalInFloorMountBody Is Nothing And Not adOrthogonalFaceNormalInFloorMountBody Is Nothing Then
    '                        'Check the orthogonality of two vectors
    '                        If FnCheckOrthogalityOfTwoVectors(adOrthogonalFaceNormalInFloorMountBody, adMatingFaceNormalInFloorMountBody) Then
    '                            adRotationMatrix = FnGetRotationMatrixOfGivenTwoOrthogonalVectors(adOrthogonalFaceNormalInFloorMountBody, adMatingFaceNormalInFloorMountBody)
    '                            If Not adRotationMatrix Is Nothing Then
    '                                'CHeck if the rotation matrix is consistant
    '                                If FnCheckIfRotationMatrixIsConsistent(adRotationMatrix) Then
    '                                    'Delete the Pre-existing B_FLOOR_MOUNT_VIEW
    '                                    If FnChkIfModelingViewPresent(objPart, B_FLOOR_MOUNT_VIEW) Then
    '                                        sRefreshLayout(objPart, B_FLOOR_MOUNT_VIEW)
    '                                        sDeleteModellingView(objPart, B_FLOOR_MOUNT_VIEW)
    '                                        sSetDefaultLayout(objPart)
    '                                    End If
    '                                    'Create the modelling view B_FLOOR_MOUNT_VIEW
    '                                    sCreateCustomModellingView(objPart, adRotationMatrix(0), adRotationMatrix(1), adRotationMatrix(2), _
    '                                                                adRotationMatrix(3), adRotationMatrix(4), adRotationMatrix(5), _
    '                                                                adRotationMatrix(6), adRotationMatrix(7), adRotationMatrix(8), _
    '                                                                B_FLOOR_MOUNT_VIEW, TEMP_SCALE)
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End If

    'End Sub
    'Function to filter all the body and collect only Floor Mount body in a frame/Base component
    'Function FnCollectAllFloorMountBodys(objPart As Part) As Body()
    '    Dim aoFlatBodys() As DisplayableObject = Nothing
    '    Dim aoFloorMountBody() As Body = Nothing

    '    'Create a dictionary to carry all the dowel hole parameters
    '    sPopulateDowelHoleDictionary()

    '    'Get all the body in part which has the shape attribute as flat
    '    aoFlatBodys = FnGetBodyObjectByAttributes(objPart, SHAPE, FLAT)
    '    If Not aoFlatBodys Is Nothing Then
    '        For Each objFlatBody As Body In aoFlatBodys
    '            'Check if the body has no dowel hole
    '            If Not FnCheckIfBodyHasDowelHole(objPart, objFlatBody) Then
    '                'Check if the body has only one hole
    '                If FnCheckIfBodyHasOnlyOneHole(objPart, objFlatBody) Then
    '                    'Check if the QTY in body is greater than or equal to MINIMUM_NUM_OF_FLOOR_MOUNT_BODY
    '                    If FnGetStringUserAttribute(objFlatBody, QTY) >= MINIMUM_NUM_OF_FLOOR_MOUNT_BODY Then
    '                        If aoFloorMountBody Is Nothing Then
    '                            ReDim Preserve aoFloorMountBody(0)
    '                            aoFloorMountBody(0) = objFlatBody
    '                        Else
    '                            ReDim Preserve aoFloorMountBody(UBound(aoFloorMountBody) + 1)
    '                            aoFloorMountBody(UBound(aoFloorMountBody)) = objFlatBody
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next
    '    End If
    '    FnCollectAllFloorMountBodys = aoFloorMountBody
    'End Function

    'Get Body objects by attributes
    Public Function FnGetBodyObjectByAttributes(objPart As Part, sAttrTitle As String, sAttrValue As String,
                                          Optional ByVal sAttrType As String = "String") As DisplayableObject()
        Dim arule As String = ""
        If sAttrType = "String" Then
            arule = "mqc_selectEntitiesWithFilters(" &
                "select_by_entity_type, { SOLID_BODY }, " &
                "select_by_attribute, {{ String, """ & sAttrTitle & """, """ & sAttrValue & """, """ & sAttrValue & """ }}, " &
                "ignore_entity_occurrence?, False)"
        End If
        Dim ruleName As String = "VECTRA_Rule"
        objPart.RuleManager.CreateDynamicRule("root:", ruleName, "Any", arule, "")
        Dim theObj As Object = objPart.RuleManager.Evaluate("root:" & ruleName & ":")
        objPart.RuleManager.DeleteDynamicRule("root:", ruleName)

        If Not theObj Is Nothing Then
            Dim found() As DisplayableObject = ConvertTagListToDisplayableObject(theObj)
            FnGetBodyObjectByAttributes = found
        Else
            FnGetBodyObjectByAttributes = Nothing
        End If
    End Function

    Sub sPopulateDowelHoleDictionary()
        'Key - Diameter
        'Value - Depth : {Attribute,Hole Type}
        Dim dDepthDict1 As Dictionary(Of String, String()) = Nothing
        Dim dDepthDict2 As Dictionary(Of String, String()) = Nothing
        Dim dDepthDict3 As Dictionary(Of String, String()) = Nothing
        Dim dDepthDict4 As Dictionary(Of String, String()) = Nothing
        Dim dDepthDict21 As Dictionary(Of String, String()) = Nothing
        Dim dDepthDict22 As Dictionary(Of String, String()) = Nothing
        Dim dDepthDict23 As Dictionary(Of String, String()) = Nothing


        dDepthDict1 = New Dictionary(Of String, String())
        dDepthDict2 = New Dictionary(Of String, String())
        dDepthDict3 = New Dictionary(Of String, String())
        dDepthDict4 = New Dictionary(Of String, String())
        dDepthDict21 = New Dictionary(Of String, String())
        dDepthDict22 = New Dictionary(Of String, String())
        dDepthDict23 = New Dictionary(Of String, String())

        _dictDowelHoles = New Dictionary(Of String, Dictionary(Of String, String()))

        '###########################################  DOWEL HOLES #######################################################
        dDepthDict1.Add("0", {"DIA 6 H6", DOWEL_HOLE})
        dDepthDict1.Add("9", {"DIA 6 H6 9 DEEP", DOWEL_HOLE})
        _dictDowelHoles.Add("6", dDepthDict1)

        dDepthDict2.Add("0", {"DIA 8 H6", DOWEL_HOLE})
        dDepthDict2.Add("12", {"DIA 8 H6 12 DEEP", DOWEL_HOLE})
        _dictDowelHoles.Add("8", dDepthDict2)

        dDepthDict3.Add("0", {"DIA 10 H6", DOWEL_HOLE})
        dDepthDict3.Add("15", {"DIA 10 H6 15 DEEP", DOWEL_HOLE})
        _dictDowelHoles.Add("10", dDepthDict3)

        dDepthDict4.Add("0", {"DIA 12 H6", DOWEL_HOLE})
        dDepthDict4.Add("18", {"DIA 12 H6 18 DEEP", DOWEL_HOLE})
        _dictDowelHoles.Add("12", dDepthDict4)

        dDepthDict22.Add("0", {"DIA 16 H6", DOWEL_HOLE})
        dDepthDict22.Add("24", {"DIA 16 H6 24 DEEP", DOWEL_HOLE})
        _dictDowelHoles.Add("16", dDepthDict22)

        dDepthDict21.Add("0", {"DIA 18 H6", DOWEL_HOLE})
        dDepthDict21.Add("27", {"DIA 18 H6 27 DEEP", DOWEL_HOLE})
        _dictDowelHoles.Add("18", dDepthDict21)

        dDepthDict23.Add("0", {"DIA 20 H6", DOWEL_HOLE})
        dDepthDict23.Add("30", {"DIA 20 H6 30 DEEP", DOWEL_HOLE})
        _dictDowelHoles.Add("20", dDepthDict23)
    End Sub
    'Function to check if the hole is a dowel hole
    Function FnCheckIfBodyHasDowelHole(objPart As Part, objBody As Body) As Boolean
        Dim dHoleParameters() As Double = Nothing

        If Not objBody Is Nothing Then
            For Each objCylFace As Face In objBody.GetFaces()
                If objCylFace.SolidFaceType = Face.FaceType.Cylindrical Then
                    If FnCheckIftheFaceIsAHoleFace(objPart, objCylFace) Then
                        dHoleParameters = FnDetermineDepthandDiaforSimpleHoles(objPart, objCylFace)
                        'Diameter will be 0.0 if it is not a hole face
                        If dHoleParameters(0) <> 0.0 Then
                            'Check if the hole parameters match the values with the populated dictionary value
                            If _dictDowelHoles.ContainsKey(dHoleParameters(0).ToString) Then
                                FnCheckIfBodyHasDowelHole = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        End If
        FnCheckIfBodyHasDowelHole = False
    End Function
    'Function to check if the body has only one hole face
    Function FnCheckIfBodyHasOnlyOneHole(objPart As Part, objBody As Body) As Boolean
        Dim iNumHole As Integer = 0

        If Not objBody Is Nothing Then
            For Each objCylFace As Face In objBody.GetFaces()
                If objCylFace.SolidFaceType = Face.FaceType.Cylindrical Then
                    If FnCheckIftheFaceIsAHoleFace(objPart, objCylFace) Then
                        iNumHole = iNumHole + 1
                    End If
                End If
            Next
        End If
        If iNumHole <> 0 Then
            If iNumHole = 1 Then
                FnCheckIfBodyHasOnlyOneHole = True
            Else
                FnCheckIfBodyHasOnlyOneHole = False
            End If
        Else
            FnCheckIfBodyHasOnlyOneHole = False
        End If
    End Function
    'To check if a cylindrical face is a hole face
    Public Function FnCheckIftheFaceIsAHoleFace(ByVal objPart As Part, ByVal ObjCylFace As Face) As Boolean
        Dim objEdgeVrt1 As Point3d = Nothing
        Dim objEdgeVrt2 As Point3d = Nothing
        For Each objEdge As Edge In ObjCylFace.GetEdges()
            If objEdge.SolidEdgeType = Edge.EdgeType.Circular Then
                objEdge.GetVertices(objEdgeVrt1, objEdgeVrt2)
                If objEdgeVrt1.Equals(objEdgeVrt2) Then
                    If Not FnIsEdgeConvex(objPart, objEdge) Then
                        FnCheckIftheFaceIsAHoleFace = True
                        Exit Function
                    End If
                    'For Plasma cut holes will have a slit for plasma machine to enter
                    'The included hole angle should be more than 270 degrees
                Else
                    If ((FnGetArcInfo(objEdge.Tag).limits(1) - FnGetArcInfo(objEdge.Tag).limits(0)) * 180) / PI > 270 Then
                        If Not FnIsEdgeConvex(objPart, objEdge) Then
                            FnCheckIftheFaceIsAHoleFace = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
        FnCheckIftheFaceIsAHoleFace = False
    End Function
    Function FnGetArcInfo(ByVal edgeTag As Tag) As NXOpen.UF.UFEval.Arc
        Dim theUFSession As UFSession = UFSession.GetUFSession()
        Dim arc_evaluator As System.IntPtr
        Dim arc_data As NXOpen.UF.UFEval.Arc = Nothing

        theUFSession.Eval.Initialize(edgeTag, arc_evaluator)
        theUFSession.Eval.AskArc(arc_evaluator, arc_data)
        theUFSession.Eval.Free(arc_evaluator)

        FnGetArcInfo = arc_data
    End Function
    'For blind and through simple holes
    Public Function FnDetermineDepthandDiaforSimpleHoles(objPart As Part, ByVal objHoleFace As Face) As Double()
        Dim bHoleEdge As Boolean = False
        Dim dHoleDia As Double = 0.0
        Dim dHoleDepth As Double = 0.0
        Dim bBlindHole As Boolean = False
        Dim dEdgeCenter1 As Double() = Nothing
        Dim dEdgeCenter2 As Double() = Nothing
        Dim aoHoleEdges As Edge() = Nothing
        Dim bCounterSunkHole As Boolean = False
        Dim bCounterBoreHole As Boolean = False
        Dim objAssocConEdge1 As Edge = Nothing
        Dim objAssocConEdge2 As Edge = Nothing
        Dim bFirstConicalFaceFound As Boolean = False
        Dim bSecondConicalFaceFound As Boolean = False
        Dim iCountOfCircularEdges As Integer = 0
        'Special case of a Blind Hole with a Pin Hole at the bottom
        Dim objConnConicalFace As Face = Nothing
        Dim bProbableBlindHoleWithPinHole As Boolean = False

        'Hole face can have 3 edges as there may be a intersecting hole
        'In case of intersecting hole, there is hole on the hole face
        'The Edge produced by an intersecting hole is an intersection edge
        For Each objEdge As Edge In objHoleFace.GetEdges()
            If objEdge.SolidEdgeType = Edge.EdgeType.Circular Then
                If aoHoleEdges Is Nothing Then
                    ReDim Preserve aoHoleEdges(0)
                    aoHoleEdges(0) = objEdge
                Else
                    ReDim Preserve aoHoleEdges(UBound(aoHoleEdges) + 1)
                    aoHoleEdges(UBound(aoHoleEdges)) = objEdge
                End If
            End If
        Next

        If Not aoHoleEdges Is Nothing Then
            If aoHoleEdges(0).SolidEdgeType = Edge.EdgeType.Circular Then
                For Each objConnectedFace As Face In aoHoleEdges(0).GetFaces()
                    If objConnectedFace.SolidFaceType = Face.FaceType.Conical Then
                        'The countersunk hole will also have conical face , 
                        'however in case of the counter sunk hole the conical face will have 2 edges, hence those should not be considered
                        If UBound(objConnectedFace.GetEdges().ToArray) = 0 Then
                            bBlindHole = True
                            Exit For
                        Else
                            'Check if the conical faces have 2 edges (Chamfer cut on through holes)
                            If UBound(objConnectedFace.GetEdges().ToArray) = 1 Then
                                objConnConicalFace = objConnectedFace
                                bProbableBlindHoleWithPinHole = True

                                'CODE ADDED - 6/25/16 - Amitabh - Distinguish between chamfer cut and counter-sink holes
                                'Measure the distance between the two edges of the conical face
                                'Chamfer cut on through holes will have the distance between the 2 edges forming the conical face <2mm
                                If FnComputeMinDistance(objPart, objConnectedFace.GetEdges().ToArray(0), objConnectedFace.GetEdges().ToArray(1)) > 2 Then
                                    bFirstConicalFaceFound = True
                                Else
                                    bFirstConicalFaceFound = False
                                End If
                                Exit For
                            End If
                        End If
                    End If
                Next

                If Not bBlindHole Then
                    'The other hole edge might be an intersection edge in which case the intersection edge should not be considered
                    If UBound(aoHoleEdges) > 0 Then
                        For Each objConnectedFace As Face In aoHoleEdges(1).GetFaces()
                            If objConnectedFace.SolidFaceType = Face.FaceType.Conical Then
                                If UBound(objConnectedFace.GetEdges().ToArray) = 0 Then
                                    bBlindHole = True
                                    Exit For
                                Else
                                    'Check if the conical faces have 2 edges (Chamfer cut on through holes)
                                    If UBound(objConnectedFace.GetEdges().ToArray) = 1 Then
                                        'Blind hole with a Pin hole at the bottom will just have one connected conical face with two circular edges
                                        If Not bProbableBlindHoleWithPinHole Then
                                            objConnConicalFace = objConnectedFace
                                            bProbableBlindHoleWithPinHole = True
                                        Else
                                            'As both the ends of the cylindrical face are connected to conical edges hence it cannot be a blind hole with Pin hole at the bottom
                                            bProbableBlindHoleWithPinHole = False
                                        End If
                                        'CODE ADDED - 6/25/16 - Amitabh - Distinguish between chamfer cut and counter-sink holes
                                        'Measure the distance between the two edges of the conical face
                                        'Chamfer cut on through holes will have the distance between the 2 edges forming the conical face <2mm
                                        If FnComputeMinDistance(objPart, objConnectedFace.GetEdges().ToArray(0), objConnectedFace.GetEdges().ToArray(1)) > 2 Then
                                            bSecondConicalFaceFound = True
                                        Else
                                            bSecondConicalFaceFound = False
                                        End If
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If

                'It is only a counter sunk hole  if one of the face connected to the cylindrical face is conical.
                'In case of chamfer cut holes both the faces connected to the cylindrical face will be conical
                If ((Not bSecondConicalFaceFound) And bFirstConicalFaceFound) Or (bSecondConicalFaceFound And (Not bFirstConicalFaceFound)) Then
                    'Check if the hole is a blind hole with a pin hole at the bottom or a Type 3 Blind Hole
                    'In that case the smaller of the two circular edges of the connected conical face should also lie on a planar face
                    If bProbableBlindHoleWithPinHole Then
                        If FnCheckForBlindHoleWithPinHoleAtBottom(objConnConicalFace) Or FnCheckForType3BlindHole(objConnConicalFace) Then
                            bBlindHole = True
                        Else
                            bCounterSunkHole = True
                        End If
                    Else
                        bCounterSunkHole = True
                    End If
                End If

                If (Not bCounterSunkHole) Or (Not bBlindHole) Then
                    'Check if it is Counter bore hole in which case a planar face will have two concentric circular edges
                    For Each objConnFace As Face In aoHoleEdges(0).GetFaces()
                        If objConnFace.SolidFaceType = Face.FaceType.Planar Then
                            If UBound(objConnFace.GetEdges().ToArray) = 1 Then
                                If objConnFace.GetEdges().ToArray(0).SolidEdgeType = Edge.EdgeType.Circular And
                                    objConnFace.GetEdges().ToArray(1).SolidEdgeType = Edge.EdgeType.Circular Then
                                    'Check if the edge centres match
                                    dEdgeCenter1 = FnGetArcInfo(objConnFace.GetEdges().ToArray(0).Tag).center
                                    dEdgeCenter2 = FnGetArcInfo(objConnFace.GetEdges().ToArray(1).Tag).center
                                    'Check if they are concentric
                                    If dEdgeCenter1(0) = dEdgeCenter2(0) And dEdgeCenter1(1) = dEdgeCenter2(1) And
                                        dEdgeCenter1(2) = dEdgeCenter2(2) Then
                                        bCounterBoreHole = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                        If bCounterBoreHole Then
                            Exit For
                        End If
                    Next
                End If

                If Not bCounterBoreHole Or Not bCounterSunkHole Or Not bBlindHole Then
                    'Check if it is Counter bore hole in which case a planar face will have two concentric circular edges
                    'The other hole edge might be an intersection edge in which case the intersection edge should not be considered
                    If UBound(aoHoleEdges) > 0 Then
                        For Each objConnFace As Face In aoHoleEdges(1).GetFaces()
                            If objConnFace.SolidFaceType = Face.FaceType.Planar Then
                                If UBound(objConnFace.GetEdges().ToArray) = 1 Then
                                    If objConnFace.GetEdges().ToArray(0).SolidEdgeType = Edge.EdgeType.Circular And
                                    objConnFace.GetEdges().ToArray(1).SolidEdgeType = Edge.EdgeType.Circular Then
                                        'Check if the edge centres match
                                        dEdgeCenter1 = FnGetArcInfo(objConnFace.GetEdges().ToArray(0).Tag).center
                                        dEdgeCenter2 = FnGetArcInfo(objConnFace.GetEdges().ToArray(1).Tag).center
                                        'Check if they are concentric
                                        If dEdgeCenter1(0) = dEdgeCenter2(0) And dEdgeCenter1(1) = dEdgeCenter2(1) And
                                            dEdgeCenter1(2) = dEdgeCenter2(2) Then
                                            bCounterBoreHole = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                            If bCounterBoreHole Then
                                Exit For
                            End If
                        Next
                    End If
                End If

            End If
            If bBlindHole Then
                dEdgeCenter1 = FnGetArcInfo(aoHoleEdges(0).Tag).center
                dEdgeCenter2 = FnGetArcInfo(aoHoleEdges(1).Tag).center
                dHoleDepth = Sqrt(Pow(dEdgeCenter1(0) - dEdgeCenter2(0), 2) + Pow(dEdgeCenter1(1) - dEdgeCenter2(1), 2) +
                                Pow(dEdgeCenter1(2) - dEdgeCenter2(2), 2))
            End If
            If Not bCounterSunkHole Then
                'Get the edge radius
                dHoleDia = (FnGetArcInfo(aoHoleEdges(0).Tag).radius) * 2
                FnDetermineDepthandDiaforSimpleHoles = {dHoleDia, dHoleDepth}
            ElseIf bCounterBoreHole Or bCounterSunkHole Then
                'Retun a zero value to indicate that it is not a simple hole
                FnDetermineDepthandDiaforSimpleHoles = {0.0, 0.0}
            End If
        Else
            FnDetermineDepthandDiaforSimpleHoles = {dHoleDia, dHoleDepth}
        End If
    End Function
    'To ascertain special type of blind holes with a pin hole at the bottom
    Public Function FnCheckForBlindHoleWithPinHoleAtBottom(objConicalFace As Face) As Boolean
        Dim dEdge1Dia As Double = 0.0
        Dim dEdge2Dia As Double = 0.0
        'The conical face should have excatly two connecting edges
        If UBound(objConicalFace.GetEdges().ToArray) = 1 Then
            'Get the dia of both the connected edges
            If objConicalFace.GetEdges().ToArray(0).SolidEdgeType = Edge.EdgeType.Circular Then
                dEdge1Dia = FnGetArcInfo(objConicalFace.GetEdges().ToArray(0).Tag).radius * 2
            End If
            If objConicalFace.GetEdges().ToArray(1).SolidEdgeType = Edge.EdgeType.Circular Then
                dEdge2Dia = FnGetArcInfo(objConicalFace.GetEdges().ToArray(1).Tag).radius * 2
            End If
            If dEdge1Dia <> 0.0 And dEdge2Dia <> 0.0 Then
                If dEdge1Dia < dEdge2Dia Then
                    'The smaller dia edge should be connected to a planar face
                    'The bigger edge dia should be connected to a cylindrical face
                    For Each objConnFace As Face In objConicalFace.GetEdges().ToArray(0).GetFaces()
                        If objConnFace.SolidFaceType = Face.FaceType.Planar Then
                            For Each objConn2Face As Face In objConicalFace.GetEdges().ToArray(1).GetFaces()
                                If objConn2Face.SolidFaceType = Face.FaceType.Cylindrical Then
                                    FnCheckForBlindHoleWithPinHoleAtBottom = True
                                    Exit Function
                                End If
                            Next
                        End If
                    Next
                ElseIf dEdge2Dia < dEdge1Dia Then
                    'The smaller dia edge should be connected to a planar face
                    'The bigger edge dia should be connected to a cylindrical face
                    For Each objConnFace As Face In objConicalFace.GetEdges().ToArray(1).GetFaces()
                        If objConnFace.SolidFaceType = Face.FaceType.Planar Then
                            For Each objConn2Face As Face In objConicalFace.GetEdges().ToArray(0).GetFaces()
                                If objConn2Face.SolidFaceType = Face.FaceType.Cylindrical Then
                                    FnCheckForBlindHoleWithPinHoleAtBottom = True
                                    Exit Function
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        End If
        FnCheckForBlindHoleWithPinHoleAtBottom = False
    End Function
    'Type 3 Blind Hole 
    'The connical face in the blind hole will have a circular edge and an intersection edge
    Public Function FnCheckForType3BlindHole(objConicalFace As Face) As Boolean
        'The conical face should have excatly two connecting edges
        If UBound(objConicalFace.GetEdges().ToArray) = 1 Then
            'Get the dia of both the connected edges
            If objConicalFace.GetEdges().ToArray(0).SolidEdgeType = Edge.EdgeType.Circular Then
                If objConicalFace.GetEdges().ToArray(1).SolidEdgeType = Edge.EdgeType.Intersection Then
                    FnCheckForType3BlindHole = True
                    Exit Function
                End If
            End If
            If objConicalFace.GetEdges().ToArray(1).SolidEdgeType = Edge.EdgeType.Circular Then
                If objConicalFace.GetEdges().ToArray(0).SolidEdgeType = Edge.EdgeType.Intersection Then
                    FnCheckForType3BlindHole = True
                    Exit Function
                End If
            End If
        End If
        FnCheckForType3BlindHole = False
    End Function
    'Function to get the planar face which is mating to other solid body
    Function FnGetMatingFaceInASolidBody(objPart As Part, objBody As Body) As Face

        If Not objBody Is Nothing Then
            For Each objPlanarFace As Face In objBody.GetFaces()
                If objPlanarFace.SolidFaceType = Face.FaceType.Planar Then
                    If FnCheckIftheMatingFaceIsAHoleFace(objPart, objPlanarFace) Then
                        For Each objBodyToCompare As Body In objPart.Bodies()
                            If Not objBody Is objBodyToCompare Then
                                If objBodyToCompare.IsSolidBody Then
                                    'Get the Planar face in a objBody which is mating to any of the other solid body in a Frame / Base component
                                    If FnComputeMinDistance(objPart, objPlanarFace, objBodyToCompare) = 0.0 Then
                                        FnGetMatingFaceInASolidBody = objPlanarFace
                                        Exit Function
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End If
        FnGetMatingFaceInASolidBody = Nothing
    End Function
    'To check if the mating face has a slot or hole
    'To check if a cylindrical face is a hole face
    Public Function FnCheckIftheMatingFaceIsAHoleFace(ByVal objPart As Part, ByVal ObjPlanarFace As Face) As Boolean
        Dim objEdgeVrt1 As Point3d = Nothing
        Dim objEdgeVrt2 As Point3d = Nothing
        For Each objEdge As Edge In ObjPlanarFace.GetEdges()
            If objEdge.SolidEdgeType = Edge.EdgeType.Circular Then
                objEdge.GetVertices(objEdgeVrt1, objEdgeVrt2)
                If objEdgeVrt1.Equals(objEdgeVrt2) Then
                    If Not FnIsEdgeConvex(objPart, objEdge) Then
                        FnCheckIftheMatingFaceIsAHoleFace = True
                        Exit Function
                    End If
                    'For Plasma cut holes will have a slit for plasma machine to enter
                    'The included hole angle should be more than 270 degrees
                Else
                    If ((FnGetArcInfo(objEdge.Tag).limits(1) - FnGetArcInfo(objEdge.Tag).limits(0)) * 180) / PI > 270 Then
                        If Not FnIsEdgeConvex(objPart, objEdge) Then
                            FnCheckIftheMatingFaceIsAHoleFace = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next

        ''Go for slot assesment if no holes are found
        'If FnCheckIfSlotArePresentOnTheMatingFace(objPart, ObjPlanarFace) Then
        '    FnCheckIftheMatingFaceIsAHoleOrASlotFace = True
        '    Exit Function
        'End If

        FnCheckIftheMatingFaceIsAHoleFace = False
    End Function

    Public Function FnGetFaceNormalVector(ByVal ObjFace As Face) As Double()
        Dim iFaceType As Integer = 0
        Dim adCenterPoint(2) As Double
        Dim adDir(2) As Double
        Dim adBox(5) As Double
        Dim dRadius As Double = 0.0
        Dim dRadData As Double = 0.0
        Dim iNormDir As Integer = 0
        'Dim adVector As Vector3d

        FnGetUFSession.Modl.AskFaceData(ObjFace.Tag, iFaceType, adCenterPoint, adDir, adBox, dRadius, dRadData, iNormDir)
        'adVector = New Vector3d(adDir(0), adDir(1), adDir(2))
        FnGetFaceNormalVector = adDir
    End Function

    'Function to get the smallest edge in a face
    Function FnGetSmallestEdgeInAFace(objPlannarFace As Face) As Edge
        Dim dSmallestEdgeLength As Double = 0.0
        Dim dEdgeLength As Double = 0.0
        Dim objSmallestEdge As Edge = Nothing

        If Not objPlannarFace Is Nothing Then
            For Each objLinearEdge As Edge In objPlannarFace.GetEdges()
                If objLinearEdge.SolidEdgeType = Edge.EdgeType.Linear Then
                    dEdgeLength = objLinearEdge.GetLength()
                    If dSmallestEdgeLength = 0.0 Then
                        dSmallestEdgeLength = dEdgeLength
                        objSmallestEdge = objLinearEdge
                    Else
                        If dEdgeLength < dSmallestEdgeLength Then
                            dSmallestEdgeLength = dEdgeLength
                            objSmallestEdge = objLinearEdge
                        End If
                    End If
                End If
            Next
        End If
        FnGetSmallestEdgeInAFace = objSmallestEdge
    End Function

    'Function to get the orthogonal face to the reference face in a given Body
    Public Function FnGetOrthogonalFaceToRefFaceInABody(ByVal objBody As Body, ByVal objRefFace As Face) As Face
        Dim adFaceNormalOfRefFace() As Double = Nothing
        Dim adFaceNormalOfOrthogonalFace() As Double = Nothing
        If Not objBody Is Nothing Then
            For Each objPlanarFace As Face In objBody.GetFaces()
                If objPlanarFace.SolidFaceType = Face.FaceType.Planar Then
                    adFaceNormalOfRefFace = FnGetFaceNormalVector(objRefFace)
                    adFaceNormalOfOrthogonalFace = FnGetFaceNormalVector(objPlanarFace)
                    If FnCheckOrthogalityOfTwoVectors(adFaceNormalOfRefFace, adFaceNormalOfOrthogonalFace) Then
                        FnGetOrthogonalFaceToRefFaceInABody = objPlanarFace
                        Exit Function
                    End If
                End If
            Next
        End If
        FnGetOrthogonalFaceToRefFaceInABody = Nothing
    End Function

    'Function to Change the View Type (Either Modeling or Drafting View)
    '1 = Modeling VIew
    '2 = DrawingView
    Sub sChangeApplication(iApplicationType As Integer)
        Dim viewtype As Integer = 0
        FnGetUFSession.Draw.AskDisplayState(viewtype)
        'Change to modeling
        If iApplicationType = 1 Then
            If viewtype = 2 Then
                FnGetNxSession.ApplicationSwitchImmediate("UG_APP_MODELING")
                FnGetUFSession.Draw.SetDisplayState(1)
            End If
            'Change to drafting
        ElseIf iApplicationType = 2 Then
            If viewtype = 1 Then
                FnGetNxSession.ApplicationSwitchImmediate("UG_APP_DRAFTING")
                FnGetNxSession.Parts.Work.Drafting.EnterDraftingApplication()
                FnGetUFSession.Draw.SetDisplayState(2)
            End If
        End If
        'FnGetUFSession.Disp.SetDisplay(UFConstants.UF_DISP_SUPPRESS_DISPLAY)
    End Sub
    'Refresh model layout
    Sub sRefreshLayout(objPart As Part, sViewName As String)
        sChangeApplication(1)
        Dim objViewLayout As Layout = Nothing
        'Store the view layout correesponding to this view
        For Each objLayout As Layout In objPart.Layouts()
            For Each objView As ModelingView In objLayout.GetViews()
                If objView.Name.ToUpper = sViewName.ToUpper Then
                    objViewLayout = objLayout
                    Exit For
                End If
            Next
        Next

        'Display another layout other than the layout to which this view belongs to
        For Each objLayout As Layout In objPart.Layouts()
            If Not objLayout Is objViewLayout Then
                If Not objLayout.DisplayStatus Then
                    objLayout.Open()
                    Exit For
                End If
            End If
        Next

        'Display this view layout
        If Not objViewLayout Is Nothing Then
            If Not objViewLayout.DisplayStatus Then
                objViewLayout.Open()
            End If
        End If
    End Sub
    'CODE ADDED - 4/21/16 - Amitabh - Display correct view layout
    Sub sSetDefaultLayout(objPart As Part)
        'Default Layout
        Dim objDefaultLayout As Layout = CType(objPart.Layouts.FindObject("L1"), Layout)
        If Not objDefaultLayout Is Nothing Then
            If objDefaultLayout.DisplayStatus = False Then
                objDefaultLayout.Open()
            End If
        End If
    End Sub
    Public Sub sCreateCustomModellingView(ByVal objPart As Part, ByVal Xx As Double, ByVal Xy As Double, ByVal Xz As Double,
                                      ByVal Yx As Double, ByVal Yy As Double, ByVal Yz As Double, ByVal Zx As Double,
                                      ByVal Zy As Double, ByVal Zz As Double, ByVal sCustViewName As String, ByVal dScale As Double,
                                      Optional ByVal sDefaultRefView As String = "FRONT")
        Dim viewType As Integer = 0
        '1 = modeling view
        '2 = drawing view
        'other = error
        FnGetUFSession.Draw.AskDisplayState(viewType)

        'if drawing sheet shown, change to modeling view
        If viewType = 2 Then
            FnGetUFSession.Draw.SetDisplayState(1)
        End If

        Dim objView As View = Nothing
        Dim rotMatrix As Matrix3x3 = Nothing

        rotMatrix.Xx = Xx
        rotMatrix.Xy = Xy
        rotMatrix.Xz = Xz
        rotMatrix.Yx = Yx
        rotMatrix.Yy = Yy
        rotMatrix.Yz = Yz
        rotMatrix.Zx = Zx
        rotMatrix.Zy = Zy
        rotMatrix.Zz = Zz
        Dim translation As Point3d = New Point3d(0, 0, 0)

        'To catch the error if the view is not a model view
        If Not objPart.ModelingViews.WorkView Is Nothing Then
            objPart.ModelingViews.WorkView.SetRotationTranslationScale(rotMatrix, translation, dScale)
            objView = objPart.Views.SaveAs(objPart.ModelingViews.WorkView, sCustViewName, False, False)
        Else
            For Each objMdVw As ModelingView In objPart.ModelingViews
                If Not objMdVw Is Nothing Then
                    If objMdVw.Name.ToUpper = sDefaultRefView Then
                        sReplaceViewInLayout(objPart, objMdVw)
                        objMdVw.SetRotationTranslationScale(rotMatrix, translation, dScale)
                        objPart.Views.SaveAs(objMdVw, sCustViewName, False, False)
                        Exit For
                    End If
                End If
            Next
        End If

        'reset initial view state
        'If viewType = 1 Or viewType = 2 Then
        '    FnGetUFSession.Draw.SetDisplayState(viewType)
        'End If
    End Sub

    'Code added - SHanmugam - Jan-23-2017
    'Function to create standard B_FLOOR_MOUNT view for the frames

    Public Sub sCreateStandardFloorMountView(ByVal objPart As Part, ByVal sFloorMountViewName As String, ByVal sPrimaryViewName As String)

        Dim baseView As ModelingView = Nothing
        Dim currentView As ModelingView = Nothing
        Dim rotMatPrimaryView As Matrix3x3 = Nothing
        Dim origin As Point3d = Nothing
        Dim origin1 As Point3d = Nothing
        Dim vector As Vector3d
        Dim vector1 As Vector3d
        Dim iAngle As Integer = 0
        Dim objView As View = Nothing

        'Get the sPrimaryViewName as base view
        If Not sPrimaryViewName Is Nothing Then
            If FnChkIfModelingViewPresent(objPart, sPrimaryViewName) Then
                baseView = objPart.ModelingViews.FindObject(sPrimaryViewName)
            End If
            currentView = objPart.ModelingViews.WorkView
            If Not currentView Is baseView Then
                sReplaceViewInLayout(objPart, baseView)
            End If
            rotMatPrimaryView = FnGetRotationMatrix(objPart)
            origin = objPart.ModelingViews.WorkView.AbsoluteOrigin
            origin1 = New Point3d(origin.X, origin.Y, origin.Z)
            vector = New Vector3d(rotMatPrimaryView.Xx, rotMatPrimaryView.Xy, rotMatPrimaryView.Xz)
            vector1 = New Vector3d(vector.X, vector.Y, vector.Z)
            iAngle = -90

            FnViewType(objPart)

            If objPart.ModelingViews.WorkView Is Nothing Then
                objPart.ModelingViews.WorkView.Rotate(origin1, vector1, iAngle)
                objView = objPart.Views.SaveAs(objPart.ModelingViews.WorkView, sFloorMountViewName, True, False)
            Else
                For Each objModelView As ModelingView In objPart.ModelingViews
                    If objModelView.Name.ToUpper = sPrimaryViewName Then
                        sReplaceViewInLayout(objPart, objModelView)
                        objModelView.Rotate(origin1, vector1, iAngle)
                        objModelView.Fit()
                        objPart.Views.SaveAs(objModelView, sFloorMountViewName, True, False)
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub
    'Code added May-08-2017

    '****************************************************************************************************************************************************
    'Function to compute optimal Rotation Matrix
    'Description        : Function to compute optimal Rotation Matrix from the Part (alternate logic based on rank and 
    '                     cummulative face area)
    'Function Name      : FnComputeOptimalRotationMatrixAlt
    'Input Parameter    : objPart
    'Output Parameter   : adOptimalRotationMatrix(8)
    '****************************************************************************************************************************************************
    Function FnComputeOptimalRotationMatrixAlt(ByVal objBody As Body) As Double()
        'Steps
        'This function only considers Planar faces
        '1. Find a pair of orthogonal faces and figure out the 3rd direction using cross product. Obtain a rotation matrix
        '2. For each orientation determine the number of peripheral faces aligned in X, Y and Z directions to the obtained rotation 
        '   matrix directions.
        '3. Compute Rank for this oreintation (X Aligned Peripheral Faces Present = 1, X Aligned Peripheral Faces Not Present = 0, 
        '   similarly for Y and Z as well
        '4. Possible (Min Rank = 2 and Max Rank = 3)
        '5. Compute the volume of bounding box 
        '6. Priority is given to Rank
        '7. In case of a tie, priority is given to the minimum bounding box volume
        '8. Output is optimal rotation matrix
        Dim adVecObjFace As Double() = Nothing
        Dim adVecObjFaceToCompare As Double() = Nothing
        Dim aoFacesChecked() As String = Nothing
        Dim sFacePairToBeChecked As String = Nothing
        Dim bIsFacePairAlreadyChecked As Boolean = False
        Dim adRotationMatrix(8) As Double
        Dim adOptimalRotationMatrix() As Double = Nothing
        Dim dictOfRotMat As Dictionary(Of Double, Double()) = Nothing
        Dim aoAllMachinedFaceWithHoles As Face() = Nothing
        Dim bIsMachiningFaceFound As Boolean = False
        Dim sShapeValue As String = ""
        Dim objAttNX As NXObject = Nothing

        dictOfRotMat = New Dictionary(Of Double, Double())
        'Clean memory
        _aoStructBodyOrientationInfo = Nothing
        _iOrientationIndex = 0


        'Code added May-10-2018
        'New logic discussed between Shanmugam and Pradeep.
        'We should consider the machined face which has holes for B_LCS computation.
        'Collect all the Machined face which has holes corresponding to this given solid body
        'Referring component F56500103928801260001
        aoAllMachinedFaceWithHoles = FnCollectMachinedFaceWithHoles(objBody)
        If Not aoAllMachinedFaceWithHoles Is Nothing Then
            bIsMachiningFaceFound = True
        Else
            bIsMachiningFaceFound = False
        End If

        If Not objBody Is Nothing Then
            If objBody.IsSolidBody And Not objBody.IsBlanked Then
                For Each objFace As Face In objBody.GetFaces()
                    If objFace.SolidFaceType = Face.FaceType.Planar Then
                        If FnChkIfPlanarFaceHasAtleastOneLinearEdge(objFace) Then
                            If (FnGetStringUserAttribute(objFace, NC_Contact_FACE_ATTRIBUTE) = "") Then
                                For Each objFaceToCompare As Face In objBody.GetFaces()
                                    If objFaceToCompare.SolidFaceType = Face.FaceType.Planar Then
                                        If FnChkIfPlanarFaceHasAtleastOneLinearEdge(objFaceToCompare) Then
                                            If (FnGetStringUserAttribute(objFaceToCompare, NC_Contact_FACE_ATTRIBUTE) = "") Then
                                                If objFace.Tag <> objFaceToCompare.Tag Then
                                                    sFacePairToBeChecked = objFace.Tag & "_" & objFaceToCompare.Tag
                                                    'Check if the face pairs are already checked or not.
                                                    If Not aoFacesChecked Is Nothing Then
                                                        bIsFacePairAlreadyChecked = False
                                                        If aoFacesChecked.Contains(sFacePairToBeChecked) Then
                                                            bIsFacePairAlreadyChecked = True
                                                        End If
                                                    End If
                                                    'Proceed only when the sFacePairtobechecked is not repeated
                                                    If Not bIsFacePairAlreadyChecked Then
                                                        adVecObjFace = FnGetFaceNormalVector(objFace)
                                                        adVecObjFaceToCompare = FnGetFaceNormalVector(objFaceToCompare)
                                                        bIsFacePairAlreadyChecked = True
                                                        If FnCheckOrthogalityOfTwoVectors(adVecObjFace, adVecObjFaceToCompare) Then
                                                            'when both vectors are orthogonal, get the rotational matrix
                                                            'adRotationMatrix = FnGetRotationMatrixOfGivenTwoOrthogonalVectors(adVecObjFace, adVecObjFaceToCompare)
                                                            adRotationMatrix = FnCreateRotationMatrixFromTwoVec(adVecObjFace, adVecObjFaceToCompare)
                                                            'Check if the rotation matrix obtained is consistent.
                                                            If FnCheckIfRotationMatrixIsConsistent(adRotationMatrix) Then
                                                                'Code added Oct-12-2017
                                                                'Process only unique rotation matrix
                                                                'Check if the rotation matrix is unique.
                                                                'Component mb663276 took 2hours without this validation check. To minimize the number of computation, we have added this check.
                                                                If FnCheckIfMatrixIsUnique(adRotationMatrix, dictOfRotMat) Then
                                                                    sCreateModelViewWithTrailOrentation(FnGetNxSession.Parts.Work, adRotationMatrix)
                                                                    'Populate orientation info for this orientation
                                                                    sPopulateOrientationInfo(objBody, adRotationMatrix, aoAllMachinedFaceWithHoles)
                                                                    dictOfRotMat.Add(dictOfRotMat.Count, adRotationMatrix)
                                                                End If
                                                            End If
                                                        End If
                                                        'Add the face to the FacePair checked
                                                        If bIsFacePairAlreadyChecked Then
                                                            If aoFacesChecked Is Nothing Then
                                                                ReDim Preserve aoFacesChecked(1)
                                                                aoFacesChecked(0) = objFace.Tag & "_" & objFaceToCompare.Tag
                                                                aoFacesChecked(1) = objFaceToCompare.Tag & "_" & objFace.Tag
                                                            Else
                                                                ReDim Preserve aoFacesChecked(UBound(aoFacesChecked) + 2)
                                                                aoFacesChecked(UBound(aoFacesChecked) - 1) = objFace.Tag & "_" & objFaceToCompare.Tag
                                                                aoFacesChecked(UBound(aoFacesChecked)) = objFaceToCompare.Tag & "_" & objFace.Tag
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
                'Validation added on May-26-2017
                If Not _aoStructBodyOrientationInfo Is Nothing Then
                    'Code modified on Nov-14-2018
                    'For TUBG's there is a seperate logic used to identify the optimal rotation matrix
                    If (_sOemName = DAIMLER_OEM_NAME) Or (_sOemName = FIAT_OEM_NAME) Then
                        If Not objBody.OwningComponent Is Nothing Then
                            objAttNX = objBody.OwningComponent
                        Else
                            objAttNX = objBody
                        End If
                    ElseIf (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                        objAttNX = objBody
                    End If
                    sShapeValue = FnGetStringUserAttribute(objAttNX, _SHAPE_ATTR_NAME)
                    If sShapeValue <> "" Then
                        If sShapeValue.ToUpper.Contains("TUBG") Then
                            adOptimalRotationMatrix = FnDetermineOptimalRotationMatrixForTubg(objBody, _aoStructBodyOrientationInfo)
                        Else
                            'Identify the optimal matrix
                            adOptimalRotationMatrix = FnDetermineOptimalMatrix(_aoStructBodyOrientationInfo, bIsMachiningFaceFound)
                        End If
                    Else
                        'Identify the optimal matrix
                        adOptimalRotationMatrix = FnDetermineOptimalMatrix(_aoStructBodyOrientationInfo, bIsMachiningFaceFound)
                    End If
                Else
                    adOptimalRotationMatrix = Nothing
                End If
            End If
        End If
        FnComputeOptimalRotationMatrixAlt = adOptimalRotationMatrix

    End Function

    Sub sCreateModelViewWithTrailOrentation(ObjPart As Part, adRotationMatrix As Double())
        'TODO-
        Dim mTrailOrientation As Matrix3x3 = Nothing
        Dim sModelViewName As String = "'"
        sModelViewName = "B_LCS_" & iTrailOrentationCount
        mTrailOrientation = FnConvertRotMatValueToMatrix(adRotationMatrix)
        If FnChkIfModelingViewPresent(ObjPart, sModelViewName) Then
            sRefreshLayout(ObjPart, sModelViewName)
            sDeleteModellingView(ObjPart, sModelViewName)
            sSetDefaultLayout(ObjPart)
        End If
        sCreateModelingView(ObjPart, mTrailOrientation, sModelViewName)
        iTrailOrentationCount = iTrailOrentationCount + 1
    End Sub
    'COde added May-22-2017
    Sub sPopulateOrientationForWeldment(objPart As Part)

        Dim aoMachinedFace() As DisplayableObject = Nothing
        Dim aoAllMachinedFace() As Face = Nothing
        Dim aoOrthoFace() As Face = Nothing
        Dim adFaceVec() As Double = Nothing
        Dim adOrthoFaceVec() As Double = Nothing
        Dim adRotationMatrix() As Double = Nothing
        Dim aoFacesChecked() As String = Nothing
        Dim sFacePairToBeChecked As String = Nothing
        Dim bIsFacePairAlreadyChecked As Boolean = False
        Dim adOptimalRotMat() As Double = Nothing
        Dim aoAllMachinedFaceWithHoles() As Face = Nothing

        'Collect all the machined face in the part
        aoMachinedFace = FnGetFaceObjectByAttributes(objPart, _FINISHTOLERANCE_ATTR_NAME, _FINISH_TOL_VALUE2)
        'COnvert displayable object to face
        aoAllMachinedFace = FnConvertDisplayObjectToFace(aoMachinedFace)

        'Code added May-10-2018
        'New logic added : Consider only the machined face which has holes on it.
        aoAllMachinedFaceWithHoles = FnCollectFaceWhichHasHolesOnIt(aoAllMachinedFace)

        If Not aoAllMachinedFaceWithHoles Is Nothing Then
            'If there is atleast a machined face, populate orthogonal faces with respect these machined faces
            'Orthogonal face may be a machined face as well.
            'This is done to maximise the number of aligned machined faces to LCS (to ease manufacturing)
            ''_aoStructPartOrientationInfo = Nothing
            ''_iPartOrientationIndex = 0

            'Weldment has machined face
            For Each objFace As Face In aoAllMachinedFaceWithHoles
                If objFace.SolidFaceType = Face.FaceType.Planar Then
                    'Validation added, that the face shoule not be a NC Part COntact Face
                    'If Not (FnGetStringUserAttribute(objFace, NC_Contact_FACE_ATTRIBUTE) = NC_PCF_ATTR_VALUE) Then
                    If (FnGetStringUserAttribute(objFace, NC_Contact_FACE_ATTRIBUTE) = "") Then
                        aoOrthoFace = FnCollectOrthogonalPlanarFacesToRefFace(objPart, objFace)
                        If Not aoOrthoFace Is Nothing Then
                            For Each objOrthoFace As Face In aoOrthoFace
                                'Validation added, that the face shoule not be a NC Part COntact Face
                                'If Not (FnGetStringUserAttribute(objOrthoFace, NC_Contact_FACE_ATTRIBUTE) = NC_PCF_ATTR_VALUE) Then
                                If (FnGetStringUserAttribute(objOrthoFace, NC_Contact_FACE_ATTRIBUTE) = "") Then
                                    adFaceVec = FnGetFaceNormalVector(objFace)
                                    adOrthoFaceVec = FnGetFaceNormalVector(objOrthoFace)
                                    'Check if both the faces normal direction aligns with any of the previously computed orientation
                                    'This is done to uniquely identify face pairs which will result in a unique rotation matrix
                                    If FnChkIfFaceVecIsUnique(adFaceVec, adOrthoFaceVec) Then
                                        'when both vectors are orthogonal, get the rotational matrix
                                        adRotationMatrix = FnCreateRotationMatrixFromTwoVec(adFaceVec, adOrthoFaceVec)
                                        'Check if the rotation matrix obtained is consistent.
                                        If FnCheckIfRotationMatrixIsConsistent(adRotationMatrix) Then
                                            'Populate orientation info for this orientation
                                            sPopulateWeldmentPartOrientationInfo(adRotationMatrix, aoAllMachinedFaceWithHoles)
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If

            Next
        Else
            'Weldment doesnot have a machined face
            For Each objBody As Body In objPart.Bodies()
                If objBody.IsSolidBody And (Not objBody.IsBlanked) Then
                    For Each objFace As Face In objBody.GetFaces()
                        If objFace.SolidFaceType = Face.FaceType.Planar Then
                            'Validation added, that the face shoule not be a NC Part COntact Face
                            'If Not (FnGetStringUserAttribute(objFace, NC_Contact_FACE_ATTRIBUTE) = NC_PCF_ATTR_VALUE) Then
                            If (FnGetStringUserAttribute(objFace, NC_Contact_FACE_ATTRIBUTE) = "") Then
                                For Each objBodyToCompare As Body In objPart.Bodies()
                                    If objBodyToCompare.IsSolidBody And (Not objBodyToCompare.IsBlanked) Then
                                        If Not objBody Is objBodyToCompare Then
                                            For Each objFaceToCompare As Face In objBodyToCompare.GetFaces()
                                                If objFaceToCompare.SolidFaceType = Face.FaceType.Planar Then
                                                    'Validation added, that the face shoule not be a NC Part COntact Face
                                                    'If Not (FnGetStringUserAttribute(objFaceToCompare, NC_Contact_FACE_ATTRIBUTE) = NC_PCF_ATTR_VALUE) Then
                                                    If (FnGetStringUserAttribute(objFaceToCompare, NC_Contact_FACE_ATTRIBUTE) = "") Then
                                                        If objFace.Tag <> objFaceToCompare.Tag Then
                                                            sFacePairToBeChecked = objFace.Tag & "_" & objFaceToCompare.Tag
                                                            'Check if the face pairs are already checked or not.
                                                            If Not aoFacesChecked Is Nothing Then
                                                                bIsFacePairAlreadyChecked = False
                                                                For Each sFacePair As String In aoFacesChecked
                                                                    If sFacePair = sFacePairToBeChecked Then
                                                                        bIsFacePairAlreadyChecked = True
                                                                        Exit For
                                                                    End If
                                                                Next
                                                            End If

                                                            'Proceed only when the sFacePairtobechecked is not repeated
                                                            If Not bIsFacePairAlreadyChecked Then
                                                                adFaceVec = FnGetFaceNormalVector(objFace)
                                                                adOrthoFaceVec = FnGetFaceNormalVector(objFaceToCompare)
                                                                'Check if both the faces normal direction aligns with any of the previously computed orientation
                                                                'This is done to uniquely identify face pairs which will result in a unique rotation matrix
                                                                If FnChkIfFaceVecIsUnique(adFaceVec, adOrthoFaceVec) Then
                                                                    If FnCheckOrthogalityOfTwoVectors(adFaceVec, adOrthoFaceVec) Then
                                                                        'when both vectors are orthogonal, get the rotational matrix
                                                                        adRotationMatrix = FnCreateRotationMatrixFromTwoVec(adFaceVec, adOrthoFaceVec)
                                                                        'Check if the rotation matrix obtained is consistent.
                                                                        If FnCheckIfRotationMatrixIsConsistent(adRotationMatrix) Then
                                                                            'Populate orientation info for this orientation
                                                                            sPopulateWeldmentPartOrientationInfo(adRotationMatrix, aoAllMachinedFaceWithHoles)
                                                                        End If
                                                                    End If
                                                                End If
                                                                'Add the face to the FacePair checked
                                                                If bIsFacePairAlreadyChecked Then
                                                                    If aoFacesChecked Is Nothing Then
                                                                        ReDim Preserve aoFacesChecked(1)
                                                                        aoFacesChecked(0) = objFace.Tag & "_" & objFaceToCompare.Tag
                                                                        aoFacesChecked(1) = objFaceToCompare.Tag & "_" & objFace.Tag
                                                                    Else
                                                                        ReDim Preserve aoFacesChecked(UBound(aoFacesChecked) + 2)
                                                                        aoFacesChecked(UBound(aoFacesChecked) - 1) = objFace.Tag & "_" & objFaceToCompare.Tag
                                                                        aoFacesChecked(UBound(aoFacesChecked)) = objFaceToCompare.Tag & "_" & objFace.Tag
                                                                    End If
                                                                End If
                                                            End If

                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            Next
        End If

    End Sub
    'Check if both the faces normal direction aligns with any of the previously computed orientation
    'This is done to uniquely identify face pairs which will result in a unique rotation matrix
    Function FnChkIfFaceVecIsUnique(adVec1() As Double, adVec2() As Double) As Boolean
        Dim adDCX(2) As Double
        Dim adDCY(2) As Double
        Dim adDCZ(2) As Double
        Dim bVec1MatchFound As Boolean = False
        Dim bVec2MatchFound As Boolean = False

        If _iPartOrientationIndex > 1 Then
            For iIndex As Integer = 0 To UBound(_aoStructPartOrientationInfo)
                bVec1MatchFound = False
                bVec2MatchFound = False
                adDCX(0) = _aoStructPartOrientationInfo(iIndex).xx
                adDCX(1) = _aoStructPartOrientationInfo(iIndex).xy
                adDCX(2) = _aoStructPartOrientationInfo(iIndex).xz

                adDCY(0) = _aoStructPartOrientationInfo(iIndex).yx
                adDCY(1) = _aoStructPartOrientationInfo(iIndex).yy
                adDCY(2) = _aoStructPartOrientationInfo(iIndex).yz

                adDCZ(0) = _aoStructPartOrientationInfo(iIndex).zx
                adDCZ(1) = _aoStructPartOrientationInfo(iIndex).zy
                adDCZ(2) = _aoStructPartOrientationInfo(iIndex).zz

                If FnParallelAntiParallelCheck(adDCX, adVec1) Or FnParallelAntiParallelCheck(adDCY, adVec1) Or
                    FnParallelAntiParallelCheck(adDCZ, adVec1) Then
                    bVec1MatchFound = True
                End If
                If FnParallelAntiParallelCheck(adDCX, adVec2) Or FnParallelAntiParallelCheck(adDCY, adVec2) Or
                    FnParallelAntiParallelCheck(adDCZ, adVec2) Then
                    bVec2MatchFound = True
                End If
                If bVec1MatchFound And bVec2MatchFound Then
                    Exit For
                End If
            Next
            If bVec1MatchFound And bVec2MatchFound Then
                FnChkIfFaceVecIsUnique = False
                Exit Function
            Else
                FnChkIfFaceVecIsUnique = True
                Exit Function
            End If
        End If
        FnChkIfFaceVecIsUnique = True
    End Function
    'Code modified on Mar-08-2018
    'Function to collect the orthogonal planar faces to a given reference face in a part
    Function FnCollectOrthogonalPlanarFacesToRefFace(objPart As Part, objFace As Face) As Face()
        Dim adVecObjFace() As Double = Nothing
        Dim adVecObjOrthoFace() As Double = Nothing
        Dim aoOrthogonalFace() As Face = Nothing
        Dim aoAllChildComp() As Component = Nothing
        Dim objChildPart As Part = Nothing
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim bIsValidSolidBody As Boolean = False
        Dim objOccBody As Body = Nothing
        Dim aoAllValidBody() As Body = Nothing

        If Not objPart Is Nothing Then
            aoAllChildComp = FnGetAllComponentsInSession()
            If Not aoAllChildComp Is Nothing Then
                For Each objChildComp As Component In aoAllChildComp
                    objChildPart = FnGetPartFromComponent(objChildComp)
                    If Not objChildPart Is Nothing Then
                        FnLoadPartFully(objChildPart)
                        aoAllValidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)
                        If Not aoAllValidBody Is Nothing Then
                            For Each objBody As Body In aoAllValidBody
                                If Not objBody Is Nothing Then
                                    If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                        If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                            objOccBody = objBody
                                        Else
                                            objOccBody = CType(objChildComp.FindOccurrence(objBody), Body)
                                        End If
                                    ElseIf (_sOemName = DAIMLER_OEM_NAME) Then
                                        If _sDivision = TRUCK_DIVISION Then
                                            If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                                objOccBody = objBody
                                            Else
                                                objOccBody = CType(objChildComp.FindOccurrence(objBody), Body)
                                            End If
                                        ElseIf _sDivision = CAR_DIVISION Then
                                            If FnCheckIfThisIsAChildCompInWeldment(objChildComp, _sOemName) Then
                                                'WEldment in car
                                                objOccBody = CType(objChildComp.FindOccurrence(objBody), Body)
                                            Else
                                                'Component in car
                                                objOccBody = objBody
                                            End If
                                        End If
                                    ElseIf (_sOemName = FIAT_OEM_NAME) Then
                                        'Check if the component is a child component in weldment
                                        If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                            'Fiat component
                                            objOccBody = objBody
                                        Else
                                            'This is a Fiat Weldment child component
                                            objOccBody = CType(objChildComp.FindOccurrence(objBody), Body)
                                        End If
                                    End If

                                    If Not objOccBody Is Nothing Then
                                        For Each objOrthoFace As Face In objOccBody.GetFaces()
                                            If objOrthoFace.SolidFaceType = Face.FaceType.Planar Then
                                                If objFace.Tag <> objOrthoFace.Tag Then
                                                    adVecObjFace = FnGetFaceNormalVector(objFace)
                                                    adVecObjOrthoFace = FnGetFaceNormalVector(objOrthoFace)
                                                    If FnCheckOrthogalityOfTwoVectors(adVecObjFace, adVecObjOrthoFace) Then
                                                        If aoOrthogonalFace Is Nothing Then
                                                            ReDim Preserve aoOrthogonalFace(0)
                                                            aoOrthogonalFace(0) = objOrthoFace
                                                        Else
                                                            ReDim Preserve aoOrthogonalFace(UBound(aoOrthogonalFace) + 1)
                                                            aoOrthogonalFace(UBound(aoOrthogonalFace)) = objOrthoFace
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next
                                    End If

                                End If
                            Next
                        End If
                    End If
                Next
            Else
                aoAllValidBody = FnGetValidBodyForOEM(objPart, _sOemName)
                If Not aoAllValidBody Is Nothing Then
                    For Each objBody As Body In aoAllValidBody
                        If Not objBody Is Nothing Then
                            For Each objOrthoFace As Face In objBody.GetFaces()
                                If objOrthoFace.SolidFaceType = Face.FaceType.Planar Then
                                    If objFace.Tag <> objOrthoFace.Tag Then
                                        adVecObjFace = FnGetFaceNormalVector(objFace)
                                        adVecObjOrthoFace = FnGetFaceNormalVector(objOrthoFace)
                                        If FnCheckOrthogalityOfTwoVectors(adVecObjFace, adVecObjOrthoFace) Then
                                            If aoOrthogonalFace Is Nothing Then
                                                ReDim Preserve aoOrthogonalFace(0)
                                                aoOrthogonalFace(0) = objOrthoFace
                                            Else
                                                ReDim Preserve aoOrthogonalFace(UBound(aoOrthogonalFace) + 1)
                                                aoOrthogonalFace(UBound(aoOrthogonalFace)) = objOrthoFace
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        End If
        FnCollectOrthogonalPlanarFacesToRefFace = aoOrthogonalFace
    End Function

    Public Function FnCreateRotationMatrixFromTwoVec(ByVal adVect1() As Double, ByVal adVec2() As Double) As Double()
        Dim dMatrix(8) As Double

        FnGetUFSession.Mtx3.Initialize(adVect1, adVec2, dMatrix)
        FnCreateRotationMatrixFromTwoVec = dMatrix
    End Function
    'Populate orientation info for each body orientation
    Sub sPopulateOrientationInfo(objBody As Body, ByRef adRotationMatrix() As Double, ByRef aoAllMachinedFace() As Face, Optional bIsRoundShape As Boolean = False)
        Dim iCountXAlignedPeripheralFace As Integer = 0
        Dim iCountYAlignedPeripheralFace As Integer = 0
        Dim iCountZAlignedPeripheralFace As Integer = 0
        Dim iRank As Integer = 0
        Dim objPart As Part = Nothing
        Dim dBoundingBoxVolume As Double = 0
        Dim adBoundingBoxDistance() As Double = Nothing
        Dim aoMisAlignedMachinedFace() As Face = Nothing
        Dim iCountMisAlignedMachinedFace As Integer = 0
        Dim iCountNumOfAlignedFace As Integer = 0
        Dim iCountNumOfAlignedFaceWithHoles As Integer = 0

        objPart = FnGetNxSession.Parts.Work
        If Not objBody Is Nothing Then
            If objBody.IsSolidBody Then
                If Not objBody.IsBlanked Then
                    'For Round Shape, we do not calculate Peripheral planar faces. We just store the rotation matrix for overall part rotation matrix computation
                    If Not bIsRoundShape Then
                        For Each objFace As Face In objBody.GetFaces()
                            If objFace.SolidFaceType = Face.FaceType.Planar Then
                                If FnGetFaceAlignmentDir(objPart, adRotationMatrix, objFace) = "X" Then
                                    If FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt(objPart, objBody, objFace, adRotationMatrix, "X") Then
                                        iCountXAlignedPeripheralFace = iCountXAlignedPeripheralFace + 1
                                        If iCountXAlignedPeripheralFace = 1 Then
                                            iRank = iRank + 1
                                        End If
                                    End If
                                ElseIf FnGetFaceAlignmentDir(objPart, adRotationMatrix, objFace) = "Y" Then
                                    If FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt(objPart, objBody, objFace, adRotationMatrix, "Y") Then
                                        iCountYAlignedPeripheralFace = iCountYAlignedPeripheralFace + 1
                                        If iCountYAlignedPeripheralFace = 1 Then
                                            iRank = iRank + 1
                                        End If
                                    End If
                                ElseIf FnGetFaceAlignmentDir(objPart, adRotationMatrix, objFace) = "Z" Then
                                    If FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt(objPart, objBody, objFace, adRotationMatrix, "Z") Then
                                        iCountZAlignedPeripheralFace = iCountZAlignedPeripheralFace + 1
                                        If iCountZAlignedPeripheralFace = 1 Then
                                            iRank = iRank + 1
                                        End If
                                    End If
                                End If
                                'Code added Jan-20-2018
                                'Get the number of Aligned face
                                If FnCheckIfFaceIsAligned(objFace, adRotationMatrix) Then
                                    iCountNumOfAlignedFace = iCountNumOfAlignedFace + 1
                                    'Code added Nov-14-2018
                                    'Get the count of Aligned faces which has holes on it
                                    If FnCheckIfthePlanarFaceHasHolesOnIt(objFace) Then
                                        iCountNumOfAlignedFaceWithHoles = iCountNumOfAlignedFaceWithHoles + 1
                                    End If
                                End If
                            End If
                        Next
                        adBoundingBoxDistance = FnGetBoundingBoxDistace(objBody, adRotationMatrix)
                        dBoundingBoxVolume = Round(adBoundingBoxDistance(0) * adBoundingBoxDistance(1) * adBoundingBoxDistance(2), 2)
                        'COde added May-10-2018
                        If Not aoAllMachinedFace Is Nothing Then
                            aoMisAlignedMachinedFace = FnGetColOfMisAlignedMachinedFace(aoAllMachinedFace, adRotationMatrix)
                            If Not aoMisAlignedMachinedFace Is Nothing Then
                                iCountMisAlignedMachinedFace = aoMisAlignedMachinedFace.Length()
                            End If
                        End If

                        'Populate the information for this orientation matrix
                        ReDim Preserve _aoStructBodyOrientationInfo(_iOrientationIndex)
                        _aoStructBodyOrientationInfo(_iOrientationIndex).xx = adRotationMatrix(0)
                        _aoStructBodyOrientationInfo(_iOrientationIndex).xy = adRotationMatrix(1)
                        _aoStructBodyOrientationInfo(_iOrientationIndex).xz = adRotationMatrix(2)
                        _aoStructBodyOrientationInfo(_iOrientationIndex).yx = adRotationMatrix(3)
                        _aoStructBodyOrientationInfo(_iOrientationIndex).yy = adRotationMatrix(4)
                        _aoStructBodyOrientationInfo(_iOrientationIndex).yz = adRotationMatrix(5)
                        _aoStructBodyOrientationInfo(_iOrientationIndex).zx = adRotationMatrix(6)
                        _aoStructBodyOrientationInfo(_iOrientationIndex).zy = adRotationMatrix(7)
                        _aoStructBodyOrientationInfo(_iOrientationIndex).zz = adRotationMatrix(8)
                        _aoStructBodyOrientationInfo(_iOrientationIndex).iCountXAlignedPeripheralFaces = iCountXAlignedPeripheralFace
                        _aoStructBodyOrientationInfo(_iOrientationIndex).iCountYAlignedPeripheralFaces = iCountYAlignedPeripheralFace
                        _aoStructBodyOrientationInfo(_iOrientationIndex).iCountZAlignedPeripheralFaces = iCountZAlignedPeripheralFace
                        _aoStructBodyOrientationInfo(_iOrientationIndex).dBoundingBoxVolume = dBoundingBoxVolume
                        _aoStructBodyOrientationInfo(_iOrientationIndex).iRank = iRank
                        _aoStructBodyOrientationInfo(_iOrientationIndex).iCountMisAlignedMachinedFaces = iCountMisAlignedMachinedFace
                        _aoStructBodyOrientationInfo(_iOrientationIndex).iCountAlignedFace = iCountNumOfAlignedFace
                        _aoStructBodyOrientationInfo(_iOrientationIndex).iCountAlignedFaceWithHoles = iCountNumOfAlignedFaceWithHoles
                        _iOrientationIndex = _iOrientationIndex + 1
                    End If

                    'Code added Oct-30-2017
                    'When storing data for Part LCS, computation should involve all solid body.
                    sPopulateWeldmentPartOrientationInfo(adRotationMatrix, aoAllMachinedFace)
                End If
            End If
        End If
    End Sub
    'Code modified on Dec-26-2017
    'Code modified to work with Daimler weldments
    'Populate orientation info for entire part in case of a weldment
    'This function is specifically written to compare orthogonal faces across multiple bodies
    Sub sPopulateWeldmentPartOrientationInfo(ByRef adRotationMatrix() As Double, ByRef aoAllMachinedFace() As Face)
        Dim iCountXAlignedPeripheralFace As Integer = 0
        Dim iCountYAlignedPeripheralFace As Integer = 0
        Dim iCountZAlignedPeripheralFace As Integer = 0
        Dim iRank As Integer = 0
        Dim objPart As Part = Nothing
        Dim dBoundingBoxVolume As Double = 0
        Dim adBoundingBoxDistance() As Double = Nothing
        Dim aoMisAlignedMachinedFace() As Face = Nothing
        Dim iCountMisAlignedMachinedFaces As Integer = 0
        Dim aoAllChildComp() As Component = Nothing
        Dim objChildPart As Part = Nothing
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim bIsValidSolidBody As Boolean = False
        Dim objOccBody As Body = Nothing
        Dim iCountNumOfAlignedFace As Integer = 0
        Dim iCountNumOfAlignedFaceWithHoles As Integer = 0
        Dim aoAllValidBody() As Body = Nothing

        objPart = FnGetNxSession.Parts.Work
        If Not objPart Is Nothing Then
            aoAllChildComp = FnGetAllComponentsInSession()
            If Not aoAllChildComp Is Nothing Then
                For Each objChildComp As Component In aoAllChildComp
                    objChildPart = FnGetPartFromComponent(objChildComp)
                    If Not objChildPart Is Nothing Then
                        FnLoadPartFully(objChildPart)
                        aoAllValidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)
                        If Not aoAllValidBody Is Nothing Then
                            For Each objBody As Body In aoAllValidBody

                                If Not objBody Is Nothing Then
                                    If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                        If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                            objOccBody = objBody
                                        Else
                                            objOccBody = CType(objChildComp.FindOccurrence(objBody), Body)
                                        End If
                                    ElseIf (_sOemName = DAIMLER_OEM_NAME) Then
                                        If _sDivision = TRUCK_DIVISION Then
                                            If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                                objOccBody = objBody
                                            Else
                                                objOccBody = CType(objChildComp.FindOccurrence(objBody), Body)
                                            End If
                                        ElseIf _sDivision = CAR_DIVISION Then
                                            If FnCheckIfThisIsAChildCompInWeldment(objChildComp, _sOemName) Then
                                                'WEldment in car
                                                objOccBody = CType(objChildComp.FindOccurrence(objBody), Body)
                                            Else
                                                'Component in car
                                                objOccBody = objBody
                                            End If
                                        End If
                                    ElseIf (_sOemName = FIAT_OEM_NAME) Then
                                        'Check if the component is a child component in weldment
                                        If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                            'Fiat component
                                            objOccBody = objBody
                                        Else
                                            'This is a Fiat Weldment child component
                                            objOccBody = CType(objChildComp.FindOccurrence(objBody), Body)
                                        End If
                                    End If


                                    If Not objOccBody Is Nothing Then
                                        For Each objFace As Face In objOccBody.GetFaces()
                                            If objFace.SolidFaceType = Face.FaceType.Planar Then
                                                If FnGetFaceAlignmentDir(objPart, adRotationMatrix, objFace) = "X" Then
                                                    If FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt(objPart, objOccBody, objFace, adRotationMatrix, "X") Then
                                                        iCountXAlignedPeripheralFace = iCountXAlignedPeripheralFace + 1
                                                        If iCountXAlignedPeripheralFace = 1 Then
                                                            iRank = iRank + 1
                                                        End If
                                                    End If
                                                ElseIf FnGetFaceAlignmentDir(objPart, adRotationMatrix, objFace) = "Y" Then
                                                    If FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt(objPart, objOccBody, objFace, adRotationMatrix, "Y") Then
                                                        iCountYAlignedPeripheralFace = iCountYAlignedPeripheralFace + 1
                                                        If iCountYAlignedPeripheralFace = 1 Then
                                                            iRank = iRank + 1
                                                        End If
                                                    End If
                                                ElseIf FnGetFaceAlignmentDir(objPart, adRotationMatrix, objFace) = "Z" Then
                                                    If FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt(objPart, objOccBody, objFace, adRotationMatrix, "Z") Then
                                                        iCountZAlignedPeripheralFace = iCountZAlignedPeripheralFace + 1
                                                        If iCountZAlignedPeripheralFace = 1 Then
                                                            iRank = iRank + 1
                                                        End If
                                                    End If
                                                End If
                                                'Code added Jan-20-2018
                                                'Get the number of Aligned face
                                                If FnCheckIfFaceIsAligned(objFace, adRotationMatrix) Then
                                                    iCountNumOfAlignedFace = iCountNumOfAlignedFace + 1
                                                    'Code added Nov-14-2018
                                                    'Get the number of Aligned face which has Holes on it.
                                                    If FnCheckIfthePlanarFaceHasHolesOnIt(objFace) Then
                                                        iCountNumOfAlignedFaceWithHoles = iCountNumOfAlignedFaceWithHoles + 1
                                                    End If
                                                End If

                                            End If
                                        Next
                                    End If

                                End If
                            Next
                        End If
                    End If
                Next
            Else
                aoAllValidBody = FnGetValidBodyForOEM(objPart, _sOemName)


                If Not aoAllValidBody Is Nothing Then
                    For Each objBody As Body In aoAllValidBody
                        For Each objFace As Face In objBody.GetFaces()
                            If objFace.SolidFaceType = Face.FaceType.Planar Then
                                If FnGetFaceAlignmentDir(objPart, adRotationMatrix, objFace) = "X" Then
                                    If FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt(objPart, objBody, objFace, adRotationMatrix, "X") Then
                                        iCountXAlignedPeripheralFace = iCountXAlignedPeripheralFace + 1
                                        If iCountXAlignedPeripheralFace = 1 Then
                                            iRank = iRank + 1
                                        End If
                                    End If
                                ElseIf FnGetFaceAlignmentDir(objPart, adRotationMatrix, objFace) = "Y" Then
                                    If FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt(objPart, objBody, objFace, adRotationMatrix, "Y") Then
                                        iCountYAlignedPeripheralFace = iCountYAlignedPeripheralFace + 1
                                        If iCountYAlignedPeripheralFace = 1 Then
                                            iRank = iRank + 1
                                        End If
                                    End If
                                ElseIf FnGetFaceAlignmentDir(objPart, adRotationMatrix, objFace) = "Z" Then
                                    If FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt(objPart, objBody, objFace, adRotationMatrix, "Z") Then
                                        iCountZAlignedPeripheralFace = iCountZAlignedPeripheralFace + 1
                                        If iCountZAlignedPeripheralFace = 1 Then
                                            iRank = iRank + 1
                                        End If
                                    End If
                                End If
                                'Code added Jan-20-2018
                                'Get the number of Aligned face
                                If FnCheckIfFaceIsAligned(objFace, adRotationMatrix) Then
                                    iCountNumOfAlignedFace = iCountNumOfAlignedFace + 1
                                    'Code added Nov-14-2018
                                    'Get the number of Aligned face which has Holes on it.
                                    If FnCheckIfthePlanarFaceHasHolesOnIt(objFace) Then
                                        iCountNumOfAlignedFaceWithHoles = iCountNumOfAlignedFaceWithHoles + 1
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If

            End If

            'Find the global extents X Y Z for all solid body in a weldment
            adBoundingBoxDistance = FnFindGlobalExtentsXYZForWeldment(objPart, adRotationMatrix)
            dBoundingBoxVolume = Round(adBoundingBoxDistance(0) * adBoundingBoxDistance(1) * adBoundingBoxDistance(2), 2)

            If Not aoAllMachinedFace Is Nothing Then
                aoMisAlignedMachinedFace = FnGetColOfMisAlignedMachinedFace(aoAllMachinedFace, adRotationMatrix)
                If Not aoMisAlignedMachinedFace Is Nothing Then
                    iCountMisAlignedMachinedFaces = aoMisAlignedMachinedFace.Length()
                End If
            End If
            'Populate the information of this orientation matrix to get the Over all optimal rotation matrix of the Part
            ReDim Preserve _aoStructPartOrientationInfo(_iPartOrientationIndex)
            _aoStructPartOrientationInfo(_iPartOrientationIndex).xx = adRotationMatrix(0)
            _aoStructPartOrientationInfo(_iPartOrientationIndex).xy = adRotationMatrix(1)
            _aoStructPartOrientationInfo(_iPartOrientationIndex).xz = adRotationMatrix(2)
            _aoStructPartOrientationInfo(_iPartOrientationIndex).yx = adRotationMatrix(3)
            _aoStructPartOrientationInfo(_iPartOrientationIndex).yy = adRotationMatrix(4)
            _aoStructPartOrientationInfo(_iPartOrientationIndex).yz = adRotationMatrix(5)
            _aoStructPartOrientationInfo(_iPartOrientationIndex).zx = adRotationMatrix(6)
            _aoStructPartOrientationInfo(_iPartOrientationIndex).zy = adRotationMatrix(7)
            _aoStructPartOrientationInfo(_iPartOrientationIndex).zz = adRotationMatrix(8)
            _aoStructPartOrientationInfo(_iPartOrientationIndex).iCountXAlignedPeripheralFaces = iCountXAlignedPeripheralFace
            _aoStructPartOrientationInfo(_iPartOrientationIndex).iCountYAlignedPeripheralFaces = iCountYAlignedPeripheralFace
            _aoStructPartOrientationInfo(_iPartOrientationIndex).iCountZAlignedPeripheralFaces = iCountZAlignedPeripheralFace
            _aoStructPartOrientationInfo(_iPartOrientationIndex).dBoundingBoxVolume = dBoundingBoxVolume
            _aoStructPartOrientationInfo(_iPartOrientationIndex).iRank = iRank
            _aoStructPartOrientationInfo(_iPartOrientationIndex).iCountMisAlignedMachinedFaces = iCountMisAlignedMachinedFaces
            _aoStructPartOrientationInfo(_iPartOrientationIndex).iCountAlignedFace = iCountNumOfAlignedFace
            _aoStructPartOrientationInfo(_iPartOrientationIndex).iCountAlignedFaceWithHoles = iCountNumOfAlignedFaceWithHoles
            _iPartOrientationIndex = _iPartOrientationIndex + 1

        End If
    End Sub

    '****************************************************************************************************************************************************
    'function to Get the dot product value of two variables(either 2 vectors or vector and a point) 
    'Description        : function to Get the dot product value of two variables(either 2 vectors or vector and a point)
    'Function Name      : FnGetDotProduct
    'Input Parameter    : adEdgeCenter,adFaceVec
    'Output Parameter   : double
    '****************************************************************************************************************************************************
    Public Function FnGetDotProduct(ByVal adCenter As Double(), ByVal adFaceVec As Double()) As Double
        Dim dDotProdSum As Double = 0

        For i = LBound(adCenter) To UBound(adCenter)
            dDotProdSum = dDotProdSum + (adCenter(i) * adFaceVec(i))
        Next

        FnGetDotProduct = dDotProdSum
    End Function

    'Function to collect the collection of all misaligned machined face
    Public Function FnGetColOfMisAlignedMachinedFace(ByVal aoAllMachinedFace() As Face, ByVal adRotMatrix() As Double) As Face()
        Dim adFaceVec As Double() = Nothing
        Dim adXDirCos(2) As Double
        Dim adYDirCos(2) As Double
        Dim adZDirCos(2) As Double
        Dim dDotProductX As Double = Nothing
        Dim dDotProductY As Double = Nothing
        Dim dDotProductZ As Double = Nothing
        Dim bIsFaceMisAligned As Boolean = False
        Dim aoColOfMisAlignedMachinedFace As Face() = Nothing

        adXDirCos(0) = adRotMatrix(0)
        adXDirCos(1) = adRotMatrix(1)
        adXDirCos(2) = adRotMatrix(2)

        adYDirCos(0) = adRotMatrix(3)
        adYDirCos(1) = adRotMatrix(4)
        adYDirCos(2) = adRotMatrix(5)

        adZDirCos(0) = adRotMatrix(6)
        adZDirCos(1) = adRotMatrix(7)
        adZDirCos(2) = adRotMatrix(8)
        'If the face vectors satisfies, neither dotproduct nor cross product, then the face is said to be misaligned face.
        If Not aoAllMachinedFace Is Nothing Then
            For Each objFace As Face In aoAllMachinedFace
                bIsFaceMisAligned = False
                If objFace.SolidFaceType = Face.FaceType.Planar Then
                    adFaceVec = FnGetFaceNormalVector(objFace)
                    If Not adFaceVec Is Nothing Then
                        dDotProductX = Abs(Round(FnGetDotProduct(adXDirCos, adFaceVec), 1))
                        dDotProductY = Abs(Round(FnGetDotProduct(adYDirCos, adFaceVec), 1))
                        dDotProductZ = Abs(Round(FnGetDotProduct(adZDirCos, adFaceVec), 1))
                        bIsFaceMisAligned = False
                        If dDotProductX = 0.0 Or dDotProductX = 1.0 Then
                            If dDotProductY = 0.0 Or dDotProductY = 1.0 Then
                                If dDotProductZ = 0.0 Or dDotProductZ = 1.0 Then
                                    bIsFaceMisAligned = False
                                Else
                                    bIsFaceMisAligned = True
                                End If
                            Else
                                bIsFaceMisAligned = True
                            End If
                        Else
                            bIsFaceMisAligned = True
                        End If
                    End If
                End If
                If bIsFaceMisAligned Then
                    If aoColOfMisAlignedMachinedFace Is Nothing Then
                        ReDim Preserve aoColOfMisAlignedMachinedFace(0)
                        aoColOfMisAlignedMachinedFace(0) = objFace
                    Else
                        ReDim Preserve aoColOfMisAlignedMachinedFace(UBound(aoColOfMisAlignedMachinedFace) + 1)
                        aoColOfMisAlignedMachinedFace(UBound(aoColOfMisAlignedMachinedFace)) = objFace
                    End If
                End If
            Next
        End If
        FnGetColOfMisAlignedMachinedFace = aoColOfMisAlignedMachinedFace
    End Function
    '****************************************************************************************************************************************************
    'Function to get the Bounding box distance x,y,z 
    'Description        : Function to get the Bounding box distance x,y,z 
    'Function Name      : FnGetBoundingBoxDistance
    'Input Parameter    : adRotationMatrix
    'Output Parameter   : adDistace
    '****************************************************************************************************************************************************
    Public Function FnGetBoundingBoxDistace(ByVal objBody As Body, ByVal adRotationMatrix As Double()) As Double()

        Dim matrixTag As Tag = NXOpen.Tag.Null
        Dim csysTag As Tag = NXOpen.Tag.Null
        Dim csysOrigin As Double() = {0, 0, 0}
        Dim adMin_corner(2) As Double
        Dim adDirections(2, 2) As Double
        Dim adDistance(2) As Double

        FnGetUFSession.Csys.CreateMatrix(adRotationMatrix, matrixTag)
        FnGetUFSession.Csys.CreateCsys(csysOrigin, matrixTag, csysTag)
        FnGetUFSession.Modl.AskBoundingBoxExact(objBody.Tag, csysTag, adMin_corner, adDirections, adDistance)

        FnGetBoundingBoxDistace = adDistance

        'Delete the created CSYS
        SDeleteObjects({NXObjectManager.Get(csysTag)})
    End Function

    'Get the alignment direction for the face
    Function FnGetFaceAlignmentDir(objPart As Part, ByRef adRotationMatrix() As Double, objFace As Face) As String
        Dim dirFace1 As Direction = Nothing

        dirFace1 = objPart.Directions.CreateDirection(objFace, Sense.Forward, SmartObject.UpdateOption.WithinModeling)
        If Not adRotationMatrix Is Nothing Then
            If FnParallelAntiParallelCheck({adRotationMatrix(0), adRotationMatrix(1), adRotationMatrix(2)},
                                           {dirFace1.Vector.X, dirFace1.Vector.Y, dirFace1.Vector.Z}) Then
                FnGetFaceAlignmentDir = "X"
            ElseIf FnParallelAntiParallelCheck({adRotationMatrix(3), adRotationMatrix(4), adRotationMatrix(5)},
                                           {dirFace1.Vector.X, dirFace1.Vector.Y, dirFace1.Vector.Z}) Then
                FnGetFaceAlignmentDir = "Y"
            ElseIf FnParallelAntiParallelCheck({adRotationMatrix(6), adRotationMatrix(7), adRotationMatrix(8)},
                                       {dirFace1.Vector.X, dirFace1.Vector.Y, dirFace1.Vector.Z}) Then
                FnGetFaceAlignmentDir = "Z"
            Else
                FnGetFaceAlignmentDir = ""
            End If
        End If
    End Function
    'Code modified May-22-2017
    'Determine optimal matrix from among the various orientations
    'first try to minimize the number of misaligned faces then if there is a tie use the usual tie-breaker (rank and then volume)
    Function FnDetermineOptimalMatrix(ByRef aoStructOrientationInfo() As structOrientationInfo,
                                      Optional bMachinedFacesPresent As Boolean = False) As Double()

        Dim iIndexofMinVol As Integer = 0
        Dim dMinBoundingBoxVol As Double = 0
        Dim dBoundingBoxVol As Double = 0
        Dim aoOrientationMatrix As Double() = Nothing

        'Validation added on Oct 23 2020
        'Get the less volume orientation of the all the possible orientaton
        If Not aoStructOrientationInfo Is Nothing Then
            For iIndex As Integer = 0 To UBound(aoStructOrientationInfo)
                dBoundingBoxVol = aoStructOrientationInfo(iIndex).dBoundingBoxVolume

                If dMinBoundingBoxVol = 0 Then
                    dMinBoundingBoxVol = dBoundingBoxVol
                    iIndexofMinVol = iIndex
                    aoOrientationMatrix = {aoStructOrientationInfo(iIndexofMinVol).xx,
                                               aoStructOrientationInfo(iIndexofMinVol).xy,
                                               aoStructOrientationInfo(iIndexofMinVol).xz,
                                               aoStructOrientationInfo(iIndexofMinVol).yx,
                                               aoStructOrientationInfo(iIndexofMinVol).yy,
                                               aoStructOrientationInfo(iIndexofMinVol).yz,
                                               aoStructOrientationInfo(iIndexofMinVol).zx,
                                               aoStructOrientationInfo(iIndexofMinVol).zy,
                                               aoStructOrientationInfo(iIndexofMinVol).zz}
                ElseIf dBoundingBoxVol < dMinBoundingBoxVol Then
                    dMinBoundingBoxVol = dBoundingBoxVol
                    iIndexofMinVol = iIndex
                    aoOrientationMatrix = {aoStructOrientationInfo(iIndexofMinVol).xx,
                                                aoStructOrientationInfo(iIndexofMinVol).xy,
                                                aoStructOrientationInfo(iIndexofMinVol).xz,
                                                aoStructOrientationInfo(iIndexofMinVol).yx,
                                                aoStructOrientationInfo(iIndexofMinVol).yy,
                                                aoStructOrientationInfo(iIndexofMinVol).yz,
                                                aoStructOrientationInfo(iIndexofMinVol).zx,
                                                aoStructOrientationInfo(iIndexofMinVol).zy,
                                                aoStructOrientationInfo(iIndexofMinVol).zz}
                End If
            Next
        End If
        FnDetermineOptimalMatrix = aoOrientationMatrix
    End Function

    'Function to check if the given face is a periperal face along the given direction
    Function FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt(ByVal objPart As Part, ByVal objBody As Body, ByVal objFace As Face,
                                                         adRotationMatrix() As Double, Optional sDirection As String = Nothing) As Boolean
        '1. Obtain Bounding Box 'Min Point' and 'Extents' along each axis of View rotation matrix (Using Bounding Box API)
        '2. Compute 8 Bounding Box corners.
        '3. Rotate Bounding Box corners into View. (Use "fnGetRotatedVector" ).
        '4. Compute Min Coord, Max Coord (eg. Min X, Max X, etc) of Rotated Bounding Box Corners.
        '5. Rotate face center of input face. (Use "fnGetRotatedVector" ).
        '6. Compare Rotated Face Center Coord with Min Coord / Max Coord.
        '7. If Rotated Face Center Coord matches with either Min Coord or Max Coord, then the input face is peripheral.

        Dim adBoundingBoxCorner(7, 2) As Double
        Dim adFaceCenter() As Double = Nothing
        Dim adVector1(2) As Double
        Dim adRotatedBoundingBoxCorner(7, 2) As Double
        Dim adRotatedBBVector() As Double = Nothing
        Dim adRotatedFaceCenterVec() As Double = Nothing
        Dim adAllXValues(7) As Double
        Dim adAllYValues(7) As Double
        Dim adAllZValues(7) As Double
        Dim dRotatedBBMinX As Double = -1
        Dim dRotatedBBMaxX As Double = -1
        Dim dRotatedBBMinY As Double = -1
        Dim dRotatedBBMaxY As Double = -1
        Dim dRotatedBBMinZ As Double = -1
        Dim dRotatedBBMaxZ As Double = -1

        'Compute Bounding Box 8 corner coordinates
        adBoundingBoxCorner = FnGetBoundingBoxCorners(objBody, adRotationMatrix)
        'Get the face center of face which is to be checked
        adFaceCenter = FnGetFaceCenter(objFace)

        If Not adBoundingBoxCorner Is Nothing Then
            'Rotate each Bounding Box corners into the view
            For iIndex As Integer = 0 To 7
                adVector1(0) = adBoundingBoxCorner(iIndex, 0)
                adVector1(1) = adBoundingBoxCorner(iIndex, 1)
                adVector1(2) = adBoundingBoxCorner(iIndex, 2)

                adRotatedBBVector = FnGetRotatedVector(adRotationMatrix, adVector1)

                'adRotatedBoundingBoxCorner(iIndex, 0) = adRotatedVector(0)
                'adRotatedBoundingBoxCorner(iIndex, 1) = adRotatedVector(1)
                'adRotatedBoundingBoxCorner(iIndex, 2) = adRotatedVector(2)
                'Collect All the X, Y ,Z co-ordinates
                adAllXValues(iIndex) = adRotatedBBVector(0)
                adAllYValues(iIndex) = adRotatedBBVector(1)
                adAllZValues(iIndex) = adRotatedBBVector(2)
            Next

            'Find the minimum and Maximum X,Y,Z values from the rotated Bounding Box COrner values
            Array.Sort(adAllXValues)
            dRotatedBBMinX = adAllXValues(0)
            dRotatedBBMaxX = adAllXValues(7)

            Array.Sort(adAllYValues)
            dRotatedBBMinY = adAllYValues(0)
            dRotatedBBMaxY = adAllYValues(7)

            Array.Sort(adAllZValues)
            dRotatedBBMinZ = adAllZValues(0)
            dRotatedBBMaxZ = adAllZValues(7)

            'Rotate the Face center co-ordinates into the view
            adRotatedFaceCenterVec = FnGetRotatedVector(adRotationMatrix, adFaceCenter)


        End If
        'When direction is not specified check for all direction X,Y,Z
        If sDirection Is Nothing Then
            If ((Round(dRotatedBBMinX, 3) = Round(adRotatedFaceCenterVec(0), 3)) Or (Round(dRotatedBBMaxX, 3) = Round(adRotatedFaceCenterVec(0), 3))) Then
                FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt = True
                Exit Function
            ElseIf ((Round(dRotatedBBMinY, 3) = Round(adRotatedFaceCenterVec(1), 3)) Or (Round(dRotatedBBMaxY, 3) = Round(adRotatedFaceCenterVec(1), 3))) Then
                FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt = True
                Exit Function
            ElseIf ((Round(dRotatedBBMinZ, 3) = Round(adRotatedFaceCenterVec(2), 3)) Or (Round(dRotatedBBMaxZ, 3) = Round(adRotatedFaceCenterVec(2), 3))) Then
                FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt = True
                Exit Function
            End If
        Else
            If sDirection.ToUpper = "X" Then
                'Check if bounding box minX or MaxX is equal to Facecenter X value
                If ((Round(dRotatedBBMinX, 3) = Round(adRotatedFaceCenterVec(0), 3)) Or (Round(dRotatedBBMaxX, 3) = Round(adRotatedFaceCenterVec(0), 3))) Then
                    FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt = True
                    Exit Function
                End If
            End If
            If sDirection.ToUpper = "Y" Then
                'Check if bounding box minY or MaxY is equal to Facecenter Y value
                If ((Round(dRotatedBBMinY, 3) = Round(adRotatedFaceCenterVec(1), 3)) Or (Round(dRotatedBBMaxY, 3) = Round(adRotatedFaceCenterVec(1), 3))) Then
                    FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt = True
                    Exit Function
                End If
            End If
            If sDirection.ToUpper = "Z" Then
                'Check if bounding box minZ or MaxZ is equal to Facecenter Z value
                If ((Round(dRotatedBBMinZ, 3) = Round(adRotatedFaceCenterVec(2), 3)) Or (Round(dRotatedBBMaxZ, 3) = Round(adRotatedFaceCenterVec(2), 3))) Then
                    FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt = True
                    Exit Function
                End If
            End If
        End If
        FnChkIfFaceIsPeripheralFaceOnBoundingBoxDirAlt = False
    End Function
    Function FnGetRotatedVector(adRotationMatrix() As Double, adInputVector() As Double) As Double()
        Dim adRotationMat4X4(15) As Double
        Dim adInverseRotationMat4X4(15) As Double
        Dim adInverseRotationMat3X3(8) As Double
        Dim adMultipliedVector(2) As Double

        'Convert Matrix3X3 to Matrix4X4. This conversion is used, since NXAPI doesnot contain Inverse in 3X3 matrix
        FnGetUFSession.Mtx3.Mtx4(adRotationMatrix, adRotationMat4X4)
        'Inverse the Matrix4X4
        FnGetUFSession.Mtx4.Invert(adRotationMat4X4, adInverseRotationMat4X4)
        'Convert Matrix4X4 to Matrix3X3 by eliminating last row and last column data
        adInverseRotationMat3X3(0) = adInverseRotationMat4X4(0)
        adInverseRotationMat3X3(1) = adInverseRotationMat4X4(1)
        adInverseRotationMat3X3(2) = adInverseRotationMat4X4(2)

        adInverseRotationMat3X3(3) = adInverseRotationMat4X4(4)
        adInverseRotationMat3X3(4) = adInverseRotationMat4X4(5)
        adInverseRotationMat3X3(5) = adInverseRotationMat4X4(6)

        adInverseRotationMat3X3(6) = adInverseRotationMat4X4(8)
        adInverseRotationMat3X3(7) = adInverseRotationMat4X4(9)
        adInverseRotationMat3X3(8) = adInverseRotationMat4X4(10)

        'Code modified on July-14-2017
        'VecMultiply function did not give the correct result as intended for NC COmponent mb258636,35,41,42
        'So we are relying on manual matrix multiplication
        'FnGetUFSession.Mtx3.VecMultiply(adInputVector, adInverseRotationMat3X3, adMultipliedVector)

        adMultipliedVector(0) = adInputVector(0) * adInverseRotationMat3X3(0) + adInputVector(1) * adInverseRotationMat3X3(3) + adInputVector(2) * adInverseRotationMat3X3(6)
        adMultipliedVector(1) = adInputVector(0) * adInverseRotationMat3X3(1) + adInputVector(1) * adInverseRotationMat3X3(4) + adInputVector(2) * adInverseRotationMat3X3(7)
        adMultipliedVector(2) = adInputVector(0) * adInverseRotationMat3X3(2) + adInputVector(1) * adInverseRotationMat3X3(5) + adInputVector(2) * adInverseRotationMat3X3(8)

        FnGetRotatedVector = adMultipliedVector
    End Function

    'Function to get the bounding box 8 corner co-ordinates
    Function FnGetBoundingBoxCorners(objBody As Body, adRotationMatrix() As Double) As Double(,)
        Dim dXExtent As Double = 0
        Dim dYExtent As Double = 0
        Dim dZExtent As Double = 0
        Dim adBoundingBoxDistance() As Double = Nothing
        Dim adXAxisDCS(2) As Double
        Dim adYAxisDCS(2) As Double
        Dim adZAxisDCS(2) As Double
        Dim adBoundingBoxCorner(7, 2) As Double
        Dim matrixTag As Tag = NXOpen.Tag.Null
        Dim csysTag As Tag = NXOpen.Tag.Null
        Dim csysOrigin As Double() = {0, 0, 0}
        Dim adMin_corner(2) As Double
        Dim adDirections(2, 2) As Double
        Dim adDistance(2) As Double

        FnGetUFSession.Csys.CreateMatrix(adRotationMatrix, matrixTag)
        FnGetUFSession.Csys.CreateCsys(csysOrigin, matrixTag, csysTag)
        FnGetUFSession.Modl.AskBoundingBoxExact(objBody.Tag, csysTag, adMin_corner, adDirections, adDistance)

        'Delete the created CSYS
        SDeleteObjects({NXObjectManager.Get(csysTag)})

        dXExtent = adDistance(0)
        dYExtent = adDistance(1)
        dZExtent = adDistance(2)

        adXAxisDCS(0) = adRotationMatrix(0)
        adXAxisDCS(1) = adRotationMatrix(1)
        adXAxisDCS(2) = adRotationMatrix(2)

        adYAxisDCS(0) = adRotationMatrix(3)
        adYAxisDCS(1) = adRotationMatrix(4)
        adYAxisDCS(2) = adRotationMatrix(5)

        adZAxisDCS(0) = adRotationMatrix(6)
        adZAxisDCS(1) = adRotationMatrix(7)
        adZAxisDCS(2) = adRotationMatrix(8)

        For iIndex As Integer = 0 To 2
            adBoundingBoxCorner(0, iIndex) = adMin_corner(iIndex)
            adBoundingBoxCorner(1, iIndex) = adMin_corner(iIndex) + (dXExtent * adXAxisDCS(iIndex))
            adBoundingBoxCorner(2, iIndex) = adMin_corner(iIndex) + (dYExtent * adYAxisDCS(iIndex))
            adBoundingBoxCorner(3, iIndex) = adMin_corner(iIndex) + (dZExtent * adZAxisDCS(iIndex))
            adBoundingBoxCorner(4, iIndex) = adMin_corner(iIndex) + (dXExtent * adXAxisDCS(iIndex)) + (dYExtent * adYAxisDCS(iIndex))
            adBoundingBoxCorner(5, iIndex) = adMin_corner(iIndex) + (dYExtent * adYAxisDCS(iIndex)) + (dZExtent * adZAxisDCS(iIndex))
            adBoundingBoxCorner(6, iIndex) = adMin_corner(iIndex) + (dZExtent * adZAxisDCS(iIndex)) + (dXExtent * adXAxisDCS(iIndex))
            adBoundingBoxCorner(7, iIndex) = adMin_corner(iIndex) + (dXExtent * adXAxisDCS(iIndex)) + (dYExtent * adYAxisDCS(iIndex)) + (dZExtent * adZAxisDCS(iIndex))
        Next
        FnGetBoundingBoxCorners = adBoundingBoxCorner
    End Function
    ''****************************************************************************************************************************************************
    ''Function to Get Least Value in an array
    ''Description        : Function to Get Least Value in an array
    ''Function Name      : FnGetLeastValueInArray
    ''Input Parameter    : adValues
    ''Output Parameter   : dLeastValue
    ''****************************************************************************************************************************************************
    'Function FnGetLeastValueInArray(ByVal adValues As Double()) As Double
    '    Dim dLeastValue As Double = 0

    '    If Not adValues Is Nothing Then
    '        For Each dValue As Double In adValues
    '            If dLeastValue = 0 Then
    '                dLeastValue = dValue
    '            Else
    '                If dValue < dLeastValue Then
    '                    dLeastValue = dValue
    '                End If
    '            End If
    '        Next
    '    End If
    '    FnGetLeastValueInArray = dLeastValue
    'End Function


    ''****************************************************************************************************************************************************
    ''Function to Get Maximum Value in an array
    ''Description        : Function to Get Maximum Value in an array
    ''Function Name      : FnGetMaxValueInArray
    ''Input Parameter    : adValues
    ''Output Parameter   : dMaxValue
    ''****************************************************************************************************************************************************
    'Function FnGetMaxValueInArray(ByVal adValues As Double()) As Double
    '    Dim dMaxValue As Double = Nothing

    '    If Not adValues Is Nothing Then
    '        For Each dValue As Double In adValues
    '            If dMaxValue = Nothing Then
    '                dMaxValue = dValue
    '            Else
    '                If dValue > dMaxValue Then
    '                    dMaxValue = dValue
    '                Else
    '                    dMaxValue = dMaxValue
    '                End If
    '            End If
    '        Next
    '    End If
    '    FnGetMaxValueInArray = dMaxValue
    'End Function

    'Function to write the optimal rotation matix of each solid body of a part in a Body_LCS sheet and rotation matrix of entire part in Model View Cosines sheet
    Sub sWriteLCSInfo(objPart As Part)
        'Dim aoAllComp() As Component = Nothing
        Dim objChildPart As Part = Nothing
        Dim objOccBody As Body = Nothing
        Dim adOptimalRotationMat() As Double = Nothing
        Dim adPartOptimalRotMat() As Double = Nothing
        Dim sBodyName As String = ""
        Dim iLCSRowStart As Integer = 0
        Dim iRowFilledInModelViewDC As Integer = 0
        Dim iColLCSBodyName As Integer = 0
        Dim iColModelViewName As String = ""
        Dim iColDCSXx As Integer = 0
        Dim iColDCSXy As Integer = 0
        Dim iColDCSXz As Integer = 0
        Dim iColDCSYx As Integer = 0
        Dim iColDCSYy As Integer = 0
        Dim iColDCSYz As Integer = 0
        Dim iColDCSZx As Integer = 0
        Dim iColDCSZy As Integer = 0
        Dim iColDCSZz As Integer = 0
        Dim aoMachinedFace() As DisplayableObject = Nothing
        Dim bIsMachinedFacePresent As Boolean = False
        Dim adIdentityMatrix(8) As Double
        Dim sPartName As String = ""
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim objFeatureGroup As Features.FeatureGroup = Nothing
        Dim aoAllMembers() As Features.Feature = Nothing
        Dim bIsValidSolidBody As Boolean = False
        '_dictBodyOptimalRotMat = New Dictionary(Of String, Double())
        Dim aoAllValidBodys() As Body = Nothing
        Dim mTrailOrientation As Matrix3x3 = Nothing

        sPartName = objPart.Leaf.ToString()
        iLCSRowStart = 2
        iColLCSBodyName = 1
        iColModelViewName = 1
        iColDCSXx = 2
        iColDCSXy = 3
        iColDCSXz = 4
        iColDCSYx = 5
        iColDCSYy = 6
        iColDCSYz = 7
        iColDCSZx = 8
        iColDCSZy = 9
        iColDCSZz = 10

        'aoAllComp = FnGetAllComponentsInSession()
        sWriteToLogFile("Started populating LCS info for: " & objPart.Leaf.ToUpper)

        aoAllValidBodys = FnGetValidBodyForOEM(objPart, _sOemName)
        If Not aoAllValidBodys Is Nothing Then
            For Each objBody As Body In aoAllValidBodys
                If Not objBody Is Nothing Then
                    'cODE ADDED APR-09-2020
                    If (FnGetStringUserAttribute(objBody, _SHAPE_ATTR_NAME) <> WIRE_MESH_SHAPE) And
                                (Not FnChkIfBodyIsMesh(objPart, objBody)) Then
                        sBodyName = objBody.JournalIdentifier
                        If FnChkIfBodyHasOrthogonalFace(objBody) Then
                            'Normal computation of optimal rotation matrix
                            adOptimalRotationMat = FnComputeOptimalRotationMatrixAlt(objBody)
                        Else
                            'Round Logic
                            adOptimalRotationMat = FnComputeOptimalRotationMatForRound(objPart, objBody)
                        End If

                        'Populate the optimal rotation matrix of each solid body in Body_LCS sheet
                        If Not adOptimalRotationMat Is Nothing Then
                            sWriteToLogFile("Populate computed LCS for Body " & sBodyName & " in " & objPart.Leaf.ToUpper)
                            'FnGetNxSession.ListingWindow.Open()
                            'FnGetNxSession.ListingWindow.WriteFullline(adOptimalRotationMat(0).ToString() & "  " & adOptimalRotationMat(1).ToString() & "  " & adOptimalRotationMat(2).ToString())
                            iLCSRowStart = iLCSRowStart + 1
                            '_dictBodyOptimalRotMat.Add(sBodyName, adOptimalRotationMat)
                        Else
                            sWriteToLogFile("Populate identity matrix LCS for Body " & sBodyName & " in " & objPart.Leaf.ToUpper)
                            sPopulateIdentityMatrix(sBodyName, iLCSRowStart, bIsBodyLCS:=True)
                            iLCSRowStart = iLCSRowStart + 1
                            '_dictBodyOptimalRotMat.Add(sBodyName, {1, 0, 0, 0, 1, 0, 0, 0, 1})
                        End If
                    End If
                End If
            Next
        End If
        'Code added May-22-2017
        'Compute and Populate the orientation info accross the body. This is needed only for weldment
        'If FnChkPartisWeldment(objPart) Then
        If (_sOemName = DAIMLER_OEM_NAME) Then
            If FnCheckIfThisIsAWeldment(sPartName) Then
                sPopulateOrientationForWeldment(objPart)
            End If
        ElseIf (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
            If FnChkPartisWeldment(objPart) Then
                sPopulateOrientationForWeldment(objPart)
            End If
        ElseIf (_sOemName = FIAT_OEM_NAME) Then
            If FnChkPartisWeldmentBasedOnAttr(objPart) Then
                sPopulateOrientationForWeldment(objPart)
            End If
        End If

        'COde added May-10-2018
        'For all components and Weldments check for the presence of machined face
        'CHeck for the presence of Machined faces
        aoMachinedFace = FnGetFaceObjectByAttributes(objPart, _FINISHTOLERANCE_ATTR_NAME, _FINISH_TOL_VALUE2)
        Dim aoAllMachinedFace() As Face = Nothing
        Dim aoAllMachinedFaceWithHoles() As Face = Nothing

        aoAllMachinedFace = FnConvertDisplayObjectToFace(aoMachinedFace)
        If Not aoAllMachinedFace Is Nothing Then
            aoAllMachinedFaceWithHoles = FnCollectFaceWhichHasHolesOnIt(aoAllMachinedFace)
            If Not aoAllMachinedFaceWithHoles Is Nothing Then
                bIsMachinedFacePresent = True
            Else
                bIsMachinedFacePresent = False
            End If
        Else
            bIsMachinedFacePresent = False
        End If

        'Computing the optimal rotation matrix for the overall part as well on similar lines.
        adPartOptimalRotMat = FnDetermineOptimalMatrix(_aoStructPartOrientationInfo, bIsMachinedFacePresent)
        'Populate the optimal rotation matrix of entire part in Model View Cosines sheet.
        If Not adPartOptimalRotMat Is Nothing Then
            sWriteToLogFile("Populate computed LCS for Part " & objPart.Leaf.ToUpper)
            'FnGetNxSession.ListingWindow.WriteFullline(adPartOptimalRotMat(0).ToString() & "  " & adPartOptimalRotMat(1).ToString() & "  " & adPartOptimalRotMat(2).ToString())

        Else
            sWriteToLogFile("Populate identity matrix LCS for Part " & objPart.Leaf.ToUpper)
            sPopulateIdentityMatrix(Nothing, 0, bIsBodyLCS:=False)
        End If
        'Clean memory
        _aoStructPartOrientationInfo = Nothing
        _iPartOrientationIndex = 0
        _aoStructBodyOrientationInfo = Nothing
        _iOrientationIndex = 0

        '********************************** TEST CODE TO CREATE LCS FOR PART *****************************************************

        If _bCreatePartLCSOrientation Then
            If Not adPartOptimalRotMat Is Nothing Then
                'Bug fixed on July-14-2017
                sCreateCustomModelViewForOptimalRotMat(objPart, adPartOptimalRotMat, PART_LCS_VIEW_NAME)
            Else
                'Create identity matrix and orient the part to identity matrix
                'Identity Matrix
                adIdentityMatrix(0) = 1
                adIdentityMatrix(1) = 0
                adIdentityMatrix(2) = 0
                adIdentityMatrix(3) = 0
                adIdentityMatrix(4) = 1
                adIdentityMatrix(5) = 0
                adIdentityMatrix(6) = 0
                adIdentityMatrix(7) = 0
                adIdentityMatrix(8) = 1
                sCreateCustomModelViewForOptimalRotMat(objPart, adIdentityMatrix, PART_LCS_VIEW_NAME)
            End If
        End If
        '**************************************************************************************************************************
    End Sub

    Function FnConvertRotMatValueToMatrix(adRotMatrix() As Double) As Matrix3x3
        Dim aoMatrix As Matrix3x3 = Nothing
        If Not adRotMatrix Is Nothing Then
            aoMatrix.Xx = adRotMatrix(0)
            aoMatrix.Xy = adRotMatrix(1)
            aoMatrix.Xz = adRotMatrix(2)
            aoMatrix.Yx = adRotMatrix(3)
            aoMatrix.Yy = adRotMatrix(4)
            aoMatrix.Yz = adRotMatrix(5)
            aoMatrix.Zx = adRotMatrix(6)
            aoMatrix.Zy = adRotMatrix(7)
            aoMatrix.Zz = adRotMatrix(8)
        End If
        FnConvertRotMatValueToMatrix = aoMatrix
    End Function

    'Populate LCS information of body to excel file
    Sub sPopulateBodyLCSToExcelFile(sBodyName As String, ByRef adOptimalRotationMat() As Double, iLCSRowStart As Integer, sSheetName As String)
        Dim iColLCSBodyName As Integer = 0
        Dim iColDCSXx As Integer = 0
        Dim iColDCSXy As Integer = 0
        Dim iColDCSXz As Integer = 0
        Dim iColDCSYx As Integer = 0
        Dim iColDCSYy As Integer = 0
        Dim iColDCSYz As Integer = 0
        Dim iColDCSZx As Integer = 0
        Dim iColDCSZy As Integer = 0
        Dim iColDCSZz As Integer = 0

        iColLCSBodyName = 1
        iColDCSXx = 2
        iColDCSXy = 3
        iColDCSXz = 4
        iColDCSYx = 5
        iColDCSYy = 6
        iColDCSYz = 7
        iColDCSZx = 8
        iColDCSZy = 9
        iColDCSZz = 10

        SWriteValueToCell(_objWorkBk, sSheetName, iLCSRowStart, iColLCSBodyName, sBodyName)

        SWriteValueToCell(_objWorkBk, sSheetName, iLCSRowStart, iColDCSXx, adOptimalRotationMat(0).ToString)
        SWriteValueToCell(_objWorkBk, sSheetName, iLCSRowStart, iColDCSXy, adOptimalRotationMat(1).ToString)
        SWriteValueToCell(_objWorkBk, sSheetName, iLCSRowStart, iColDCSXz, adOptimalRotationMat(2).ToString)

        SWriteValueToCell(_objWorkBk, sSheetName, iLCSRowStart, iColDCSYx, adOptimalRotationMat(3).ToString)
        SWriteValueToCell(_objWorkBk, sSheetName, iLCSRowStart, iColDCSYy, adOptimalRotationMat(4).ToString)
        SWriteValueToCell(_objWorkBk, sSheetName, iLCSRowStart, iColDCSYz, adOptimalRotationMat(5).ToString)

        SWriteValueToCell(_objWorkBk, sSheetName, iLCSRowStart, iColDCSZx, adOptimalRotationMat(6).ToString)
        SWriteValueToCell(_objWorkBk, sSheetName, iLCSRowStart, iColDCSZy, adOptimalRotationMat(7).ToString)
        SWriteValueToCell(_objWorkBk, sSheetName, iLCSRowStart, iColDCSZz, adOptimalRotationMat(8).ToString)

    End Sub
    'Populate Part LCS to the excel file
    Sub sPopulatePartLCSToExcelFile(ByRef adPartOptimalRotMat() As Double)
        Dim iRowFilledInModelViewDC As Integer = 0
        Dim iColModelViewName As String = ""
        Dim iColDCSXx As Integer = 0
        Dim iColDCSXy As Integer = 0
        Dim iColDCSXz As Integer = 0
        Dim iColDCSYx As Integer = 0
        Dim iColDCSYy As Integer = 0
        Dim iColDCSYz As Integer = 0
        Dim iColDCSZx As Integer = 0
        Dim iColDCSZy As Integer = 0
        Dim iColDCSZz As Integer = 0

        iColModelViewName = 1
        iColDCSXx = 2
        iColDCSXy = 3
        iColDCSXz = 4
        iColDCSYx = 5
        iColDCSYy = 6
        iColDCSYz = 7
        iColDCSZx = 8
        iColDCSZy = 9
        iColDCSZz = 10

        'Get the last row filled detail in Model view cosine sheet. As it might contain datas.
        iRowFilledInModelViewDC = FnGetNumberofRows(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, 1, 1)

        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColModelViewName, PART_LCS_VIEW_NAME)

        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSXx, adPartOptimalRotMat(0).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSXy, adPartOptimalRotMat(1).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSXz, adPartOptimalRotMat(2).ToString)

        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSYx, adPartOptimalRotMat(3).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSYy, adPartOptimalRotMat(4).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSYz, adPartOptimalRotMat(5).ToString)

        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSZx, adPartOptimalRotMat(6).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSZy, adPartOptimalRotMat(7).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSZz, adPartOptimalRotMat(8).ToString)
    End Sub

    Sub sPopulateIdentityMatrix(sBodyName As String, iLCSRowStart As Integer, bIsBodyLCS As Boolean)

        Dim adIdentityMatrix(8) As Double
        Dim iColLCSBodyName As Integer = 0
        Dim iColDCSXx As Integer = 0
        Dim iColDCSXy As Integer = 0
        Dim iColDCSXz As Integer = 0
        Dim iColDCSYx As Integer = 0
        Dim iColDCSYy As Integer = 0
        Dim iColDCSYz As Integer = 0
        Dim iColDCSZx As Integer = 0
        Dim iColDCSZy As Integer = 0
        Dim iColDCSZz As Integer = 0
        Dim iRowFilledInModelViewDC As Integer = 0
        Dim iColModelViewName As String = ""

        'Identity Matrix
        adIdentityMatrix(0) = 1
        adIdentityMatrix(1) = 0
        adIdentityMatrix(2) = 0
        adIdentityMatrix(3) = 0
        adIdentityMatrix(4) = 1
        adIdentityMatrix(5) = 0
        adIdentityMatrix(6) = 0
        adIdentityMatrix(7) = 0
        adIdentityMatrix(8) = 1

        iColLCSBodyName = 1
        iColDCSXx = 2
        iColDCSXy = 3
        iColDCSXz = 4
        iColDCSYx = 5
        iColDCSYy = 6
        iColDCSYz = 7
        iColDCSZx = 8
        iColDCSZy = 9
        iColDCSZz = 10
        iColModelViewName = 1
        If bIsBodyLCS Then
            SWriteValueToCell(_objWorkBk, BODY_LCS_SHEET_NAME, iLCSRowStart, iColLCSBodyName, sBodyName)

            SWriteValueToCell(_objWorkBk, BODY_LCS_SHEET_NAME, iLCSRowStart, iColDCSXx, adIdentityMatrix(0).ToString)
            SWriteValueToCell(_objWorkBk, BODY_LCS_SHEET_NAME, iLCSRowStart, iColDCSXy, adIdentityMatrix(1).ToString)
            SWriteValueToCell(_objWorkBk, BODY_LCS_SHEET_NAME, iLCSRowStart, iColDCSXz, adIdentityMatrix(2).ToString)

            SWriteValueToCell(_objWorkBk, BODY_LCS_SHEET_NAME, iLCSRowStart, iColDCSYx, adIdentityMatrix(3).ToString)
            SWriteValueToCell(_objWorkBk, BODY_LCS_SHEET_NAME, iLCSRowStart, iColDCSYy, adIdentityMatrix(4).ToString)
            SWriteValueToCell(_objWorkBk, BODY_LCS_SHEET_NAME, iLCSRowStart, iColDCSYz, adIdentityMatrix(5).ToString)

            SWriteValueToCell(_objWorkBk, BODY_LCS_SHEET_NAME, iLCSRowStart, iColDCSZx, adIdentityMatrix(6).ToString)
            SWriteValueToCell(_objWorkBk, BODY_LCS_SHEET_NAME, iLCSRowStart, iColDCSZy, adIdentityMatrix(7).ToString)
            SWriteValueToCell(_objWorkBk, BODY_LCS_SHEET_NAME, iLCSRowStart, iColDCSZz, adIdentityMatrix(8).ToString)
        Else
            'Get the last row filled detail in Model view cosine sheet. As it might contain datas.
            iRowFilledInModelViewDC = FnGetNumberofRows(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, 1, 1)

            SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColModelViewName, PART_LCS_VIEW_NAME)

            SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSXx, adIdentityMatrix(0).ToString)
            SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSXy, adIdentityMatrix(1).ToString)
            SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSXz, adIdentityMatrix(2).ToString)

            SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSYx, adIdentityMatrix(3).ToString)
            SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSYy, adIdentityMatrix(4).ToString)
            SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSYz, adIdentityMatrix(5).ToString)

            SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSZx, adIdentityMatrix(6).ToString)
            SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSZy, adIdentityMatrix(7).ToString)
            SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSZz, adIdentityMatrix(8).ToString)

        End If


    End Sub
    '****************************************************************************************************************************************************
    'function to Check if there is Bent in the part
    'Description        : Check if there is Bent
    'Function Name      : FnChkIfPartIsBent
    'Input Parameter    : objPart
    'Output Parameter   : Boolean
    '****************************************************************************************************************************************************
    Public Function FnChkIfBodyIsBent(ByVal objPart As Part, ByVal objBody As Body) As Boolean
        Dim aoNonHoleCylFace As Face() = Nothing
        Dim objNonHoleCylFacePair As Face = Nothing
        'For Each objBody As Body In objPart.Bodies()
        If FnGetStringUserAttribute(objBody, _SHAPE_ATTR_NAME) = FLAT Then
            aoNonHoleCylFace = FnGetColOfNonHoleCylFaces(objPart, objBody)
            For Each objFace As Face In objBody.GetFaces()
                If objFace.SolidFaceType = Face.FaceType.Cylindrical Then
                    If Not FnCheckIftheFaceIsAHoleFace(objPart, objFace) Then
                        If Not aoNonHoleCylFace Is Nothing Then
                            If Not FnGetConcentricCylFace(objPart, objFace) Is Nothing Then
                                objNonHoleCylFacePair = FnGetConcentricCylFace(objPart, objFace)
                                If FnChkIfConcentricCylFace(objFace, objNonHoleCylFacePair, BENT_BRACKET_CONCENTRIC_FACES_DISTANCE_TOLERANCE) Then
                                    FnChkIfBodyIsBent = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        'Next
        FnChkIfBodyIsBent = False
    End Function
    '****************************************************************************************************************************************************
    'function to get the Collection of Non-Hole cylinderical faces from the given solid body
    'Description        : Return the Shape attribute of the given body
    'Function Name      : FnGetColOfNonHoleCylFaces
    'Input Parameter    : objBody - Solid body whoes face are to be determined
    'Output Parameter   : collection of Non-hole cylinderical faces
    '****************************************************************************************************************************************************
    Public Function FnGetColOfNonHoleCylFaces(ByVal objPart As Part, ByVal objBody As Body) As Face()
        Dim objEdgeVrt1 As Point3d = Nothing
        Dim objEdgeVrt2 As Point3d = Nothing
        Dim aoNonHoleCylFaces() As Face = Nothing

        If Not objBody Is Nothing Then
            For Each objface As Face In objBody.GetFaces()
                If objface.SolidFaceType = Face.FaceType.Cylindrical Then
                    If Not FnCheckIftheFaceIsAHoleFace(objPart, objface) Then
                        For Each objEdge As Edge In objface.GetEdges()
                            If objEdge.SolidEdgeType = Edge.EdgeType.Circular Or objEdge.SolidEdgeType = Edge.EdgeType.Elliptical Then
                                'Get the Vertices of a circular edge
                                objEdge.GetVertices(objEdgeVrt1, objEdgeVrt2)
                                'When the Vertices are not equal, Its a non-hole cylinderical face
                                If Not objEdgeVrt1.Equals(objEdgeVrt2) Then
                                    If aoNonHoleCylFaces Is Nothing Then
                                        ReDim Preserve aoNonHoleCylFaces(0)
                                        aoNonHoleCylFaces(0) = objface
                                    Else
                                        ReDim Preserve aoNonHoleCylFaces(UBound(aoNonHoleCylFaces) + 1)
                                        aoNonHoleCylFaces(UBound(aoNonHoleCylFaces)) = objface
                                    End If
                                    Exit For
                                End If
                            End If
                        Next
                    End If

                End If
            Next
        End If
        FnGetColOfNonHoleCylFaces = aoNonHoleCylFaces
    End Function
    '****************************************************************************************************************************************************
    'function to Get Concentric cylindrical face
    'Description        : Get Concentric cylindrical face
    'Function Name      : FnGetConcentricCylFace
    'Input Parameter    : objRefCylFace – Reference Cylindrical face object
    'Output Parameter   : concentric cylindrical face
    '****************************************************************************************************************************************************
    Public Function FnGetConcentricCylFace(ByVal objPart As Part, ByVal objRefCylFace As Face) As Face
        'Dim objPart As Part = Nothing
        Dim dFaceCenter As Double() = Nothing
        Dim dRefCylFaceCenter As Double() = Nothing
        Dim dEdgeCenter As Double() = Nothing
        Dim dRefEdgeCenter As Double() = Nothing
        Dim objBody As Body = Nothing

        objBody = objRefCylFace.GetBody()
        If objBody.IsSolidBody Then

            For Each objFace As Face In objBody.GetFaces()
                If objFace.SolidFaceType = Face.FaceType.Cylindrical Then
                    If Not FnCheckIftheFaceIsAHoleFace(objPart, objFace) Then
                        If objFace.Tag <> objRefCylFace.Tag Then

                            dFaceCenter = FnGetFaceCenter(objFace)
                            dRefCylFaceCenter = FnGetFaceCenter(objRefCylFace)
                            'If Round(dFaceCenter(0), 6) = Round(dRefCylFaceCenter(0), 6) And _
                            '    Round(dFaceCenter(1), 6) = Round(dRefCylFaceCenter(1), 6) And _
                            '    Round(dFaceCenter(2), 6) = Round(dRefCylFaceCenter(2), 6) Then
                            '    FnGetConcentricCylFace = objFace
                            '    Exit Function
                            'End If
                            If FnChkIfConcentricCylFace(objFace, objRefCylFace, BENT_BRACKET_CONCENTRIC_FACES_DISTANCE_TOLERANCE) Then
                                FnGetConcentricCylFace = objFace
                                Exit Function
                            End If
                        End If
                    End If

                End If
            Next
            'Next
        End If
        'End If
        '    End If
        'End If
        FnGetConcentricCylFace = Nothing
    End Function

    '****************************************************************************************************************************************************
    'function to Check if they are concentric face
    'Description        : Check if they are concentric face
    'Function Name      : FnChkIfConcentricCylFace
    'Input Parameter    : face1 and face2
    'Output Parameter   : Boolean
    '****************************************************************************************************************************************************
    Public Function FnChkIfConcentricCylFace(ByVal objFace1 As Face, ByVal objFace2 As Face, dTolerance As Double) As Boolean

        Dim dFaceCenter1 As Double() = Nothing
        Dim dFaceCenter2 As Double() = Nothing
        'Dim dAxisVecFace1 As Double() = Nothing
        'Dim dAxisVecFace2 As Double() = Nothing
        'Dim objLargerCylFace As Face = Nothing
        'Dim objCircularEdge As Edge = Nothing
        'Dim adCenterPointOfCircularEdge As Double() = Nothing

        'Dim adRotationMatrix() As Double = Nothing
        'Dim bIgnoreCheckAlongXAxis As Boolean = False
        'Dim bIgnoreCheckAlongYAxis As Boolean = False
        'Dim bIgnoreCheckAlongZAxis As Boolean = False
        Dim objPoint1 As Point3d = Nothing
        Dim objPoint2 As Point3d = Nothing
        'Dim objCenterPoint As Point3d = Nothing
        Dim objPointOnFaceCenter1 As NXOpen.Point = Nothing
        Dim objPointOnFaceCenter2 As NXOpen.Point = Nothing
        'Dim objCenterPointOfCirEdge As NXOpen.Point = Nothing
        Dim objPart As Part = Nothing
        'Dim objVectorBetCirEdgeCenterAndFaceCenter1 As Vector3d = Nothing
        'Dim objVectorBetCirEdgeCenterAndFaceCenter2 As Vector3d = Nothing
        Dim adAxisVecFace1 As Double() = Nothing
        Dim adAxisVecFace2 As Double() = Nothing
        Dim dDistance As Double = -1

        objPart = objFace1.OwningPart

        dFaceCenter1 = FnGetFaceCenter(objFace1)
        dFaceCenter2 = FnGetFaceCenter(objFace2)
        adAxisVecFace1 = FnGetAxisVecCylFace(objPart, objFace1)
        adAxisVecFace2 = FnGetAxisVecCylFace(objPart, objFace2)

        'Logic changed on Mar-17-2017
        '1.	Check if the given two cylindrical face are concave and convex and are within the tolerance of 12.5mm
        '2.	If both the cylindrical face centers have the same points, then they are considered as concentric
        '3. If not Check if Axis vectors of two cylindircal face are parallel anti parallel.
        '4. Then check if the distance between two face centers are within 5mm tolerance (CONCENTRIC_DISTANCE_TOLERANCE_BET_FACE_CENTERS)
        '5. If so ascertains it as a Concentric faces
        If FnCheckIfFaceAreOppositeCurvature(objFace1, objFace2) Then
            'Check if the faces are within the tolerance value.
            If FnChkIfFacesAreWithinToleranceDistance(objFace1, objFace2, dTolerance) Then
                If (Abs(Round(dFaceCenter1(0), 2) - Round(dFaceCenter2(0), 2)) = 0 And
                     Abs(Round(dFaceCenter1(1), 2) - Round(dFaceCenter2(1), 2)) = 0 And
                     Abs(Round(dFaceCenter1(2), 2) - Round(dFaceCenter2(2), 2)) = 0) Then
                    FnChkIfConcentricCylFace = True
                    Exit Function
                Else
                    'Check if the axis vectors are parallel antiparallel
                    If FnParallelAntiParallelCheck(adAxisVecFace1, adAxisVecFace2, bConcentricCheck:=True) Then
                        objPoint1 = New Point3d(dFaceCenter1(0), dFaceCenter1(1), dFaceCenter1(2))
                        objPoint2 = New Point3d(dFaceCenter2(0), dFaceCenter2(1), dFaceCenter2(2))
                        objPointOnFaceCenter1 = objPart.Points.CreatePoint(objPoint1)
                        objPointOnFaceCenter2 = objPart.Points.CreatePoint(objPoint2)

                        dDistance = FnGetDistanceBetweenObjects(objPart, objPointOnFaceCenter1, objPointOnFaceCenter2)
                        If dDistance <= CONCENTRIC_DISTANCE_TOLERANCE_BET_FACE_CENTERS Then
                            FnChkIfConcentricCylFace = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If



        'Code commented on Mar-17-2017
        ' ''Logic changed on Mar-15-2017
        ' ''1.	Check if the given two cylindrical face are concave and convex and are within the tolerance of 12.5mm
        ' ''2.	If both the cylindrical face centers have the same points, then they are considered as concentric
        ' ''3.	If not, Get the largest Cylindrical face among the given faces.
        ' ''4.	Choose the Circular edge from largest Cylindrical face whose edge radius is equal to the radius of this biggest cylindrical face.
        ' ''5.	Find the midpoint of the circular edge
        ' ''6.	Also Find the face centers of both the cylindrical faces
        ' ''7.	Create vector1 connecting Circular edge of largest cylindrical face and Face center of cylindrical face1
        ' ''8.	Create vector2 connecting Circular edge of largest cylindrical face and Face center of cylindrical face2
        ' ''9.	Angle between these two vectors should be within 10 degrees tolerance, which constitute Concentric Cylindrical faces.

        ''If FnCheckIfFaceAreOppositeCurvature(objFace1, objFace2) Then
        ''    'Check if the faces are within the tolerance value.
        ''    If FnChkIfFacesAreWithinToleranceDistance(objFace1, objFace2, dTolerance) Then
        ''        If (Abs(Round(dFaceCenter1(0), 2) - Round(dFaceCenter2(0), 2)) = 0 And _
        ''             Abs(Round(dFaceCenter1(1), 2) - Round(dFaceCenter2(1), 2)) = 0 And _
        ''             Abs(Round(dFaceCenter1(2), 2) - Round(dFaceCenter2(2), 2)) = 0) Then
        ''            FnChkIfConcentricCylFace = True
        ''            Exit Function

        ''        Else
        ''            'Get the largest cylindrical face among both the face.
        ''            objLargerCylFace = FnGetLargerCylFace(objFace1, objFace2)
        ''            If Not objLargerCylFace Is Nothing Then
        ''                'Get the cylindrical edge from the larger cylindrical face, whose radius is equal to the radius of larger cyl face.
        ''                objCircularEdge = FnGetCircularEdgeWithRadiusEqualToFaceRadius(objLargerCylFace)
        ''                If Not objCircularEdge Is Nothing Then
        ''                    'Get the center point of the circular edge
        ''                    adCenterPointOfCircularEdge = FnGetArcInfo(objCircularEdge.Tag).center
        ''                    If Not adCenterPointOfCircularEdge Is Nothing Then
        ''                        'Create vector between, center point of edge to each of the face centers.
        ''                        If (Not dFaceCenter1 Is Nothing) And (Not dFaceCenter2 Is Nothing) Then
        ''                            objPart = objFace1.OwningPart

        ''                            objCenterPoint = New Point3d(adCenterPointOfCircularEdge(0), adCenterPointOfCircularEdge(1), adCenterPointOfCircularEdge(2))
        ''                            objCenterPointOfCirEdge = objPart.Points.CreatePoint(objCenterPoint)

        ''                            objPoint1 = New Point3d(dFaceCenter1(0), dFaceCenter1(1), dFaceCenter1(2))
        ''                            objPoint2 = New Point3d(dFaceCenter2(0), dFaceCenter2(1), dFaceCenter2(2))
        ''                            objPointOnFaceCenter1 = objPart.Points.CreatePoint(objPoint1)
        ''                            objPointOnFaceCenter2 = objPart.Points.CreatePoint(objPoint2)
        ''                            objVectorBetCirEdgeCenterAndFaceCenter1 = FnGetVectorByTwoPoints(objPart, objCenterPointOfCirEdge, objPointOnFaceCenter1)
        ''                            objVectorBetCirEdgeCenterAndFaceCenter2 = FnGetVectorByTwoPoints(objPart, objCenterPointOfCirEdge, objPointOnFaceCenter2)

        ''                            adAxisVector1 = ({objVectorBetCirEdgeCenterAndFaceCenter1.X, objVectorBetCirEdgeCenterAndFaceCenter1.Y, objVectorBetCirEdgeCenterAndFaceCenter1.Z})
        ''                            adAxisVector2 = ({objVectorBetCirEdgeCenterAndFaceCenter2.X, objVectorBetCirEdgeCenterAndFaceCenter2.Y, objVectorBetCirEdgeCenterAndFaceCenter2.Z})

        ''                            'Check if the vectors are parallel or anti-parallel
        ''                            If FnParallelAntiParallelCheck(adAxisVector1, adAxisVector2) Then
        ''                                FnChkIfConcentricCylFace = True
        ''                                Exit Function
        ''                            End If
        ''                        End If

        ''                    End If
        ''                End If
        ''            End If
        ''        End If
        ''    End If
        ''End If

        'Code Commented on Mar-15-2017. Flaw in logic
        ''Code added Shanmugam Feb-09-2017
        ''If both the face centers has the same points, then they are considered as concentric.
        ''If not Create the vector between the two face center points. 
        ''Check if the vector formed between two points and cylindrical axis vector of face are parallel. If so they are concentric face
        'If FnCheckIfFaceAreOppositeCurvature(objFace1, objFace2) Then
        '    If (Abs(Round(dFaceCenter1(0), 2) - Round(dFaceCenter2(0), 2)) = 0 And _
        '        Abs(Round(dFaceCenter1(1), 2) - Round(dFaceCenter2(1), 2)) = 0 And _
        '        Abs(Round(dFaceCenter1(2), 2) - Round(dFaceCenter2(2), 2)) = 0) Then

        '        FnChkIfConcentricCylFace = True
        '        Exit Function
        '    Else
        '        'Create a vector between two face center points
        '        If (Not dFaceCenter1 Is Nothing) And (Not dFaceCenter2 Is Nothing) Then
        '            objPart = objFace1.OwningPart

        '            objPoint1 = New Point3d(dFaceCenter1(0), dFaceCenter1(1), dFaceCenter1(2))
        '            objPoint2 = New Point3d(dFaceCenter2(0), dFaceCenter2(1), dFaceCenter2(2))
        '            objPointOnFaceCenter1 = objPart.Points.CreatePoint(objPoint1)
        '            objPointOnFaceCenter2 = objPart.Points.CreatePoint(objPoint2)
        '            objVectorBetweenFaceCenter = FnGetVectorByTwoPoints(objPart, objPointOnFaceCenter1, objPointOnFaceCenter2)
        '            dAxisBetweenFaceCenter = ({objVectorBetweenFaceCenter.X, objVectorBetweenFaceCenter.Y, objVectorBetweenFaceCenter.Z})
        '            'Check if the vectors are parallel or anti-parallel
        '            If FnParallelAntiParallelCheck(dAxisVecFace1, dAxisBetweenFaceCenter) Then
        '                FnChkIfConcentricCylFace = True
        '                Exit Function
        '            End If
        '        End If
        '    End If
        'End If

        ''Code commented on Feb-09-2017. Flaw in logic. 
        ' ''Check if these two axis vectors are parallel
        ''If FnParallelCheck(dAxisVecFace1, dAxisVecFace2) Then

        ''    If FnGetCylFaceRadius(objFace1) <> FnGetCylFaceRadius(objFace2) Then
        ''        'Get the rotation matrix
        ''        adRotationMatrix = FnComputeOptimalRotationMatrix(objFace1.GetBody())
        ''        If Not adRotationMatrix Is Nothing Then
        ''            'COmpare the axis vector of face with the rotation matrix and check which axis of the rotation matrix is to be ignored.

        ''            If ((Round(adRotationMatrix(0), 6)) = (Round(dAxisVecFace1(0), 6))) And _
        ''                ((Round(adRotationMatrix(1), 6)) = (Round(dAxisVecFace1(1), 6))) And _
        ''                ((Round(adRotationMatrix(2), 6)) = (Round(dAxisVecFace1(2), 6))) Then
        ''                bIgnoreCheckAlongXAxis = True
        ''            End If
        ''            If ((Round(adRotationMatrix(3), 6)) = (Round(dAxisVecFace1(0), 6))) And _
        ''                ((Round(adRotationMatrix(4), 6)) = (Round(dAxisVecFace1(1), 6))) And _
        ''                ((Round(adRotationMatrix(5), 6)) = (Round(dAxisVecFace1(2), 6))) Then
        ''                bIgnoreCheckAlongYAxis = True
        ''            End If
        ''            If ((Round(adRotationMatrix(6), 6)) = (Round(dAxisVecFace1(0), 6))) And _
        ''                ((Round(adRotationMatrix(7), 6)) = (Round(dAxisVecFace1(1), 6))) And _
        ''                ((Round(adRotationMatrix(8), 6)) = (Round(dAxisVecFace1(2), 6))) Then
        ''                bIgnoreCheckAlongZAxis = True
        ''            End If
        ''            If bIgnoreCheckAlongXAxis Then
        ''                If (Abs(Round(dFaceCenter1(1), 2) - Round(dFaceCenter2(1), 2)) <= FACE_CENTER_TOLERANCE And _
        ''                    Abs(Round(dFaceCenter1(2), 2) - Round(dFaceCenter2(2), 2)) <= FACE_CENTER_TOLERANCE) Then
        ''                    FnChkIfConcentricCylFace = True
        ''                    Exit Function
        ''                End If
        ''            End If
        ''            If bIgnoreCheckAlongYAxis Then
        ''                If (Abs(Round(dFaceCenter1(0), 2) - Round(dFaceCenter2(0), 2)) <= FACE_CENTER_TOLERANCE And _
        ''                    Abs(Round(dFaceCenter1(2), 2) - Round(dFaceCenter2(2), 2)) <= FACE_CENTER_TOLERANCE) Then
        ''                    FnChkIfConcentricCylFace = True
        ''                    Exit Function
        ''                End If
        ''            End If
        ''            If bIgnoreCheckAlongZAxis Then
        ''                If (Abs(Round(dFaceCenter1(0), 2) - Round(dFaceCenter2(0), 2)) <= FACE_CENTER_TOLERANCE And _
        ''                    Abs(Round(dFaceCenter1(1), 2) - Round(dFaceCenter2(1), 2)) <= FACE_CENTER_TOLERANCE) Then
        ''                    FnChkIfConcentricCylFace = True
        ''                    Exit Function
        ''                End If
        ''            End If
        ''        End If
        ''    End If
        ''End If

        'Code commented
        ''instead of face center being equal, a tolerance of 1 is added.
        ''if two face centers are within the tolerance of 1 then they are considered as the concentric cylindrical face.
        'If (Abs(Round(dFaceCenter1(0), 2) - Round(dFaceCenter2(0), 2)) <= FACE_CENTER_TOLERANCE And _
        '    Abs(Round(dFaceCenter1(1), 2) - Round(dFaceCenter2(1), 2)) <= FACE_CENTER_TOLERANCE And _
        '    Abs(Round(dFaceCenter1(2), 2) - Round(dFaceCenter2(2), 2)) <= FACE_CENTER_TOLERANCE) Then

        '    FnChkIfConcentricCylFace = True
        '    Exit Function
        'End If
        FnChkIfConcentricCylFace = False
    End Function

    'Function to check if the face are within tolerance distance
    Function FnChkIfFacesAreWithinToleranceDistance(objFace1 As Face, objFace2 As Face, dTolerance As Double) As Boolean
        Dim dDistanceBetFaces As Double = -1

        If (Not objFace1 Is Nothing) And (Not objFace2 Is Nothing) Then
            dDistanceBetFaces = FnGetDistanceBetweenObjects(objFace1.OwningPart, objFace1, objFace2)
            If dDistanceBetFaces <> -1 Then
                If dDistanceBetFaces <= dTolerance Then
                    FnChkIfFacesAreWithinToleranceDistance = True
                    Exit Function
                End If
            End If
        End If
        FnChkIfFacesAreWithinToleranceDistance = False
    End Function
    'Function to check if the body has orthogonal faces
    Function FnChkIfBodyHasOrthogonalFace(objBody As Body) As Boolean
        Dim adVecObjFace As Double() = Nothing
        Dim adVecObjFaceToCompare As Double() = Nothing

        If Not objBody Is Nothing Then
            If objBody.IsSolidBody And (Not objBody.IsBlanked) Then
                For Each objFace As Face In objBody.GetFaces()
                    If objFace.SolidFaceType = Face.FaceType.Planar Then
                        'Validation added, that the face shoule not be a NC Part COntact Face
                        'If Not (FnGetStringUserAttribute(objFace, NC_Contact_FACE_ATTRIBUTE) = NC_PCF_ATTR_VALUE) Then
                        If (FnGetStringUserAttribute(objFace, NC_Contact_FACE_ATTRIBUTE) = "") Then
                            For Each objFaceToCompare As Face In objBody.GetFaces()
                                If objFaceToCompare.SolidFaceType = Face.FaceType.Planar Then
                                    'Validation added, that the face shoule not be a NC Part COntact Face
                                    'If Not (FnGetStringUserAttribute(objFaceToCompare, NC_Contact_FACE_ATTRIBUTE) = NC_PCF_ATTR_VALUE) Then
                                    If (FnGetStringUserAttribute(objFaceToCompare, NC_Contact_FACE_ATTRIBUTE) = "") Then
                                        If objFace.Tag <> objFaceToCompare.Tag Then

                                            adVecObjFace = FnGetFaceNormalVector(objFace)
                                            adVecObjFaceToCompare = FnGetFaceNormalVector(objFaceToCompare)

                                            If FnCheckOrthogalityOfTwoVectors(adVecObjFace, adVecObjFaceToCompare) Then
                                                FnChkIfBodyHasOrthogonalFace = True
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
        FnChkIfBodyHasOrthogonalFace = False
    End Function

    Function FnComputeOptimalRotationMatForRound(objPart As Part, objBody As Body) As Double()
        'This function will be called for solid body whoes shape is ROUND and which doesnot have any orthogonal planar face.

        '1. Compute the larger cylindrical face from the solid body
        '2. Cylindrical axis of this larger cylindrical face forms first vector.
        '3. Second vector is formed by joining the mid point and the vertex of this larger cylindrical face(one of the circular edge)
        '4. Check if these two vector are orthogonal.
        '5. Compute the third vector using formulae
        '6. These three vector forms the rotation matrix
        Dim objMaxDiaCylFace As Face = Nothing
        Dim adCylFaceAxisVec() As Double = Nothing
        Dim adSecondVec() As Double = Nothing
        Dim objVertex1 As Point3d = Nothing
        Dim objVertex2 As Point3d = Nothing
        Dim objCreatedVrtPt1 As NXOpen.Point = Nothing
        Dim adCenter() As Double = Nothing
        Dim objCenterPoint As Point3d = Nothing
        Dim objCreatedCenterPt As NXOpen.Point = Nothing
        Dim adRotationMatrix() As Double = Nothing

        If Not objBody Is Nothing Then
            If objBody.IsSolidBody Then
                'Check if Shape atribute is ROUND or RND TUBG
                If (FnGetStringUserAttribute(objBody, _SHAPE_ATTR_NAME) = ROUND_SHAPE) Or (FnGetStringUserAttribute(objBody, _SHAPE_ATTR_NAME) = ROUND_TUBING_SHAPE) Then
                    'Get the maximum dia cylindrical face from the body
                    objMaxDiaCylFace = FnGetMaxDiaFace(objBody)
                    If Not objMaxDiaCylFace Is Nothing Then
                        'Get the cylindrical axis of maximum dia cylindrical face (First Vector along cylindrical axis)
                        adCylFaceAxisVec = FnGetAxisVecCylFace(objPart, objMaxDiaCylFace)
                        If Not adCylFaceAxisVec Is Nothing Then

                            For Each objCirEdge As Edge In objMaxDiaCylFace.GetEdges()
                                If objCirEdge.SolidEdgeType = Edge.EdgeType.Circular Then
                                    'Find one of the vertex point on the circular edge of Maximum dia cylindrical face.
                                    objCirEdge.GetVertices(objVertex1, objVertex2)
                                    objCreatedVrtPt1 = objPart.Points.CreatePoint(objVertex1)
                                    'Get the center point of the circular edge for whose vertex point was indentified.
                                    adCenter = FnGetArcInfo(objCirEdge.Tag).center
                                    objCenterPoint = New Point3d(adCenter(0), adCenter(1), adCenter(2))
                                    objCreatedCenterPt = objPart.Points.CreatePoint(objCenterPoint)
                                    Exit For
                                End If
                            Next
                            'Create the second vector which is formed from the center point of circular edge and Vertex point of circular edge. 
                            'This circular edge are obtained from maximum dia cylindrical face
                            adSecondVec = FnGetVectorDirCosByTwoPoints(objPart, objCreatedCenterPt, objCreatedVrtPt1)
                            If Not adSecondVec Is Nothing Then
                                'Check the orthogonality of two vectors
                                If FnCheckOrthogalityOfTwoVectors(adCylFaceAxisVec, adSecondVec) Then
                                    'Compute the third vector from NX API
                                    adRotationMatrix = FnCreateRotationMatrixFromTwoVec(adSecondVec, adCylFaceAxisVec)
                                    'Store the rotation matirx in a structure to be used in over all rotation matrix computation of part
                                    sPopulateOrientationInfo(objBody, adRotationMatrix, Nothing, bIsRoundShape:=True)
                                    FnComputeOptimalRotationMatForRound = adRotationMatrix
                                    Exit Function
                                End If
                            End If
                        End If

                    End If
                End If
            End If
        End If
        FnComputeOptimalRotationMatForRound = adRotationMatrix
    End Function
    'Function to get the maximum dia cylindrical face in a given body
    Public Function FnGetMaxDiaFace(ByVal objBody As Body) As Face
        Dim dFaceDia As Double = Nothing
        Dim dFaceMaxDia As Double = Nothing
        Dim objMaxDiaFace As Face = Nothing

        If Not objBody Is Nothing Then
            If objBody.IsSolidBody Then
                For Each objCylFace As Face In objBody.GetFaces()
                    If objCylFace.SolidFaceType = Face.FaceType.Cylindrical Then
                        dFaceDia = Round((FnGetCylFaceRadius(objCylFace) * 2), 1)
                        If dFaceMaxDia = Nothing Then
                            dFaceMaxDia = dFaceDia
                            objMaxDiaFace = objCylFace
                        Else
                            If dFaceDia > dFaceMaxDia Then
                                dFaceMaxDia = dFaceDia
                                objMaxDiaFace = objCylFace
                            End If
                        End If
                    End If
                Next
            End If
        End If
        FnGetMaxDiaFace = objMaxDiaFace
    End Function
    Function FnConvertDisplayObjectToFace(aoAllDisplayObject() As DisplayableObject) As Face()
        Dim objFace As Face = Nothing
        Dim aoAllFace() As Face = Nothing

        If Not aoAllDisplayObject Is Nothing Then
            For Each objDisp As DisplayableObject In aoAllDisplayObject
                objFace = CType(objDisp, Face)
                If aoAllFace Is Nothing Then
                    ReDim Preserve aoAllFace(0)
                    aoAllFace(0) = objFace
                Else
                    ReDim Preserve aoAllFace(UBound(aoAllFace) + 1)
                    aoAllFace(UBound(aoAllFace)) = objFace
                End If
            Next
        End If
        FnConvertDisplayObjectToFace = aoAllFace
    End Function
    'Function to populate the Minimum and Maximum X, Y, Z extentends for all solid body
    Sub sPopulateMinMaxXYZForAllBody(ByVal objPart As Part, ByRef adRotationMatrix() As Double)

        Dim objBody As Body = Nothing
        Dim sBodyName As String = ""
        Dim aCompCol() As Component = FnGetAllComponentsInSession()
        Dim adMinMaxXYZ() As Double = Nothing
        Dim aoPartDesignMembers() As NXOpen.Features.Feature = Nothing
        Dim bIsValidSolidBody As Boolean = False
        Dim aoAllValidBody() As Body = Nothing

        If Not aCompCol Is Nothing Then
            For Each objComp As Component In aCompCol
                If Not FnGetPartFromComponent(objComp) Is Nothing Then
                    FnLoadPartFully(FnGetPartFromComponent(objComp))

                    aoAllValidBody = FnGetValidBodyForOEM(FnGetPartFromComponent(objComp), _sOemName)
                    If Not aoAllValidBody Is Nothing Then
                        For Each body As Body In aoAllValidBody
                            If Not body Is Nothing Then
                                If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        sBodyName = body.JournalIdentifier
                                        objBody = body
                                    Else
                                        sBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                        objBody = CType(objComp.FindOccurrence(body), Body)
                                    End If
                                ElseIf (_sOemName = DAIMLER_OEM_NAME) Then
                                    If _sDivision = TRUCK_DIVISION Then
                                        If objComp Is objPart.ComponentAssembly.RootComponent Then
                                            sBodyName = body.JournalIdentifier
                                            objBody = body
                                        Else
                                            sBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                            objBody = CType(objComp.FindOccurrence(body), Body)
                                        End If
                                    ElseIf _sDivision = CAR_DIVISION Then
                                        If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                            'Weldment in car
                                            sBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                            objBody = CType(objComp.FindOccurrence(body), Body)
                                        Else
                                            'Component in car
                                            sBodyName = body.JournalIdentifier
                                            objBody = body
                                        End If
                                    End If
                                ElseIf (_sOemName = FIAT_OEM_NAME) Then
                                    'Check if the component is a child component in weldment
                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        'Fiat component
                                        objBody = body
                                        sBodyName = objBody.JournalIdentifier
                                    Else
                                        'This is a Fiat Weldment child component
                                        objBody = CType(objComp.FindOccurrence(body), Body)
                                        If Not objBody Is Nothing Then
                                            'when populating body name, give Body journal identifier and Child component journalidentifier
                                            sBodyName = objBody.JournalIdentifier & " " & objComp.JournalIdentifier
                                        End If
                                    End If
                                End If

                                If Not objBody Is Nothing Then
                                    'Find the bounding box extends minimum and maximum X Y Z values for a body in given orientation
                                    adMinMaxXYZ = FnFindMinMaxXYZForBody(objBody, adRotationMatrix)
                                    If Not adMinMaxXYZ Is Nothing Then
                                        ReDim Preserve _aoStructMinMaxXYZ(_iSolidBodyIndex)
                                        _aoStructMinMaxXYZ(_iSolidBodyIndex).sBodyName = sBodyName
                                        _aoStructMinMaxXYZ(_iSolidBodyIndex).dMinX = adMinMaxXYZ(0)
                                        _aoStructMinMaxXYZ(_iSolidBodyIndex).dMaxX = adMinMaxXYZ(1)
                                        _aoStructMinMaxXYZ(_iSolidBodyIndex).dMinY = adMinMaxXYZ(2)
                                        _aoStructMinMaxXYZ(_iSolidBodyIndex).dMaxY = adMinMaxXYZ(3)
                                        _aoStructMinMaxXYZ(_iSolidBodyIndex).dMinZ = adMinMaxXYZ(4)
                                        _aoStructMinMaxXYZ(_iSolidBodyIndex).dMaxZ = adMinMaxXYZ(5)
                                        _iSolidBodyIndex = _iSolidBodyIndex + 1
                                    End If
                                End If
                            End If

                        Next
                    End If
                End If
            Next
        Else

            aoAllValidBody = FnGetValidBodyForOEM(objPart, _sOemName)
            If Not aoAllValidBody Is Nothing Then
                For Each body As Body In aoAllValidBody
                    If Not body Is Nothing Then
                        objBody = body
                        sBodyName = body.JournalIdentifier
                        adMinMaxXYZ = FnFindMinMaxXYZForBody(objBody, adRotationMatrix)
                        If Not adMinMaxXYZ Is Nothing Then
                            ReDim Preserve _aoStructMinMaxXYZ(_iSolidBodyIndex)
                            _aoStructMinMaxXYZ(_iSolidBodyIndex).sBodyName = sBodyName
                            _aoStructMinMaxXYZ(_iSolidBodyIndex).dMinX = adMinMaxXYZ(0)
                            _aoStructMinMaxXYZ(_iSolidBodyIndex).dMaxX = adMinMaxXYZ(1)
                            _aoStructMinMaxXYZ(_iSolidBodyIndex).dMinY = adMinMaxXYZ(2)
                            _aoStructMinMaxXYZ(_iSolidBodyIndex).dMaxY = adMinMaxXYZ(3)
                            _aoStructMinMaxXYZ(_iSolidBodyIndex).dMinZ = adMinMaxXYZ(4)
                            _aoStructMinMaxXYZ(_iSolidBodyIndex).dMaxZ = adMinMaxXYZ(5)
                            _iSolidBodyIndex = _iSolidBodyIndex + 1
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    'Function to find the Minimum and Maximum X Y Z extends for a body
    Function FnFindMinMaxXYZForBody(objBody As Body, ByRef adRotationMatrix() As Double) As Double()

        '1.Find the 8 Cornor points of solid body
        '2.Rotate all the 8 cornor points with respect to rotationmatrix
        '3.Find the Minimum and Maximum X Y Z extents for this solid body

        Dim adBoundingBoxCorner(7, 2) As Double
        Dim adVector1(2) As Double
        Dim adRotatedBBVector() As Double = Nothing
        Dim adAllXValues(7) As Double
        Dim adAllYValues(7) As Double
        Dim adAllZValues(7) As Double
        Dim adMinMaxXYZ(5) As Double

        'Compute Bounding Box 8 corner coordinates
        adBoundingBoxCorner = FnGetBoundingBoxCorners(objBody, adRotationMatrix)
        If Not adBoundingBoxCorner Is Nothing Then
            'Rotate each Bounding Box corners into the view
            For iIndex As Integer = 0 To 7
                adVector1(0) = adBoundingBoxCorner(iIndex, 0)
                adVector1(1) = adBoundingBoxCorner(iIndex, 1)
                adVector1(2) = adBoundingBoxCorner(iIndex, 2)

                adRotatedBBVector = FnGetRotatedVector(adRotationMatrix, adVector1)

                'Collect All the X, Y ,Z co-ordinates
                adAllXValues(iIndex) = adRotatedBBVector(0)
                adAllYValues(iIndex) = adRotatedBBVector(1)
                adAllZValues(iIndex) = adRotatedBBVector(2)
            Next

            'Find the minimum and Maximum X,Y,Z values from the rotated Bounding Box COrner values
            Array.Sort(adAllXValues)

            adMinMaxXYZ(0) = adAllXValues(0)
            adMinMaxXYZ(1) = adAllXValues(7)

            Array.Sort(adAllYValues)
            adMinMaxXYZ(2) = adAllYValues(0)
            adMinMaxXYZ(3) = adAllYValues(7)

            Array.Sort(adAllZValues)
            adMinMaxXYZ(4) = adAllZValues(0)
            adMinMaxXYZ(5) = adAllZValues(7)
            FnFindMinMaxXYZForBody = adMinMaxXYZ
            Exit Function
        End If
        FnFindMinMaxXYZForBody = Nothing
    End Function

    'Function to find the global extents X Y Z for all solid body in a weldment
    Function FnFindGlobalExtentsXYZForWeldment(objPart As Part, ByRef adRotationMatrix() As Double) As Double()

        Dim dMinXExtent As Double = 0
        Dim dMinX As Double = 0
        Dim dMinYExtent As Double = 0
        Dim dMinY As Double = 0
        Dim dMinZExtent As Double = 0
        Dim dMinZ As Double = 0
        Dim dMaxX As Double = 0
        Dim dMaxXExtent As Double = 0
        Dim dMaxY As Double = 0
        Dim dMaxYExtent As Double = 0
        Dim dMaxZ As Double = 0
        Dim dMaxZExtent As Double = 0
        Dim dXExtent As Double = 0
        Dim dYExtent As Double = 0
        Dim dZExtent As Double = 0
        Dim adExtent(2) As Double

        'Populate Minimum and Maximum X Y Z extends for all solid body in a weldment
        _aoStructMinMaxXYZ = Nothing
        _iSolidBodyIndex = 0
        sPopulateMinMaxXYZForAllBody(objPart, adRotationMatrix)

        If _iSolidBodyIndex > 0 Then
            For iIndex As Integer = 0 To UBound(_aoStructMinMaxXYZ)
                'Find the Minimum X extents
                dMinX = _aoStructMinMaxXYZ(iIndex).dMinX
                If iIndex = 0 Then
                    dMinXExtent = dMinX
                Else
                    If dMinX < dMinXExtent Then
                        dMinXExtent = dMinX
                    End If
                End If
                'Find the Minimum Y extents
                dMinY = _aoStructMinMaxXYZ(iIndex).dMinY
                If iIndex = 0 Then
                    dMinYExtent = dMinY
                Else
                    If dMinY < dMinYExtent Then
                        dMinYExtent = dMinY
                    End If
                End If
                'Find the Minimum Z extents
                dMinZ = _aoStructMinMaxXYZ(iIndex).dMinZ
                If iIndex = 0 Then
                    dMinZExtent = dMinZ
                Else
                    If dMinZ < dMinZExtent Then
                        dMinZExtent = dMinZ
                    End If
                End If

                'Find the Maximum X extents
                dMaxX = _aoStructMinMaxXYZ(iIndex).dMaxX
                If iIndex = 0 Then
                    dMaxXExtent = dMaxX
                Else
                    If dMaxX > dMaxXExtent Then
                        dMaxXExtent = dMaxX
                    End If
                End If
                'Find the Maximum Y Extents
                dMaxY = _aoStructMinMaxXYZ(iIndex).dMaxY
                If iIndex = 0 Then
                    dMaxYExtent = dMaxY
                Else
                    If dMaxY > dMaxYExtent Then
                        dMaxYExtent = dMaxY
                    End If
                End If
                'Find the Maximum Z Extents
                dMaxZ = _aoStructMinMaxXYZ(iIndex).dMaxZ
                If iIndex = 0 Then
                    dMaxZExtent = dMaxZ
                Else
                    If dMaxZ > dMaxZExtent Then
                        dMaxZExtent = dMaxZ
                    End If
                End If
            Next
            'Find the X Y Z Extents from Minimum and Maximum X Y Z extents
            dXExtent = dMaxXExtent - dMinXExtent
            dYExtent = dMaxYExtent - dMinYExtent
            dZExtent = dMaxZExtent - dMinZExtent

            adExtent(0) = dXExtent
            adExtent(1) = dYExtent
            adExtent(2) = dZExtent
        End If
        FnFindGlobalExtentsXYZForWeldment = adExtent
    End Function
    'Code modified on May-03-2018
    'Create a custom model view based on the optimal rotation matrix
    Sub sCreateCustomModelViewForOptimalRotMat(objpart As Part, adOptimalRotMat() As Double, sViewName As String)
        Dim sBodyName As String = ""
        Dim sModelViewName As String = ""

        If Not adOptimalRotMat Is Nothing Then
            sModelViewName = sViewName
            'Change application to modeling view
            sChangeApplication(1)
            'Check if modeling view is present
            If FnChkIfModelingViewPresent(objpart, sModelViewName) Then
                sRefreshLayout(objpart, sModelViewName)
                sDeleteModellingView(objpart, sModelViewName)
                sSetDefaultLayout(objpart)
            End If
            sCreateCustomModellingView(objpart, adOptimalRotMat(0), adOptimalRotMat(1), adOptimalRotMat(2), adOptimalRotMat(3),
                                        adOptimalRotMat(4), adOptimalRotMat(5), adOptimalRotMat(6), adOptimalRotMat(7), adOptimalRotMat(8),
                                        sModelViewName, dScale:=1)
            'Code commented on Sep-01-2017
            'If there is no drawing sheet switch api will throw an error
            'Change application to drafting view
            'sChangeApplication(2)
        End If
    End Sub

    Public Sub sCreateModelingView(ByVal objPart As Part, ByVal rotmatrix As Matrix3x3, ByVal sCustViewName As String,
                                    Optional ByVal sDefaultRefView As String = "FRONT")
        sChangeApplication(1)
        Dim objView As NXOpen.View = Nothing
        Dim translation As Point3d = New Point3d(0, 0, 0)
        Dim dScale As Double = 1
        If objPart.ModelingViews.WorkView Is Nothing Then
            objPart.ModelingViews.WorkView.SetRotationTranslationScale(rotmatrix, translation, dScale)
            objView = objPart.Views.SaveAs(objPart.ModelingViews.WorkView, sCustViewName, False, False)
        Else
            For Each objModelView As ModelingView In objPart.ModelingViews
                If Not objModelView Is Nothing Then
                    If objModelView.Name.ToUpper = sDefaultRefView Then
                        sReplaceViewInLayout(objPart, objModelView)
                        objModelView.SetRotationTranslationScale(rotmatrix, translation, dScale)
                        'objModelView.Fit()
                        'objModelView.RenderingStyle = View.RenderingStyleType.Shaded
                        objPart.Views.SaveAs(objModelView, sCustViewName, False, False)
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub

    'Code added Jun-29-2017
    'Function to check if part is to be details or not. Based on this PSD will be executed on this part
    Function FnCheckIfPartIsToBeDetailed(objPart As Part) As Boolean
        Dim sDetailValue As String = ""
        If Not objPart Is Nothing Then
            FnLoadPartFully(objPart)
            sDetailValue = FnGetStringUserAttribute(objPart, PART_DETAIL_ATTRIBUTE)
            If sDetailValue <> "" Then
                If sDetailValue.ToUpper = PART_DETAIL_YES.ToUpper Then
                    FnCheckIfPartIsToBeDetailed = True
                    sWriteToLogFile(objPart.Leaf.ToUpper & " Part needs to be detailed")
                    Exit Function
                End If
            End If
        End If
        sWriteToLogFile(objPart.Leaf.ToUpper & " Part not to be detailed")

        FnCheckIfPartIsToBeDetailed = False
    End Function

    'Code added Oct-05-2017
    'Function to read Stock Size value from model and write it to an attribute in Model VECTRA category
    Sub sAddStockSizeAttribute(objPart As Part, aoAllCompInSession() As Component, sOemName As String)
        Dim sPartName As String = ""
        Dim sStockFromClient As String = ""
        Dim sModifiedStockSize As String = ""
        Dim objChildPart As Part = Nothing
        Dim bCheckStockSize As Boolean = False
        Dim bIsWeldment As Boolean = False
        Dim objGeoComp As Component = Nothing
        Dim objGeoPart As Part = Nothing
        Dim objparentContainerComp As Component = Nothing
        Dim objParentContainerPart As Part = Nothing

        If Not objPart Is Nothing Then
            If (sOemName = DAIMLER_OEM_NAME) Or (sOemName = FIAT_OEM_NAME) Then
                sPartName = objPart.Leaf.ToString().ToUpper
                If sOemName = DAIMLER_OEM_NAME Then
                    If FnCheckIfThisIsAWeldment(sPartName) Then
                        bIsWeldment = True

                    Else
                        'Code added Apr-03-2019
                        'For the latest Daimler project V297, Rhomass and Werkstoff attributes will be in GEO level. 
                        'For previous project these two attributes were in container level
                        'objPart is a container part. get the immediate child component to it and fetch the attribute

                        If Not objPart.ComponentAssembly.RootComponent Is Nothing Then
                            If Not objPart.ComponentAssembly.RootComponent.GetChildren Is Nothing Then
                                objGeoComp = objPart.ComponentAssembly.RootComponent.GetChildren(0)
                                If Not objGeoComp Is Nothing Then
                                    objGeoPart = FnGetPartFromComponent(objGeoComp)
                                    If Not objGeoPart Is Nothing Then
                                        FnLoadPartFully(objGeoPart)
                                        'Component
                                        'To get the Rhomass from GEO level, use objGeoPart
                                        'TO get the Rhomass from Container lever use objPart
                                        sStockFromClient = FnGetStringUserAttribute(objGeoPart, _CLIENT_STOCK_SIZE_ATTR)
                                        If sStockFromClient <> "" Then
                                            sModifiedStockSize = FnSegregateStockSize(sStockFromClient, sOemName)
                                            If sModifiedStockSize <> "" Then
                                                sSetStringUserAttribute(objPart, _STOCK_SIZE_METRIC, sModifiedStockSize)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                    End If
                ElseIf sOemName = FIAT_OEM_NAME Then
                    If FnChkPartisWeldmentBasedOnAttr(objPart) Then
                        bIsWeldment = True
                    Else
                        'Component
                        sStockFromClient = FnGetStringUserAttribute(objPart, _CLIENT_STOCK_SIZE_ATTR)
                        If sStockFromClient <> "" Then
                            sModifiedStockSize = FnSegregateStockSize(sStockFromClient, sOemName)
                            If sModifiedStockSize <> "" Then
                                sSetStringUserAttribute(objPart, _STOCK_SIZE_METRIC, sModifiedStockSize)
                            End If
                        End If
                    End If
                End If
                If bIsWeldment Then
                    'Welded Component
                    If Not aoAllCompInSession Is Nothing Then
                        For Each objChildComp As Component In aoAllCompInSession
                            bCheckStockSize = False
                            If sOemName = DAIMLER_OEM_NAME Then
                                If _sDivision = CAR_DIVISION Then
                                    'In case of Car divison, attribute will be added to the container level for component as well as weldment (first project)
                                    'For Daimler V297 project, Rhomass were added to the GEo level and not to then container level
                                    'If FnCheckIfThisIsAChildCompContainerInWeldment(objChildComp.DisplayName.ToUpper) Then
                                    If FnCheckIfThisIsAChildCompInWeldment(objChildComp, sOemName) Then
                                        'objChildComp is a GEO COmponent
                                        'Get the Parent container component for the Geo component


                                        objparentContainerComp = objChildComp.Parent
                                        If Not objparentContainerComp Is Nothing Then
                                            objParentContainerPart = FnGetPartFromComponent(objparentContainerComp)
                                            If Not objParentContainerPart Is Nothing Then
                                                objChildPart = FnGetPartFromComponent(objChildComp)
                                                If Not objChildPart Is Nothing Then
                                                    sStockFromClient = FnGetStringUserAttribute(objChildPart, _CLIENT_STOCK_SIZE_ATTR)
                                                    If sStockFromClient <> "" Then
                                                        sModifiedStockSize = FnSegregateStockSize(sStockFromClient, sOemName)
                                                        If sModifiedStockSize <> "" Then
                                                            sSetStringUserAttribute(objParentContainerPart, _STOCK_SIZE_METRIC, sModifiedStockSize)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                        'bCheckStockSize = True
                                    End If
                                ElseIf _sDivision = TRUCK_DIVISION Then
                                    'bCheckStockSize = True
                                    objChildPart = FnGetPartFromComponent(objChildComp)
                                    If Not objChildPart Is Nothing Then
                                        sStockFromClient = FnGetStringUserAttribute(objChildPart, _CLIENT_STOCK_SIZE_ATTR)
                                        If sStockFromClient <> "" Then
                                            sModifiedStockSize = FnSegregateStockSize(sStockFromClient, sOemName)
                                            If sModifiedStockSize <> "" Then
                                                sSetStringUserAttribute(objChildPart, _STOCK_SIZE_METRIC, sModifiedStockSize)
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf sOemName = FIAT_OEM_NAME Then
                                If FnCheckIfThisIsAChildCompInWeldment(objChildComp, sOemName) Then
                                    bCheckStockSize = True
                                    objChildPart = FnGetPartFromComponent(objChildComp)
                                    If Not objChildPart Is Nothing Then
                                        sStockFromClient = FnGetStringUserAttribute(objChildPart, _CLIENT_STOCK_SIZE_ATTR)
                                        If sStockFromClient <> "" Then
                                            sModifiedStockSize = FnSegregateStockSize(sStockFromClient, sOemName)
                                            If sModifiedStockSize <> "" Then
                                                sSetStringUserAttribute(objChildPart, _STOCK_SIZE_METRIC, sModifiedStockSize)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            'Code commented on Apr-03-2019
                            'Based on OEM, the fetching of the attribute is done at different level
                            'If bCheckStockSize Then
                            '    objChildPart = FnGetPartFromComponent(objChildComp)
                            '    If Not objChildPart Is Nothing Then
                            '        sStockFromClient = FnGetStringUserAttribute(objChildPart, _CLIENT_STOCK_SIZE_ATTR)
                            '        If sStockFromClient <> "" Then
                            '            sModifiedStockSize = FnSegregateStockSize(sStockFromClient, sOemName)
                            '            If sModifiedStockSize <> "" Then
                            '                sSetStringUserAttribute(objChildPart, _STOCK_SIZE_METRIC, sModifiedStockSize)
                            '            End If
                            '        End If
                            '    End If
                            'End If
                        Next
                    End If
                    'Code commented on Apr-03-2019
                    'Based on OEM, the fetching of the attribute is done at different level
                    ''Else
                    ''    'Component
                    ''    sStockFromClient = FnGetStringUserAttribute(objPart, _CLIENT_STOCK_SIZE_ATTR)
                    ''    If sStockFromClient <> "" Then
                    ''        sModifiedStockSize = FnSegregateStockSize(sStockFromClient, sOemName)
                    ''        If sModifiedStockSize <> "" Then
                    ''            sSetStringUserAttribute(objPart, _STOCK_SIZE_METRIC, sModifiedStockSize)
                    ''        End If
                    ''    End If
                End If
            End If
        End If
    End Sub

    Function FnSegregateStockSize(sStockFromClient As String, sOemName As String) As String
        Dim asStockValues() As String = Nothing
        Dim sFirstValue As String = Nothing
        Dim sSecondValue As String = Nothing
        Dim sThirdValue As String = Nothing
        Dim sFourthValue As String = Nothing
        Dim sModifiedStockSize As String = ""
        Dim dictOfChannelABValues As Dictionary(Of String, String) = Nothing
        Dim dictOfBeamABValues As Dictionary(Of String, String) = Nothing

        If Not sStockFromClient Is Nothing Then
            If (sOemName = DAIMLER_OEM_NAME) Then
                dictOfChannelABValues = FnChannelABStockValuesForDaimler()
                dictOfBeamABValues = FnBeamABValuesForDaimler()


                If (_sSupplierName = COMAU_NAME) Then
                    'Try catch added on Apr-19-2018
                    Try
                        asStockValues = Split(sStockFromClient.ToUpper, "X")
                        If Not asStockValues Is Nothing Then

                            If asStockValues.Length = 2 Then
                                'Stock size will be of Round stock size
                                'Validation added Apr-19-2018
                                If asStockValues(0).ToUpper.Contains("RD") Then
                                    sFirstValue = Split(asStockValues(0), "RD")(1) & "MM"
                                    sSecondValue = asStockValues(1) & "MM"

                                    sModifiedStockSize = sFirstValue & " DIA" & " X " & sSecondValue
                                Else
                                    sFirstValue = asStockValues(0) & "MM"
                                    sSecondValue = asStockValues(1) & "MM"

                                    sModifiedStockSize = sFirstValue & " DIA" & " X " & sSecondValue
                                End If
                            ElseIf (Not asStockValues(0) Is Nothing) And (Not asStockValues(1) Is Nothing) And (Not asStockValues(2) Is Nothing) Then
                                sFirstValue = asStockValues(0) & " MM "
                                sSecondValue = asStockValues(1) & " MM "

                                If asStockValues(2).ToUpper.Contains("L") Then
                                    sThirdValue = Split(asStockValues(2).ToUpper, "L")(0) & " MM W/T "
                                    sFourthValue = Split(asStockValues(2).ToUpper, "L=")(1) & " MM "
                                ElseIf asStockValues.Length = 4 Then
                                    sThirdValue = asStockValues(2) & " MM "
                                    sFourthValue = asStockValues(3) & " MM "
                                Else
                                    sThirdValue = asStockValues(2) & " MM "
                                    sFourthValue = Nothing
                                End If

                                If sFourthValue Is Nothing Then
                                    sModifiedStockSize = sFirstValue & "X " & sSecondValue & "X " & sThirdValue
                                Else
                                    sModifiedStockSize = sFirstValue & "X " & sSecondValue & "X " & sThirdValue & "X " & sFourthValue
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        sModifiedStockSize = ""
                    End Try
                ElseIf (_sSupplierName = VALIANT_NAME) Then
                    'Code added Aug-14-2020
                    Try
                        asStockValues = Split(sStockFromClient.ToUpper, "X")
                        If Not asStockValues Is Nothing Then

                            If asStockValues.Length = 2 Then
                                'Stock size will be of Round stock size
                                'Validation added Apr-19-2018

                                If (asStockValues(0).ToUpper.Contains("RD")) Or (asStockValues(0).ToUpper.Contains("Ø")) Then
                                    If (asStockValues(1).ToUpper.Contains("L")) Then
                                        'ROUND TUBD
                                        'DIA X W/T X MM
                                        If (asStockValues(0).ToUpper.Contains("RD")) Then
                                            sFirstValue = Split(asStockValues(0), "RD")(1) & "MM"
                                        ElseIf (asStockValues(0).ToUpper.Contains("Ø")) Then
                                            sFirstValue = Split(asStockValues(0), "Ø")(1) & "MM"
                                        End If
                                        sSecondValue = Split(asStockValues(1).ToUpper, "L")(0) & " MM W/T "
                                        'Code added Aug-28-2020
                                        If (asStockValues(1).ToUpper.Contains("L=")) Then
                                            sThirdValue = Split(asStockValues(1).ToUpper, "L=")(1) & " MM "
                                        ElseIf (asStockValues(1).ToUpper.Contains("L =")) Then
                                            sThirdValue = Split(asStockValues(1).ToUpper, "L =")(1) & " MM "
                                        End If

                                        sModifiedStockSize = sFirstValue & " DIA" & " X " & sSecondValue & " X " & sThirdValue
                                    Else
                                        'ROUND 
                                        'DIA X MM
                                        If (asStockValues(0).ToUpper.Contains("RD")) Then
                                            sFirstValue = Split(asStockValues(0), "RD")(1) & "MM"
                                        ElseIf (asStockValues(0).ToUpper.Contains("Ø")) Then
                                            sFirstValue = Split(asStockValues(0), "Ø")(1) & "MM"
                                        End If

                                        sSecondValue = asStockValues(1) & "MM"

                                        sModifiedStockSize = sFirstValue & " DIA" & " X " & sSecondValue
                                    End If

                                ElseIf (asStockValues(0).ToUpper.Contains("CH")) Then
                                    'CH80X8L=200
                                    'Here the first and Second value are obtained from the table for Channel based on first value.
                                    sFirstValue = Split(asStockValues(0), "CH")(1)
                                    If sFirstValue <> "" Then
                                        If dictOfChannelABValues.ContainsKey(sFirstValue) Then
                                            sSecondValue = dictOfChannelABValues(sFirstValue)
                                        End If
                                        sThirdValue = Split(asStockValues(1).ToUpper, "L")(0) & " MM W/T "
                                        If (asStockValues(1).ToUpper.Contains("L=")) Then
                                            sFourthValue = Split(asStockValues(1).ToUpper, "L=")(1) & " MM "
                                        ElseIf (asStockValues(1).ToUpper.Contains("L =")) Then
                                            sFourthValue = Split(asStockValues(1).ToUpper, "L =")(1) & " MM "
                                        End If

                                        sModifiedStockSize = sFirstValue & "MM X " & sSecondValue & " MM X " & sThirdValue & "X " & sFourthValue
                                    End If
                                ElseIf (asStockValues(0).ToUpper.Contains("SB")) Then
                                    'SB80X8L=200
                                    'Here the first and Second value are obtained from the table for Beam based on first value.
                                    sFirstValue = Split(asStockValues(0), "SB")(1)
                                    If sFirstValue <> "" Then
                                        If dictOfBeamABValues.ContainsKey(sFirstValue) Then
                                            sSecondValue = dictOfBeamABValues(sFirstValue)
                                        End If
                                        sThirdValue = Split(asStockValues(1).ToUpper, "L")(0) & " MM W/T "
                                        If (asStockValues(1).ToUpper.Contains("L=")) Then
                                            sFourthValue = Split(asStockValues(1).ToUpper, "L=")(1) & " MM "
                                        ElseIf (asStockValues(1).ToUpper.Contains("L =")) Then
                                            sFourthValue = Split(asStockValues(1).ToUpper, "L =")(1) & " MM "
                                        End If

                                        sModifiedStockSize = sFirstValue & "MM X " & sSecondValue & " MM X " & sThirdValue & "X " & sFourthValue
                                    End If
                                Else
                                    sFirstValue = asStockValues(0) & "MM"
                                    sSecondValue = asStockValues(1) & "MM"

                                    sModifiedStockSize = sFirstValue & " DIA" & " X " & sSecondValue
                                End If
                            ElseIf (Not asStockValues(0) Is Nothing) And (Not asStockValues(1) Is Nothing) And (Not asStockValues(2) Is Nothing) Then
                                If asStockValues(0).ToUpper.Contains("CH") Then
                                    'CH100X50X5L=660
                                    sFirstValue = Split(asStockValues(1), "CH")(0) & " MM "
                                    sSecondValue = asStockValues(1) & " MM "
                                ElseIf asStockValues(0).ToUpper.Contains("SB") Then
                                    'SB100X100X6L=300
                                    sFirstValue = Split(asStockValues(0), "SB")(1) & " MM "
                                    sSecondValue = asStockValues(1) & " MM "
                                Else
                                    '180X200X20 or 100X80X4L=200
                                    sFirstValue = asStockValues(0) & " MM "
                                    sSecondValue = asStockValues(1) & " MM "
                                End If


                                If asStockValues(2).ToUpper.Contains("L") Then
                                    sThirdValue = Split(asStockValues(2).ToUpper, "L")(0) & " MM W/T "
                                    If (asStockValues(2).ToUpper.Contains("L=")) Then
                                        sFourthValue = Split(asStockValues(2).ToUpper, "L=")(1) & " MM "
                                    ElseIf (asStockValues(2).ToUpper.Contains("L =")) Then
                                        sFourthValue = Split(asStockValues(2).ToUpper, "L =")(1) & " MM "
                                    End If

                                ElseIf asStockValues.Length = 4 Then
                                    sThirdValue = asStockValues(2) & " MM "
                                    sFourthValue = asStockValues(3) & " MM "
                                Else
                                    sThirdValue = asStockValues(2) & " MM "
                                    sFourthValue = Nothing
                                End If

                                If sFourthValue Is Nothing Then
                                    'sModifiedStockSize = sFirstValue & "X " & sSecondValue & "X " & sThirdValue
                                    'Code modified on Aug-25-2020
                                    'Valiant Daimler will be giving the thickness value at the very last value.
                                    ' Rearrage the smaller value to be in the front for Pradeep
                                    Dim asReArrange(2) As Double
                                    asReArrange(0) = Split(sFirstValue, " MM ")(0)
                                    asReArrange(1) = Split(sSecondValue, " MM ")(0)
                                    asReArrange(2) = Split(sThirdValue, " MM ")(0)
                                    Array.Sort(asReArrange)

                                    sModifiedStockSize = asReArrange(0) & " MM X " & asReArrange(1) & " MM X " & asReArrange(2) & " MM"
                                Else
                                    sModifiedStockSize = sFirstValue & "X " & sSecondValue & "X " & sThirdValue & "X " & sFourthValue
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        sWriteToLogFile("Error in converting the Stock Format")
                        sWriteToLogFile(ex.Message)
                        sWriteToLogFile(ex.StackTrace)
                        sModifiedStockSize = ""
                    End Try
                End If

            ElseIf (sOemName = FIAT_OEM_NAME) Then
                Try
                    asStockValues = Split(sStockFromClient.ToUpper, "X")
                    If Not asStockValues Is Nothing Then

                        If asStockValues.Length = 2 Then
                            'Stock size will be of Round stock size
                            'Validation added Apr-19-2018
                            If asStockValues(0).ToUpper.Contains("Ø") Then
                                If asStockValues(0).ToUpper.Contains("""") Then
                                    'Φ inch x mm
                                    sFirstValue = Split(asStockValues(0), "Ø")(1).Trim()
                                Else
                                    'Φ mm x mm
                                    sFirstValue = Split(asStockValues(0), "Ø")(1).Trim() & "MM"
                                End If
                                sSecondValue = asStockValues(1).Trim() & "MM"
                                sModifiedStockSize = sFirstValue & " DIA" & " X " & sSecondValue
                            Else
                                If asStockValues(0).ToUpper.Contains("""") Then
                                    'inch x mm
                                    sFirstValue = asStockValues(0).Trim()
                                Else
                                    'mm x mm
                                    sFirstValue = asStockValues(0).Trim() & "MM"
                                End If
                                sSecondValue = asStockValues(1).Trim() & "MM"
                                sModifiedStockSize = sFirstValue & " DIA" & " X " & sSecondValue
                            End If
                        ElseIf (Not asStockValues(0) Is Nothing) And (Not asStockValues(1) Is Nothing) And (Not asStockValues(2) Is Nothing) Then


                            If asStockValues(2).ToUpper.Contains("L") Then
                                'Square tube, Rectangular tube
                                sFirstValue = asStockValues(0).Trim() & " MM "
                                sSecondValue = asStockValues(1).Trim() & " MM "
                                sThirdValue = Split(asStockValues(2).Trim().ToUpper, "L")(0) & " MM W/T "
                                sFourthValue = Split(asStockValues(2).Trim().ToUpper, "L=")(1) & " MM "
                            ElseIf asStockValues(0).Contains("Ø") Then
                                'Round Tubing
                                If (asStockValues(0).ToUpper.Contains("""")) Then
                                    'Φ inch x mm x mm
                                    sFirstValue = Split(asStockValues(0), "Ø")(1).Trim()
                                Else
                                    'Φ mm x mm x mm
                                    sFirstValue = Split(asStockValues(0), "Ø")(1).Trim() & "MM"
                                End If
                                sSecondValue = asStockValues(1).Trim() & " MM "
                                sThirdValue = asStockValues(1).Trim() & " MM "
                                sFourthValue = Nothing

                            ElseIf asStockValues(1).ToUpper.Contains("L") Then
                                'L PLATE(REGULAR)
                                If asStockValues(0).ToUpper.Contains("""") Then
                                    'inch x inch  L=mm
                                    sFirstValue = asStockValues(0)
                                    sSecondValue = Split(asStockValues(1), "L=")(0).Trim()
                                    sThirdValue = Split(asStockValues(1), "L=")(0).Trim()
                                    sFourthValue = Split(asStockValues(1), "L=")(1).Trim()
                                Else
                                    'mm x mm  L=mm
                                    sFirstValue = asStockValues(0) & "MM"
                                    sSecondValue = Split(asStockValues(1), "L=")(0).Trim() & "MM"
                                    sThirdValue = Split(asStockValues(1), "L=")(0).Trim() & "MM"
                                    sFourthValue = Split(asStockValues(1), "L=")(1).Trim() & "MM"
                                End If
                            Else
                                'BLOCK, PLATE
                                If asStockValues(2).ToUpper.Contains("""") Then
                                    'mm x mm x  inch
                                    sFirstValue = asStockValues(2).Trim()
                                Else
                                    'mm x mm x  mm
                                    sFirstValue = asStockValues(2).Trim() & "MM"
                                End If
                                sSecondValue = asStockValues(1).Trim() & " MM "
                                sThirdValue = asStockValues(0).Trim() & " MM "
                                sFourthValue = Nothing
                            End If

                            If sFourthValue Is Nothing Then
                                sModifiedStockSize = sFirstValue & "X " & sSecondValue & "X " & sThirdValue
                            Else
                                sModifiedStockSize = sFirstValue & "X " & sSecondValue & "X " & sThirdValue & "X " & sFourthValue
                            End If
                        End If
                    End If
                Catch ex As Exception
                    sModifiedStockSize = ""
                End Try
            End If
        End If
        FnSegregateStockSize = sModifiedStockSize
    End Function

    Public Sub sWriteComponentDataInBodyNameSheet(sConfigFolderPath As String)
        Dim theSession As Session = Session.GetSession()
        Dim objPart As Part = theSession.Parts.Work
        Dim iDeatailNos As String = ""
        Dim sShape As String = ""
        Dim sToolClass As String = ""
        Dim iRowStart As Integer
        Dim sStockSize As String = ""
        Dim sPMat As String = ""
        Dim s3DErrDesc As String = ""
        Dim sBodyName As String = ""
        Dim sPartName As String = ""
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim bIsValidSolidBody As Boolean = False
        Dim aoListOfAllSolidBodies As List(Of Body) = Nothing
        Dim objChildComp As Component = Nothing
        Dim objChildPart As Part = Nothing
        Dim objNXToFetchAttribute As NXObject = Nothing

        iRowStart = BODY_INFO_START_ROW_WRITE

        If Not objPart Is Nothing Then
            'Collect solid Body from the component, based on the division (CAR/TRUCK)
            aoListOfAllSolidBodies = FnCollectComponentSolidBody(objPart)
            If Not aoListOfAllSolidBodies Is Nothing Then
                For Each objbody As Body In aoListOfAllSolidBodies
                    'Validation added Dec-19-2017
                    If objbody.IsSolidBody Then

                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyName, objbody.JournalIdentifier.ToString())
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColComponentName, objPart.Leaf.ToString())
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyLayer, objbody.Layer.ToString())

                        'Get the object from which attributes are fetched.
                        'In Case of GM attribtues are fetched from body
                        'in case of Chrysler attributes are fetched from body
                        'In case of Daimler attribtues are fetched from Part
                        If (_sOemName = GM_OEM_NAME) Then
                            objNXToFetchAttribute = objbody
                            sSetGMToolkitAttributes(objPart, objbody, False, True)
                        ElseIf (_sOemName = CHRYSLER_OEM_NAME) Then
                            objNXToFetchAttribute = objbody
                        ElseIf (_sOemName = DAIMLER_OEM_NAME) Then
                            objNXToFetchAttribute = objPart
                            'In case of Daimler Stock Size attributes, we rearrange and fetch the Stock Size attribute in correct format
                            sAddStockSizeAttribute(objPart, Nothing, _sOemName)
                        ElseIf (_sOemName = FIAT_OEM_NAME) Then
                            'Attributes will be added at the Part level in case of FIAT Components
                            objNXToFetchAttribute = objPart
                            'In case of FIAT Stock Size attributes, we rearrange and fetch the Stock Size attribute in correct format
                            sAddStockSizeAttribute(objPart, Nothing, _sOemName)
                        ElseIf (_sOemName = GESTAMP_OEM_NAME) Then
                            objNXToFetchAttribute = objbody
                        End If

                        If Not objNXToFetchAttribute Is Nothing Then
                            'Shape
                            sShape = FnGetStringUserAttribute(objNXToFetchAttribute, _SHAPE_ATTR_NAME)
                            'Code added Aug-16-2018
                            'Shapes are categorized and mapped.
                            'FLAT,SQUARE,PLATE are mapped as FLAT
                            'RECT TUBG and SQUARE TUBG are mapped as RECT TUBG
                            If sShape <> "" Then
                                If (sShape.ToUpper = PLATE) Or (sShape.ToUpper = SQUARE) Then
                                    sShape = FLAT
                                End If
                                If (sShape.ToUpper = SQUARE_TUBG) Then
                                    sShape = RECT_TUBG
                                End If
                            End If
                            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyShape, sShape)

                            'Tool Class
                            If (_sOemName = GM_OEM_NAME) Or (_sOemName = DAIMLER_OEM_NAME) Or (_sOemName = FIAT_OEM_NAME) Then
                                sToolClass = FnGetStringUserAttribute(objNXToFetchAttribute, _TOOL_CLASS)
                                If sToolClass <> "" Then
                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, sToolClass)
                                Else
                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, NOTAPPLICABLE)
                                End If
                            ElseIf (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                'Code added Aug-17-2018
                                'For Chrysler these attributes are read at part level and not from body level
                                sToolClass = FnGetStringUserAttribute(objPart, _TOOL_CLASS)
                                If sToolClass <> "" Then
                                    'Mapping is need for this attribute.
                                    If sToolClass.ToUpper = "P" Then
                                        sToolClass = "COMM"
                                    ElseIf sToolClass.ToUpper = "M/P" Then
                                        sToolClass = "ALT STD"
                                    ElseIf sToolClass.ToUpper = "NC" Then
                                        If FnGetStringUserAttribute(objPart, _TOOL_ID) <> "" Then
                                            sToolClass = "ALT STD"
                                        End If
                                    End If
                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, sToolClass)
                                End If
                            End If


                            'Stock Size
                            sStockSize = FnGetStringUserAttribute(objNXToFetchAttribute, _STOCK_SIZE_METRIC)
                            If _sOemName = GM_OEM_NAME Then
                                If sStockSize = "" Then
                                    sStockSize = FnGetStringUserAttribute(objNXToFetchAttribute, _STOCK_SIZE)
                                End If
                                If sStockSize = "" Then
                                    sStockSize = FnGetStringUserAttribute(objNXToFetchAttribute, _TOOL_ID)
                                End If

                            ElseIf (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                If sStockSize = "" Then
                                    sStockSize = FnGetStringUserAttribute(objNXToFetchAttribute, _STOCK_SIZE)
                                End If
                                'Code added Aug-16-2018
                                'For Standard and Alt Standard components, get the NAAMS number as Stock Size values.
                                If sStockSize = "" Then
                                    sStockSize = FnGetStringUserAttribute(objPart, _TOOL_ID)
                                End If
                                If sStockSize = "" Then
                                    'sStockSize = FnGetStringUserAttribute(objPart, "DB_ALT_PURCH")
                                    sStockSize = FnGetStringUserAttribute(objPart, _ALTPURCH)
                                End If
                            Else
                                If sStockSize = "" Then
                                    sStockSize = FnGetStringUserAttribute(objNXToFetchAttribute, _TOOL_ID)
                                End If
                            End If

                            If sStockSize <> "" Then
                                'Update the STOCK SIZE METRIC Information in the data
                                SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, STOCK_SIZE_METRIC_INFO_ROW_NOS,
                                                    STOCK_SIZE_METRIC_INFO_COLUMN_NOS, sStockSize)
                                'Add the stock size in the body name sheet.
                                SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColStockSize, sStockSize)
                            Else
                                'Populate the 3D exception report in case of missing stock size in the component
                                s3DErrDesc = "Stock size is missing in the 3D model part " & objPart.Leaf.ToString
                                SWrite(s3DErrDesc, Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, THREE_DIMENSIONAL_MODEL_EXCEPTION_REPORT_FILE_NAME))
                            End If

                            'Material
                            'Code Change made on Apr-03-2019
                            'For Daimler car V297 project, Material (WERKSTOFF) attribute is fetched at GEO PART level and not at container level.
                            If (_sDivision = CAR_DIVISION) Then
                                If Not objbody.OwningPart Is Nothing Then
                                    Dim objGeoPart As Part = Nothing
                                    objGeoPart = CType(objbody.OwningPart, Part)
                                    If Not objGeoPart Is Nothing Then
                                        FnLoadPartFully(objGeoPart)
                                        sPMat = FnGetStringUserAttribute(objGeoPart, _P_MAT)
                                    End If
                                End If
                            Else
                                sPMat = FnGetStringUserAttribute(objNXToFetchAttribute, _P_MAT)
                            End If

                            If sPMat <> "" Then
                                SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColPMat, sPMat)
                            End If

                            'Part Name
                            sPartName = FnGetStringUserAttribute(objNXToFetchAttribute, _PART_NAME)
                            If sPartName <> "" Then
                                SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColCompDBPartName, sPartName)
                            End If
                            iRowStart = iRowStart + 1
                        End If
                    End If
                Next
            End If
        End If

    End Sub
    'Code to Populate Weldment body information in Body Names sheet
    Public Sub sWriteWeldmentDataInBodyNameSheet(ByVal objPart As Part, sConfigFolderPath As String, aoAllCompInSession() As Component)
        Dim sStockSize As String = ""
        Dim iRowStart As Integer = 0
        Dim sShape As String = ""
        Dim sBodyName As String = ""
        Dim bSubComponent As Boolean = False
        Dim s3DErrDesc As String = ""
        'Dim sFolderName As String = ""
        Dim sToolClass As String = ""
        Dim sPMat As String = ""
        Dim objChildPart As Part = Nothing
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim bIsValidSolidBody As Boolean = False
        Dim objNXToFetchAttribute As NXObject = Nothing
        Dim aoAllValidBody() As Body = Nothing
        Dim objParentComp As Component = Nothing
        'For computing the exact view bounds of the bodies
        'Dim min_corner(2) As Double
        'Dim directions(2, 2) As Double
        'Dim distances(2) As Double

        iRowStart = BODY_INFO_START_ROW_WRITE
        Dim dictStockSizeCompData As Dictionary(Of String, NXObject()) = Nothing
        dictStockSizeCompData = New Dictionary(Of String, NXObject())

        If Not aoAllCompInSession Is Nothing Then
            sAddStockSizeAttribute(objPart, aoAllCompInSession, _sOemName)
            For Each objComp As Component In aoAllCompInSession
                objChildPart = FnGetPartFromComponent(objComp)
                If Not objChildPart Is Nothing Then
                    FnLoadPartFully(objChildPart)

                    aoAllValidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)
                    If Not aoAllValidBody Is Nothing Then
                        For Each objbody As Body In aoAllValidBody
                            If Not objbody Is Nothing Then
                                'Check if the body belongs to the root comp or the sub assembly , assign a unique name as reuired by core algo
                                'by joining the body name with the component instance tag
                                If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                    objNXToFetchAttribute = objbody
                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        sBodyName = objbody.JournalIdentifier
                                        bSubComponent = False
                                        If (_sOemName = GM_OEM_NAME) Then
                                            sSetGMToolkitAttributes(objPart, objbody, False, False)
                                        End If
                                        sStockSize = FnGetStringUserAttribute(objbody, _STOCK_SIZE_METRIC)
                                        If sStockSize = "" Then
                                            sStockSize = FnGetStringUserAttribute(objbody, _STOCK_SIZE)
                                        End If
                                        'COde added Aug-16-2018
                                        'In case of Chrysler, get the DB_NAAMS and DB_ALT_PURCH attribute at part level
                                        If (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                            sStockSize = FnGetStringUserAttribute(objPart, _TOOL_ID)
                                            If sStockSize = "" Then
                                                'sStockSize = FnGetStringUserAttribute(objPart, "DB_ALT_PURCH")
                                                sStockSize = FnGetStringUserAttribute(objPart, _ALTPURCH)
                                            End If
                                        Else
                                            If sStockSize = "" Then
                                                sStockSize = FnGetStringUserAttribute(objbody, _TOOL_ID)
                                            End If
                                        End If

                                    Else
                                        sBodyName = CType(objComp.FindOccurrence(objbody), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                        bSubComponent = True
                                        'Add the GM Toolkit Attributes
                                        If (_sOemName = GM_OEM_NAME) Then
                                            sSetGMToolkitAttributes(objPart, objComp, True, False)
                                        End If
                                        sStockSize = FnGetStringUserAttribute(objComp, _STOCK_SIZE_METRIC)
                                        If sStockSize = "" Then
                                            sStockSize = FnGetStringUserAttribute(objComp, _STOCK_SIZE)
                                        End If
                                        'COde added Aug-16-2018
                                        'In case of Chrysler, get the DB_NAAMS and DB_ALT_PURCH attribute at part level
                                        If (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                            If sStockSize = "" Then
                                                sStockSize = FnGetStringUserAttribute(objChildPart, _TOOL_ID)
                                            End If
                                            If sStockSize = "" Then
                                                'sStockSize = FnGetStringUserAttribute(objChildPart, "DB_ALT_PURCH")
                                                sStockSize = FnGetStringUserAttribute(objChildPart, _ALTPURCH)
                                            End If
                                        Else
                                            If sStockSize = "" Then
                                                sStockSize = FnGetStringUserAttribute(objComp, _TOOL_ID)
                                            End If
                                        End If

                                    End If
                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColComponentName, objComp.DisplayName)
                                ElseIf _sOemName = DAIMLER_OEM_NAME Then
                                    'In case of Car division, the component attributes would be added at the container level.
                                    If _sDivision = TRUCK_DIVISION Then
                                        objNXToFetchAttribute = objComp
                                        objParentComp = objComp
                                    ElseIf _sDivision = CAR_DIVISION Then
                                        objParentComp = objComp.Parent
                                        objNXToFetchAttribute = objParentComp
                                    End If

                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        sBodyName = objbody.JournalIdentifier
                                        sStockSize = FnGetStringUserAttribute(objPart, _STOCK_SIZE_METRIC)
                                    Else
                                        sBodyName = CType(objComp.FindOccurrence(objbody), Body).JournalIdentifier & " " & objNXToFetchAttribute.JournalIdentifier
                                        'Add the DB_PART_NAME info for all the child components in a weldment
                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColCompDBPartName,
                                                                                FnGetCompAttribute(objNXToFetchAttribute, "String", _PART_NAME))
                                        sStockSize = FnGetStringUserAttribute(objNXToFetchAttribute, _STOCK_SIZE_METRIC)
                                    End If
                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColComponentName, objParentComp.Parent.DisplayName)

                                ElseIf _sOemName = FIAT_OEM_NAME Then
                                    objNXToFetchAttribute = objComp
                                    objParentComp = objComp.Parent

                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        sBodyName = objbody.JournalIdentifier
                                        sStockSize = FnGetStringUserAttribute(objPart, _STOCK_SIZE_METRIC)
                                    Else
                                        sBodyName = CType(objComp.FindOccurrence(objbody), Body).JournalIdentifier & " " & objComp.JournalIdentifier

                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColCompDBPartName,
                                                                                FnGetCompAttribute(objParentComp, "String", _PART_NAME))
                                        sStockSize = FnGetStringUserAttribute(objNXToFetchAttribute, _STOCK_SIZE_METRIC)
                                    End If
                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColComponentName, objParentComp.DisplayName)
                                End If

                                'Tool Class
                                If (_sOemName = GM_OEM_NAME) Or (_sOemName = DAIMLER_OEM_NAME) Or (_sOemName = FIAT_OEM_NAME) Then
                                    sToolClass = FnGetStringUserAttribute(objNXToFetchAttribute, _TOOL_CLASS)
                                    If sToolClass <> "" Then
                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, sToolClass)
                                    Else
                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, NOTAPPLICABLE)
                                    End If
                                ElseIf (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                    'Code added Aug-17-2018
                                    'For Chrysler these attributes are read at part level and not from body level
                                    sToolClass = FnGetStringUserAttribute(objPart, _TOOL_CLASS)
                                    If sToolClass <> "" Then
                                        'Mapping is need for this attribute.
                                        If sToolClass.ToUpper = "P" Then
                                            sToolClass = "COMM"
                                        ElseIf sToolClass.ToUpper = "M/P" Then
                                            sToolClass = "ALT STD"
                                        ElseIf sToolClass.ToUpper = "NC" Then
                                            If FnGetStringUserAttribute(objPart, _TOOL_ID) <> "" Then
                                                sToolClass = "ALT STD"
                                            End If
                                        End If
                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, sToolClass)
                                    End If
                                End If
                                'Check whether it is some body which is not suppose to be generated in sweep data
                                'CODE MODIFIED - 6/13/16 - Amitabh - Ignore WIRE MESH bodies with respect to the SHAPE attribute
                                If (FnGetStringUserAttribute(objNXToFetchAttribute, _SHAPE_ATTR_NAME) <> WIRE_MESH_SHAPE) Then

                                    If sStockSize <> "" Then
                                        If (_sOemName = GM_OEM_NAME) Or (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                                            If Not dictStockSizeCompData.ContainsKey(sStockSize) Then
                                                If Not bSubComponent Then
                                                    dictStockSizeCompData.Add(sStockSize, {objbody})
                                                Else
                                                    dictStockSizeCompData.Add(sStockSize, {objComp})
                                                End If
                                            Else
                                                If Not bSubComponent Then
                                                    ReDim Preserve dictStockSizeCompData(sStockSize)(UBound(dictStockSizeCompData(sStockSize)) + 1)
                                                    dictStockSizeCompData(sStockSize)(UBound(dictStockSizeCompData(sStockSize))) = objbody
                                                Else
                                                    ReDim Preserve dictStockSizeCompData(sStockSize)(UBound(dictStockSizeCompData(sStockSize)) + 1)
                                                    dictStockSizeCompData(sStockSize)(UBound(dictStockSizeCompData(sStockSize))) = objComp
                                                End If
                                            End If
                                        ElseIf (_sOemName = DAIMLER_OEM_NAME) Or (_sOemName = FIAT_OEM_NAME) Then
                                            If Not dictStockSizeCompData.ContainsKey(sStockSize) Then
                                                dictStockSizeCompData.Add(sStockSize, {objNXToFetchAttribute})
                                            Else
                                                ReDim Preserve dictStockSizeCompData(sStockSize)(UBound(dictStockSizeCompData(sStockSize)) + 1)
                                                dictStockSizeCompData(sStockSize)(UBound(dictStockSizeCompData(sStockSize))) = objNXToFetchAttribute
                                            End If
                                        End If

                                    End If

                                    'Update Value to the cell
                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyName, sBodyName)

                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyLayer, objbody.Layer.ToString)
                                    sShape = FnGetStringUserAttribute(objNXToFetchAttribute, _SHAPE_ATTR_NAME)
                                    'Code added Aug-16-2018
                                    'Shapes are categorized and mapped.
                                    'FLAT,SQUARE,PLATE are mapped as FLAT
                                    'RECT TUBG and SQUARE TUBG are mapped as RECT TUBG
                                    If sShape <> "" Then
                                        If (sShape.ToUpper = PLATE) Or (sShape.ToUpper = SQUARE) Then
                                            sShape = FLAT
                                        End If
                                        If (sShape.ToUpper = SQUARE_TUBG) Then
                                            sShape = RECT_TUBG
                                        End If
                                    End If
                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyShape, sShape)
                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColStockSize, sStockSize)

                                    'Material
                                    'Code Change made on Apr-03-2019
                                    'For Daimler car V297 project, Material (WERKSTOFF) attribute is fetched at GEO PART level and not at container level.
                                    If (_sDivision = CAR_DIVISION) Then
                                        'If Not objbody.OwningPart Is Nothing Then
                                        '    Dim objGeoPart As Part = Nothing
                                        '    objGeoPart = CType(objbody.OwningPart, Part)
                                        '    If Not objGeoPart Is Nothing Then
                                        '        FnLoadPartFully(objGeoPart)
                                        sPMat = FnGetStringUserAttribute(objComp, _P_MAT)
                                        '    End If
                                        'End If
                                    Else
                                        sPMat = FnGetStringUserAttribute(objNXToFetchAttribute, _P_MAT)
                                    End If

                                    ' sPMat = FnGetStringUserAttribute(objNXToFetchAttribute, _P_MAT)
                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColPMat, sPMat)
                                    If sStockSize = "" Then
                                        'Populate the 3D exception report in case of missing stock size in the component
                                        s3DErrDesc = "Stock size is missing in the 3D model part " & objPart.Leaf.ToString
                                        SWrite(s3DErrDesc, Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, THREE_DIMENSIONAL_MODEL_EXCEPTION_REPORT_FILE_NAME))
                                    End If
                                    iRowStart = iRowStart + 1
                                End If
                                'End If

                            End If
                        Next
                    End If
                End If
            Next
        Else
            For Each objbody As Body In FnGetNxSession.Parts.Work.Bodies()
                'Check whether the body is a solid body
                'Only pick bodies which are in layer 1 (other side may also be present in the same part which need not be detailed) - 26/2/2014
                If objbody.IsSolidBody Then 'And (body.Layer = 1 Or (body.Layer >= 92 And body.Layer <= 184)) Then
                    sSetStatus("Collecting attribute data for " & objbody.JournalIdentifier.ToUpper)
                    'sToolClass = FnGetBodyAttribute(objbody, "String", _TOOL_CLASS)
                    'If Not sToolClass = "" Then
                    '    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, sToolClass)
                    'End If

                    If (_sOemName = GM_OEM_NAME) Or (_sOemName = DAIMLER_OEM_NAME) Then
                        sToolClass = FnGetStringUserAttribute(objbody, _TOOL_CLASS)
                    ElseIf (_sOemName = DAIMLER_OEM_NAME) Or (_sOemName = FIAT_OEM_NAME) Then
                        sToolClass = FnGetStringUserAttribute(objPart, _TOOL_CLASS)

                    ElseIf (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                        'Code added Aug-17-2018
                        'For Chrysler these attributes are read at part level and not from body level
                        sToolClass = FnGetStringUserAttribute(objPart, _TOOL_CLASS)
                        If sToolClass <> "" Then
                            'Mapping is need for this attribute.
                            If sToolClass.ToUpper = "P" Then
                                sToolClass = "COMM"
                            ElseIf sToolClass.ToUpper = "M/P" Then
                                sToolClass = "ALT STD"
                            ElseIf sToolClass.ToUpper = "NC" Then
                                If FnGetStringUserAttribute(objPart, _TOOL_ID) <> "" Then
                                    sToolClass = "ALT STD"
                                End If
                            End If
                        End If
                    End If
                    If sToolClass <> "" Then
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, sToolClass)
                    Else
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, NOTAPPLICABLE)
                    End If

                    'Check whether it is some body which is not suppose to be generated in sweep data
                    'CODE MODIFIED - 6/13/16 - Amitabh - Ignore WIRE MESH bodies with respect to the SHAPE attribute
                    If (FnGetStringUserAttribute(objbody, _SHAPE_ATTR_NAME) <> WIRE_MESH_SHAPE) And
                                (Not FnChkIfBodyIsMesh(objPart, objbody)) Then
                        'Add the GM Toolkit Attributes
                        sSetGMToolkitAttributes(objPart, objbody, False, False)
                        'Add this body to the collection of solid bodies
                        'sStoreSolidBody(Body)

                        'If FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("M") Then
                        'To Output the stock size
                        sStockSize = FnGetBodyAttribute(objbody, "String", _STOCK_SIZE_METRIC)
                        'Check if the STOCK_SIZE attribute is present if the above attribute is absent
                        If sStockSize = "" Then
                            sStockSize = FnGetBodyAttribute(objbody, "String", _STOCK_SIZE)
                        End If
                        'Code added Aug-16-2018

                        If (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                            If sStockSize = "" Then
                                sStockSize = FnGetStringUserAttribute(objPart, _TOOL_ID)
                            End If
                            If sStockSize = "" Then
                                'sStockSize = FnGetStringUserAttribute(objPart, "DB_ALT_PURCH")
                                sStockSize = FnGetStringUserAttribute(objPart, _ALTPURCH)
                            End If
                        Else
                            If sStockSize = "" Then
                                sStockSize = FnGetBodyAttribute(objbody, "String", _TOOL_ID)
                            End If
                        End If

                        If sStockSize <> "" Then
                            If Not dictStockSizeCompData.ContainsKey(sStockSize) Then
                                dictStockSizeCompData.Add(sStockSize, {objbody})
                            Else
                                ReDim Preserve dictStockSizeCompData(sStockSize)(UBound(dictStockSizeCompData(sStockSize)) + 1)
                                dictStockSizeCompData(sStockSize)(UBound(dictStockSizeCompData(sStockSize))) = objbody
                            End If
                        End If
                        'Get the exact bounding box of the bodies
                        'FnGetUFSession.Modl.AskBoundingBoxExact(body.Tag, NXOpen.Tag.Null, min_corner, directions, distances)

                        'Update Value to the cell
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyName, objbody.JournalIdentifier.ToString)
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColComponentName, objPart.Leaf.ToString())
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyLayer, objbody.Layer.ToString)
                        sShape = FnGetBodyAttribute(objbody, "String", _SHAPE_ATTR_NAME)
                        'Code added Aug-16-2018
                        'Shapes are categorized and mapped.
                        'FLAT,SQUARE,PLATE are mapped as FLAT
                        'RECT TUBG and SQUARE TUBG are mapped as RECT TUBG
                        If sShape <> "" Then
                            If (sShape.ToUpper = PLATE) Or (sShape.ToUpper = SQUARE) Then
                                sShape = FLAT
                            End If
                            If (sShape.ToUpper = SQUARE_TUBG) Then
                                sShape = RECT_TUBG
                            End If
                        End If
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyShape, sShape)
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColStockSize, sStockSize)
                        sPMat = FnGetBodyAttribute(objbody, "String", _P_MAT)
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColPMat, sPMat)
                        If sStockSize = "" Then
                            'Populate the 3D exception report in case of missing stock size
                            'sFolderName = Split(Split(objPart.FullPath, sConfigFolderPath & "\")(1), "\")(0)
                            s3DErrDesc = "Stock size is missing in the body " & objbody.JournalIdentifier.ToString & " in the 3D model part " & objPart.Leaf.ToString
                            SWrite(s3DErrDesc, Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, THREE_DIMENSIONAL_MODEL_EXCEPTION_REPORT_FILE_NAME))
                        End If

                        iRowStart = iRowStart + 1
                    End If
                End If
            Next
        End If

        'Now Check for any physical differences and then renumber sub details if required
        sSetStatus("Analysing solid bodies for sub-details ")
        If (_sOemName = FIAT_OEM_NAME) Then
            'In case of fiat, Sub detail number attribute will be created based on the Sub child component number.
            'For Fiat physical dia check is not needed.
            sCreateSubDetNumForFiat(objPart)
        Else
            sCheckForPhysicalDifferencesInSubDetailBasedOnFaceArea(objPart, dictStockSizeCompData)
        End If

    End Sub

    ''Function To check if the part file belong to any weldment
    'Function FnCheckIfPartFileIsNotWithinWeldment(objPart As Part) As Boolean
    '    Dim sPartName As String = ""
    '    Dim asPartFileNames() As String = Nothing

    '    If Not objPart Is Nothing Then
    '        sPartName = objPart.Leaf.ToString().ToUpper
    '        If FnCheckFileExists(Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, LIST_OF_PARTS_IN_WELDMENT_FILE_NAME)) Then
    '            asPartFileNames = FnReadFullFile(Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, LIST_OF_PARTS_IN_WELDMENT_FILE_NAME))
    '            If Not asPartFileNames Is Nothing Then
    '                For Each sPartFileName In asPartFileNames
    '                    If sPartFileName.ToUpper = sPartName.ToUpper Then
    '                        FnCheckIfPartFileIsNotWithinWeldment = True
    '                        Exit Function
    '                    End If
    '                Next
    '            End If
    '        End If
    '    End If
    '    FnCheckIfPartFileIsNotWithinWeldment = False
    'End Function

    'COde added Oct-12-2017
    'Function to check if the given rotation matrix is unique
    ' 1. Check if the given rotation matrix is Parallel anti-parallel to the unique rotation matrix present in the dictionary
    ' 2. If rotation matrix is found to be parallel anti parallel, then neglect that rotation matrix. 
    Function FnCheckIfMatrixIsUnique(adRotMat() As Double, dictRotMat As Dictionary(Of Double, Double())) As Boolean
        Dim bIsVector1ParallelAntiParallel As Boolean = False
        Dim bIsVector2ParallelAntiParallel As Boolean = False
        Dim bIsVector3ParallelAntiParallel As Boolean = False

        If Not dictRotMat Is Nothing Then
            For Each adUniqRotMat In dictRotMat.Values
                bIsVector1ParallelAntiParallel = False
                bIsVector2ParallelAntiParallel = False
                bIsVector3ParallelAntiParallel = False

                If FnParallelAntiParallelCheck({adUniqRotMat(0), adUniqRotMat(1), adUniqRotMat(2)}, {adRotMat(0), adRotMat(1), adRotMat(2)}) Then
                    bIsVector1ParallelAntiParallel = True
                ElseIf FnParallelAntiParallelCheck({adUniqRotMat(0), adUniqRotMat(1), adUniqRotMat(2)}, {adRotMat(3), adRotMat(4), adRotMat(5)}) Then
                    bIsVector1ParallelAntiParallel = True
                ElseIf FnParallelAntiParallelCheck({adUniqRotMat(0), adUniqRotMat(1), adUniqRotMat(2)}, {adRotMat(6), adRotMat(7), adRotMat(8)}) Then
                    bIsVector1ParallelAntiParallel = True
                End If

                If FnParallelAntiParallelCheck({adUniqRotMat(3), adUniqRotMat(4), adUniqRotMat(5)}, {adRotMat(0), adRotMat(1), adRotMat(2)}) Then
                    bIsVector2ParallelAntiParallel = True
                ElseIf FnParallelAntiParallelCheck({adUniqRotMat(3), adUniqRotMat(4), adUniqRotMat(5)}, {adRotMat(3), adRotMat(4), adRotMat(5)}) Then
                    bIsVector2ParallelAntiParallel = True
                ElseIf FnParallelAntiParallelCheck({adUniqRotMat(3), adUniqRotMat(4), adUniqRotMat(5)}, {adRotMat(6), adRotMat(7), adRotMat(8)}) Then
                    bIsVector2ParallelAntiParallel = True
                End If

                If FnParallelAntiParallelCheck({adUniqRotMat(6), adUniqRotMat(7), adUniqRotMat(8)}, {adRotMat(0), adRotMat(1), adRotMat(2)}) Then
                    bIsVector3ParallelAntiParallel = True
                ElseIf FnParallelAntiParallelCheck({adUniqRotMat(6), adUniqRotMat(7), adUniqRotMat(8)}, {adRotMat(3), adRotMat(4), adRotMat(5)}) Then
                    bIsVector3ParallelAntiParallel = True
                ElseIf FnParallelAntiParallelCheck({adUniqRotMat(6), adUniqRotMat(7), adUniqRotMat(8)}, {adRotMat(6), adRotMat(7), adRotMat(8)}) Then
                    bIsVector3ParallelAntiParallel = True
                End If

                If bIsVector1ParallelAntiParallel And bIsVector2ParallelAntiParallel And bIsVector3ParallelAntiParallel Then
                    FnCheckIfMatrixIsUnique = False
                    Exit Function
                End If
            Next
        End If
        FnCheckIfMatrixIsUnique = True
    End Function

    'Function to check if a Planar face has atleast one linear edge
    Function FnChkIfPlanarFaceHasAtleastOneLinearEdge(objFace As Face) As Boolean
        If Not objFace Is Nothing Then
            For Each objEdge As Edge In objFace.GetEdges()
                If objEdge.SolidEdgeType = Edge.EdgeType.Linear Then
                    FnChkIfPlanarFaceHasAtleastOneLinearEdge = True
                    Exit Function
                End If
            Next
        End If
        FnChkIfPlanarFaceHasAtleastOneLinearEdge = False
    End Function

    'Function to collect all the feature members of Part Design feature group
    Function FnCollectAllMembersOfFeatureGroup(objPart As Part, sFeatureGroupName As String) As Feature()
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim objFeatureGroup As Features.FeatureGroup = Nothing
        Dim aoAllMembers() As Features.Feature = Nothing

        If Not objPart.Features Is Nothing Then
            For Each objFeature As Features.Feature In objPart.Features
                'Collect Feature groups
                Try
                    objFeatureGroup = CType(objFeature, FeatureGroup)
                    If Not objFeatureGroup Is Nothing Then
                        'Check if the feature group name is Part_Design
                        If objFeatureGroup.Name.ToUpper.Contains(sFeatureGroupName) Then
                            'Get all the members of part_design feature group
                            objFeatureGroup.GetMembers(aoAllMembers)
                            If Not aoAllMembers Is Nothing Then
                                For Each objMember In aoAllMembers

                                    'Neglect EXTRACT_FACE and EXTRACT_BODY
                                    If objMember.FeatureType <> "EXTRACT_FACE" And objMember.FeatureType <> "EXTRACT_BODY" Then
                                        If aoPartDesignMembers Is Nothing Then
                                            ReDim Preserve aoPartDesignMembers(0)
                                            aoPartDesignMembers(0) = objMember
                                        Else
                                            ReDim Preserve aoPartDesignMembers(UBound(aoPartDesignMembers) + 1)
                                            aoPartDesignMembers(UBound(aoPartDesignMembers)) = objMember
                                        End If
                                    End If
                                Next
                            End If
                            Exit For
                        End If
                    End If
                Catch ex As Exception

                End Try
            Next
        End If
        FnCollectAllMembersOfFeatureGroup = aoPartDesignMembers
    End Function

    'Code added Dec-05-2017
    'Collect all the Relief Cut Face (Single Diamond Finish Tolerance Faces)
    Function FnCollectReliefCutFace(objPart As Part) As Face()
        Dim aoReliefCutFinishTol() As DisplayableObject = Nothing
        Dim objReliefCutFace As Face = Nothing
        Dim asReliefCutFace() As Face = Nothing
        'Dim sFaceName As String = ""

        If Not objPart Is Nothing Then
            aoReliefCutFinishTol = FnGetFaceObjectByAttributes(objPart, _FINISHTOLERANCE_ATTR_NAME, _RELIEF_CUT_FACE_TOLERANCE_VALUE)
            If Not aoReliefCutFinishTol Is Nothing Then
                For Each objReliefCutObject As DisplayableObject In aoReliefCutFinishTol
                    objReliefCutFace = CType(objReliefCutObject, Face)
                    If Not objReliefCutFace Is Nothing Then
                        If asReliefCutFace Is Nothing Then
                            ReDim Preserve asReliefCutFace(0)
                            asReliefCutFace(0) = objReliefCutFace
                        Else
                            If Not asReliefCutFace.Contains(objReliefCutFace) Then
                                ReDim Preserve asReliefCutFace(UBound(asReliefCutFace) + 1)
                                asReliefCutFace(UBound(asReliefCutFace)) = objReliefCutFace
                            End If
                        End If

                    End If
                Next
            End If
        End If
        FnCollectReliefCutFace = asReliefCutFace
    End Function
    'Populate all the associated machined face to the reference Relief cut face.
    Sub sPopulateAssociatedMachinedFaceToReliefCutFace(objPart As Part)

        Dim aoReliefCutFinishTolFace() As Face = Nothing
        Dim iCountLastRowFilled As Integer = 0
        Dim aTAdjFaceTag() As Tag = Nothing
        Dim objAssociatedFace As Face = Nothing
        Dim sMachinedFaceNameInSheet As String = ""
        Dim bAssociatedMachiningFaceFound As Boolean = False
        Dim objPerpendicularMachinedFace As Face = Nothing

        'Collect all Relief cut Finish tolerance face
        aoReliefCutFinishTolFace = FnCollectReliefCutFace(objPart)
        If Not aoReliefCutFinishTolFace Is Nothing Then
            For Each objReliefCutFace As Face In aoReliefCutFinishTolFace
                'Validation added on May-12-2018
                'In Daimler tool, even the cylindrical face were given machining tolerance of 6.4
                'Ref component F56500103929301270005
                If objReliefCutFace.SolidFaceType = Face.FaceType.Planar Then
                    bAssociatedMachiningFaceFound = False
                    objPerpendicularMachinedFace = FnGetPerpendicularWallToMatingFace(objPart, objReliefCutFace)
                    If Not objPerpendicularMachinedFace Is Nothing Then
                        'collect the associated face which has finish tolerance value
                        FnGetUFSession.Modl.AskAdjacFaces(objReliefCutFace.Tag, aTAdjFaceTag)
                        If Not aTAdjFaceTag Is Nothing Then
                            For Each objAdjFaceTag As Tag In aTAdjFaceTag
                                objAssociatedFace = CType(NXObjectManager.Get(objAdjFaceTag), Face)
                                If Not objAssociatedFace Is Nothing Then
                                    'Check if the associated face is a machined face to the relief cut face
                                    If FnGetStringUserAttribute(objAssociatedFace, _FINISHTOLERANCE_ATTR_NAME) <> "" Then
                                        If FnGetStringUserAttribute(objAssociatedFace, _FINISHTOLERANCE_ATTR_NAME).Contains(FINISHTOLATTRVALUE) Then
                                            If objPerpendicularMachinedFace Is objAssociatedFace Then
                                                bAssociatedMachiningFaceFound = True
                                                'Get the corresponding objRelief cut face name on the misc info sheet and write the associated machined face
                                                iCountLastRowFilled = FnGetNumberofRows(_objWorkBk, MISCINFOSHEETNAME, 1, 1)
                                                If iCountLastRowFilled > 2 Then
                                                    For iIndex As Integer = 3 To iCountLastRowFilled
                                                        sMachinedFaceNameInSheet = FnReadSingleRowForColumn(_objWorkBk, MISCINFOSHEETNAME, MISC_INFO_START_COL_WRITE, iIndex)
                                                        If sMachinedFaceNameInSheet <> "" Then
                                                            If sMachinedFaceNameInSheet.ToUpper = objReliefCutFace.Name.ToUpper Then
                                                                SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, iIndex, 3, objAssociatedFace.Name.ToUpper)
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            End If
                                            If bAssociatedMachiningFaceFound Then
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            Next


        End If
    End Sub
    'When we encounter a double-diamond face (Face A), we assess if there exists a Planar Face (Face B):
    '        a. which shares a linear edge with Face A
    '        b. whose normal vector is orthogonal to that of Face A
    '        c. Find pair of edges, one in each face, which are PERPENDICULAR to each other & share a vertex ("P") between them									

    Function FnGetPerpendicularWallToMatingFace(objPart As Part, objFaceA As Face) As Face
        Dim objFaceB As Face = Nothing
        Dim adPointP() As Double = Nothing
        Dim adA(2) As Double
        Dim adB(2) As Double
        Dim adFaceNormalA() As Double = Nothing
        Dim adFaceNormalB() As Double = Nothing
        Dim objCommonEdge As Edge = Nothing
        Dim objVrt1 As Point3d = Nothing
        Dim objVrt2 As Point3d = Nothing
        Dim adMidPt() As Double = Nothing
        Dim inside1 As Integer = 0
        Dim inside2 As Integer = 0
        Dim iRoundingPrecission As Integer = 2

        'Get all the linear edges in FACE A
        For Each objEdge As Edge In objFaceA.GetEdges()
            If objEdge.SolidEdgeType = Edge.EdgeType.Linear Then
                'Check if there is atleast one planar face common to this edge
                For Each objConnFace As Face In objEdge.GetFaces()
                    If (objConnFace.SolidFaceType = Face.FaceType.Planar) And (Not objConnFace Is objFaceA) Then
                        objFaceB = objConnFace

                        'Check if the normal vector of FACE B is orthogonal to FACE A
                        'Dot product should be 0
                        If Round(FnComputeDotProduct(FnGetDirVecOfFace(objPart, objFaceA), FnGetDirVecOfFace(objPart, objFaceB)), iRoundingPrecission) = 0 Then
                            '1. Find common edge between two perpendicular faces
                            objCommonEdge = FnFindCommonEdgeToFaces(objFaceA, objFaceB)
                            '2. Compute mid point "P" of the common edge
                            objCommonEdge.GetVertices(objVrt1, objVrt2)
                            adMidPt = {(objVrt1.X + objVrt2.X) / 2, (objVrt1.Y + objVrt2.Y) / 2, (objVrt1.Z + objVrt2.Z) / 2}
                            '3. Compute points "A", "B" at a distance of "0.1mm" from "P" in the direction of each face normal
                            'A = P + (0.1) X F1 (component wise)			
                            'B = P + (0.1) X F2 (component wise)
                            'Get the FACE NORMAL
                            adFaceNormalA = FnGetFaceNormal(objFaceA)
                            adFaceNormalB = FnGetFaceNormal(objFaceB)
                            For i = 0 To 2
                                adA(i) = adMidPt(i) + (0.1 * adFaceNormalA(i))
                                adB(i) = adMidPt(i) + (0.1 * adFaceNormalB(i))
                            Next
                            '1 = point is inside the body 
                            '2 = point is outside the body 
                            '3 = point is on the body

                            '4. Check if A and B each are contained on one of the input Faces
                            FnGetUFSession.Modl.AskPointContainment(adA, objFaceA.GetBody().Tag, inside1)
                            FnGetUFSession.Modl.AskPointContainment(adB, objFaceA.GetBody().Tag, inside2)
                            If (inside1 = 3 Or inside1 = 1) And (inside2 = 3 Or inside2 = 1) Then
                                ''Assign the finish tolerance
                                'sAttributeFace(objFaceB.Prototype, FINISHTOLATTRNAME, NO_MATING_FACE_TOLERANCE_VALUE)
                                'sUpdateAttributesInModel()
                                FnGetPerpendicularWallToMatingFace = objFaceB
                                Exit Function
                            Else
                                'Compute new points in the direction opposite to other face
                                For i = 0 To 2
                                    adA(i) = adA(i) - (0.1 * adFaceNormalB(i))
                                    adB(i) = adB(i) - (0.1 * adFaceNormalA(i))
                                Next
                                '4. Check if A and B each are contained on one of the input Faces
                                FnGetUFSession.Modl.AskPointContainment(adA, objFaceA.GetBody().Tag, inside1)
                                FnGetUFSession.Modl.AskPointContainment(adB, objFaceA.GetBody().Tag, inside2)
                                If (inside1 = 3 Or inside1 = 1) And (inside2 = 3 Or inside2 = 1) Then
                                    ''Assign the finish tolerance
                                    'sAttributeFace(objFaceB.Prototype, FINISHTOLATTRNAME, NO_MATING_FACE_TOLERANCE_VALUE)
                                    'sUpdateAttributesInModel()
                                    FnGetPerpendicularWallToMatingFace = objFaceB
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next
        FnGetPerpendicularWallToMatingFace = Nothing
    End Function
    'Compute the dot product
    Public Function FnComputeDotProduct(ByVal adLine1Vec() As Double, ByVal adLine2Vec() As Double) As Double
        Dim dDotProduct As Double = 0.0
        Dim adUnitVecLine1() As Double = FnGetDirCosOfVector(adLine1Vec)
        Dim adUnitVecLine2() As Double = FnGetDirCosOfVector(adLine2Vec)
        dDotProduct = (adUnitVecLine1(0) * adUnitVecLine2(0)) + (adUnitVecLine1(1) * adUnitVecLine2(1)) + (adUnitVecLine1(2) * adUnitVecLine2(2))
        FnComputeDotProduct = dDotProduct
    End Function
    'Get the direction cosine of a vector
    Public Function FnGetDirCosOfVector(ByVal adVec() As Double) As Double()
        Dim adUnitVec(2) As Double
        Dim dMagOfVec As Double = FnGetMagnitudeOfVector(adVec)
        adUnitVec(0) = adVec(0) / dMagOfVec
        adUnitVec(1) = adVec(1) / dMagOfVec
        adUnitVec(2) = adVec(2) / dMagOfVec
        FnGetDirCosOfVector = adUnitVec
    End Function
    'Get the magnitude of vector
    Public Function FnGetMagnitudeOfVector(ByVal adVec() As Double) As Double
        FnGetMagnitudeOfVector = Sqrt(Pow(adVec(0), 2) + Pow(adVec(1), 2) + Pow(adVec(2), 2))
    End Function
    Function FnFindCommonEdgeToFaces(objFaceA As Face, objFaceB As Face) As Edge
        Dim objStartEdgeVrtA As Point3d = Nothing
        Dim objEndEdgeVrtA As Point3d = Nothing
        Dim objStartEdgeVrtB As Point3d = Nothing
        Dim objEndEdgeVrtB As Point3d = Nothing
        Dim objPointP As Point3d = Nothing
        Dim PA(2) As Double
        Dim PB(2) As Double
        Dim objPtA As Point3d = Nothing
        Dim objPtB As Point3d = Nothing

        For Each objEdgeA As Edge In objFaceA.GetEdges()
            If objEdgeA.SolidEdgeType = Edge.EdgeType.Linear Then
                For Each objEdgeB As Edge In objFaceB.GetEdges()
                    If objEdgeB.SolidEdgeType = Edge.EdgeType.Linear Then
                        If objEdgeA.Equals(objEdgeB) Then
                            FnFindCommonEdgeToFaces = objEdgeA
                            Exit Function
                        End If
                    End If
                Next
            End If
        Next
        FnFindCommonEdgeToFaces = Nothing
    End Function
    'Get the direction vector of a face
    Public Function FnGetDirVecOfFace(ByVal objPart As Part, ByVal ObjFace As Face) As Double()
        Dim dir As Direction = Nothing
        dir = objPart.Directions.CreateDirection(ObjFace, Sense.Forward, SmartObject.UpdateOption.WithinModeling)
        FnGetDirVecOfFace = {dir.Vector.X, dir.Vector.Y, dir.Vector.Z}
    End Function
    'Get the face center
    Public Function FnGetFaceNormal(objFace As Face) As Double()
        Dim iFaceType As Integer = 0
        Dim adCenterPoint(2) As Double
        Dim adDir(2) As Double
        Dim adBox(5) As Double
        Dim dRadius As Double = 0.0
        Dim dRadData As Double = 0.0
        Dim iNormDir As Integer = 0

        'Get the face center , radius in case of cylindrical face and normal direction
        FnGetUFSession.Modl.AskFaceData(objFace.Tag, iFaceType, adCenterPoint,
                                    adDir, adBox, dRadius, dRadData, iNormDir)
        FnGetFaceNormal = adDir
    End Function

    'Code added Jan-04-2018
    'To check for the presence of a text in a given text file 
    Function FnCheckIfFileIsAlreadyProcessed(sStatusTextFilePath As String, sValueToCheck As String) As Boolean
        Dim asContent() As String = Nothing
        If FnCheckFileExists(sStatusTextFilePath) Then
            'Read the full file
            asContent = FnReadFullFile(sStatusTextFilePath)
            If Not asContent Is Nothing Then
                'Read each line
                For Each sLine As String In asContent
                    If sLine.ToUpper.Contains(sValueToCheck.ToUpper) Then
                        FnCheckIfFileIsAlreadyProcessed = True
                        Exit Function
                    End If
                Next
            Else
                FnCheckIfFileIsAlreadyProcessed = False
                Exit Function
            End If
        Else
            FnCheckIfFileIsAlreadyProcessed = False
            Exit Function
        End If
        FnCheckIfFileIsAlreadyProcessed = False
        Exit Function
    End Function

    'Code added Jan-20-2018
    'Function to Get the number of Aligned Face
    Function FnGetNumOfAlignedFace(objBody As Body, adRotationMat() As Double) As Integer
        Dim adFaceNormal() As Double = Nothing
        Dim bIsFaceAligned As Boolean = False
        Dim iCountNumOfAlignedFace As Integer = 0
        'Loop through all the face in a body and check for Parallel Anti-Parallel face, to identify the aligned face
        If Not objBody Is Nothing Then
            For Each objFace As Face In objBody.GetFaces()
                bIsFaceAligned = False

                If objFace.SolidFaceType = Face.FaceType.Planar Then
                    adFaceNormal = FnGetFaceNormal(objFace)
                    If FnParallelAntiParallelCheck(adFaceNormal, {adRotationMat(0), adRotationMat(1), adRotationMat(2)}) Then
                        bIsFaceAligned = True
                    End If
                    If FnParallelAntiParallelCheck(adFaceNormal, {adRotationMat(3), adRotationMat(4), adRotationMat(5)}) Then
                        bIsFaceAligned = True
                    End If
                    If FnParallelAntiParallelCheck(adFaceNormal, {adRotationMat(6), adRotationMat(7), adRotationMat(8)}) Then
                        bIsFaceAligned = True
                    End If
                    If bIsFaceAligned Then
                        iCountNumOfAlignedFace = iCountNumOfAlignedFace + 1
                    End If
                End If
            Next
        End If
        FnGetNumOfAlignedFace = iCountNumOfAlignedFace
    End Function

    'Function to check if the given face is Aligned in Rotation matrix dir
    Function FnCheckIfFaceIsAligned(objFace As Face, adRotationMat() As Double) As Boolean
        Dim bIsFaceAligned As Boolean = False
        Dim adFaceNormal() As Double = Nothing

        If Not objFace Is Nothing Then
            If objFace.SolidFaceType = Face.FaceType.Planar Then
                adFaceNormal = FnGetFaceNormal(objFace)
                If FnParallelAntiParallelCheck(adFaceNormal, {adRotationMat(0), adRotationMat(1), adRotationMat(2)}) Then
                    bIsFaceAligned = True
                End If
                If FnParallelAntiParallelCheck(adFaceNormal, {adRotationMat(3), adRotationMat(4), adRotationMat(5)}) Then
                    bIsFaceAligned = True
                End If
                If FnParallelAntiParallelCheck(adFaceNormal, {adRotationMat(6), adRotationMat(7), adRotationMat(8)}) Then
                    bIsFaceAligned = True
                End If
            End If
        End If
        FnCheckIfFaceIsAligned = bIsFaceAligned
    End Function

    'Code added Feb-27-2018
    Sub sPopulateConfigSheet(objpart As Part, sOEMName As String)
        Dim sDesSource As String = ""
        Dim sAuto2DRunDate As String = ""
        Dim sCorrectDesignSource As String = ""

        'Update the Attribute in the 3D part with the information in the CONFIG Tab
        If FnGetPartAttribute(objpart, "String", DESIGN_SOURCE) = "" Then
            sDesSource = FnReadSingleRowForColumn(_objWorkBk, CONFIGSHEETNAME, 2, 2)
            If Not sDesSource Is Nothing Then
                sCorrectDesignSource = FnGetMappingSupplierName(sOEMName, sDesSource)
                sSetStringUserAttribute(objpart, DESIGN_SOURCE, sCorrectDesignSource)
                sUpdateAttributesInModel()
            End If
        End If
        'Code added to write OEM name
        SWriteValueToCell(_objWorkBk, CONFIGSHEETNAME, 1, 2, sOEMName)

        'Code added May-10-2017
        SWriteValueToCell(_objWorkBk, CONFIGSHEETNAME, 3, 2, MODULE_VERSION)
        'Code added Jun-12-2017
        'Added an attribute  B_AUTO2D_RUN_DATE to collect the date and time the sweep data was created.
        'This attribute will be present in the config sheet of AutoDim and to the property in the part level.
        sAuto2DRunDate = DateAndTime.Now.ToString("yyyyMMddHHmm")

        If sAuto2DRunDate <> "" Then
            'SWriteValueToCell(_objWorkBk, CONFIGSHEETNAME, 4, 1, AUTO2D_RUN_DATE_TIME_ATTR)
            SWriteValueToCell(_objWorkBk, CONFIGSHEETNAME, 4, 2, sAuto2DRunDate)
            sSetStringUserAttribute(objpart, AUTO2D_RUN_DATE_TIME_ATTR, sAuto2DRunDate)
        End If
        'Code added Feb-22-2018
        'Tool Name information is need for Data mining algorithm to be written by Algo team
        'Fetch Tool name information for Tool_Folder_Path text file
        If _sToolNamefromConfigFile <> "" Then
            Dim sToolName As String = ""
            sToolName = Split(_sToolNamefromConfigFile, ".PRT")(0)
            SWriteValueToCell(_objWorkBk, CONFIGSHEETNAME, 5, 2, sToolName)
        End If
        'Code added Feb-28-2018
        'Populate the DIVISION name in config sheet
        If _sOemName = DAIMLER_OEM_NAME Then
            SWriteValueToCell(_objWorkBk, CONFIGSHEETNAME, 6, 2, _sDivision)
        Else
            SWriteValueToCell(_objWorkBk, CONFIGSHEETNAME, 6, 2, "Not Applicable")
        End If

    End Sub
    'Function cofigured to work for all OEM
    Sub sDeleteOldSweepDataInformation(objPart As Part, aoAllCompInSession() As Component)
        'TO DELETE OLD SWEEP DATA EDGE AND FACE NAME RECORDS
        'Identify if old edge and face records exist in the 3D model
        'Also delete any attributes added by VECTRA in any previous runs

        Dim objBody As Body = Nothing
        Dim bDeleteEdgeFaceRecords As Boolean = False

        If Not aoAllCompInSession Is Nothing Then
            For Each objComp As Component In aoAllCompInSession
                If FnGetStringUserAttribute(objPart, VECTRA_SDR) <> VECTRA_SDR_RUN_YES.ToString Then
                    If Not FnGetPartFromComponent(objComp) Is Nothing Then
                        FnLoadPartFully(FnGetPartFromComponent(objComp))
                        For Each body As Body In FnGetPartFromComponent(objComp).Bodies()
                            If FnGetStringUserAttribute(objPart, VECTRA_SDR) <> VECTRA_SDR_RUN_YES.ToString Then
                                If _sOemName = DAIMLER_OEM_NAME Then
                                    If _sDivision = TRUCK_DIVISION Then
                                        If objComp Is objPart.ComponentAssembly.RootComponent Then
                                            objBody = body
                                        Else
                                            objBody = CType(objComp.FindOccurrence(body), Body)
                                        End If
                                    ElseIf _sDivision = CAR_DIVISION Then
                                        If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                            objBody = CType(objComp.FindOccurrence(body), Body)
                                        Else
                                            objBody = body
                                        End If
                                    End If
                                    'Code added Nov-07-2018
                                    'Modified to work for FIAT OEM
                                ElseIf _sOemName = FIAT_OEM_NAME Then
                                    If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                        objBody = CType(objComp.FindOccurrence(body), Body)
                                    Else
                                        objBody = body
                                    End If
                                Else
                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        objBody = body
                                    Else
                                        objBody = CType(objComp.FindOccurrence(body), Body)
                                    End If
                                End If

                                If Not objBody Is Nothing Then
                                    For Each face As Face In objBody.GetFaces()
                                        If FnGetStringUserAttribute(objPart, VECTRA_SDR) <> VECTRA_SDR_RUN_YES.ToString Then
                                            If face.Name.Contains("FACE") Then
                                                sSetStringUserAttribute(objPart, VECTRA_SDR, VECTRA_SDR_RUN_YES)
                                                sUpdateAttributesInModel()
                                                'objPart.SetAttribute(VECTRA_SDR, VECTRA_SDR_RUN_YES)
                                                Exit For
                                            End If
                                            For Each objEdge As Edge In face.GetEdges()
                                                If objEdge.Name.Contains("BODY EDGE") Then
                                                    sSetStringUserAttribute(objPart, VECTRA_SDR, VECTRA_SDR_RUN_YES)
                                                    sUpdateAttributesInModel()
                                                    'objPart.SetAttribute(VECTRA_SDR, VECTRA_SDR_RUN_YES)
                                                    Exit For
                                                End If
                                            Next
                                        Else
                                            Exit For
                                        End If
                                    Next
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    End If
                Else
                    Exit For
                End If
            Next
        Else
            For Each obody As Body In objPart.Bodies()
                If FnGetStringUserAttribute(objPart, VECTRA_SDR) <> VECTRA_SDR_RUN_YES.ToString Then
                    For Each face As Face In obody.GetFaces()
                        If FnGetStringUserAttribute(objPart, VECTRA_SDR) <> VECTRA_SDR_RUN_YES.ToString Then
                            If face.Name.Contains("FACE") Then
                                sSetStringUserAttribute(objPart, VECTRA_SDR, VECTRA_SDR_RUN_YES)
                                sUpdateAttributesInModel()
                                'objPart.SetAttribute(VECTRA_SDR, VECTRA_SDR_RUN_YES)
                                Exit For
                            End If
                            For Each objEdge As Edge In face.GetEdges()
                                If objEdge.Name.Contains("BODY EDGE") Then
                                    sSetStringUserAttribute(objPart, VECTRA_SDR, VECTRA_SDR_RUN_YES)
                                    sUpdateAttributesInModel()
                                    'objPart.SetAttribute(VECTRA_SDR, VECTRA_SDR_RUN_YES)
                                    Exit For
                                End If
                            Next
                        Else
                            Exit For
                        End If
                    Next
                End If
            Next
        End If

        sSetStatus("Deleting old sweep data...please wait")
        'To check if the sweep data is already run
        If FnGetStringUserAttribute(objPart, VECTRA_SDR) = VECTRA_SDR_RUN_YES.ToString Then
            'Delete all the existing face names and edge names
            If Not aoAllCompInSession Is Nothing Then
                For Each objComp As Component In aoAllCompInSession
                    If Not FnGetPartFromComponent(objComp) Is Nothing Then
                        FnLoadPartFully(FnGetPartFromComponent(objComp))
                        For Each obody As Body In FnGetPartFromComponent(objComp).Bodies()
                            If _sOemName = DAIMLER_OEM_NAME Then
                                If _sDivision = TRUCK_DIVISION Then
                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        objBody = obody
                                    Else
                                        objBody = CType(objComp.FindOccurrence(obody), Body)
                                    End If
                                ElseIf _sDivision = CAR_DIVISION Then
                                    If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                        objBody = CType(objComp.FindOccurrence(obody), Body)
                                    Else
                                        objBody = obody
                                    End If
                                End If
                                'Code added Nov-07-2018
                                'Modified to work for FIAT OEM
                            ElseIf _sOemName = FIAT_OEM_NAME Then
                                If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                    objBody = CType(objComp.FindOccurrence(obody), Body)
                                Else
                                    objBody = obody
                                End If
                            Else
                                If objComp Is objPart.ComponentAssembly.RootComponent Then
                                    objBody = obody
                                Else
                                    objBody = CType(objComp.FindOccurrence(obody), Body)
                                End If
                            End If
                            bDeleteEdgeFaceRecords = False
                            If Not objBody Is Nothing Then
                                For Each face As Face In objBody.GetFaces()
                                    'CODE ADDED - 5/13/16 - Amitabh - To identify if the edge name or face name is assigned seperately to the
                                    'occourence or inherited from the prototype edge or face
                                    'In case the edge or face name is inherited from the prototype edge or face, this name cannot be deleted 
                                    'in the occourence as this name is inherited
                                    If face.IsOccurrence Then
                                        If face.Name.ToUpper = face.Prototype.Name.ToUpper Then
                                            bDeleteEdgeFaceRecords = False
                                        Else
                                            bDeleteEdgeFaceRecords = True
                                        End If
                                    Else
                                        bDeleteEdgeFaceRecords = True
                                    End If
                                    If bDeleteEdgeFaceRecords Then
                                        sChangeObjectNames(objPart, face, "")
                                    End If
                                    bDeleteEdgeFaceRecords = False
                                    For Each objEdge As Edge In face.GetEdges()
                                        If objEdge.IsOccurrence Then
                                            If objEdge.Name.ToUpper = objEdge.Prototype.Name.ToUpper Then
                                                bDeleteEdgeFaceRecords = False
                                            Else
                                                bDeleteEdgeFaceRecords = True
                                            End If
                                        Else
                                            bDeleteEdgeFaceRecords = True
                                        End If
                                        If bDeleteEdgeFaceRecords Then
                                            sChangeObjectNames(objPart, objEdge, "")
                                        End If
                                    Next
                                Next
                            End If
                        Next
                    End If
                Next
            Else
                For Each obody As Body In objPart.Bodies()
                    For Each face As Face In obody.GetFaces()
                        sChangeObjectNames(objPart, face, "")
                        For Each objEdge As Edge In face.GetEdges()
                            sChangeObjectNames(objPart, objEdge, "")
                        Next
                    Next
                Next
            End If
        End If

        sSetStatus("Old sweep data deleted successfully")
    End Sub

    'Function to populate all the face information to the face vec tab
    Sub sPopulateFaceInfoInFaceVecTab(objPart As Part, objChildComp As Component, objBody As Body, objFace As Face, iRowFaceVecStart As Integer, sSheetName As String)
        Dim adCenterPoint(2) As Double
        Dim adDir(2) As Double
        Dim adBox(5) As Double
        Dim dRadius As Double = 0.0
        Dim dRadData As Double = 0.0
        Dim iNormDir As Integer = 0
        Dim dir As Direction = Nothing
        Dim iFaceType As Integer = 0
        Dim objGeomProp As GeometricAnalysis.GeometricProperties = Nothing
        Dim objGeomPropFace As GeometricAnalysis.GeometricProperties.Face = Nothing
        Dim sHoleSize As String = ""
        Dim objRefPoint1 As Point3d = Nothing
        Dim objRefPoint2 As Point3d = Nothing
        Dim sPreFab As String = ""
        Dim sFlameCutFace As String = ""
        '_sStartTime = DateTime.Now
        'Populate the FACE VEC TAB
        FnGetUFSession.Modl.AskFaceData(objFace.Tag, iFaceType, adCenterPoint, adDir, adBox, dRadius, dRadData, iNormDir)
        'sCalculateTimeForAPI("Ask Face Data API ")
        'Get the Face Center using a different API
        'Code added on 5/30/2016
        'Code modified on Jun-01-2020
        'FOr CONICAL and REVOLUTION Face get the face center using Ask Face Props API
        ' The earlier code API was not able to fetch the face center for Offset / Blending / Parametric type faces hence using an altermate API for these type of faces
        'Code updated on Jun-29-2020
        'DO NOT CHANGE THE LOGIC FOR CONICAL FACE, SINCE PRADEEP WANTED THE FACE CENTER TO BE ON THE AXIS FOR CONICAL FACE
        If objFace.SolidFaceType.ToString.ToUpper.Contains("OFFSET") Or objFace.SolidFaceType.ToString.ToUpper.Contains("BLENDING") Or
            objFace.SolidFaceType.ToString.ToUpper.Contains("PARAMETRIC") Or 'objFace.SolidFaceType.ToString.ToUpper.Contains("CONICAL") Or
            objFace.SolidFaceType.ToString.ToUpper.Contains("REVOLUTION") Then
            adCenterPoint = FnGetFaceCenter(objFace)
        ElseIf objFace.SolidFaceType.ToString.ToUpper.Contains("CONICAL") Then
            Try
                Dim adConicalFaceCenter() As Double = Nothing

                adConicalFaceCenter = FnGetFaceCenterOfConicalFace(objFace)
                If Not adConicalFaceCenter Is Nothing Then
                    sWriteToLogFile("New Logic used for Conical Face " & objFace.Name)
                    adCenterPoint = adConicalFaceCenter
                End If

            Catch ex As Exception
                sWriteToLogFile("Error encountered in computing Face center for Conical Face " & objFace.Name)
                sWriteToLogFile(ex.Message)
                sWriteToLogFile(ex.StackTrace)
                sWriteToLogFile("ASK Face Data API is used to get the Face center of conical face")
            End Try

        End If
        '_sStartTime = DateTime.Now
        SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceCenterX, adCenterPoint(0).ToString)
        SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceCenterY, adCenterPoint(1).ToString)
        SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceCenterZ, adCenterPoint(2).ToString)
        'Only for cylindrical face
        If iFaceType = 16 Then
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceRadius, dRadius.ToString)
        End If
        SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceDirection, iNormDir.ToString)
        'Code modified on Nov-13-2018
        'In case of FIAT, when populating the face names of Burnout body. Some faces will be same as that of Final Part body.
        'To distinguish that Face which is common in Final Part and Burnout Body, we prepend "FP_" infront of the face names only in sweep data and not in 3D model.
        If sSheetName = BURNOUT_FACEVEC_SHEET Then
            If _sOemName = FIAT_OEM_NAME Then
                If Not objFace.Name.ToUpper.StartsWith("BO_") Then
                    SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceNameFaceVec, "FP_" & objFace.Name)
                Else
                    SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceNameFaceVec, objFace.Name)
                End If
            Else
                SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceNameFaceVec, objFace.Name)
            End If
        Else
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceNameFaceVec, objFace.Name)
        End If

        Try
            dir = objPart.Directions.CreateDirection(objFace, Sense.Forward, SmartObject.UpdateOption.WithinModeling)
            'SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceNameFaceVec, objFace.Name)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceType, objFace.SolidFaceType.ToString)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorX, dir.Vector.X.ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorX)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorY, dir.Vector.Y.ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorY)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorZ, dir.Vector.Z.ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorZ)
            'Compute the face area as well
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceArea,
                              Round(FnCalculateFaceArea(objPart, objFace), 6).ToString)
        Catch ex As Exception
            objGeomProp = objPart.AnalysisManager.CreateGeometricPropertiesObject()
            objFace.GetEdges(0).GetVertices(objRefPoint1, objRefPoint2)
            objGeomProp.GetFaceProperties(objFace, objRefPoint1, objGeomPropFace)
            'SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceNameFaceVec, objFace.Name)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceType, objFace.SolidFaceType.ToString)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorX, objGeomPropFace.NormalInWcs.X.ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorX)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorY, objGeomPropFace.NormalInWcs.Y.ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorY)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorZ, objGeomPropFace.NormalInWcs.Z.ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowFaceVecStart, iColVectorZ)
            'Compute the face area as well
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFaceArea,
                              Round(FnCalculateFaceArea(objPart, objFace), 6).ToString)
            objGeomProp.Destroy()
            'For Testing
            'SWriteValueToCell(_objWorkBk, SHEETFACEVECTORDETAILS, iRowStart, iColHoleSize + 1, "***")
        End Try
        'sCalculateTimeForAPI("Popualte AskFaceData Info in excel sheet ")
        '_sStartTime = DateTime.Now
        'Populate all Face property in Face Vec Sheet
        sPopulateFacePropertyInFaceVecTab(objFace, iRowFaceVecStart, sSheetName)
        'sCalculateTimeForAPI("Populate FaceProperty from Attribute ")
    End Sub
    'Code modified on Feb-28-2018
    Sub sPopulateFacePropertyInFaceVecTab(objFace As Face, iRowFaceVecStart As Integer, sSheetName As String)
        Dim sHoleSize As String = ""
        Dim sPreFab As String = ""
        Dim sFlameCutFace As String = ""
        Dim sFeatName As String = ""
        Dim sMappedFeatName As String = ""

        'Populate Hole_Size attribute
        sHoleSize = FnGetFaceAttribute(objFace, "String", HOLE_SIZE)
        If Not sHoleSize = "" Then
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColHoleSize, sHoleSize)
        End If

        'Populate Pre Fab Attribute
        sPreFab = FnGetFaceAttribute(objFace, "String", PRE_FAB_HOLE_ATTR_TITLE)
        If sPreFab = PRE_FAB_HOLE_ATTR_VALUE Then
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColPreFab, sPreFab)
        End If

        'Code added Jan-20-2018
        'Fetch the flame cut attribute from face and populate it
        sFlameCutFace = FnGetFaceAttribute(objFace, "String", FLAME_CUT_FACE_ATTR)
        If sFlameCutFace = FLAME_CUT_FACE_ATTR_VALUE Then
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFlameCutFace, sFlameCutFace)
        End If
        'Code commented on Jan-30-2020
        'Populate the mapped Featname
        'Populate Feat Name attribute
        'sFeatName = FnGetFaceAttribute(objFace, "String", FEAT_NAME_FACE_ATTR)
        'If sFeatName <> "" Then
        '    SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFeatNameAttr, sFeatName)
        'End If
        'Code added Jan-30-2020
        'Populate the mapped FeatName value
        sFeatName = FnGetStringUserAttribute(objFace, FEAT_NAME_FACE_ATTR)
        sMappedFeatName = FnGetFeatNameMappingValueForSweepData(objFace)
        If sMappedFeatName <> "" Then
            SWriteValueToCell(_objWorkBk, sSheetName, iRowFaceVecStart, iColFeatNameAttr, sMappedFeatName)
        End If
    End Sub
    'Populate Finish Tolerance informaiton of a face in Misc Info Sheet
    Sub sPopulateFinishTolInfoOfFaceInMiscInfoTab(objFace As Face, ByRef iRowFinishStart As Integer)
        'hardcoding the column values based on the defined template
        Dim iColFinishFaceName As Integer = 0
        Dim iColFinishTol As Integer = 0
        Dim sFinishTolValue As String = ""

        iColFinishFaceName = 1
        iColFinishTol = 2
        Try
            'Code modified  - Shanmugam Jan-30-2017
            'Print Finish Tolerance value, only if the attribute value has MICRONS
            'If FnGetStringUserAttribute(face, FINISHTOLATTRNAME) Like FINISHTOLATTRVALUE Then
            sFinishTolValue = FnGetStringUserAttribute(objFace, _FINISHTOLERANCE_ATTR_NAME)
            If sFinishTolValue <> "" Then
                If sFinishTolValue.Contains(FINISHTOLATTRVALUE) Then
                    'Validation added Jan-22-2018
                    If objFace.Name.Trim <> "" Then
                        SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, iRowFinishStart, iColFinishFaceName, objFace.Name)
                        SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, iRowFinishStart, iColFinishTol, sFinishTolValue)
                        'FnGetNxSession.ListingWindow.Open()
                        'FnGetNxSession.ListingWindow.WriteLine("FinishTolerance Face name " & objFace.Name)
                        sWriteToLogFile("Finish Tolerance applied to Face : " & objFace.Name)
                        iRowFinishStart = iRowFinishStart + 1
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    'Populate all the edge information in F_Data sheet
    Sub sPopulateEdgeInfoInFDataTab(objPart As Part, objChildComp As Component, objBody As Body, objFace As Face, objEdge As Edge, origin() As Double, iRowStart As Integer, sSheetName As String)
        Dim arc_data As NXOpen.UF.UFEval.Arc = Nothing
        Dim arc_data_ellipse As NXOpen.UF.UFEval.Ellipse = Nothing
        Dim objPoint1 As Point3d
        Dim objPoint2 As Point3d
        Dim dStartAngle As Double = 0
        Dim dEndAngle As Double = 0

        '_sStartTime = DateTime.Now
        SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColFeatType, "BODY")
        SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeType, objEdge.SolidEdgeType.ToString)

        'Code added Nov-13-2018
        'In case of FIAT, when populating the face names of Burnout body. Some faces will be same as that of Final Part body.
        'To distinguish that Face which is common in Final Part and Burnout Body, we prepend "FP_" infront of the face names only in sweep data and not in 3D model.
        If sSheetName = BURNOUT_FDATA_SHEET Then
            If _sOemName = FIAT_OEM_NAME Then
                If Not objFace.Name.ToUpper.StartsWith("BO_") Then
                    SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColFaceName, "FP_" & objFace.Name)
                Else
                    SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColFaceName, objFace.Name)
                End If
                If Not objEdge.Name.ToUpper.StartsWith("BO_") Then
                    SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeName, "FP_" & objEdge.Name)
                Else
                    SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeName, objEdge.Name)
                End If
            Else
                SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColFaceName, objFace.Name)
                SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeName, objEdge.Name)
            End If
        Else
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColFaceName, objFace.Name)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeName, objEdge.Name)
        End If

        'Get the center point and the diameter data for circular edges
        If objEdge.SolidEdgeType.ToString = "Circular" Then
            arc_data = FnGetEdgeData(objEdge.Tag)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeDia, Round(arc_data.radius * 2, 8).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColEdgeDia)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterX, Round(arc_data.center(0) - origin(0), 8).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterX)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterY, Round(arc_data.center(1) - origin(1), 8).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterY)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterZ, Round(arc_data.center(2) - origin(2), 8).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterZ)
            'Code added Apr-12-2018
            dStartAngle = arc_data.limits(0) / DEGREE
            dEndAngle = arc_data.limits(1) / DEGREE
            'Populate Start and End Angle
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColStartAngle, dStartAngle.ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColStartAngle)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEndAngle, dEndAngle.ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColEndAngle)
            'Populate X Axis Direction Cosines
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColDCXx, arc_data.x_axis(0).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColDCXx)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColDCXy, arc_data.x_axis(1).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColDCXy)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColDCXz, arc_data.x_axis(2).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColDCXz)
            'Populate Y Axis Direction Cosines
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColDCYx, arc_data.y_axis(0).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColDCYx)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColDCYy, arc_data.y_axis(1).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColDCYy)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColDCYz, arc_data.y_axis(2).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColDCYz)
            'Code added on Aug-9-2017
            'Check the Edge Curvature
            If FnIsEdgeConvex(objPart, objEdge) Then
                'Edge is Convex
                SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeCurvature, CONVEX)
            Else
                'Edge is Concave
                SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeCurvature, CONCAVE)
            End If
        ElseIf objEdge.SolidEdgeType.ToString = "Elliptical" Then
            arc_data_ellipse = FnGetEdgeCenterForEllipse(objEdge.Tag)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeDia, Round(arc_data_ellipse.minor * 2, 8).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColEdgeDia)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterX, Round(arc_data_ellipse.center(0) - origin(0), 8).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterX)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterY, Round(arc_data_ellipse.center(1) - origin(1), 8).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterY)
            SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterZ, Round(arc_data_ellipse.center(2) - origin(2), 8).ToString)
            SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColEdgeCenterZ)
            'Code added on Aug-9-2017
            'Check the Edge Curvature
            If FnIsEdgeConvex(objPart, objEdge) Then
                'Edge is Convex
                SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeCurvature, CONVEX)
            Else
                'Edge is Concave
                SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeCurvature, CONCAVE)
            End If
        End If


        objEdge.GetVertices(objPoint1, objPoint2)

        SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColVertex1X, Round(objPoint1.X - origin(0), 8).ToString)
        SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColVertex1X)
        SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColVertex1Y, Round(objPoint1.Y - origin(1), 8).ToString)
        SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColVertex1Y)
        SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColVertex1Z, Round(objPoint1.Z - origin(2), 8).ToString)
        SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColVertex1Z)

        SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColVertex2X, Round(objPoint2.X - origin(0), 8).ToString)
        SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColVertex2X)
        SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColVertex2Y, Round(objPoint2.Y - origin(1), 8).ToString)
        SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColVertex2Y)
        SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColVertex2Z, Round(objPoint2.Z - origin(2), 8).ToString)
        SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColVertex2Z)
        'Code added Feb-22-2018
        'Populate EDge length for all edges
        SWriteValueToCell(_objWorkBk, sSheetName, iRowStart, iColEdgeLength, objEdge.GetLength)
        SChangeToNumberFormat(_objWorkBk, sSheetName, iRowStart, iColEdgeLength)
        'sCalculateTimeForAPI("Populate Edge Info :")
    End Sub

    'Function to collect component solid body
    Function FnCollectComponentSolidBody(objPart As Part) As List(Of Body)
        Dim aoListOfAllSolidBodies As List(Of Body) = Nothing
        Dim objChildComp As Component = Nothing
        Dim aoPartDesignMembers() As Features.Feature = Nothing
        Dim objChildPart As Part = Nothing
        Dim bIsValidSolidBody As Boolean = False
        Dim aoAllValidBody() As Body = Nothing
        'Code modified on Feb-28-2018

        If Not objPart Is Nothing Then
            aoListOfAllSolidBodies = New List(Of Body)

            If _sOemName = GM_OEM_NAME Then
                If Not objPart Is Nothing Then
                    For Each objBody As Body In objPart.Bodies()
                        If objBody.IsSolidBody Then
                            aoListOfAllSolidBodies.Add(objBody)
                        End If
                    Next
                End If

            ElseIf _sOemName = CHRYSLER_OEM_NAME Then
                If Not objPart Is Nothing Then
                    For Each objBody As Body In objPart.Bodies()
                        If objBody.IsSolidBody Then
                            aoListOfAllSolidBodies.Add(objBody)
                        End If
                    Next
                End If

            ElseIf _sOemName = DAIMLER_OEM_NAME Then
                If _sDivision = TRUCK_DIVISION Then
                    'In Case of Truck Component, get the collection of body from objPart directly
                    aoAllValidBody = FnGetValidBodyForOEM(objPart, _sOemName)
                    If Not aoAllValidBody Is Nothing Then
                        For Each objBody As Body In aoAllValidBody
                            aoListOfAllSolidBodies.Add(objBody)
                        Next
                    End If

                ElseIf _sDivision = CAR_DIVISION Then
                    'In Case of Car component, get the collection of body from childcomponent of objPart.
                    'ObjPart will be the component container in case of Car. So get the immediate child of it and get the solid bodies
                    If Not objPart.ComponentAssembly.RootComponent.GetChildren Is Nothing Then
                        objChildComp = objPart.ComponentAssembly.RootComponent.GetChildren(0)
                        If Not objChildComp Is Nothing Then
                            objChildPart = FnGetPartFromComponent(objChildComp)
                            If Not objChildPart Is Nothing Then
                                'aoPartDesignMembers = FnCollectAllMembersOfFeatureGroup(objChildPart, _sFeatureGroupName)
                                aoAllValidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)
                                If Not aoAllValidBody Is Nothing Then
                                    For Each objBody As Body In aoAllValidBody
                                        If objBody.IsSolidBody And objBody.Layer <> 110 Then
                                            If Not objBody.JournalIdentifier.ToUpper.Contains("EXTRACT_") Then
                                                aoListOfAllSolidBodies.Add(objBody)
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                End If
                'Code added Nov-07-2018
            ElseIf _sOemName = FIAT_OEM_NAME Then
                If Not objPart Is Nothing Then
                    For Each objBody As Body In objPart.Bodies()
                        If objBody.IsSolidBody Then
                            'Validation added on Jan-30-2019
                            If Not objBody.IsBlanked Then
                                'Check for layer 1, to avoid Intersection Curves
                                If objBody.Layer = 1 Then
                                    aoListOfAllSolidBodies.Add(objBody)
                                End If
                            End If
                        End If
                    Next
                End If
            ElseIf _sOemName = GESTAMP_OEM_NAME Then
                If Not objPart Is Nothing Then
                    For Each objBody As Body In objPart.Bodies()
                        If objBody.IsSolidBody Then
                            aoListOfAllSolidBodies.Add(objBody)
                        End If
                    Next
                End If
            End If

        End If
        FnCollectComponentSolidBody = aoListOfAllSolidBodies
    End Function

    'Populate Misc Info Sheet
    Sub sPopulateADAInfoInMiscInfoTab(objPart As Part)
        Dim objFaceAttr() As DisplayableObject = Nothing
        Dim iColCutXFace As Integer = 0
        Dim iColCutYFace As Integer = 0
        Dim iColCutZFace As Integer = 0
        Dim iBodySheetRowWrite As Integer = 0
        Dim iNCRowStart As Integer = 0
        Dim iColPartName As Integer = 0
        Dim iColValue As Integer = 0
        Dim iRowPartNameStart As Integer = 0
        Dim sValue As String
        Dim iFloorMountFaceRowStart As Integer = 0

        iNCRowStart = NC_PART_CONTACT_FACE_START_ROW_WRITE
        iColNCPartContactFace = 16
        iColPartName = 10
        iColValue = 11
        iRowPartNameStart = MISC_INFO_START_ROW_WRITE
        iColFloorMountFace = 17
        iFloorMountFaceRowStart = FLOOR_MOUNT_FACE_START_ROW_WRITE
        sValue = FnGetStringUserAttribute(objPart, _PART_NAME)
        SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, iRowPartNameStart, iColPartName, DB_PART_NAME_IN_TEMPLATE)
        SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, iRowPartNameStart, iColValue, sValue)
        iRowPartNameStart = iRowPartNameStart + 1

        'Write the part name
        SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 1, 14, objPart.Leaf)

        ' Writing Weldment Type
        If FnGetPartAttribute(objPart, "String", B_WELDMENT_TYPE) = "FRAME" Then
            SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 4, 11, "FRAME")
        ElseIf FnGetPartAttribute(objPart, "String", B_WELDMENT_TYPE) = "WELDMENT" Then
            SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 4, 11, "WELDMENT")
        ElseIf FnGetPartAttribute(objPart, "String", B_WELDMENT_TYPE) = "2-FRAME" Then
            SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 4, 11, "2-FRAME")
        ElseIf FnGetPartAttribute(objPart, "String", B_WELDMENT_TYPE) = "NON-FRAME" Then
            SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 4, 11, "NON-FRAME")
        End If

        'Writing X Machine Start Face
        '**************************** FRAME-1 ********************************************************************
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_SHORIZVALUEMC_FRAME1, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 3, 4, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Y Machining Start Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_SVERTICALVALUEMC_FRAME1, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 4, 4, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Z Machining Start Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_SLATERALVALUEMC_FRAME1, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 5, 4, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Datum Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_DATUM_FACE_FRAME1, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 6, 4, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Datum Hole
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_DATUM_HOLE_FRAME1, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 7, 4, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If


        'Writing Machine Face at Direction X
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_1ST_MC_FACE_X, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 8, 4, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Machine Face at Direction Y
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_1ST_MC_FACE_Y, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 9, 4, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        '*******************************************************************************************************************

        'CODE ADDED - 4/18/16 - Amitabh 
        '**************************** FRAME-2 ********************************************************************
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_SHORIZVALUEMC_FRAME2, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 3, 18, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Y Machining Start Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_SVERTICALVALUEMC_FRAME2, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 4, 18, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Z Machining Start Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_SLATERALVALUEMC_FRAME2, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 5, 18, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Datum Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_DATUM_FACE_FRAME2, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 6, 18, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Datum Hole
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_DATUM_HOLE_FRAME2, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 7, 18, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If


        'Writing Machine Face at Direction U
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_1ST_MC_FACE_U, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 8, 18, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Machine Face at Direction V
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_1ST_MC_FACE_V, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 9, 18, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        '*******************************************************************************************************************

        'Writing X Fab Origin Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_X_FAB, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 3, 7, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Y Fab Origin Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_Y_FAB, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 4, 7, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Z Fab Origin Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_Z_FAB, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 5, 7, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing A Fab Origin face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_U_FAB, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 6, 7, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing B Fab Origin Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_V_FAB, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 7, 7, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing C Fab Origin Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_W_FAB, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then

            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, 8, 7, CType(objDisp, Face).Name)
                    Exit For
                End If
            Next
            objFaceAttr = Nothing
        End If


        iColBodyName = FnGetColumnNumberByName(_objWorkBk, BODYSHEETNAME, BODY_NAME, BODY_NAME_START_COL_WRITE)
        iColCutXFace = FnGetColumnNumberByName(_objWorkBk, BODYSHEETNAME, CUT_X_FACE, BODY_NAME_START_COL_WRITE)
        iColCutYFace = FnGetColumnNumberByName(_objWorkBk, BODYSHEETNAME, CUT_Y_FACE, BODY_NAME_START_COL_WRITE)
        iColCutZFace = FnGetColumnNumberByName(_objWorkBk, BODYSHEETNAME, CUT_Z_FACE, BODY_NAME_START_COL_WRITE)


        'Writing X Cut Origin Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_CUT_X_FACE, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    iBodySheetRowWrite = FnFindRowNumberByColumnAndValue(_objWorkBk, BODYSHEETNAME, iColBodyName, BODY_NAME_START_COL_WRITE,
                                                                  CType(objDisp, Face).GetBody().JournalIdentifier.ToString())
                    If iBodySheetRowWrite <> 0 Then
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iBodySheetRowWrite, iColCutXFace, CType(objDisp, Face).Name)
                        Exit For
                    End If
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Y Cut Origin Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_CUT_Y_FACE, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    iBodySheetRowWrite = FnFindRowNumberByColumnAndValue(_objWorkBk, BODYSHEETNAME, iColBodyName, BODY_NAME_START_COL_WRITE,
                                                                CType(objDisp, Face).GetBody().JournalIdentifier.ToString())
                    If Not iBodySheetRowWrite = 0 Then
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iBodySheetRowWrite, iColCutYFace, CType(objDisp, Face).Name)
                        Exit For
                    End If
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Z Cut Origin Face
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, B_CUT_Z_FACE, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then

            For Each objDisp As DisplayableObject In objFaceAttr
                If CType(objDisp, Face).Name.Trim <> "" Then
                    iBodySheetRowWrite = FnFindRowNumberByColumnAndValue(_objWorkBk, BODYSHEETNAME, iColBodyName, BODY_NAME_START_COL_WRITE,
                                                                 CType(objDisp, Face).GetBody().JournalIdentifier.ToString())
                    If Not iBodySheetRowWrite = 0 Then
                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iBodySheetRowWrite, iColCutZFace, CType(objDisp, Face).Name)
                        Exit For
                    End If
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Writing Part Contact Face to NC Blocks
        'CODE COMMENTED - 4/15/16 - components other than NC can also have Part Contact Face
        'If FnChkIfPartIsNC(FnGetPartName(objPart)) Then
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, NC_Contact_FACE_ATTRIBUTE, "")
        If Not objFaceAttr Is Nothing Then
            For Each objface As DisplayableObject In objFaceAttr
                'Validation added on Jan-10-2018
                'Face which doesnot have face name and which has the PCF attribute should not be reported in the MISC_INFO sheet.
                If CType(objface, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, iNCRowStart, iColNCPartContactFace, CType(objface, Face).Name)
                    iNCRowStart = iNCRowStart + 1
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Code added May-10-2019
        'Populate FLoor Mount faces on the Sweep Data
        objFaceAttr = FnGetFaceObjectByAttributes(objPart, FLOOR_MOUNT_ATTRIBUTE, ASSIGN_ATTR)
        If Not objFaceAttr Is Nothing Then
            For Each objFace As DisplayableObject In objFaceAttr
                If CType(objFace, Face).Name.Trim <> "" Then
                    SWriteValueToCell(_objWorkBk, MISCINFOSHEETNAME, iFloorMountFaceRowStart, iColFloorMountFace, CType(objFace, Face).Name)
                    iFloorMountFaceRowStart = iFloorMountFaceRowStart + 1
                End If
            Next
            objFaceAttr = Nothing
        End If

        'Code commented on May-10-2019
        ''Code added on Sep-12-2018
        ''Get the Floor Mounting face Attribute from the face and add it to the sweep data sheet
        'objFaceAttr = FnGetFaceObjectByAttributes(objPart, FLOOR_MOUNT_ATTRIBUTE, ASSIGN_ATTR)
        'If Not objFaceAttr Is Nothing Then
        '    For Each objface As DisplayableObject In objFaceAttr
        '        If CType(objface, Face).Name.Trim <> "" Then
        '            Dim adFaceNormalVec() As Double = Nothing
        '            Dim adFloorMountDir(8) As Double
        '            adFaceNormalVec = FnGetFaceNormal(objface)
        '            If Not adFaceNormalVec Is Nothing Then
        '                adFloorMountDir(0) = 0
        '                adFloorMountDir(1) = 0
        '                adFloorMountDir(2) = 0
        '                adFloorMountDir(3) = 0
        '                adFloorMountDir(4) = 0
        '                adFloorMountDir(5) = 0
        '                adFloorMountDir(6) = adFaceNormalVec(0)
        '                adFloorMountDir(7) = adFaceNormalVec(1)
        '                adFloorMountDir(8) = adFaceNormalVec(2)
        '                sPopulateFloorMountFaceNormalToExcelFile(adFloorMountDir)
        '            End If
        '        End If
        '    Next
        '    objFaceAttr = Nothing
        'End If
    End Sub

    'Code added Feb-28-2018
    Sub sAssignPSDColumnHeaderNums()
        iColFeatType = 1
        iColFeatName = 2
        iColFaceName = 3
        iColEdgeName = 4
        iColEdgeType = 5
        iColEdgeDia = 6
        iColEdgeCenterX = 7
        iColEdgeCenterY = 8
        iColEdgeCenterZ = 9

        iColVertex1X = 10
        iColVertex1Y = 11
        iColVertex1Z = 12

        iColVertex2X = 13
        iColVertex2Y = 14
        iColVertex2Z = 15

        iColEdgeLength = 16
        'Code added Apr-12-2018
        iColStartAngle = 17
        iColEndAngle = 18
        iColDCXx = 19
        iColDCXy = 20
        iColDCXz = 21
        iColDCYx = 22
        iColDCYy = 23
        iColDCYz = 24

        iColCallout = 25
        iColEdgeCurvature = 26

        iColHoleName = 1
        iColHoleVecX = 2
        iColHoleVecY = 3
        iColHoleVecZ = 4

        iColBodyName = 1
        iColBodyShape = 2
        iColBodyDetailNos = 3
        iColBodyToolClass = 4
        iColStockSize = 8
        iColComponentName = 9
        iColBodyLayer = 25
        iColCompDBPartName = 26
        iColPMat = 27

        'Get the Face Vector details
        iColFaceNameFaceVec = 1
        iColVectorX = 3
        iColVectorY = 4
        iColVectorZ = 5
        iColFaceType = 2
        iColHoleSize = 6
        iColFaceArea = 7
        iColPreFab = 8
        iColFeatNameAttr = 9
        iColFlameCutFace = 15


        iColFaceCenterX = 10
        iColFaceCenterY = 11
        iColFaceCenterZ = 12
        iColFaceRadius = 13
        iColFaceDirection = 14
    End Sub


    'Public Sub sWriteWeldmentDataInBodyNameSheetAlt(ByVal objPart As Part, sConfigFolderPath As String, aoAllCompInSession() As Component)
    '    Dim sStockSize As String = ""
    '    Dim iRowStart As Integer = 0
    '    Dim sShape As String = ""
    '    Dim sBodyName As String = ""
    '    Dim bSubComponent As Boolean = False
    '    Dim s3DErrDesc As String = ""
    '    'Dim sFolderName As String = ""
    '    Dim sToolClass As String = ""
    '    Dim sPMat As String = ""
    '    Dim objChildPart As Part = Nothing
    '    Dim aoPartDesignMembers() As Features.Feature = Nothing
    '    Dim bIsValidSolidBody As Boolean = False
    '    'Dim objCompToFetchAttribute As Component = Nothing
    '    Dim bProcess As Boolean = False
    '    Dim aoListOfAllSolidBodies As List(Of Body) = Nothing

    '    'For computing the exact view bounds of the bodies
    '    'Dim min_corner(2) As Double
    '    'Dim directions(2, 2) As Double
    '    'Dim distances(2) As Double

    '    iRowStart = BODY_INFO_START_ROW_WRITE
    '    Dim dictStockSizeCompData As Dictionary(Of String, NXObject()) = Nothing
    '    dictStockSizeCompData = New Dictionary(Of String, NXObject())

    '    If Not aoAllCompInSession Is Nothing Then
    '        sAddStockSizeAttribute(objPart, aoAllCompInSession)
    '        For Each objComp As Component In aoAllCompInSession
    '            bProcess = False

    '            If _sDivision = TRUCK_DIVISION Then
    '                bProcess = True
    '            ElseIf _sDivision = CAR_DIVISION Then
    '                If FnCheckIfThisIsAChildCompContainerInWeldment(objComp.DisplayName.ToUpper) Then
    '                    bProcess = True
    '                Else
    '                    bProcess = False
    '                End If
    '            End If
    '            If bProcess Then
    '                objChildPart = FnGetPartFromComponent(objComp)
    '                'Collect solid Body from the component, based on the division (CAR/TRUCK)
    '                aoListOfAllSolidBodies = FnCollectComponentSolidBody(objChildPart)
    '                If Not aoListOfAllSolidBodies Is Nothing Then
    '                    If Not objChildPart Is Nothing Then
    '                        FnLoadPartFully(objChildPart)

    '                        aoPartDesignMembers = FnCollectAllMembersOfFeatureGroup(objChildPart, _sFeatureGroupName)
    '                        For Each objbody As Body In aoListOfAllSolidBodies
    '                            bIsValidSolidBody = False
    '                            If aoPartDesignMembers Is Nothing Then
    '                                bIsValidSolidBody = True
    '                            Else
    '                                For Each objPartDesignMember In aoPartDesignMembers
    '                                    If objPartDesignMember.JournalIdentifier = objbody.JournalIdentifier Then
    '                                        bIsValidSolidBody = True
    '                                    End If
    '                                Next
    '                            End If
    '                            If bIsValidSolidBody Then
    '                                'Validation added Dec-19-2017
    '                                If objbody.IsSolidBody Then

    '                                    'Check if the body belongs to the root comp or the sub assembly , assign a unique name as reuired by core algo
    '                                    'by joining the body name with the component instance tag
    '                                    'Check whether the body is a solid body
    '                                    'Only pick bodies which are in layer 1 (other side may also be present in the same part which need not be detailed) - 26/2/2014
    '                                    'If Body.IsSolidBody Then 'And (body.Layer = 1 Or (body.Layer >= 92 And body.Layer <= 184)) Then
    '                                    'sSetStatus("Collecting attribute data for " & Body.JournalIdentifier.ToUpper)
    '                                    'sToolClass = FnGetBodyAttribute(Body, "String", TOOL_CLASS)
    '                                    'If Not sToolClass = "" Then
    '                                    SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyToolClass, NOTAPPLICABLE)
    '                                    'End If
    '                                    'Check whether it is some body which is not suppose to be generated in sweep data
    '                                    'CODE MODIFIED - 6/13/16 - Amitabh - Ignore WIRE MESH bodies with respect to the SHAPE attribute
    '                                    If (FnGetStringUserAttribute(objComp, SHAPE) <> WIRE_MESH_SHAPE) Then
    '                                        If objComp Is objPart.ComponentAssembly.RootComponent Then
    '                                            sBodyName = objbody.JournalIdentifier
    '                                            ''Add this body to the collection of solid bodies
    '                                            'sStoreSolidBody(Body)

    '                                            'bSubComponent = False
    '                                            'Add the GM Toolkit Attributes
    '                                            'sSetGMToolkitAttributes(objPart, Body, False, False)
    '                                            'If FnGetBodyAttribute(body, "String", PURCH_OPTION).Contains("M") Then
    '                                            'To Output the stock size
    '                                            sStockSize = FnGetStringUserAttribute(objPart, STOCK_SIZE_METRIC)

    '                                        Else
    '                                            sBodyName = CType(objComp.FindOccurrence(objbody), Body).JournalIdentifier & " " & objComp.JournalIdentifier
    '                                            ''Add this occurence body to the collection of solid bodies
    '                                            'sStoreSolidBody(CType(objComp.FindOccurrence(Body), Body))
    '                                            'bSubComponent = True

    '                                            'Add the DB_PART_NAME info for all the child components in a weldment
    '                                            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColCompDBPartName, _
    '                                                                                    FnGetCompAttribute(objComp, "String", DB_PART_NAME))

    '                                            'Add the GM Toolkit Attributes
    '                                            'sSetGMToolkitAttributes(objPart, objComp, True, False)
    '                                            'If FnGetCompAttribute(objComp, "String", PURCH_OPTION).Contains("M") Then
    '                                            'To Output the stock size
    '                                            sStockSize = FnGetStringUserAttribute(objComp, STOCK_SIZE_METRIC)
    '                                        End If
    '                                        If sStockSize <> "" Then
    '                                            If Not dictStockSizeCompData.ContainsKey(sStockSize) Then
    '                                                dictStockSizeCompData.Add(sStockSize, {objComp})
    '                                            Else
    '                                                ReDim Preserve dictStockSizeCompData(sStockSize)(UBound(dictStockSizeCompData(sStockSize)) + 1)
    '                                                dictStockSizeCompData(sStockSize)(UBound(dictStockSizeCompData(sStockSize))) = objComp
    '                                            End If
    '                                        End If

    '                                        'Update Value to the cell
    '                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyName, sBodyName)
    '                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColComponentName, objComp.Parent.DisplayName)
    '                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyLayer, objbody.Layer.ToString)
    '                                        sShape = FnGetStringUserAttribute(objComp, SHAPE)
    '                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColBodyShape, sShape)
    '                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColStockSize, sStockSize)
    '                                        sPMat = FnGetStringUserAttribute(objComp, P_MAT)
    '                                        SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iRowStart, iColPMat, sPMat)
    '                                        If sStockSize = "" Then
    '                                            'Populate the 3D exception report in case of missing stock size in the component
    '                                            s3DErrDesc = "Stock size is missing in the 3D model part " & objPart.Leaf.ToString
    '                                            SWrite(s3DErrDesc, Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, THREE_DIMENSIONAL_MODEL_EXCEPTION_REPORT_FILE_NAME))
    '                                        End If
    '                                        iRowStart = iRowStart + 1
    '                                    End If
    '                                    'End If
    '                                End If
    '                            End If
    '                        Next
    '                        ' ''Write the names of all Parts present within the weldment. So that when running part sweep data for components these parts can be neglected
    '                        ''SWrite(objComp.DisplayName, Path.Combine(_sSweepDataOutputFolderPath, UNIT_SWEEP_DATA, _sToolFolderName, LIST_OF_PARTS_IN_WELDMENT_FILE_NAME))
    '                    End If
    '                End If
    '            End If
    '        Next
    '    End If

    '    'Now Check for any physical differences and then renumber sub details if required
    '    sSetStatus("Analysing solid bodies for sub-details ")
    '    sCheckForPhysicalDifferencesInSubDetailBasedOnFaceArea(objPart, dictStockSizeCompData)
    'End Sub

    'Code added Apr-17-2018
    'Code to populate information for BurnOut Body
    Sub sPopulateBurnOutBodyFaceAndEdgeInfo(objPart As Part, aoAllCompInSession() As Component, sOemName As String)

        'Dim aoAllCompInSession() As Component = Nothing
        Dim objChildPart As Part = Nothing
        Dim objBodyToAnalyse As Body = Nothing
        Dim sBodyName As String = ""
        Dim iCount As Integer = 0
        Dim iFaceCount As Integer = 0
        Dim iRowStart As Integer = 0
        Dim iRowFaceVecStart As Integer = 0
        Dim wcs As NXOpen.Tag = NXOpen.Tag.Null
        Dim wcs_mx As NXOpen.Tag = NXOpen.Tag.Null

        Dim origin(2) As Double
        Dim wcs_mx_vals(8) As Double
        Dim aoAllBurnOutBody() As Body = Nothing

        'aoAllCompInSession = FnGetAllComponentsInSession()
        iCount = 1
        iFaceCount = 1
        iRowStart = START_ROW_WRITE
        iRowFaceVecStart = 2

        'Coordinate system Data
        FnGetUFSession.Csys.AskWcs(wcs)
        FnGetUFSession.Csys.AskCsysInfo(wcs, wcs_mx, origin)

        If Not aoAllCompInSession Is Nothing Then
            For Each objChildComp As Component In aoAllCompInSession
                objChildPart = FnGetPartFromComponent(objChildComp)
                If Not objChildPart Is Nothing Then
                    'Load the part fully
                    FnLoadPartFully(objChildPart)
                    'Code added Nov-12-2018
                    aoAllBurnOutBody = FnGetValidBurnoutBodyForOEM(objChildPart, sOemName)
                    If Not aoAllBurnOutBody Is Nothing Then
                        For Each objBody As Body In aoAllBurnOutBody
                            If (sOemName = DAIMLER_OEM_NAME) Then
                                If _sDivision = TRUCK_DIVISION Then
                                    'Check if it is the root component or it is a sub assembly compoenent
                                    If (objChildComp Is objPart.ComponentAssembly.RootComponent) Then
                                        'Component in truck. Get Prototype body
                                        objBodyToAnalyse = objBody
                                        sBodyName = objBodyToAnalyse.JournalIdentifier
                                    Else
                                        'Weldment in truck. Get Occurrence body
                                        objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                        If Not objBodyToAnalyse Is Nothing Then
                                            sBodyName = objBodyToAnalyse.JournalIdentifier & " " & objChildComp.JournalIdentifier
                                        End If
                                    End If
                                ElseIf _sDivision = CAR_DIVISION Then
                                    'Code added Apr-05-2019
                                    'In case of Daimler, Pradeep wanted the same Final Body name to be output as the Burnout Body Name.
                                    '(Daimler has two different solid body (Final body and Burnout body))

                                    Dim aoAllValidSolidBody() As Body = Nothing
                                    Dim objFinalBody As Body = Nothing
                                    Dim objFinalOccBody As Body = Nothing

                                    aoAllValidSolidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)
                                    If Not aoAllValidSolidBody Is Nothing Then
                                        'In case of Daimler FnGetValidBodyFOrOEM function will return only one solid body.
                                        objFinalBody = aoAllValidSolidBody(0)
                                    End If

                                    'Check if the component is a child component in weldment
                                    If FnCheckIfThisIsAChildCompInWeldment(objChildComp, _sOemName) Then
                                        'Weldment in Car. Get the Occurrence body
                                        objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                        'Code commented on Apr-05-2019
                                        ''If Not objBodyToAnalyse Is Nothing Then
                                        ''    'when populating body name, give Body journal identifier and component container journalidentifier
                                        ''    sBodyName = objBodyToAnalyse.JournalIdentifier & " " & objChildComp.Parent.JournalIdentifier
                                        ''End If
                                        If Not objFinalBody Is Nothing Then
                                            objFinalOccBody = CType(objChildComp.FindOccurrence(objFinalBody), Body)
                                            If Not objFinalOccBody Is Nothing Then
                                                sBodyName = objFinalOccBody.JournalIdentifier & " " & objChildComp.Parent.JournalIdentifier
                                            End If
                                        End If
                                    Else
                                        'Component in Car. Get the Prototype Body
                                        'Code modified on May-14-2018
                                        objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                        If Not objBodyToAnalyse Is Nothing Then
                                            objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                            'sBodyName = objBody.JournalIdentifier
                                        End If
                                        If Not objFinalBody Is Nothing Then
                                            sBodyName = objFinalBody.JournalIdentifier
                                        End If

                                    End If
                                End If
                            ElseIf (sOemName = FIAT_OEM_NAME) Then
                                'Check if the component is a child component in weldment
                                If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                    'Fiat component
                                    objBodyToAnalyse = objBody
                                    sBodyName = objBodyToAnalyse.JournalIdentifier
                                    'Code modified on Jan-22-2019
                                    'Pradeep wanted the Body Name to be the parent body name itself.
                                    'cODE MODIFIED ON DEC -28-2018
                                    'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                                    'sBodyName = "BO_" & sBodyName

                                Else
                                    'This is a Fiat Weldment child component
                                    objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                    If Not objBodyToAnalyse Is Nothing Then
                                        'when populating body name, give Body journal identifier and Child component journalidentifier
                                        sBodyName = objBodyToAnalyse.JournalIdentifier & " " & objChildComp.JournalIdentifier
                                        'Code modified on Jan-22-2019
                                        'Pradeep wanted the Body Name to be the parent body name itself.
                                        'cODE MODIFIED ON DEC -28-2018
                                        'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                                        'sBodyName = "BO_" & sBodyName
                                    End If
                                End If
                            End If
                            If Not objBodyToAnalyse Is Nothing Then
                                'sStoreSolidBody(objBodyToAnalyse)
                                sSetStatus("Collecting data for " & sBodyName.ToUpper)
                                'Check whether the body is a solid body and the solid body should be on a range of layers
                                If objBodyToAnalyse.IsSolidBody Then 'And (body.Layer = 1 Or (body.Layer >= 92 And body.Layer <= 184)) Then
                                    For Each objFaceToAnalyse As Face In objBodyToAnalyse.GetFaces()
                                        If (sOemName = DAIMLER_OEM_NAME) Then
                                            If _bIsComponent Then
                                                objFaceToAnalyse.Prototype.SetName("BO_FACE " & iFaceCount.ToString())
                                            Else
                                                objFaceToAnalyse.SetName("BO_FACE " & iFaceCount.ToString())
                                            End If
                                            iFaceCount = iFaceCount + 1
                                        ElseIf (sOemName = FIAT_OEM_NAME) Then
                                            If Not objFaceToAnalyse.Name Is Nothing Then
                                                If Not objFaceToAnalyse.Name.ToUpper.Contains("FACE") Then
                                                    objFaceToAnalyse.SetName("BO_FACE " & iFaceCount.ToString())
                                                    iFaceCount = iFaceCount + 1
                                                End If
                                            Else
                                                objFaceToAnalyse.SetName("BO_FACE " & iFaceCount.ToString())
                                                iFaceCount = iFaceCount + 1
                                            End If
                                        End If

                                        'Populate Face information in Face Vec sheet
                                        sPopulateFaceInfoInFaceVecTab(objPart, objChildComp, objBodyToAnalyse, objFaceToAnalyse, iRowFaceVecStart, BURNOUT_FACEVEC_SHEET)
                                        iRowFaceVecStart = iRowFaceVecStart + 1

                                        ''Populate Face attributes in MiscInfo Sheet
                                        'sPopulateFinishTolInfoOfFaceInMiscInfoTab(objFaceToAnalyse, iRowFinishStart)

                                        For Each objEdge As Edge In objFaceToAnalyse.GetEdges()
                                            SWriteValueToCell(_objWorkBk, BURNOUT_FDATA_SHEET, iRowStart, iColFeatName, sBodyName)
                                            If (sOemName = DAIMLER_OEM_NAME) Then
                                                If _bIsComponent Then
                                                    If Not objEdge.Prototype.Name Is Nothing Then
                                                        If Not objEdge.Prototype.Name.Contains("BO_BODY EDGE") Then
                                                            objEdge.Prototype.SetName("BO_BODY EDGE " & iCount.ToString)
                                                            iCount = iCount + 1
                                                        End If
                                                    Else
                                                        objEdge.Prototype.SetName("BO_BODY EDGE " & iCount.ToString)
                                                        iCount = iCount + 1
                                                    End If
                                                Else
                                                    If Not objEdge.Name Is Nothing Then
                                                        If Not objEdge.Name.Contains("BO_BODY EDGE") Then
                                                            objEdge.SetName("BO_BODY EDGE " & iCount.ToString)
                                                            iCount = iCount + 1
                                                        Else
                                                            'CODE ADDED - 5/13/16 - Amitabh - To add unique edge names to occurrence edges
                                                            If objEdge.IsOccurrence Then
                                                                If objEdge.Name.ToUpper = objEdge.Prototype.Name.ToUpper Then
                                                                    objEdge.SetName("BO_BODY EDGE " & iCount.ToString)
                                                                    iCount = iCount + 1
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        objEdge.SetName("BO_BODY EDGE " & iCount.ToString)
                                                        iCount = iCount + 1
                                                    End If
                                                End If
                                            ElseIf (sOemName = FIAT_OEM_NAME) Then
                                                If Not objEdge.Name Is Nothing Then
                                                    If Not objEdge.Name.ToUpper.Contains("BODY EDGE") Then
                                                        objEdge.SetName("BO_BODY EDGE " & iCount.ToString)
                                                        iCount = iCount + 1
                                                    End If
                                                Else
                                                    objEdge.SetName("BO_BODY EDGE " & iCount.ToString)
                                                    iCount = iCount + 1
                                                End If
                                            End If

                                            'Populate EDge information in F_Data sheet
                                            sPopulateEdgeInfoInFDataTab(objPart, objChildComp, objBodyToAnalyse, objFaceToAnalyse, objEdge, origin, iRowStart, BURNOUT_FDATA_SHEET)
                                            iRowStart = iRowStart + 1
                                        Next
                                        'iCount = iCount + 1

                                    Next
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Else
            'Code added Nov-12-2018
            aoAllBurnOutBody = FnGetValidBurnoutBodyForOEM(objPart, sOemName)
            If Not aoAllBurnOutBody Is Nothing Then
                For Each objbody As Body In aoAllBurnOutBody
                    'Check whether the body is a solid body
                    If objbody.IsSolidBody Then ' And (body.Layer = 1 Or (body.Layer >= 92 And body.Layer <= 184)) Then
                        sSetStatus("Collecting data for " & objbody.JournalIdentifier.ToUpper)
                        For Each objFace As Face In objbody.GetFaces()
                            If (sOemName = DAIMLER_OEM_NAME) Then
                                objFace.SetName("BO_FACE " & iFaceCount.ToString())
                                iFaceCount = iFaceCount + 1
                            ElseIf (sOemName = FIAT_OEM_NAME) Then
                                If Not objFace.Name Is Nothing Then
                                    If Not objFace.Name.ToUpper.Contains("FACE") Then
                                        objFace.SetName("BO_FACE " & iFaceCount.ToString())
                                        iFaceCount = iFaceCount + 1
                                    End If
                                Else
                                    objFace.SetName("BO_FACE " & iFaceCount.ToString())
                                    iFaceCount = iFaceCount + 1
                                End If
                            End If

                            'Populate Face information in Face Vec sheet
                            sPopulateFaceInfoInFaceVecTab(objPart, Nothing, objbody, objFace, iRowFaceVecStart, BURNOUT_FACEVEC_SHEET)
                            iRowFaceVecStart = iRowFaceVecStart + 1

                            ''Populate Face attributes in MiscInfo Sheet
                            'sPopulateFinishTolInfoOfFaceInMiscInfoTab(objFace, iRowFinishStart)

                            For Each objEdge As Edge In objFace.GetEdges()
                                'Code modified on Jan-22-2019
                                'Pradeep wanted the Body Name to be the parent body name itself.
                                'cODE MODIFIED ON DEC -28-2018
                                'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                                'SWriteValueToCell(_objWorkBk, BURNOUT_FDATA_SHEET, iRowStart, iColFeatName, "BO_" & objbody.JournalIdentifier)
                                SWriteValueToCell(_objWorkBk, BURNOUT_FDATA_SHEET, iRowStart, iColFeatName, objbody.JournalIdentifier)
                                If (sOemName = DAIMLER_OEM_NAME) Then
                                    If Not objEdge.Name Is Nothing Then
                                        If Not objEdge.Name.Contains("BO_BODY EDGE") Then
                                            objEdge.SetName("BO_BODY EDGE " & iCount.ToString)
                                            iCount = iCount + 1
                                        End If
                                    Else
                                        objEdge.SetName("BO_BODY EDGE " & iCount.ToString)
                                        iCount = iCount + 1
                                    End If
                                ElseIf (sOemName = FIAT_OEM_NAME) Then
                                    If Not objEdge.Name Is Nothing Then
                                        If Not objEdge.Name.Contains("BODY EDGE") Then
                                            objEdge.SetName("BO_BODY EDGE " & iCount.ToString)
                                            iCount = iCount + 1
                                        End If
                                    Else
                                        objEdge.SetName("BO_BODY EDGE " & iCount.ToString)
                                        iCount = iCount + 1
                                    End If
                                End If
                                'Populate EDge information in F_Data sheet
                                sPopulateEdgeInfoInFDataTab(objPart, Nothing, objbody, objFace, objEdge, origin, iRowStart, BURNOUT_FDATA_SHEET)
                                iRowStart = iRowStart + 1
                            Next
                            'iCount = iCount + 1
                        Next
                        'End If
                    End If
                Next
            End If
        End If
    End Sub
    'Populate BurnOutBody Information in Burnout BodyNames sheet
    Sub sPopulateBurnoutBodyDataInBodyNameSheet(ByVal objPart As Part, aoAllCOmpInSession() As Component, sOemName As String)

        Dim iRowStart As Integer = 0
        Dim sShape As String = ""
        Dim sBodyName As String = ""
        Dim bSubComponent As Boolean = False
        Dim sToolClass As String = ""
        Dim sPMat As String = ""
        Dim objChildPart As Part = Nothing
        Dim objCompToFetchAttribute As Component = Nothing
        'Dim aoAllCompInSession() As Component = Nothing
        Dim bPopulateInExcel As Boolean = False
        Dim sStockSize As String = ""
        Dim aoAllBurnOutBody() As Body = Nothing

        iRowStart = BODY_INFO_START_ROW_WRITE
        'aoAllCompInSession = FnGetAllComponentsInSession()

        If Not aoAllCOmpInSession Is Nothing Then
            For Each objComp As Component In aoAllCOmpInSession
                objChildPart = FnGetPartFromComponent(objComp)
                If Not objChildPart Is Nothing Then
                    FnLoadPartFully(objChildPart)
                    'Code added Nov-13-2018
                    aoAllBurnOutBody = FnGetValidBurnoutBodyForOEM(objChildPart, sOemName)
                    If Not aoAllBurnOutBody Is Nothing Then
                        For Each objbody As Body In aoAllBurnOutBody
                            bPopulateInExcel = False
                            If objbody.IsSolidBody Then
                                If (sOemName = DAIMLER_OEM_NAME) Then
                                    'In case of Car division, the component attributes would be added at the container level.
                                    If _sDivision = TRUCK_DIVISION Then
                                        objCompToFetchAttribute = objComp
                                    ElseIf _sDivision = CAR_DIVISION Then
                                        objCompToFetchAttribute = objComp.Parent
                                    End If
                                    'Code added Apr-05-2019
                                    'In case of Daimler, Pradeep wanted the same Final Body name to be output as the Burnout Body Name.
                                    '(Daimler has two different solid body (Final body and Burnout body))

                                    Dim aoAllValidSolidBody() As Body = Nothing
                                    Dim objFinalBody As Body = Nothing

                                    aoAllValidSolidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)
                                    If Not aoAllValidSolidBody Is Nothing Then
                                        'In case of Daimler FnGetValidBodyFOrOEM function will return only one solid body.
                                        objFinalBody = aoAllValidSolidBody(0)
                                    End If
                                    If Not FnCheckIfThisIsAWeldment(objPart.Leaf.ToString) Then
                                        'sBodyName = objbody.JournalIdentifier
                                        If Not objFinalBody Is Nothing Then
                                            sBodyName = objFinalBody.JournalIdentifier
                                        End If
                                        SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColComponentName, objCompToFetchAttribute.DisplayName)
                                        bPopulateInExcel = True
                                    Else
                                        If Not (objComp.FindOccurrence(objbody)) Is Nothing Then
                                            If Not objFinalBody Is Nothing Then
                                                sBodyName = CType(objComp.FindOccurrence(objFinalBody), Body).JournalIdentifier & " " & objCompToFetchAttribute.JournalIdentifier
                                            End If
                                            'sBodyName = CType(objComp.FindOccurrence(objbody), Body).JournalIdentifier & " " & objCompToFetchAttribute.JournalIdentifier
                                            SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColCompDBPartName,
                                                                                    FnGetCompAttribute(objCompToFetchAttribute, "String", _PART_NAME))
                                            SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColComponentName, objCompToFetchAttribute.Parent.DisplayName)
                                            bPopulateInExcel = True
                                        End If
                                    End If
                                ElseIf (sOemName = FIAT_OEM_NAME) Then
                                    objCompToFetchAttribute = objComp
                                    'Check if the component is a child component in weldment
                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        'Fiat component
                                        sBodyName = objbody.JournalIdentifier
                                        'Code modified on Jan-22-2019
                                        'Pradeep wanted the Body Name to be the parent body name itself.
                                        'cODE MODIFIED ON DEC -28-2018
                                        'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                                        'sBodyName = "BO_" & sBodyName
                                        SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColComponentName, objCompToFetchAttribute.DisplayName)
                                        bPopulateInExcel = True
                                    Else
                                        'This is a Fiat Weldment child component
                                        If Not CType(objComp.FindOccurrence(objbody), Body) Is Nothing Then
                                            'when populating body name, give Body journal identifier and Child component journalidentifier
                                            sBodyName = CType(objComp.FindOccurrence(objbody), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                            'Code modified on Jan-22-2019
                                            'Pradeep wanted the Body Name to be the parent body name itself.
                                            'cODE MODIFIED ON DEC -28-2018
                                            'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                                            'sBodyName = "BO_" & sBodyName
                                            SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColCompDBPartName,
                                                                                    FnGetCompAttribute(objCompToFetchAttribute, "String", _PART_NAME))
                                            SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColComponentName, objCompToFetchAttribute.Parent.DisplayName)
                                            bPopulateInExcel = True
                                        End If
                                    End If
                                End If
                                If bPopulateInExcel Then
                                    SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColBodyToolClass, NOTAPPLICABLE)

                                    'Update Value to the cell
                                    SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColBodyName, sBodyName)

                                    SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColBodyLayer, objbody.Layer.ToString)
                                    sShape = FnGetStringUserAttribute(objCompToFetchAttribute, _SHAPE_ATTR_NAME)
                                    'Code added Aug-16-2018
                                    'Shapes are categorized and mapped.
                                    'FLAT,SQUARE,PLATE are mapped as FLAT
                                    'RECT TUBG and SQUARE TUBG are mapped as RECT TUBG
                                    If sShape <> "" Then
                                        If (sShape.ToUpper = PLATE) Or (sShape.ToUpper = SQUARE) Then
                                            sShape = FLAT
                                        End If
                                        If (sShape.ToUpper = SQUARE_TUBG) Then
                                            sShape = RECT_TUBG
                                        End If
                                    End If
                                    SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColBodyShape, sShape)
                                    'Code modified on Apr-03-2019
                                    'WERKSTOFF or material attribute for Daimler car is changed to read at GEO level (Earlier it was read at container level)
                                    If (_sDivision = CAR_DIVISION) Then
                                        sPMat = FnGetStringUserAttribute(objComp, _P_MAT)
                                    Else
                                        sPMat = FnGetStringUserAttribute(objCompToFetchAttribute, _P_MAT)
                                    End If

                                    SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColPMat, sPMat)
                                    sStockSize = FnGetStringUserAttribute(objCompToFetchAttribute, _STOCK_SIZE_METRIC)
                                    SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColStockSize, sStockSize)
                                    iRowStart = iRowStart + 1
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Else
            'Code added Nov-12-2018
            aoAllBurnOutBody = FnGetValidBurnoutBodyForOEM(objPart, sOemName)
            If Not aoAllBurnOutBody Is Nothing Then
                For Each objbody As Body In aoAllBurnOutBody
                    'Check whether the body is a solid body
                    If objbody.IsSolidBody Then ' And
                        sBodyName = objbody.JournalIdentifier
                        'Code modified on Jan-22-2019
                        'Pradeep wanted the Body Name to be the parent body name itself.
                        'cODE MODIFIED ON DEC -28-2018
                        'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                        'sBodyName = "BO_" & sBodyName
                        SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColComponentName, objPart.Leaf.ToString)
                        'If bPopulateInExcel Then
                        SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColBodyToolClass, NOTAPPLICABLE)

                        'Update Value to the cell
                        SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColBodyName, sBodyName)

                        SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColBodyLayer, objbody.Layer.ToString)
                        sShape = FnGetStringUserAttribute(objPart, _SHAPE_ATTR_NAME)
                        'Code added Aug-16-2018
                        'Shapes are categorized and mapped.
                        'FLAT,SQUARE,PLATE are mapped as FLAT
                        'RECT TUBG and SQUARE TUBG are mapped as RECT TUBG
                        If sShape <> "" Then
                            If (sShape.ToUpper = PLATE) Or (sShape.ToUpper = SQUARE) Then
                                sShape = FLAT
                            End If
                            If (sShape.ToUpper = SQUARE_TUBG) Then
                                sShape = RECT_TUBG
                            End If
                        End If
                        SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColBodyShape, sShape)
                        sPMat = FnGetStringUserAttribute(objPart, _P_MAT)
                        SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColPMat, sPMat)
                        sStockSize = FnGetStringUserAttribute(objPart, _STOCK_SIZE_METRIC)
                        SWriteValueToCell(_objWorkBk, BURNOUT_BODYNAMES_SHEET, iRowStart, iColStockSize, sStockSize)
                        iRowStart = iRowStart + 1
                        'End If
                    End If
                Next
            End If
        End If
    End Sub

    'To get the bounding box of a body aligned to a given model view
    Sub sWriteBoundingBoxofBurnOutBodiesInAModelView(ByVal objPart As Part, aoAllCompInSession() As Component, ByVal sModelViewName As String, ByVal sSheetNameToWrite As String, sOemName As String)
        'For computing the exact view bounds of the bodies
        Dim min_corner(2) As Double
        Dim directions(2, 2) As Double
        Dim distances(2) As Double
        Dim adBoundingBox() As Double = Nothing
        Dim objBody As Body = Nothing
        Dim sBodyName As String = ""
        Dim sBodyNameCompare As String = ""
        Dim iRowToWrite As Integer = 2
        Dim iNosOfFilledRows As Integer = 0
        Dim iColBodyName As Integer = 0

        Dim matrixTag As Tag = NXOpen.Tag.Null
        Dim csysTag As Tag = NXOpen.Tag.Null
        Dim objChildPart As Part = Nothing
        Dim bPopulateBBInfo As Boolean = False
        Dim aoAllValidBody() As Body = Nothing

        For Each objModelView As ModelingView In objPart.ModelingViews
            If objModelView.Name.ToUpper = sModelViewName.ToUpper Then
                'Make this view as the work view
                sReplaceViewInLayout(objPart, objModelView)
                Dim admatrixValues As Double() = {objModelView.Matrix.Xx, objModelView.Matrix.Xy, objModelView.Matrix.Xz,
                                                  objModelView.Matrix.Yx, objModelView.Matrix.Yy, objModelView.Matrix.Yz,
                                                  objModelView.Matrix.Zx, objModelView.Matrix.Zy, objModelView.Matrix.Zz}
                'Create the CSYS matrix
                UFSession.GetUFSession().Csys.CreateMatrix(admatrixValues, matrixTag)
                UFSession.GetUFSession.Csys.CreateCsys({0, 0, 0}, matrixTag, csysTag)
                Exit For
            End If
        Next

        iColBodyName = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, BODY_NAME, 1)
        iNosOfFilledRows = FnGetNumberofRows(_objWorkBk, sSheetNameToWrite, 1, 1)

        iColMinPointX = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, MIN_POINTX, 1)
        iColMinPointY = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, MIN_POINTY, 1)
        iColMinPointZ = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, MIN_POINTZ, 1)
        iColVectorXX = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORXX, 1)
        iColVectorXY = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORXY, 1)
        iColVectorXZ = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORXZ, 1)
        iColVectorYX = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORYX, 1)
        iColVectorYY = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORYY, 1)
        iColVectorYZ = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORYZ, 1)
        iColVectorZX = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORZX, 1)
        iColVectorZY = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORZY, 1)
        iColVectorZZ = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, VECTORZZ, 1)
        iColMagX = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, Magnitude_X, 1)
        iColMagY = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, Magnitude_Y, 1)
        iColMagZ = FnGetColumnNumberByName(_objWorkBk, sSheetNameToWrite, Magnitude_Z, 1)

        If Not aoAllCompInSession Is Nothing Then
            For Each objComp As Component In aoAllCompInSession
                objChildPart = FnGetPartFromComponent(objComp)
                If Not objChildPart Is Nothing Then
                    FnLoadPartFully(objChildPart)
                    aoAllValidBody = FnGetValidBurnoutBodyForOEM(objChildPart, sOemName)
                    If Not aoAllValidBody Is Nothing Then
                        For Each body As Body In aoAllValidBody
                            If body.IsSolidBody Then
                                If sOemName = DAIMLER_OEM_NAME Then
                                    If _sDivision = TRUCK_DIVISION Then
                                        If objComp Is objPart.ComponentAssembly.RootComponent Then
                                            'Component in truck
                                            sBodyName = body.JournalIdentifier
                                            objBody = body
                                        Else
                                            'Weldment in truck
                                            objBody = objComp.FindOccurrence(body)
                                            If Not objBody Is Nothing Then
                                                sBodyName = CType(objComp.FindOccurrence(body), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                                'sBodyName = objComp.DisplayName.ToUpper
                                                objBody = CType(objComp.FindOccurrence(body), Body)
                                            End If
                                        End If
                                    ElseIf _sDivision = CAR_DIVISION Then
                                        'Code added Apr-05-2019
                                        'In case of Daimler, Pradeep wanted the same Final Body name to be output as the Burnout Body Name.
                                        '(Daimler has two different solid body (Final body and Burnout body))

                                        Dim aoAllValidSolidBody() As Body = Nothing
                                        Dim objFinalBody As Body = Nothing

                                        aoAllValidSolidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)
                                        If Not aoAllValidSolidBody Is Nothing Then
                                            'In case of Daimler FnGetValidBodyFOrOEM function will return only one solid body.
                                            objFinalBody = aoAllValidSolidBody(0)
                                        End If

                                        'Check if the component is a child component in weldment
                                        If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                            objBody = objComp.FindOccurrence(body)
                                            If Not objBody Is Nothing Then
                                                'Weldment in car. Get the occurrence body
                                                'sBodyName = CType(objBody, Body).JournalIdentifier & " " & objComp.Parent.JournalIdentifier
                                                objBody = CType(objComp.FindOccurrence(body), Body)
                                            End If
                                            If Not objFinalBody Is Nothing Then
                                                sBodyName = CType(objFinalBody, Body).JournalIdentifier & " " & objComp.Parent.JournalIdentifier
                                            End If
                                        Else
                                            'Component in car. Get the prototype body
                                            objBody = objComp.FindOccurrence(body)
                                            If Not objBody Is Nothing Then
                                                'sBodyName = body.JournalIdentifier
                                                objBody = CType(objComp.FindOccurrence(body), Body)
                                            End If
                                            If Not objFinalBody Is Nothing Then
                                                sBodyName = objFinalBody.JournalIdentifier
                                            End If
                                        End If
                                    End If
                                ElseIf (sOemName = FIAT_OEM_NAME) Then
                                    'Check if the component is a child component in weldment
                                    If objComp Is objPart.ComponentAssembly.RootComponent Then
                                        'Fiat component
                                        objBody = body
                                        sBodyName = objBody.JournalIdentifier
                                        'Code modified on Jan-22-2019
                                        'Pradeep wanted the Body Name to be the parent body name itself.
                                        'cODE MODIFIED ON DEC -28-2018
                                        'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                                        'sBodyName = "BO_" & sBodyName
                                    Else
                                        'This is a Fiat Weldment child component
                                        objBody = CType(objComp.FindOccurrence(body), Body)
                                        If Not objBody Is Nothing Then
                                            'when populating body name, give Body journal identifier and Child component journalidentifier
                                            sBodyName = objBody.JournalIdentifier & " " & objComp.JournalIdentifier
                                            'Code modified on Jan-22-2019
                                            'Pradeep wanted the Body Name to be the parent body name itself.
                                            'cODE MODIFIED ON DEC -28-2018
                                            'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                                            'sBodyName = "BO_" & sBodyName
                                        End If
                                    End If
                                End If
                                If Not objBody Is Nothing Then
                                    FnGetUFSession.Modl.AskBoundingBoxExact(objBody.Tag, csysTag, min_corner, directions, distances)
                                    'Find the row to update
                                    For iloopIndex = 2 To iNosOfFilledRows
                                        bPopulateBBInfo = False
                                        'Fetch the Body Name number already assigned
                                        sBodyNameCompare = FnReadSingleRowForColumn(_objWorkBk, sSheetNameToWrite, iColBodyName, iloopIndex)
                                        If sBodyNameCompare.ToUpper = sBodyName.ToUpper Then
                                            iRowToWrite = iloopIndex
                                            bPopulateBBInfo = True
                                            Exit For
                                        End If
                                    Next
                                    If bPopulateBBInfo Then
                                        'To get the bounding box exact for each body
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointX, min_corner(0).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointY, min_corner(1).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointZ, min_corner(2).ToString)

                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXX, directions(0, 0).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXY, directions(0, 1).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXZ, directions(0, 2).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagX, distances(0).ToString)

                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYX, directions(1, 0).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYY, directions(1, 1).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYZ, directions(1, 2).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagY, distances(1).ToString)

                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZX, directions(2, 0).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZY, directions(2, 1).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZZ, directions(2, 2).ToString)
                                        SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagZ, distances(2).ToString)
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Else
            aoAllValidBody = FnGetValidBurnoutBodyForOEM(objPart, sOemName)
            If Not aoAllValidBody Is Nothing Then
                For Each body As Body In aoAllValidBody
                    'Check whether the body is a solid body
                    If body.IsSolidBody Then
                        objBody = body
                        sBodyName = body.JournalIdentifier
                        'Code modified on Jan-22-2019
                        'Pradeep wanted the Body Name to be the parent body name itself.
                        'cODE MODIFIED ON DEC -28-2018
                        'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                        'sBodyName = "BO_" & sBodyName
                        FnGetUFSession.Modl.AskBoundingBoxExact(objBody.Tag, csysTag, min_corner, directions, distances)

                        'Find the row to update
                        For iloopIndex = 2 To iNosOfFilledRows
                            bPopulateBBInfo = False
                            'Fetch the Body Name number already assigned
                            sBodyNameCompare = FnReadSingleRowForColumn(_objWorkBk, sSheetNameToWrite, iColBodyName, iloopIndex)
                            If sBodyNameCompare.ToUpper = sBodyName.ToUpper Then
                                iRowToWrite = iloopIndex
                                bPopulateBBInfo = True
                                Exit For
                            End If
                        Next
                        If bPopulateBBInfo Then
                            'To get the bounding box exact for each body
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointX, min_corner(0).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointY, min_corner(1).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMinPointZ, min_corner(2).ToString)

                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXX, directions(0, 0).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXY, directions(0, 1).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorXZ, directions(0, 2).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagX, distances(0).ToString)

                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYX, directions(1, 0).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYY, directions(1, 1).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorYZ, directions(1, 2).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagY, distances(1).ToString)

                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZX, directions(2, 0).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZY, directions(2, 1).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColVectorZZ, directions(2, 2).ToString)
                            SWriteValueToCell(_objWorkBk, sSheetNameToWrite, iRowToWrite, iColMagZ, distances(2).ToString)
                        End If
                    End If

                Next
            End If
        End If

        'Delete the created CSYS
        SDeleteObjects({NXObjectManager.Get(csysTag)})

    End Sub

    'Compute B_Burnout_LCS orientation for BurnOutBody
    '(Logic given by Pradeep)
    '* Identify Fillet curves on the Burnout body and save axis direction cosines of each of them
    '* Among them, compute the direction along which most fillet curves are aligned (Dir 1)
    '* From the Optimal LCS rotation matrix for the same Body, identify a direction (Dir 2) which is orthogonal to fillet curve direction 
    '* Burnout View LCS can be arranged as follows:
    '  (Dir 2, Dir 1 x Dir 2, Dir 1)

    Sub sComputeLCSForBurnOutBody(ByVal objPart As Part, aoAllCOmpInSession() As Component, sOemName As String)

        Dim objChildPart As Part = Nothing
        Dim objCompToFetchAttribute As Component = Nothing
        Dim aoNonHoleCylFace() As Face = Nothing
        Dim adDirVector1() As Double = Nothing
        Dim adDirVector2() As Double = Nothing
        Dim adCrossPrdtVector() As Double = Nothing
        Dim adBurnOutBodyLCS(8) As Double
        Dim sBodyname As String = ""
        Dim iLCSRowStart As Integer = 0
        Dim iIndex As Integer = 1
        Dim adOptimalLCS() As Double = Nothing
        Dim objBodyToAnalyze As Body = Nothing
        Dim aoAllBurnOutBody() As Body = Nothing

        iLCSRowStart = 2

        If Not aoAllCOmpInSession Is Nothing Then
            For Each objComp As Component In aoAllCOmpInSession
                objChildPart = FnGetPartFromComponent(objComp)
                If Not objChildPart Is Nothing Then
                    FnLoadPartFully(objChildPart)
                    'bURNOUT BODY WILL HAVE FLAT AS THE SHAPE ATTRIBUTE
                    If FnGetStringUserAttribute(objChildPart, _SHAPE_ATTR_NAME) <> ROUND_SHAPE Then
                        aoAllBurnOutBody = FnGetValidBurnoutBodyForOEM(objChildPart, sOemName)
                        If Not aoAllBurnOutBody Is Nothing Then
                            For Each objbody As Body In aoAllBurnOutBody
                                objBodyToAnalyze = Nothing
                                If objbody.IsSolidBody Then
                                    If (sOemName = DAIMLER_OEM_NAME) Then
                                        If _sDivision = TRUCK_DIVISION Then
                                            If objComp Is objPart.ComponentAssembly.RootComponent Then
                                                'Component in truck
                                                sBodyname = objbody.JournalIdentifier
                                                objBodyToAnalyze = objbody
                                            Else
                                                'Weldment in truck
                                                objBodyToAnalyze = objComp.FindOccurrence(objbody)
                                                If Not objBodyToAnalyze Is Nothing Then
                                                    sBodyname = CType(objComp.FindOccurrence(objbody), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                                    'sBodyName = objComp.DisplayName.ToUpper
                                                    objBodyToAnalyze = CType(objComp.FindOccurrence(objbody), Body)
                                                End If
                                            End If
                                        ElseIf _sDivision = CAR_DIVISION Then
                                            'Code added Apr-05-2019
                                            'In case of Daimler, Pradeep wanted the same Final Body name to be output as the Burnout Body Name.
                                            '(Daimler has two different solid body (Final body and Burnout body))

                                            Dim aoAllValidSolidBody() As Body = Nothing
                                            Dim objFinalBody As Body = Nothing

                                            aoAllValidSolidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)
                                            If Not aoAllValidSolidBody Is Nothing Then
                                                'In case of Daimler FnGetValidBodyFOrOEM function will return only one solid body.
                                                objFinalBody = aoAllValidSolidBody(0)
                                            End If

                                            'Check if the component is a child component in weldment
                                            If FnCheckIfThisIsAChildCompInWeldment(objComp, _sOemName) Then
                                                'Weldment in car. Get the occurrence body
                                                objBodyToAnalyze = objComp.FindOccurrence(objbody)
                                                If Not objBodyToAnalyze Is Nothing Then
                                                    'sBodyname = CType(objComp.FindOccurrence(objbody), Body).JournalIdentifier & " " & objComp.Parent.JournalIdentifier
                                                    objBodyToAnalyze = CType(objComp.FindOccurrence(objbody), Body)
                                                End If
                                                If Not objFinalBody Is Nothing Then
                                                    sBodyname = CType(objComp.FindOccurrence(objFinalBody), Body).JournalIdentifier & " " & objComp.Parent.JournalIdentifier
                                                End If
                                            Else
                                                'Component in car. Get the prototype body
                                                'sBodyname = objbody.JournalIdentifier
                                                objBodyToAnalyze = objbody
                                                If Not objFinalBody Is Nothing Then
                                                    sBodyname = objFinalBody.JournalIdentifier
                                                End If
                                            End If
                                        End If
                                    ElseIf (sOemName = FIAT_OEM_NAME) Then
                                        If objComp Is objPart.ComponentAssembly.RootComponent Then
                                            'Component in Fiat
                                            sBodyname = objbody.JournalIdentifier
                                            objBodyToAnalyze = objbody
                                            'Code modified on Jan-22-2019
                                            'Pradeep wanted the Body Name to be the parent body name itself.
                                            'cODE MODIFIED ON DEC -28-2018
                                            'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                                            'sBodyname = "BO_" & sBodyname
                                        Else
                                            'Weldment in Fiat
                                            objBodyToAnalyze = objComp.FindOccurrence(objbody)
                                            If Not objBodyToAnalyze Is Nothing Then
                                                sBodyname = CType(objComp.FindOccurrence(objbody), Body).JournalIdentifier & " " & objComp.JournalIdentifier
                                                'sBodyName = objComp.DisplayName.ToUpper
                                                objBodyToAnalyze = CType(objComp.FindOccurrence(objbody), Body)
                                                'Code modified on Jan-22-2019
                                                'Pradeep wanted the Body Name to be the parent body name itself.
                                                'cODE MODIFIED ON DEC -28-2018
                                                'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                                                'sBodyname = "BO_" & sBodyname
                                            End If
                                        End If
                                    End If

                                    If Not objBodyToAnalyze Is Nothing Then
                                        'Collect Non-Hole Cylindrical Face from Burnout Body
                                        aoNonHoleCylFace = FnGetColOfNonHoleCylFaces(objChildPart, objBodyToAnalyze)
                                        If Not aoNonHoleCylFace Is Nothing Then
                                            'Among them, compute the direction along which most fillet curves are aligned (Dir 1)
                                            'Get the most number of identical vector
                                            adDirVector1 = FnGetIdenticalVectorFromCylFaces(objPart, aoNonHoleCylFace)
                                            If Not adDirVector1 Is Nothing Then
                                                adOptimalLCS = FnComputeOptimalRotationMatrixAlt(objBodyToAnalyze)
                                                If Not adOptimalLCS Is Nothing Then
                                                    'Get the Dir(2) which is any orthogonal direction from LCS computation to the given dir(1)
                                                    adDirVector2 = FnGetOrthogonalVectorToRefVectorFromOptimalRotMatrix(adDirVector1, adOptimalLCS)
                                                    'adDirVector2 = FnGetOrthogonalVectorToRefVectorInGivenSolidBody(objbodytoAnalyze, adDirVector1)
                                                    If Not adDirVector2 Is Nothing Then
                                                        adCrossPrdtVector = FnGetCrossProductVector(adDirVector1, adDirVector2)
                                                        If Not adCrossPrdtVector Is Nothing Then
                                                            adBurnOutBodyLCS(0) = adDirVector2(0)
                                                            adBurnOutBodyLCS(1) = adDirVector2(1)
                                                            adBurnOutBodyLCS(2) = adDirVector2(2)
                                                            adBurnOutBodyLCS(3) = adCrossPrdtVector(0)
                                                            adBurnOutBodyLCS(4) = adCrossPrdtVector(1)
                                                            adBurnOutBodyLCS(5) = adCrossPrdtVector(2)
                                                            adBurnOutBodyLCS(6) = adDirVector1(0)
                                                            adBurnOutBodyLCS(7) = adDirVector1(1)
                                                            adBurnOutBodyLCS(8) = adDirVector1(2)

                                                            sPopulateBodyLCSToExcelFile(sBodyname, adBurnOutBodyLCS, iLCSRowStart, "BO_Body_LCS")
                                                            iLCSRowStart = iLCSRowStart + 1
                                                            sCreateCustomModelViewForOptimalRotMat(objPart, adBurnOutBodyLCS, "BO_LCS_" & sBodyname)
                                                            'iIndex = iIndex + 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                'Clean memory
                                _aoStructPartOrientationInfo = Nothing
                                _iPartOrientationIndex = 0
                                _aoStructBodyOrientationInfo = Nothing
                                _iOrientationIndex = 0
                            Next
                        End If
                    End If
                End If
            Next
        Else
            'Code added on Nov-13-2018
            aoAllBurnOutBody = FnGetValidBurnoutBodyForOEM(objPart, sOemName)
            If Not aoAllBurnOutBody Is Nothing Then
                For Each body As Body In aoAllBurnOutBody
                    If FnGetStringUserAttribute(body, _SHAPE_ATTR_NAME) <> ROUND_SHAPE Then
                        'Check whether the body is a solid body
                        If body.IsSolidBody Then
                            sBodyname = body.JournalIdentifier
                            'Code modified on Jan-22-2019
                            'Pradeep wanted the Body Name to be the parent body name itself.
                            'cODE MODIFIED ON DEC -28-2018
                            'It is pradeep's requirement. Burnout body should get the similar names of final body, with prefix Bo
                            'sBodyname = "BO_" & sBodyname
                            'Collect Non-Hole Cylindrical Face from Burnout Body
                            aoNonHoleCylFace = FnGetColOfNonHoleCylFaces(objPart, body)
                            If Not aoNonHoleCylFace Is Nothing Then
                                'Among them, compute the direction along which most fillet curves are aligned (Dir 1)
                                'Get the most number of identical vector
                                adDirVector1 = FnGetIdenticalVectorFromCylFaces(objPart, aoNonHoleCylFace)
                                If Not adDirVector1 Is Nothing Then
                                    adOptimalLCS = FnComputeOptimalRotationMatrixAlt(body)
                                    If Not adOptimalLCS Is Nothing Then
                                        'Get the Dir(2) which is any orthogonal direction from LCS computation to the given dir(1)
                                        adDirVector2 = FnGetOrthogonalVectorToRefVectorFromOptimalRotMatrix(adDirVector1, adOptimalLCS)
                                        'adDirVector2 = FnGetOrthogonalVectorToRefVectorInGivenSolidBody(body, adDirVector1)
                                        If Not adDirVector2 Is Nothing Then
                                            adCrossPrdtVector = FnGetCrossProductVector(adDirVector1, adDirVector2)
                                            If Not adCrossPrdtVector Is Nothing Then
                                                adBurnOutBodyLCS(0) = adDirVector2(0)
                                                adBurnOutBodyLCS(1) = adDirVector2(1)
                                                adBurnOutBodyLCS(2) = adDirVector2(2)
                                                adBurnOutBodyLCS(3) = adCrossPrdtVector(0)
                                                adBurnOutBodyLCS(4) = adCrossPrdtVector(1)
                                                adBurnOutBodyLCS(5) = adCrossPrdtVector(2)
                                                adBurnOutBodyLCS(6) = adDirVector1(0)
                                                adBurnOutBodyLCS(7) = adDirVector1(1)
                                                adBurnOutBodyLCS(8) = adDirVector1(2)

                                                sPopulateBodyLCSToExcelFile(sBodyname, adBurnOutBodyLCS, iLCSRowStart, "BO_Body_LCS")
                                                iLCSRowStart = iLCSRowStart + 1
                                                sCreateCustomModelViewForOptimalRotMat(objPart, adBurnOutBodyLCS, "BO_LCS_" & sBodyname)
                                                'iIndex = iIndex + 1
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                        End If
                        'Clean memory
                        _aoStructPartOrientationInfo = Nothing
                        _iPartOrientationIndex = 0
                        _aoStructBodyOrientationInfo = Nothing
                        _iOrientationIndex = 0
                    End If
                Next
            End If
        End If
    End Sub
    'Function to get the orthogonal vector to the given vector in a given solid body
    Function FnGetOrthogonalVectorToRefVectorInGivenSolidBody(objBody As Body, adVector1() As Double) As Double()
        Dim adVectorToCompare() As Double = Nothing

        If Not objBody Is Nothing Then
            For Each objFace As Face In objBody.GetFaces()
                If objFace.SolidFaceType = Face.FaceType.Planar Then
                    adVectorToCompare = FnGetFaceNormalVector(objFace)
                    If Not adVectorToCompare Is Nothing Then
                        If FnCheckOrthogalityOfTwoVectors(adVector1, adVectorToCompare) Then
                            FnGetOrthogonalVectorToRefVectorInGivenSolidBody = adVectorToCompare
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
        FnGetOrthogonalVectorToRefVectorInGivenSolidBody = Nothing
    End Function
    'Function to get the orthogonal vector to the ref vector from a given rotation matrix
    Function FnGetOrthogonalVectorToRefVectorFromOptimalRotMatrix(adRefVector() As Double, adOptimalLCS() As Double) As Double()
        Dim adVector1ToCOmpare(2) As Double
        Dim adVector2ToCOmpare(2) As Double
        Dim adVector3ToCOmpare(2) As Double
        If Not adOptimalLCS Is Nothing Then
            adVector1ToCOmpare(0) = adOptimalLCS(0)
            adVector1ToCOmpare(1) = adOptimalLCS(1)
            adVector1ToCOmpare(2) = adOptimalLCS(2)

            adVector2ToCOmpare(0) = adOptimalLCS(3)
            adVector2ToCOmpare(1) = adOptimalLCS(4)
            adVector2ToCOmpare(2) = adOptimalLCS(5)

            adVector3ToCOmpare(0) = adOptimalLCS(6)
            adVector3ToCOmpare(1) = adOptimalLCS(7)
            adVector3ToCOmpare(2) = adOptimalLCS(8)

            If FnCheckOrthogalityOfTwoVectors(adRefVector, adVector1ToCOmpare) Then
                FnGetOrthogonalVectorToRefVectorFromOptimalRotMatrix = adVector1ToCOmpare
                Exit Function
            End If
            If FnCheckOrthogalityOfTwoVectors(adRefVector, adVector2ToCOmpare) Then
                FnGetOrthogonalVectorToRefVectorFromOptimalRotMatrix = adVector2ToCOmpare
                Exit Function
            End If
            If FnCheckOrthogalityOfTwoVectors(adRefVector, adVector3ToCOmpare) Then
                FnGetOrthogonalVectorToRefVectorFromOptimalRotMatrix = adVector3ToCOmpare
                Exit Function
            End If
        End If
        FnGetOrthogonalVectorToRefVectorFromOptimalRotMatrix = Nothing
    End Function


    'Function to find the most common vector(direction co-sines) from colloection of non-holecylindrical face
    Function FnGetIdenticalVectorFromCylFaces(objPart As Part, aoAllCylFace() As Face) As Double()
        Dim adVector() As Double = Nothing
        Dim dictColOfVectors As Dictionary(Of Face, Double()) = Nothing
        Dim advector1ToCompare() As Double = Nothing
        Dim adVector2ToCOmpare() As Double = Nothing
        Dim iCount As Integer = 0
        Dim dictColOfFaceAndNumOfIdenticalVector As Dictionary(Of Face, Integer) = Nothing
        Dim iMaxNumOfIdenticalVector As Integer = 0
        Dim objCylFaceWithMaxIdenticalVector As Face = Nothing

        If Not aoAllCylFace Is Nothing Then
            dictColOfVectors = New Dictionary(Of Face, Double())
            For Each objCylFace As Face In aoAllCylFace
                adVector = FnGetAxisVecCylFace(objPart, objCylFace)
                dictColOfVectors.Add(objCylFace, adVector)
            Next

            If Not dictColOfVectors Is Nothing Then
                If dictColOfVectors.Count > 0 Then
                    dictColOfFaceAndNumOfIdenticalVector = New Dictionary(Of Face, Integer)
                    For Each key In dictColOfVectors.Keys
                        iCount = 0
                        advector1ToCompare = dictColOfVectors(key)
                        For Each keyToCompare In dictColOfVectors.Keys
                            If Not key Is keyToCompare Then
                                adVector2ToCOmpare = dictColOfVectors(keyToCompare)

                                If FnParallelAntiParallelCheck(advector1ToCompare, adVector2ToCOmpare) Then
                                    iCount = iCount + 1
                                End If
                            End If
                        Next
                        dictColOfFaceAndNumOfIdenticalVector.Add(key, iCount)
                    Next

                    If Not dictColOfFaceAndNumOfIdenticalVector Is Nothing Then
                        If dictColOfFaceAndNumOfIdenticalVector.Count > 0 Then
                            For Each key In dictColOfFaceAndNumOfIdenticalVector.Keys
                                If dictColOfFaceAndNumOfIdenticalVector(key) > iMaxNumOfIdenticalVector Then
                                    iMaxNumOfIdenticalVector = dictColOfFaceAndNumOfIdenticalVector(key)
                                    objCylFaceWithMaxIdenticalVector = key
                                End If
                            Next
                        End If
                    End If
                End If
            End If

            If Not objCylFaceWithMaxIdenticalVector Is Nothing Then
                FnGetIdenticalVectorFromCylFaces = FnGetAxisVecCylFace(objPart, objCylFaceWithMaxIdenticalVector)
                Exit Function
            End If
        End If
        FnGetIdenticalVectorFromCylFaces = Nothing
    End Function
    Sub sPopulateBurnOutBodyInfo(objPart As Part, sOemName As String)
        Dim aoAllCompInSession() As Component = Nothing

        'Code added Nov-13-2018
        'In case of FIAT, there is no seperate body as a Burnout Body. It is a single body.
        'So, to get the Butnout body in FIAT, suppress all the features except, CORPO featuregroup.
        'Body which is present within the CORPO feature group is the Burnout body.
        If (sOemName = FIAT_OEM_NAME) Then
            sAddOrRemoveFeatureGroupTemporarily(objPart, bIsSuppress:=True)
        End If

        aoAllCompInSession = FnGetAllComponentsInSession()

        'If Not aoAllCompInSession Is Nothing Then
        'Populate Burnout Body Face and Edge Informaiton
        sPopulateBurnOutBodyFaceAndEdgeInfo(objPart, aoAllCompInSession, sOemName)
        'Populate BurnOut Body Names Sheet Information
        sPopulateBurnoutBodyDataInBodyNameSheet(objPart, aoAllCompInSession, sOemName)

        sWriteBoundingBoxofBurnOutBodiesInAModelView(objPart, aoAllCompInSession, PART_LCS_VIEW_NAME, BURNOUT_BODYNAMES_SHEET, sOemName)

        sComputeLCSForBurnOutBody(objPart, aoAllCompInSession, sOemName)
        'End If

        If (sOemName = FIAT_OEM_NAME) Then
            sAddOrRemoveFeatureGroupTemporarily(objPart, bIsSuppress:=False)
        End If
    End Sub
    'Code added May-10-2018
    'Collect all machined face with holes on it from a given solid body
    Function FnCollectMachinedFaceWithHoles(objBody As Body) As Face()
        Dim aoAllMachinedFace() As Face = Nothing
        If Not objBody Is Nothing Then
            For Each objFace As Face In objBody.GetFaces()
                If FnGetStringUserAttribute(objFace, _FINISHTOLERANCE_ATTR_NAME) = _FINISH_TOL_VALUE2 Then
                    'Check if the machined face has holes on it
                    If FnCheckIfthePlanarFaceHasHolesOnIt(objFace) Then
                        'Collect only the machined face which has holes on it.
                        If aoAllMachinedFace Is Nothing Then
                            ReDim Preserve aoAllMachinedFace(0)
                            aoAllMachinedFace(0) = objFace
                        Else
                            ReDim Preserve aoAllMachinedFace(UBound(aoAllMachinedFace) + 1)
                            aoAllMachinedFace(UBound(aoAllMachinedFace)) = objFace
                        End If
                    End If
                End If
            Next
        End If
        FnCollectMachinedFaceWithHoles = aoAllMachinedFace
    End Function

    'To check if the face has a hole
    Public Function FnCheckIfthePlanarFaceHasHolesOnIt(ByVal ObjPlanarFace As Face) As Boolean
        Dim objEdgeVrt1 As Point3d = Nothing
        Dim objEdgeVrt2 As Point3d = Nothing
        Dim objPart As Part = Nothing
        If Not ObjPlanarFace Is Nothing Then
            objPart = ObjPlanarFace.OwningPart
            If Not objPart Is Nothing Then
                For Each objEdge As Edge In ObjPlanarFace.GetEdges()
                    If objEdge.SolidEdgeType = Edge.EdgeType.Circular Then
                        objEdge.GetVertices(objEdgeVrt1, objEdgeVrt2)
                        If objEdgeVrt1.Equals(objEdgeVrt2) Then
                            If Not FnIsEdgeConvex(objPart, objEdge) Then
                                FnCheckIfthePlanarFaceHasHolesOnIt = True
                                Exit Function
                            End If
                            'For Plasma cut holes will have a slit for plasma machine to enter
                            'The included hole angle should be more than 270 degrees
                        Else
                            If ((FnGetArcInfo(objEdge.Tag).limits(1) - FnGetArcInfo(objEdge.Tag).limits(0)) * 180) / PI > 270 Then
                                If Not FnIsEdgeConvex(objPart, objEdge) Then
                                    FnCheckIfthePlanarFaceHasHolesOnIt = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
        FnCheckIfthePlanarFaceHasHolesOnIt = False
    End Function

    'Filter the face which has holes on it
    Function FnCollectFaceWhichHasHolesOnIt(aoAllFace() As Face) As Face()
        Dim aoAllFaceWithHolesOnIt() As Face = Nothing

        If Not aoAllFace Is Nothing Then
            For Each objface As Face In aoAllFace
                If FnCheckIfthePlanarFaceHasHolesOnIt(objface) Then
                    If aoAllFaceWithHolesOnIt Is Nothing Then
                        ReDim Preserve aoAllFaceWithHolesOnIt(0)
                        aoAllFaceWithHolesOnIt(0) = objface
                    Else
                        ReDim Preserve aoAllFaceWithHolesOnIt(UBound(aoAllFaceWithHolesOnIt) + 1)
                        aoAllFaceWithHolesOnIt(UBound(aoAllFaceWithHolesOnIt)) = objface
                    End If
                End If
            Next
        End If
        FnCollectFaceWhichHasHolesOnIt = aoAllFaceWithHolesOnIt
    End Function

    'Code created on Jun-01-2018
    'Function to get the valid solid body based on OEM
    Function FnGetValidBodyForOEM(objPart As Part, sOemName As String) As Body()
        Dim objValidBody As Body = Nothing
        Dim aoAllValidBody() As Body = Nothing
        Dim aoFeature() As Feature = Nothing
        If sOemName <> "" Then
            If sOemName = GM_OEM_NAME Then
                If Not objPart Is Nothing Then
                    For Each objBody As Body In objPart.Bodies()
                        If objBody.IsSolidBody Then
                            If Not objBody.IsBlanked Then
                                If aoAllValidBody Is Nothing Then
                                    ReDim Preserve aoAllValidBody(0)
                                    aoAllValidBody(0) = objBody
                                Else
                                    ReDim Preserve aoAllValidBody(UBound(aoAllValidBody) + 1)
                                    aoAllValidBody(UBound(aoAllValidBody)) = objBody
                                End If
                            End If
                        End If
                    Next
                End If

            ElseIf sOemName = CHRYSLER_OEM_NAME Then
                If Not objPart Is Nothing Then
                    For Each objBody As Body In objPart.Bodies()
                        If objBody.IsSolidBody Then
                            If Not objBody.IsBlanked Then
                                If aoAllValidBody Is Nothing Then
                                    ReDim Preserve aoAllValidBody(0)
                                    aoAllValidBody(0) = objBody
                                Else
                                    ReDim Preserve aoAllValidBody(UBound(aoAllValidBody) + 1)
                                    aoAllValidBody(UBound(aoAllValidBody)) = objBody
                                End If
                            End If
                        End If
                    Next
                End If
            ElseIf sOemName = DAIMLER_OEM_NAME Then
                If Not objPart Is Nothing Then
                    aoFeature = FnCollectAllMembersOfFeatureGroup(objPart, _sFeatureGroupName)
                    If Not aoFeature Is Nothing Then
                        For i = 0 To aoFeature.Length - 1
                            Try
                                Dim bdtag As Tag = NXOpen.Tag.Null
                                FnGetUFSession.Modl.AskFeatBody(aoFeature(i).Tag, bdtag)
                                objValidBody = CType(NXObjectManager.Get(bdtag), Body)
                                If objValidBody.Layer = 1 And objValidBody.IsSolidBody Then
                                    If Not objValidBody.IsBlanked Then
                                        FnGetValidBodyForOEM = {objValidBody}
                                        Exit Function
                                    End If
                                End If
                            Catch ex As Exception

                            End Try
                        Next
                    End If
                End If
            ElseIf sOemName = FIAT_OEM_NAME Then
                If Not objPart Is Nothing Then
                    For Each objBody As Body In objPart.Bodies()
                        If objBody.IsSolidBody Then
                            If Not objBody.IsBlanked Then
                                'To avoid Intersection Curves
                                If objBody.Layer = 1 Then
                                    If aoAllValidBody Is Nothing Then
                                        ReDim Preserve aoAllValidBody(0)
                                        aoAllValidBody(0) = objBody
                                    Else
                                        ReDim Preserve aoAllValidBody(UBound(aoAllValidBody) + 1)
                                        aoAllValidBody(UBound(aoAllValidBody)) = objBody
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            ElseIf sOemName = GESTAMP_OEM_NAME Then
                If Not objPart Is Nothing Then
                    For Each objBody As Body In objPart.Bodies()
                        If objBody.IsSolidBody Then
                            If Not objBody.IsBlanked Then
                                If aoAllValidBody Is Nothing Then
                                    ReDim Preserve aoAllValidBody(0)
                                    aoAllValidBody(0) = objBody
                                Else
                                    ReDim Preserve aoAllValidBody(UBound(aoAllValidBody) + 1)
                                    aoAllValidBody(UBound(aoAllValidBody)) = objBody
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If
        FnGetValidBodyForOEM = aoAllValidBody
    End Function

    'Code commented on July-05-2019
    'This function is not needed and the attributes are read from XML sheet
    ''Based on the OEM Name decide the Attribute Name
    'Sub sGetAttributeNamesBasedOnOEM(sOEMName As String)

    '    If sOEMName = GM_OEM_NAME Then
    '        _PART_NAME = "DB_PART_NAME"
    '        _SHAPE_ATTR_NAME = "SHAPE"
    '        _SUB_DETAIL_NUMBER = "SUB_DET_NUM"
    '        _P_MASS = "P_MASS"
    '        _PURCH_OPTION = "PURCH_OPTION"
    '        _STOCK_SIZE_METRIC = "STOCK_SIZE_METRIC"
    '        _STOCK_SIZE = "STOCK_SIZE"
    '        _FINISHTOLERANCE_ATTR_NAME = "FINISH_TOL"
    '        _QTY = "QTY"
    '        _P_MAT = "P_MAT"
    '        _RELIEF_CUT_FACE_TOLERANCE_VALUE = "6.3 MICRONS [V]"
    '        _FINISH_TOL_VALUE2 = "3.2 MICRONS [VV]"
    '        _BOM_ATTR = "BOM"
    '        _CLIENT_PART_NAME = "DB_PART_NAME"
    '        _TOOL_CLASS = "TOOL_CLASS"
    '        _TOOL_ID = "TOOL_ID"
    '    ElseIf sOEMName = DAIMLER_OEM_NAME Then
    '        _PART_NAME = "DETAIL_NAME"
    '        _SHAPE_ATTR_NAME = "DETAIL_SHAPE"
    '        _SUB_DETAIL_NUMBER = "SUB_DETAIL_NUMBER"
    '        _P_MASS = "MASS"
    '        _PURCH_OPTION = "PURCHASE_OPTION"
    '        _STOCK_SIZE_METRIC = "STOCK_SIZE_METRIC"
    '        _STOCK_SIZE = "STOCK_SIZE"
    '        _FINISHTOLERANCE_ATTR_NAME = "FINISH_TOL"
    '        _QTY = "QUANTITY"
    '        _RELIEF_CUT_FACE_TOLERANCE_VALUE = "6.4 MICRONS [V]"
    '        _FINISH_TOL_VALUE2 = "3.2 MICRONS [VV]"
    '        _TOOL_CLASS = "TOOL_CLASS"
    '        _TOOL_ID = "TOOL_ID"
    '        If _sDivision = TRUCK_DIVISION Then
    '            _P_MAT = "Lieferant/Werkstoff_W060"
    '            _CLIENT_STOCK_SIZE_ATTR = "Fertigmaße/Bestellbezeichnung_W060"
    '            _BOM_ATTR = "Stuecklistenrelevant_W060"
    '            _CLIENT_PART_NAME = "Benennung_W060"
    '        ElseIf _sDivision = CAR_DIVISION Then
    '            _P_MAT = "Werkstoff"
    '            _CLIENT_STOCK_SIZE_ATTR = "Rohmass"
    '            _BOM_ATTR = ""
    '            _CLIENT_PART_NAME = "NOMENCLATURE_EN"
    '        End If
    '    ElseIf sOEMName = CHRYSLER_OEM_NAME Then
    '        _PART_NAME = "DETAIL_NAME"
    '        _SHAPE_ATTR_NAME = "DETAIL_SHAPE"
    '        _SUB_DETAIL_NUMBER = "SUB_DETAIL_NUMBER"
    '        _P_MASS = "MASS"
    '        _PURCH_OPTION = "PURCHASE_OPTION"
    '        _STOCK_SIZE_METRIC = "STOCK_SIZE_METRIC"
    '        _STOCK_SIZE = "STOCK_SIZE"
    '        _FINISHTOLERANCE_ATTR_NAME = "FINISH_TOLERANCE"
    '        _QTY = "QUANTITY"
    '        _P_MAT = "DB_MATERIAL"
    '        _RELIEF_CUT_FACE_TOLERANCE_VALUE = "6 MICRONS [V]"
    '        _FINISH_TOL_VALUE2 = "3 MICRONS [VV]"
    '        _BOM_ATTR = "BOM"
    '        _CLIENT_PART_NAME = "DB_PART_NAME"
    '        _TOOL_CLASS = "DB_SOURCE"
    '        _TOOL_ID = "DB_NAAMS"
    '    ElseIf sOEMName = FIAT_OEM_NAME Then
    '        _PART_NAME = "DETAIL_NAME"
    '        _SHAPE_ATTR_NAME = "DETAIL_SHAPE"
    '        _SUB_DETAIL_NUMBER = "SUB_DETAIL_NUMBER"
    '        _P_MASS = "MASS"
    '        _PURCH_OPTION = ""
    '        _STOCK_SIZE_METRIC = "STOCK_SIZE"
    '        _STOCK_SIZE = "STOCK_SIZE"
    '        _FINISHTOLERANCE_ATTR_NAME = "FINISH_TOLERANCE"
    '        _QTY = "QUANTITY"
    '        _P_MAT = "CM_Material"
    '        _RELIEF_CUT_FACE_TOLERANCE_VALUE = "6.3 MICRONS [V]"
    '        _FINISH_TOL_VALUE2 = "3.2 MICRONS [VV]"
    '        _BOM_ATTR = "BOM"
    '        _CLIENT_PART_NAME = "CM_Comau_Description_EN"
    '        _TOOL_CLASS = ""
    '        _TOOL_ID = ""
    '        _CLIENT_STOCK_SIZE_ATTR = "CM_Stock_Size"
    '    End If

    'End Sub
    'Code added Jun-28-2019
    Sub sGetAttributeNamesBasedOnOEMandSupplier(xmlFilePath As String, sOemName As String, sSupplierName As String, sDivision As String)
        Dim xmlDoc As XDocument = Nothing
        Dim xmlElement As XElement = Nothing
        Dim aoxmlElements As XElement = Nothing
        Dim sAttributeName As String = ""
        Dim sAttributeValue As String = ""

        If xmlFilePath <> "" Then
            xmlDoc = XDocument.Load(xmlFilePath)
            If Not xmlDoc Is Nothing Then
                xmlElement = xmlDoc.Root
                If Not xmlElement Is Nothing Then
                    If (sOemName <> "") And (sSupplierName <> "") Then
                        'OEM Check
                        If xmlElement.Element(sOemName).HasElements Then
                            'SUPPLIER Check
                            If xmlElement.Element(sOemName).Element(sSupplierName).HasElements Then
                                'DIVISION Check
                                If xmlElement.Element(sOemName).Element(sSupplierName).Element(sDivision).HasElements Then
                                    'aoxmlElements = xmlElement.Element(sOemName).Element(sSupplierName).Element(sDivision).Elements("Attribute")
                                    'If Not aoxmlElements Is Nothing Then
                                    For Each oxmlElement As XElement In xmlElement.Element(sOemName).Element(sSupplierName).Element(sDivision).Elements("Attribute")
                                        sAttributeName = oxmlElement.Attribute("Name").Value
                                        If (Not sAttributeName Is Nothing) And (sAttributeName = XML_PARTNAME) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _PART_NAME = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_SHAPE) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _SHAPE_ATTR_NAME = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_SUBDETAILNUMBER) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _SUB_DETAIL_NUMBER = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_MASS) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _P_MASS = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_PURCHASEOPTION) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _PURCH_OPTION = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_STOCKSIZE) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _STOCK_SIZE = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_STOCKSIZEMETRIC) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _STOCK_SIZE_METRIC = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_FINISHTOLERANCE) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _FINISHTOLERANCE_ATTR_NAME = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_QUANTITY) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _QTY = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_MATERIAL) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _P_MAT = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_FINISHTOLVALUE1) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _RELIEF_CUT_FACE_TOLERANCE_VALUE = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_FINISHTOLVALUE2) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _FINISH_TOL_VALUE2 = sAttributeValue
                                            'ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_FINISHTOLVALUE3) Then
                                            '    sAttributeValue = oxmlElement.Attribute("Value").Value
                                            '    '*************************** NOT NEEDED FOR PSD *******************************
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_BOM) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _BOM_ATTR = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_CLIENTPARTNAME) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _CLIENT_PART_NAME = sAttributeValue
                                            'ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_DETAILSHEET1) Then
                                            '    sAttributeValue = oxmlElement.Attribute("Value").Value
                                            '    '*************************** NOT NEEDED FOR PSD *******************************
                                            'ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_LAYOUTSHEET1) Then
                                            '    sAttributeValue = oxmlElement.Attribute("Value").Value
                                            '    '*************************** NOT NEEDED FOR PSD *******************************
                                            'ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_LAYOUTSHEET2) Then
                                            '    sAttributeValue = oxmlElement.Attribute("Value").Value
                                            '    '*************************** NOT NEEDED FOR PSD *******************************
                                            'ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_LAYOUTSHEET3) Then
                                            '    sAttributeValue = oxmlElement.Attribute("Value").Value
                                            '    '*************************** NOT NEEDED FOR PSD *******************************
                                            'ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_SHEET) Then
                                            '    sAttributeValue = oxmlElement.Attribute("Value").Value
                                            '   '*************************** NOT NEEDED FOR PSD *******************************
                                            'ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_MMTOINCHCONVERSATION) Then
                                            '    sAttributeValue = oxmlElement.Attribute("Value").Value
                                            '    '*************************** NOT NEEDED FOR PSD *******************************
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_CLIENTSTOCKSIZE) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _CLIENT_STOCK_SIZE_ATTR = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_TOOLCLASS) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _TOOL_CLASS = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_TOOLID) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _TOOL_ID = sAttributeValue
                                        ElseIf (Not sAttributeName Is Nothing) And (sAttributeName = XML_ALTPURCH) Then
                                            sAttributeValue = oxmlElement.Attribute("Value").Value
                                            _ALTPURCH = sAttributeValue
                                        End If
                                    Next
                                    'End If

                                End If


                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    'Function to write message to Auto2D log file
    Sub sWriteToLogFile(sLogMessage As String)
        Dim sFolderPath As String = ""
        Dim sFilePath As String = ""
        If _sSweepDataOutputFolderPath <> "" Then
            'Create output folder path
            sFolderPath = Path.Combine(_sSweepDataOutputFolderPath, LOG_FOLDER, _sToolFolderName)
            If Not FnCheckFolderExists(sFolderPath) Then
                SCreateDirectory(sFolderPath)
            End If
            sFilePath = Path.Combine(_sSweepDataOutputFolderPath, LOG_FOLDER, _sToolFolderName, LOG_FILE)
            SWrite(DateTime.Now.ToString & Chr(9) & sLogMessage, sFilePath)
        End If
    End Sub

    'Populate floor Mount Face Normal to the excel file
    Sub sPopulateFloorMountFaceNormalToExcelFile(ByRef adPartOptimalRotMat() As Double)
        Dim iRowFilledInModelViewDC As Integer = 0
        Dim iColModelViewName As String = ""
        Dim iColDCSXx As Integer = 0
        Dim iColDCSXy As Integer = 0
        Dim iColDCSXz As Integer = 0
        Dim iColDCSYx As Integer = 0
        Dim iColDCSYy As Integer = 0
        Dim iColDCSYz As Integer = 0
        Dim iColDCSZx As Integer = 0
        Dim iColDCSZy As Integer = 0
        Dim iColDCSZz As Integer = 0

        iColModelViewName = 1
        iColDCSXx = 2
        iColDCSXy = 3
        iColDCSXz = 4
        iColDCSYx = 5
        iColDCSYy = 6
        iColDCSYz = 7
        iColDCSZx = 8
        iColDCSZy = 9
        iColDCSZz = 10

        'Get the last row filled detail in Model view cosine sheet. As it might contain datas.
        iRowFilledInModelViewDC = FnGetNumberofRows(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, 1, 1)

        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColModelViewName, FLOOR_MOUNT_DIR)

        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSXx, adPartOptimalRotMat(0).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSXy, adPartOptimalRotMat(1).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSXz, adPartOptimalRotMat(2).ToString)

        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSYx, adPartOptimalRotMat(3).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSYy, adPartOptimalRotMat(4).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSYz, adPartOptimalRotMat(5).ToString)

        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSZx, adPartOptimalRotMat(6).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSZy, adPartOptimalRotMat(7).ToString)
        SWriteValueToCell(_objWorkBk, VIEW_DIR_COS_SHEET_NAME, iRowFilledInModelViewDC + 1, iColDCSZz, adPartOptimalRotMat(8).ToString)
    End Sub
    'Loops for a specificied period of time (milliseconds)
    Private Sub wait(ByVal interval As Integer)
        Dim sw As New Stopwatch
        sw.Start()
        Do While sw.ElapsedMilliseconds < interval
            ' Allows UI to remain responsive
            'Do Nothing
            'sWriteToLogFile("waiting..")
        Loop
        sw.Stop()
    End Sub

    'Sub sCalculateTimeForAPI(sAPIName As String)

    '    _sEndTime = DateTime.Now

    '    _lRunTime = Abs(DateDiff(DateInterval.Second, _sEndTime, _sStartTime))
    '    sWriteToLogFile(sAPIName & " : " & _lRunTime.ToString & " Sec")
    'End Sub
    'Code added Oct-22-2018
    'Write the message to ProcessID.txt file present in Auto2D Log folder
    Sub sWriteProcessID(sProcessID As String)
        Dim sFolderPath As String = ""
        Dim sFilePath As String = ""
        If _sSweepDataOutputFolderPath <> "" Then
            'Create output folder path
            sFolderPath = Path.Combine(_sSweepDataOutputFolderPath, LOG_FOLDER, _sToolFolderName)
            If Not FnCheckFolderExists(sFolderPath) Then
                SCreateDirectory(sFolderPath)
            End If
            sFilePath = Path.Combine(_sSweepDataOutputFolderPath, LOG_FOLDER, _sToolFolderName, PROCESS_ID_FILE)
            SWrite(sProcessID, sFilePath)
        End If
    End Sub

    'Function to check if the part is weldment based on attribute
    Function FnChkPartisWeldmentBasedOnAttr(objPart As Part) As Boolean
        Dim sPartType As String = ""

        If Not objPart Is Nothing Then
            FnLoadPartFully(objPart)
            sPartType = FnGetStringUserAttribute(objPart, B_PART_TYPE)
            If sPartType <> "" Then
                If sPartType.ToUpper = WELDED_ASS Then
                    FnChkPartisWeldmentBasedOnAttr = True
                    Exit Function
                End If
            End If
        End If
        FnChkPartisWeldmentBasedOnAttr = False
    End Function
    'Function to get the valid Burnout solid body based on OEM
    Function FnGetValidBurnoutBodyForOEM(objPart As Part, sOemName As String) As Body()
        Dim objValidBurnOutBody As Body = Nothing
        Dim aoAllValidBurnOutBody() As Body = Nothing
        Dim aoFeature() As Feature = Nothing
        If sOemName <> "" Then
            If sOemName = DAIMLER_OEM_NAME Then
                If Not objPart Is Nothing Then
                    FnLoadPartFully(objPart)
                    For Each objBurnOutBody As Body In objPart.Bodies()
                        If objBurnOutBody.Layer = 110 Then
                            If objBurnOutBody.IsSolidBody Then
                                If aoAllValidBurnOutBody Is Nothing Then
                                    ReDim Preserve aoAllValidBurnOutBody(0)
                                    aoAllValidBurnOutBody(0) = objBurnOutBody
                                Else
                                    ReDim Preserve aoAllValidBurnOutBody(UBound(aoAllValidBurnOutBody) + 1)
                                    aoAllValidBurnOutBody(UBound(aoAllValidBurnOutBody)) = objBurnOutBody
                                End If
                            End If
                        End If
                    Next
                End If
            ElseIf sOemName = FIAT_OEM_NAME Then
                If Not objPart Is Nothing Then
                    If FnGetStringUserAttribute(objPart, BURNOUT_STRING) = BURNOUT_YES Then
                        aoFeature = FnCollectAllMembersOfFeatureGroup(objPart, BURNOUT_FEATURE_GROUP)
                        If Not aoFeature Is Nothing Then
                            For i = 0 To aoFeature.Length - 1
                                Try
                                    Dim bdtag As Tag = NXOpen.Tag.Null
                                    If TypeOf (aoFeature(i)) Is NXOpen.Features.BodyFeature Then
                                        FnGetUFSession.Modl.AskFeatBody(aoFeature(i).Tag, bdtag)
                                        objValidBurnOutBody = CType(NXObjectManager.Get(bdtag), Body)
                                        If objValidBurnOutBody.IsSolidBody Then
                                            If aoAllValidBurnOutBody Is Nothing Then
                                                ReDim Preserve aoAllValidBurnOutBody(0)
                                                aoAllValidBurnOutBody(0) = objValidBurnOutBody
                                            Else
                                                If Not aoAllValidBurnOutBody.Contains(objValidBurnOutBody) Then
                                                    ReDim Preserve aoAllValidBurnOutBody(UBound(aoAllValidBurnOutBody) + 1)
                                                    aoAllValidBurnOutBody(UBound(aoAllValidBurnOutBody)) = objValidBurnOutBody
                                                End If
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception

                                End Try
                            Next
                        End If
                    End If
                End If
            End If
        End If
        FnGetValidBurnoutBodyForOEM = aoAllValidBurnOutBody
    End Function
    'Function to suppress and unsuppress the feature groups
    Sub sAddOrRemoveFeatureGroupTemporarily(objPart As Part, bIsSuppress As Boolean)
        Dim aoAllCompInSession() As Component = Nothing
        Dim objFeatureGroup As FeatureGroup = Nothing
        Dim aoAllFeatureGroup() As FeatureGroup = Nothing
        Dim objChildPart As Part = Nothing


        If Not objPart Is Nothing Then
            aoAllCompInSession = FnGetAllComponentsInSession()
            If Not aoAllCompInSession Is Nothing Then
                For Each objChildComp As Component In aoAllCompInSession
                    objChildPart = FnGetPartFromComponent(objChildComp)
                    If Not objChildPart Is Nothing Then
                        FnLoadPartFully(objChildPart)
                        If FnGetStringUserAttribute(objChildPart, BURNOUT_STRING) = BURNOUT_YES Then
                            If Not objChildPart.Features Is Nothing Then
                                For Each objFeature As Feature In objChildPart.Features()
                                    'Collect all the feature Group
                                    Try
                                        objFeatureGroup = CType(objFeature, FeatureGroup)
                                        If Not objFeatureGroup Is Nothing Then
                                            'Collect all the feature group except corpo featuregroup
                                            If Not (objFeatureGroup.Name.ToUpper.Contains(BURNOUT_FEATURE_GROUP)) Then
                                                If aoAllFeatureGroup Is Nothing Then
                                                    ReDim Preserve aoAllFeatureGroup(0)
                                                    aoAllFeatureGroup(0) = objFeatureGroup
                                                Else
                                                    ReDim Preserve aoAllFeatureGroup(UBound(aoAllFeatureGroup) + 1)
                                                    aoAllFeatureGroup(UBound(aoAllFeatureGroup)) = objFeatureGroup
                                                End If
                                            End If
                                        End If
                                    Catch ex As Exception

                                    End Try
                                Next
                                If Not aoAllFeatureGroup Is Nothing Then
                                    Try
                                        FnGetNxSession.UpdateManager.SetDefaultUpdateFailureAction(Update.FailureOption.AcceptAll)
                                        If bIsSuppress Then
                                            objChildPart.Features.SuppressFeatures(aoAllFeatureGroup)
                                        Else
                                            objChildPart.Features.UnsuppressFeatures(aoAllFeatureGroup)
                                        End If
                                    Catch ex As Exception
                                        sWriteToLogFile("Error encountered in Suppress / UnSuppress feature group")
                                        sWriteToLogFile(ex.Message.ToString)
                                    End Try
                                    aoAllFeatureGroup = Nothing
                                End If
                            End If
                        End If
                    End If
                Next

            Else
                If Not objPart Is Nothing Then
                    FnLoadPartFully(objPart)
                    If FnGetStringUserAttribute(objPart, BURNOUT_STRING) = BURNOUT_YES Then
                        If Not objPart.Features Is Nothing Then
                            For Each objFeature As Feature In objPart.Features()
                                'Collect all the feature Group
                                Try
                                    objFeatureGroup = CType(objFeature, FeatureGroup)
                                    If Not objFeatureGroup Is Nothing Then
                                        'Collect all the feature group except corpo featuregroup
                                        If Not (objFeatureGroup.Name.ToUpper.Contains(BURNOUT_FEATURE_GROUP)) Then
                                            If aoAllFeatureGroup Is Nothing Then
                                                ReDim Preserve aoAllFeatureGroup(0)
                                                aoAllFeatureGroup(0) = objFeatureGroup
                                            Else
                                                ReDim Preserve aoAllFeatureGroup(UBound(aoAllFeatureGroup) + 1)
                                                aoAllFeatureGroup(UBound(aoAllFeatureGroup)) = objFeatureGroup
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception

                                End Try
                            Next
                            If Not aoAllFeatureGroup Is Nothing Then
                                Try
                                    FnGetNxSession.UpdateManager.SetDefaultUpdateFailureAction(Update.FailureOption.AcceptAll)
                                    If bIsSuppress Then
                                        objPart.Features.SuppressFeatures(aoAllFeatureGroup)
                                    Else
                                        objPart.Features.UnsuppressFeatures(aoAllFeatureGroup)
                                    End If
                                Catch ex As Exception
                                    sWriteToLogFile("Error encountered in Suppress / UnSuppress feature group")
                                    sWriteToLogFile(ex.Message.ToString)
                                End Try

                            End If
                        End If
                    End If
                End If
            End If

        End If
    End Sub


    'Code added on Nov-14-2018
    'Function to determine the Vectror of a co-Axial Cylindrical face in tubing
    '1. Collect all the non-Hole cylindrical face
    '2. Loop over all the non-Hole Cyl Face and Get the concentric face to it.,
    '3. Concentric face and the non Hole Cyl Face will be in opposite to curvature (i.e one in concave and other in convex)
    '4. Their Cylindrical axis will be parall / anti-Parallel to each other,
    '5. Get the Axis of this cylindrical face as the CoAxialCylindrical Face Vec

    Function FnGetCoAxialCylFaceVec(objBody As Body) As Double()
        Dim objPart As Part = Nothing
        Dim aoNonHoleCylFace() As Face = Nothing
        Dim objCylFace2 As Face = Nothing
        Dim adAxisVec() As Double = Nothing

        If Not objBody Is Nothing Then
            objPart = objBody.OwningPart
            If Not objPart Is Nothing Then
                aoNonHoleCylFace = FnGetColOfNonHoleCylFaces(objPart, objBody)
                If Not aoNonHoleCylFace Is Nothing Then
                    'Code modified on Sep26-2019
                    'Get the common Axis Vector in the Rectangular tubing
                    adAxisVec = FnGetIdenticalVectorFromCylFaces(objPart, aoNonHoleCylFace)
                    If Not adAxisVec Is Nothing Then
                        FnGetCoAxialCylFaceVec = adAxisVec
                        Exit Function
                    End If

                    'Code commented on Sep-26-2019
                    'Flaw in concentricCylFace
                    'For Each objCylFace1 As Face In aoNonHoleCylFace
                    '    If Not FnGetConcentricCylFace(objPart, objCylFace1) Is Nothing Then
                    '        objCylFace2 = FnGetConcentricCylFace(objPart, objCylFace1)
                    '        If Not objCylFace2 Is Nothing Then
                    '            If FnChkIfConcentricCylFace(objCylFace1, objCylFace2, CONCENTRIC_DISTANCE_TOLERANCE_BET_FACE_CENTERS) Then
                    '                adAxisVec = FnGetAxisVecCylFace(objPart, objCylFace1)
                    '                If Not adAxisVec Is Nothing Then
                    '                    FnGetCoAxialCylFaceVec = adAxisVec
                    '                    Exit Function
                    '                End If
                    '            End If
                    '        End If
                    '    End If
                    'Next
                End If
            End If
        End If
        FnGetCoAxialCylFaceVec = Nothing
    End Function
    'Code added Nov-14-2018
    'Function to obtain the Optimal Rotation Matrix for TUBINGS
    'For TUBING we use a seperate logic
    'Logic received from Varma -Email dated Nov-8-2018
    '1. Collect all the Trial Rotation Matrix
    '2. Get the CoAxial Cylindrical Face vector of the Tubing
    '3. Loop over all the trial rotation matrix,
    '4. CoAxial Cylindrical Face vector and the Z Axis Vector of the Rotation Matrix must be Parallel / Anti-Parallel to each other.
    '5. This orientation is choosen as the optimal Rotation Matrix
    Function FnDetermineOptimalRotationMatrixForTubg(objBody As Body, ByRef aoStructOrientationInfo() As structOrientationInfo) As Double()

        Dim adCoAxialFaceVec() As Double = Nothing
        Dim adZVecOfTrialRotMatrix(2) As Double
        Dim adOptimalRotMat(8) As Double

        If Not aoStructOrientationInfo Is Nothing Then
            If Not objBody Is Nothing Then
                adCoAxialFaceVec = FnGetCoAxialCylFaceVec(objBody)
                If Not adCoAxialFaceVec Is Nothing Then
                    For iIndex As Integer = 0 To UBound(aoStructOrientationInfo)

                        adZVecOfTrialRotMatrix(0) = aoStructOrientationInfo(iIndex).zx
                        adZVecOfTrialRotMatrix(1) = aoStructOrientationInfo(iIndex).zy
                        adZVecOfTrialRotMatrix(2) = aoStructOrientationInfo(iIndex).zz

                        If FnParallelAntiParallelCheck(adCoAxialFaceVec, adZVecOfTrialRotMatrix) Then
                            adOptimalRotMat(0) = aoStructOrientationInfo(iIndex).xx
                            adOptimalRotMat(1) = aoStructOrientationInfo(iIndex).xy
                            adOptimalRotMat(2) = aoStructOrientationInfo(iIndex).xz
                            adOptimalRotMat(3) = aoStructOrientationInfo(iIndex).yx
                            adOptimalRotMat(4) = aoStructOrientationInfo(iIndex).yy
                            adOptimalRotMat(5) = aoStructOrientationInfo(iIndex).yz
                            adOptimalRotMat(6) = aoStructOrientationInfo(iIndex).zx
                            adOptimalRotMat(7) = aoStructOrientationInfo(iIndex).zy
                            adOptimalRotMat(8) = aoStructOrientationInfo(iIndex).zz
                            FnDetermineOptimalRotationMatrixForTubg = adOptimalRotMat
                            Exit Function
                        End If
                    Next
                End If
            End If
        End If
        FnDetermineOptimalRotationMatrixForTubg = Nothing
    End Function
    'Code added Dec-24-2018
    'Function to create Sub DEtail Number for Fiat Weldments
    '1. For Fiat, use the Nomenclature number used for all the child component as the sub detail number attribute.
    '2. 281702Z_M-00000000000-020Z16015ZR-001-A00-02    (02 as the Sub detail Number)
    Sub sCreateSubDetNumForFiat(objWeldPart As Part)
        Dim aoAllChildComp As Component() = Nothing
        Dim asCompNames() As String = Nothing
        Dim sSubDetNum As String = ""
        Dim objChildPart As Part = Nothing

        If Not objWeldPart Is Nothing Then
            If Not objWeldPart.ComponentAssembly.RootComponent Is Nothing Then
                aoAllChildComp = FnGetAllComponentsInSession()
                If Not aoAllChildComp Is Nothing Then
                    For Each objChildComp As Component In aoAllChildComp
                        If Not objChildComp Is Nothing Then
                            If FnGetStringUserAttribute(objChildComp, B_PART_TYPE) = WELDED_CHILD Then
                                objChildPart = FnGetPartFromComponent(objChildComp)
                                If Not objChildPart Is Nothing Then
                                    FnLoadPartFully(objChildPart)
                                    asCompNames = Split(objChildComp.DisplayName, "-")
                                    If Not asCompNames Is Nothing Then
                                        If asCompNames.Length = 6 Then
                                            sSubDetNum = asCompNames(5)
                                            If sSubDetNum <> "" Then
                                                'Code added Jan-08-2019
                                                'Sub detail number should be 3 digit
                                                If sSubDetNum.Length = 2 Then
                                                    sSubDetNum = "0" & sSubDetNum
                                                ElseIf sSubDetNum.Length = 1 Then
                                                    sSubDetNum = "00" & sSubDetNum
                                                End If
                                                sSetIntegerUserAttribute(objChildPart, _SUB_DETAIL_NUMBER, sSubDetNum)
                                                sSetIntegerUserAttribute(objChildComp, _SUB_DETAIL_NUMBER, sSubDetNum)
                                                sWriteSubDetailNumForFiat(objChildComp, CInt(sSubDetNum))
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Sub sWriteSubDetailNumForFiat(objComp As Component, iSubDetNum As Integer)
        Dim objCompPart As Part = Nothing
        Dim bProcess As Boolean = False
        Dim sCompBodyname As String = ""
        Dim sBodyName As String = ""
        Dim bWriteSubDetNum As Boolean = False
        Dim iNosOfFilledRows As Integer

        If Not objComp Is Nothing Then
            objCompPart = FnGetPartFromComponent(objComp)
            'Get the number of rows of filled data
            iNosOfFilledRows = FnGetNumberofRows(_objWorkBk, BODYSHEETNAME, 1, BODY_INFO_START_ROW_WRITE)
            For Each Body As Body In _aoSolidBody
                bProcess = False
                bWriteSubDetNum = False
                'Code added May-14-2018
                'Validation added to check if the body is alive.
                Try
                    If Not Body Is Nothing Then
                        Dim iStatus As Integer
                        iStatus = FnGetUFSession.Obj.AskStatus(Body.Tag)
                        If iStatus = UFConstants.UF_OBJ_ALIVE Then
                            bProcess = True
                        End If
                    End If
                Catch ex As Exception
                    bProcess = False
                End Try
                If bProcess Then
                    sCompBodyName = Body.JournalIdentifier & " " & objComp.JournalIdentifier
                    For iloopIndex = BODY_INFO_START_ROW_WRITE To iNosOfFilledRows
                        sBodyName = FnReadSingleRowForColumn(_objWorkBk, BODYSHEETNAME, 1, iloopIndex)
                        If sBodyName = sCompBodyName Then
                            sSetIntegerUserAttribute(objComp, _SUB_DETAIL_NUMBER, iSubDetNum)
                            'Add the sub detail number at part level for each child component so that this value can be used at Sub Detail Callout Module
                            sSetIntegerUserAttribute(objCompPart, _SUB_DETAIL_NUMBER, iSubDetNum)
                            SWriteValueToCell(_objWorkBk, BODYSHEETNAME, iloopIndex, iColBodyDetailNos, iSubDetNum.ToString)
                            bWriteSubDetNum = True
                            Exit For
                        End If
                    Next
                End If
                If bWriteSubDetNum Then
                    Exit For
                End If
            Next
        End If
    End Sub

    'Function added Feb-20-2019
    'FUnction to check if the part is weldment based on different OEM
    Function FnCheckIfPartIsWeldmentBasedOnOEM(objPart As Part, sOemName As String) As Boolean
        Dim sPartname As String = ""

        If Not objPart Is Nothing Then
            FnLoadPartFully(objPart)
            sPartname = objPart.Leaf.ToString
            If (sOemName = GM_OEM_NAME) Or (sOemName = CHRYSLER_OEM_NAME) Or (sOemName = GESTAMP_OEM_NAME) Then
                If FnChkPartisWeldment(objPart) Then
                    FnCheckIfPartIsWeldmentBasedOnOEM = True
                    Exit Function
                End If
            ElseIf _sOemName = DAIMLER_OEM_NAME Then
                If FnCheckIfThisIsAWeldment(sPartname) Then
                    FnCheckIfPartIsWeldmentBasedOnOEM = True
                    Exit Function
                End If
            ElseIf _sOemName = FIAT_OEM_NAME Then
                If FnChkPartisWeldmentBasedOnAttr(objPart) Then
                    FnCheckIfPartIsWeldmentBasedOnOEM = True
                    Exit Function
                End If
            End If
        End If
        FnCheckIfPartIsWeldmentBasedOnOEM = False
    End Function
    'Code added Mar-29-2019
    Sub sComputePrimary2ViewOrientation(objPart As Part)
        'Following are the logic
        '(1) Identify the Construction hole face in Auxiliary frame.
        '(2) Identify all auxiliary faces.
        '         - A face is an auxiliary face if it satisfies the following criteria.
        '                 (a) Face type = "Planar"
        '                 (b) Face must be mis-aligned in B_LCS View.
        '                 (c) Face must contain at least one auxiliary hole or Slot.
        '                         -- Auxiliary hole = Hole whose axis is parallel / Anti-parallel to parent planar face + (Diameter < 20mm for weldments)
        '(3) Among the Std projections (LP, RP, BP, TP, RrP) of B_LCS, identify the projection in which Construction hole is visible.
        '         - Visibility of Construction hole = Axis of Construction hole must be parallel / Anti-parallel to Std projection's Z-axis DCs.
        '"(4) Identify common linear edge between auxilary faces (identified from step 2) and planar faces aligned parallel to and visible in Std projection identified in Step (3).
        '(i.e) We are trying to find the common edge between the auxilary faces and planar faces parallel / anti-parallel to the construction hole face)"
        '(5) Compute Axis of rotation.
        '         - Axis of rotation = DCs of common linear edge
        '(6) Compute angle of rotation.
        '         - Angle of rotation = Angle between face normal vectors of planar faces sharing common edge.
        '         - This is the angle between Auxiliary face and Construction hole view face.
        '(7) Compute rotation matrix obtained by rotating the part by angle computed in step (6) about axis of rotation obtained from Step (5).
        '        (Use <fnGetRotationMatrixAboutArbitraryVector> to compute rotation matrix)
        '(8) Compute B_Primary2 rotation matrix (Auxiliary view) using the formula below.
        '         - B_Primary2 rotation matrix = Rotation matrix obtained from Step (7) * B_LCS rotation matrix
        '(9) If auxiliary face is strictly parallel to B_Primary2's Z-axis DCs, then output B_Primary2.
        '(10) If not parallel, then repeat steps (6) to (8) by changing Angle of rotation to (-1)*Angle of rotation and output B_Primary2.

        Dim objConstHoleFace As Face = Nothing
        Dim lstOfAuxFace As List(Of Face) = Nothing
        Dim objPrimary1View As ModelingView = Nothing
        Dim adPrimary1RotMat() As Double = Nothing
        Dim objStdProjView As ModelingView = Nothing
        Dim objAdjPlanarFace As Face = Nothing
        Dim objAuxFace As Face = Nothing
        Dim objCommonEdge As Edge = Nothing
        Dim adAxisVecOfCommonEdge() As Double = Nothing
        Dim dAngleBetPlanarFaceAndAuxFace As Double = 0
        Dim adAdjPlanarFaceNormal() As Double = Nothing
        Dim adAuxFaceNormal() As Double = Nothing
        Dim objDirOfAdjPlanarFace As Direction = Nothing
        Dim objDirOfAuxFace As Direction = Nothing
        Dim adDirVecOfCommonEdge() As Double = Nothing
        Dim aDTrialRotMatrix(8) As Double
        Dim adPrimary2RotMat(8) As Double
        Dim adStdProjRotMat(8) As Double


        If Not objPart Is Nothing Then
            FnLoadPartFully(objPart)
            'Check if this part Needs Primary2 view computation
            If FnCheckIfComponentNeedsPrimary2ViewComputation(objPart) Then
                sWriteToLogFile("B_PRIMARY2 view computation is needed for this part")
                If FnChkIfModelingViewPresent(objPart, B_PRIMARY2) Then
                    sDeleteModellingView(objPart, B_PRIMARY2)
                End If
                'Get the construction Hole face
                objConstHoleFace = FnGetConstructionHoleInFrame(objPart)
                If Not objConstHoleFace Is Nothing Then
                    If FnChkIfModelingViewPresent(objPart, B_PRIMARY1) Then
                        objStdProjView = FnGetStdProjectViewHavingVisibleRefCylFace(objPart, B_PRIMARY1, objConstHoleFace)
                        If Not objStdProjView Is Nothing Then
                            objPrimary1View = FnGetModellingView(objPart, B_PRIMARY1)
                            If Not objPrimary1View Is Nothing Then
                                adPrimary1RotMat = FnGetRotMatOfAView(objPrimary1View)
                                lstOfAuxFace = FnCollectListOfAuxPlanarFaceInARefView(objPart, adPrimary1RotMat)
                                If Not lstOfAuxFace Is Nothing Then
                                    objAdjPlanarFace = FnGetAdjPlanarFaceToAuxFaceAlignedParallelToHoleFace(objPart, lstOfAuxFace, objConstHoleFace, objAuxFace)
                                    If (Not objAdjPlanarFace Is Nothing) And (Not objAuxFace Is Nothing) Then
                                        objDirOfAdjPlanarFace = FnGetDirOfFace(objPart, objAdjPlanarFace)
                                        objDirOfAuxFace = FnGetDirOfFace(objPart, objAuxFace)
                                        dAngleBetPlanarFaceAndAuxFace = FnGetAngleByVectors(objPart, objDirOfAdjPlanarFace, objDirOfAuxFace, bMinorAngle:=True)
                                        If dAngleBetPlanarFaceAndAuxFace <> 0 Then
                                            objCommonEdge = FnGetCommonEdgeBetweenTwoFaces(objAdjPlanarFace, objAuxFace)
                                            If Not objCommonEdge Is Nothing Then
                                                adDirVecOfCommonEdge = FnGetDirVecOfALinearEdge(objPart, objCommonEdge)
                                                If Not adDirVecOfCommonEdge Is Nothing Then
                                                    aDTrialRotMatrix = FnGetRotationMatrixAboutArbitraryVector(adDirVecOfCommonEdge, dAngleBetPlanarFaceAndAuxFace)
                                                    adStdProjRotMat = FnGetRotMatOfAView(objStdProjView)
                                                    FnGetUFSession.Mtx3.Multiply(adStdProjRotMat, aDTrialRotMatrix, adPrimary2RotMat)
                                                    adAuxFaceNormal = FnGetFaceNormal(objAuxFace)
                                                    If FnParallelCheck(adAuxFaceNormal, {adPrimary2RotMat(6), adPrimary2RotMat(7), adPrimary2RotMat(8)}) Then
                                                        sCreateCustomModellingView(objPart, adPrimary2RotMat(0), adPrimary2RotMat(0), adPrimary2RotMat(0),
                                                                                    adPrimary2RotMat(0), adPrimary2RotMat(0), adPrimary2RotMat(0),
                                                                                    adPrimary2RotMat(0), adPrimary2RotMat(0), adPrimary2RotMat(0), B_PRIMARY2, 1)
                                                    Else
                                                        '(10) If not parallel, then repeat steps (6) to (8) by changing Angle of rotation to (-1)*Angle of rotation and output B_Primary2.
                                                        dAngleBetPlanarFaceAndAuxFace = (-1) * dAngleBetPlanarFaceAndAuxFace
                                                        aDTrialRotMatrix = FnGetRotationMatrixAboutArbitraryVector(adDirVecOfCommonEdge, dAngleBetPlanarFaceAndAuxFace)
                                                        If Not aDTrialRotMatrix Is Nothing Then
                                                            FnGetUFSession.Mtx3.Multiply(adStdProjRotMat, aDTrialRotMatrix, adPrimary2RotMat)
                                                            If Not adPrimary2RotMat Is Nothing Then
                                                                If FnParallelCheck(adAuxFaceNormal, {adPrimary2RotMat(6), adPrimary2RotMat(7), adPrimary2RotMat(8)}) Then
                                                                    sCreateCustomModellingView(objPart, adPrimary2RotMat(0), adPrimary2RotMat(0), adPrimary2RotMat(0),
                                                                                    adPrimary2RotMat(0), adPrimary2RotMat(0), adPrimary2RotMat(0),
                                                                                    adPrimary2RotMat(0), adPrimary2RotMat(0), adPrimary2RotMat(0), B_PRIMARY2, 1)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If

                            End If
                        End If
                    End If
                End If
            Else
                sWriteToLogFile("B_PRIMARY2 view computation is not needed")
            End If
        End If
    End Sub


    'Function to check if this component needs B_PRIMARY2 view computation
    Function FnCheckIfComponentNeedsPrimary2ViewComputation(objPart As Part) As Boolean
        '1. Primary2 is needed only for Frame / Base compoonent
        '2. It is computed for the component which has the attribute B_ADA_TYPE = AUXILIARY FRAME. (attribute assigned during ADA)

        Dim sPartName As String = ""
        Dim sFrameType As String = ""

        If Not objPart Is Nothing Then
            FnLoadPartFully(objPart)
            'Check if this part is a Frame / Base component
            sPartName = FnGetPartName(objPart)
            If sPartName <> "" Then
                If (sPartName.ToUpper = "FRAME") Or (sPartName = "BASE") Then
                    'Check if this has an Auxiliary Frame attribute, which is created during ADA module
                    sFrameType = FnGetStringUserAttribute(objPart, B_ADA_TYPE)
                    If sFrameType <> "" Then
                        If sFrameType = AUXILIARY_FRAME Then
                            FnCheckIfComponentNeedsPrimary2ViewComputation = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        FnCheckIfComponentNeedsPrimary2ViewComputation = False
    End Function

    Function FnGetConstructionHoleInFrame(objPart As Part) As Face

        Dim aoDispFace() As DisplayableObject = Nothing
        Dim objConstHole As Face = Nothing

        If Not objPart Is Nothing Then
            FnLoadPartFully(objPart)
            aoDispFace = FnGetFaceObjectByAttributes(objPart, HOLE_SIZE, "DIA 6 H6 6 DEEP")
            If Not aoDispFace Is Nothing Then
                For Each objDisp As DisplayableObject In aoDispFace
                    If TypeOf (objDisp) Is Face Then
                        objConstHole = CType(objDisp, Face)
                        If Not objConstHole Is Nothing Then
                            If objConstHole.SolidFaceType = Face.FaceType.Cylindrical Then
                                FnGetConstructionHoleInFrame = objConstHole
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End If
        End If
        FnGetConstructionHoleInFrame = Nothing
    End Function

    'Function to collect all the Aux faces ina part
    Function FnCollectListOfAuxPlanarFaceInARefView(objPart As Part, adPrimary1RotMat() As Double) As List(Of Face)
        Dim lstOfValidBody As List(Of Body) = Nothing
        Dim lstOfAuxPlanarFace As List(Of Face) = Nothing

        If Not adPrimary1RotMat Is Nothing Then
            lstOfValidBody = FnGetListOfAllValidSolidBodiesFromNXPart(objPart)
            If Not lstOfValidBody Is Nothing Then
                For Each objBody As Body In lstOfValidBody
                    If Not objBody Is Nothing Then
                        If objBody.IsSolidBody Then
                            For Each objFace As Face In objBody.GetFaces()
                                If objFace.SolidFaceType = Face.FaceType.Planar Then
                                    If Not FnCheckIfFaceIsAligned(objFace, adPrimary1RotMat) Then
                                        If FnCheckIfRefPlanarFaceHasAnAuxHoleOnIt(objPart, objFace) Then
                                            If lstOfAuxPlanarFace Is Nothing Then
                                                lstOfAuxPlanarFace = New List(Of Face)
                                                lstOfAuxPlanarFace.Add(objFace)
                                            Else
                                                lstOfAuxPlanarFace.Add(objFace)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
        FnCollectListOfAuxPlanarFaceInARefView = lstOfAuxPlanarFace
    End Function

    'Function to check if the given planar face has a Aux Hole face on it.
    Function FnCheckIfRefPlanarFaceHasAnAuxHoleOnIt(objPart As Part, objRefFace As Face) As Boolean
        'A Planar Face can have an Aux HOle Face on it based on following criteria check
        '1. Check the Adjacent face(cylindrical) of the planar face.
        '2. If the Face normal of Planar face and the Cylindrical face axis are parallel-antiparallel, then Aux Hole is present on the planar face
        Dim adPlanarFaceNormal() As Double = Nothing
        Dim aTAdjFaceTag() As Tag = Nothing
        Dim objAdjHoleFace As Face = Nothing
        Dim adCylHoleAxis() As Double = Nothing

        If Not objRefFace Is Nothing Then
            adPlanarFaceNormal = FnGetFaceNormal(objRefFace)
            If FnCheckIfthePlanarFaceHasHolesOnIt(objRefFace) Then
                FnGetUFSession.Modl.AskAdjacFaces(objRefFace.Tag, aTAdjFaceTag)
                If Not aTAdjFaceTag Is Nothing Then
                    For Each objAdjFaceTag As Tag In aTAdjFaceTag
                        objAdjHoleFace = CType(NXObjectManager.Get(objAdjFaceTag), Face)
                        If Not objAdjHoleFace Is Nothing Then
                            If objAdjHoleFace.SolidFaceType = Face.FaceType.Cylindrical Then
                                adCylHoleAxis = FnGetAxisVecCylFace(objPart, objAdjHoleFace)
                                If Not adCylHoleAxis Is Nothing Then
                                    If FnParallelAntiParallelCheck(adPlanarFaceNormal, adCylHoleAxis) Then
                                        FnCheckIfRefPlanarFaceHasAnAuxHoleOnIt = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If
        FnCheckIfRefPlanarFaceHasAnAuxHoleOnIt = False
    End Function

    'Function to collect all the valid Solid body ina part file based on OEM
    Function FnGetListOfAllValidSolidBodiesFromNXPart(objPart As Part) As List(Of Body)
        Dim aoAllComps() As Component = Nothing
        Dim objChildPart As Part = Nothing
        Dim aoAllValidSolidBody() As Body = Nothing
        Dim objBodyToAnalyse As Body = Nothing
        Dim lstOfSolidBody As List(Of Body) = Nothing


        aoAllComps = FnGetAllComponentsInSession()
        If Not aoAllComps Is Nothing Then
            For Each objChildComp As Component In aoAllComps
                objChildPart = FnGetPartFromComponent(objChildComp)
                If Not objChildPart Is Nothing Then
                    'Load the part fully
                    FnLoadPartFully(objChildPart)
                    'Code added Jun-01-2018
                    'Collect the solid body from the part based on OEM
                    aoAllValidSolidBody = FnGetValidBodyForOEM(objChildPart, _sOemName)
                    If Not aoAllValidSolidBody Is Nothing Then
                        For Each objBody As Body In aoAllValidSolidBody
                            If Not objBody Is Nothing Then
                                If _sOemName = GM_OEM_NAME Or _sOemName = CHRYSLER_OEM_NAME Or _sOemName = GESTAMP_OEM_NAME Then
                                    If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                        objBodyToAnalyse = objBody
                                    Else
                                        objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                    End If
                                ElseIf _sOemName = DAIMLER_OEM_NAME Then
                                    If _sDivision = TRUCK_DIVISION Then
                                        'Check if it is the root component or it is a sub assembly compoenent
                                        If (objChildComp Is objPart.ComponentAssembly.RootComponent) Then
                                            'Component in truck. Get Prototype body
                                            objBodyToAnalyse = objBody
                                        Else
                                            'Weldment in truck. Get Occurrence body
                                            objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                        End If
                                    ElseIf _sDivision = CAR_DIVISION Then
                                        'Check if the component is a child component in weldment
                                        If FnCheckIfThisIsAChildCompInWeldment(objChildComp, _sOemName) Then
                                            'Weldment in Car. Get the Occurrence body
                                            objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                        Else
                                            'Component in Car. Get the Prototype Body
                                            'objBodyToAnalyse = objBody
                                            'Code modified on May-14-2018
                                            'Get the occurrence body. There was some mismatch between the geomety at container level and GEo level
                                            objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                        End If
                                    End If
                                    'Code added Nov-07-2018
                                    'Added Fiat OEM
                                ElseIf _sOemName = FIAT_OEM_NAME Then
                                    'Check if the component is a child component in weldment
                                    If objChildComp Is objPart.ComponentAssembly.RootComponent Then
                                        'Fiat component
                                        objBodyToAnalyse = objBody
                                    Else
                                        'This is a Fiat Weldment child component
                                        objBodyToAnalyse = CType(objChildComp.FindOccurrence(objBody), Body)
                                    End If
                                End If
                                If Not objBodyToAnalyse Is Nothing Then
                                    If lstOfSolidBody Is Nothing Then
                                        lstOfSolidBody = New List(Of Body)
                                        lstOfSolidBody.Add(objBodyToAnalyse)
                                    Else
                                        lstOfSolidBody.Add(objBodyToAnalyse)
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Else
            aoAllValidSolidBody = FnGetValidBodyForOEM(objPart, _sOemName)
            'sCalculateTimeForAPI("Identify Valid Body ")
            If Not aoAllValidSolidBody Is Nothing Then
                For Each objbody As Body In aoAllValidSolidBody
                    'Check whether the body is a solid body
                    If objbody.IsSolidBody Then
                        If lstOfSolidBody Is Nothing Then
                            lstOfSolidBody = New List(Of Body)
                            lstOfSolidBody.Add(objbody)
                        Else
                            lstOfSolidBody.Add(objbody)
                        End If
                    End If
                Next
            End If
        End If
        FnGetListOfAllValidSolidBodiesFromNXPart = lstOfSolidBody
    End Function

    'Function to get the Projection view in which the given ref Cyl face is visible
    '(Ref Cyl Face axis is parallel / antiparallel to the Std projections Z axis DC's)

    Function FnGetStdProjectViewHavingVisibleRefCylFace(objPart As Part, sRefViewName As String, objRefCylFace As Face) As ModelingView
        Dim objStdProjFV As ModelingView = Nothing
        Dim objStdProjRV As ModelingView = Nothing
        Dim objStdProjTP As ModelingView = Nothing
        Dim objStdProjRrV As ModelingView = Nothing
        Dim objStdProjBP As ModelingView = Nothing
        Dim objStdProjLP As ModelingView = Nothing
        Dim adStdProjZAxis(2) As Double
        Dim adCylFaceAxis(2) As Double
        Dim objPlanarFaceContainingCylFace As Face = Nothing
        Dim aTAdjTag() As Tag = Nothing
        Dim adPlanarVec() As Double = Nothing

        If sRefViewName <> Nothing Then
            If Not objRefCylFace Is Nothing Then

                FnGetUFSession.Modl.AskAdjacFaces(objRefCylFace.Tag, aTAdjTag)
                If Not aTAdjTag Is Nothing Then
                    For Each objTag As Tag In aTAdjTag
                        If Not NXObjectManager.Get(objTag) Is Nothing Then
                            If TypeOf NXObjectManager.Get(objTag) Is Face Then
                                objPlanarFaceContainingCylFace = CType(NXObjectManager.Get(objTag), Face)
                                If Not objPlanarFaceContainingCylFace Is Nothing Then
                                    If objPlanarFaceContainingCylFace.SolidFaceType = Face.FaceType.Planar Then
                                        adPlanarVec = FnGetFaceNormal(objPlanarFaceContainingCylFace)

                                    End If
                                End If
                            End If


                        End If
                    Next
                End If


                'Check if RefView is present in the model
                If FnChkIfModelingViewPresent(objPart, sRefViewName) Then
                    'Delete MOdeling View starting with STD_PROJ_
                    For Each objModelView As ModelingView In objPart.ModelingViews
                        If objModelView.Name.ToUpper.StartsWith("STD_PROJ_") Then
                            sDeleteModellingView(objPart, objModelView.Name)
                        End If
                    Next
                    'Create the Std Projection Views to the Ref View
                    sCreateModelingViewByName(objPart, STD_PROJ_FRONTVIEW, sRefViewName)
                    sCreateModelingViewByName(objPart, STD_PROJ_RIGHTVIEW, sRefViewName)
                    sCreateModelingViewByName(objPart, STD_PROJ_TOPVIEW, sRefViewName)
                    sCreateModelingViewByName(objPart, STD_PROJ_REARVIEW, sRefViewName)
                    sCreateModelingViewByName(objPart, STD_PROJ_LEFTVIEW, sRefViewName)
                    sCreateModelingViewByName(objPart, STD_PROJ_BOTTOMVIEW, sRefViewName)

                    If FnChkIfModelingViewPresent(objPart, STD_PROJ_FRONTVIEW) Then
                        objStdProjFV = FnGetModellingView(objPart, STD_PROJ_FRONTVIEW)
                        If Not objStdProjFV Is Nothing Then
                            adStdProjZAxis(0) = objStdProjFV.Matrix.Zx
                            adStdProjZAxis(1) = objStdProjFV.Matrix.Zy
                            adStdProjZAxis(2) = objStdProjFV.Matrix.Zz
                            If FnParallelAntiParallelCheck(adStdProjZAxis, adPlanarVec) Then
                                FnGetStdProjectViewHavingVisibleRefCylFace = objStdProjFV
                                Exit Function
                            End If
                        End If
                    End If

                    If FnChkIfModelingViewPresent(objPart, STD_PROJ_RIGHTVIEW) Then
                        objStdProjRV = FnGetModellingView(objPart, STD_PROJ_RIGHTVIEW)
                        If Not objStdProjRV Is Nothing Then
                            adStdProjZAxis(0) = objStdProjRV.Matrix.Zx
                            adStdProjZAxis(1) = objStdProjRV.Matrix.Zy
                            adStdProjZAxis(2) = objStdProjRV.Matrix.Zz
                            If FnParallelAntiParallelCheck(adStdProjZAxis, adPlanarVec) Then
                                FnGetStdProjectViewHavingVisibleRefCylFace = objStdProjRV
                                Exit Function
                            End If
                        End If
                    End If

                    If FnChkIfModelingViewPresent(objPart, STD_PROJ_TOPVIEW) Then
                        objStdProjTP = FnGetModellingView(objPart, STD_PROJ_TOPVIEW)
                        If Not objStdProjTP Is Nothing Then
                            adStdProjZAxis(0) = objStdProjTP.Matrix.Zx
                            adStdProjZAxis(1) = objStdProjTP.Matrix.Zy
                            adStdProjZAxis(2) = objStdProjTP.Matrix.Zz
                            If FnParallelAntiParallelCheck(adStdProjZAxis, adPlanarVec) Then
                                FnGetStdProjectViewHavingVisibleRefCylFace = objStdProjTP
                                Exit Function
                            End If
                        End If
                    End If

                    If FnChkIfModelingViewPresent(objPart, STD_PROJ_REARVIEW) Then
                        objStdProjRrV = FnGetModellingView(objPart, STD_PROJ_REARVIEW)
                        If Not objStdProjRrV Is Nothing Then
                            adStdProjZAxis(0) = objStdProjRrV.Matrix.Zx
                            adStdProjZAxis(1) = objStdProjRrV.Matrix.Zy
                            adStdProjZAxis(2) = objStdProjRrV.Matrix.Zz
                            If FnParallelAntiParallelCheck(adStdProjZAxis, adPlanarVec) Then
                                FnGetStdProjectViewHavingVisibleRefCylFace = objStdProjRrV
                                Exit Function
                            End If
                        End If
                    End If

                    If FnChkIfModelingViewPresent(objPart, STD_PROJ_LEFTVIEW) Then
                        objStdProjLP = FnGetModellingView(objPart, STD_PROJ_LEFTVIEW)
                        If Not objStdProjLP Is Nothing Then
                            adStdProjZAxis(0) = objStdProjLP.Matrix.Zx
                            adStdProjZAxis(1) = objStdProjLP.Matrix.Zy
                            adStdProjZAxis(2) = objStdProjLP.Matrix.Zz
                            If FnParallelAntiParallelCheck(adStdProjZAxis, adPlanarVec) Then
                                FnGetStdProjectViewHavingVisibleRefCylFace = objStdProjLP
                                Exit Function
                            End If
                        End If
                    End If

                    If FnChkIfModelingViewPresent(objPart, STD_PROJ_BOTTOMVIEW) Then
                        objStdProjBP = FnGetModellingView(objPart, STD_PROJ_BOTTOMVIEW)
                        If Not objStdProjBP Is Nothing Then
                            adStdProjZAxis(0) = objStdProjBP.Matrix.Zx
                            adStdProjZAxis(1) = objStdProjBP.Matrix.Zy
                            adStdProjZAxis(2) = objStdProjBP.Matrix.Zz
                            If FnParallelAntiParallelCheck(adStdProjZAxis, adPlanarVec) Then
                                FnGetStdProjectViewHavingVisibleRefCylFace = objStdProjBP
                                Exit Function
                            End If
                        End If
                    End If
                End If

            End If

        End If
        FnGetStdProjectViewHavingVisibleRefCylFace = Nothing
    End Function

    'Function to get the common edge between the Aux Face and the Adjacent Planar face (Adjacent planar face should have the same face normal of the Construction Hole axis)
    Function FnGetCommonEdgeBetweenTwoFaces(objFace1 As Face, objFace2 As Face) As Edge
        Dim aoFaceTag() As Tag = Nothing
        Dim objAdjFace As Face = Nothing

        If (Not objFace1 Is Nothing) And (Not objFace2 Is Nothing) Then

            For Each objEdgeOnFace1 As Edge In objFace1.GetEdges
                If objEdgeOnFace1.SolidEdgeType = Edge.EdgeType.Linear Then
                    FnGetUFSession.Modl.AskEdgeFaces(objEdgeOnFace1.Tag, aoFaceTag)
                    If Not aoFaceTag Is Nothing Then
                        For Each objFaceTag As Tag In aoFaceTag
                            If TypeOf (NXObjectManager.Get(objFaceTag)) Is Face Then
                                objAdjFace = CType(NXObjectManager.Get(objFaceTag), Face)
                                If objAdjFace.SolidFaceType = Face.FaceType.Planar Then
                                    If objAdjFace Is objFace2 Then
                                        FnGetCommonEdgeBetweenTwoFaces = objEdgeOnFace1
                                        Exit Function
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End If
        FnGetCommonEdgeBetweenTwoFaces = Nothing
    End Function
    'Function to get the correct adjacent planar face to the given Aux faces. Adjacent planar face normal should be parallel to the ref Hole axis.
    Function FnGetAdjPlanarFaceToAuxFaceAlignedParallelToHoleFace(objPart As Part, lstOfAuxFace As List(Of Face), objHoleFace As Face, ByRef objFilteredAuxFace As Face) As Face

        Dim adHoleAxisVec() As Double = Nothing
        Dim aAdjFaceTag() As Tag = Nothing
        Dim objAdjPlanarFace As Face = Nothing
        Dim adAdjPlanarFaceNormal() As Double = Nothing

        If Not lstOfAuxFace Is Nothing Then
            If Not objHoleFace Is Nothing Then
                adHoleAxisVec = FnGetAxisVecCylFace(objPart, objHoleFace)
                If Not adHoleAxisVec Is Nothing Then
                    For Each objAuxFace As Face In lstOfAuxFace
                        FnGetUFSession.Modl.AskAdjacFaces(objAuxFace.Tag, aAdjFaceTag)
                        If Not aAdjFaceTag Is Nothing Then
                            For Each objAdjTag As Tag In aAdjFaceTag
                                If TypeOf (NXObjectManager.Get(objAdjTag)) Is Face Then
                                    objAdjPlanarFace = CType(NXObjectManager.Get(objAdjTag), Face)
                                    If objAdjPlanarFace.SolidFaceType = Face.FaceType.Planar Then
                                        adAdjPlanarFaceNormal = FnGetFaceNormal(objAdjPlanarFace)
                                        If Not adAdjPlanarFaceNormal Is Nothing Then
                                            If FnParallelCheck(adAdjPlanarFaceNormal, adHoleAxisVec) Then
                                                objFilteredAuxFace = objAuxFace
                                                FnGetAdjPlanarFaceToAuxFaceAlignedParallelToHoleFace = objAdjPlanarFace
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        End If
        FnGetAdjPlanarFaceToAuxFaceAlignedParallelToHoleFace = Nothing
    End Function

    'Function to rotate a matrix about arbitary vector
    'Code converted to Vb.Net from Varma's code
    Function FnGetRotationMatrixAboutArbitraryVector(adArbVector() As Double, dAngle As Double) As Double()
        'Description: This UDF computes the rotation matrix that is required to rotate a point/vector _
        '            about an arbitrary vector by a certain angle.
        'Algorithm: http://inside.mines.edu/fs_home/gmurray/ArbitraryAxisRotation/

        '(1) Compute unit vector of input arbitrary vector.
        '(2) Rotate unit vector about the z axis so that the rotation axis lies in the xz plane.
        '(3) Rotate unit vector about the y axis so that the rotation axis lies along the z axis.
        '(4) Perform the desired rotation by alpha angle about the z axis.
        '(5) Apply the inverse of step (2).
        '(6) Apply the inverse of step (1).

        Dim dAngleInRad As Double = Nothing
        Dim adUnitVec() As Double = Nothing
        Dim dU As Double = Nothing
        Dim dV As Double = Nothing
        Dim dW As Double = Nothing
        Dim adFinalRotMat() As Double = Nothing
        Dim dVecLengthOfdUdV As Double = Nothing
        Dim adTxz(8) As Double
        Dim adTz(8) As Double
        Dim adRz(8) As Double
        Dim adInverseTxz() As Double = Nothing
        Dim adInverseTz() As Double = Nothing
        Dim adMultITxzITz(8) As Double
        Dim adMultRzTz(8) As Double
        Dim adMultITxzITzRzTz(8) As Double
        Dim adMultITxzITzRzTzTxz(8) As Double


        dAngleInRad = dAngle * (Math.PI / 180)
        If Not adArbVector Is Nothing Then
            adUnitVec = FnGetDCs(adArbVector)
            If Not adUnitVec Is Nothing Then
                dU = adUnitVec(0)
                dV = adUnitVec(1)
                dW = adUnitVec(2)
                'If the rotation axis happens to be z-axis, then output its rotation matrix directly.
                If ((Abs(dU) < ONE_DEG_TOLERANCE) And (Abs(dV) < ONE_DEG_TOLERANCE) And (Abs(dW - 1) < ONE_DEG_TOLERANCE Or Abs(dW + 1) < ONE_DEG_TOLERANCE)) Then
                    adFinalRotMat(0) = Cos(dAngleInRad)
                    adFinalRotMat(1) = -1 * Sin(dAngleInRad)
                    adFinalRotMat(2) = 0

                    adFinalRotMat(3) = Sin(dAngleInRad)
                    adFinalRotMat(4) = Cos(dAngleInRad)
                    adFinalRotMat(5) = 0

                    adFinalRotMat(6) = 0
                    adFinalRotMat(7) = 0
                    adFinalRotMat(8) = 1
                    FnGetRotationMatrixAboutArbitraryVector = adFinalRotMat
                    Exit Function
                End If

                dVecLengthOfdUdV = FnComputeVectorLength(dU, dV, 0)

                adTxz(0) = dU / dVecLengthOfdUdV
                adTxz(1) = dV / dVecLengthOfdUdV
                adTxz(2) = 0

                adTxz(3) = -1 * (dV / dVecLengthOfdUdV)
                adTxz(4) = dU / dVecLengthOfdUdV
                adTxz(5) = 0

                adTxz(6) = 0
                adTxz(7) = 0
                adTxz(8) = 1

                adTz(0) = dW
                adTz(1) = 0
                adTz(2) = -1 * dVecLengthOfdUdV

                adTz(3) = 0
                adTz(4) = 1
                adTz(5) = 0

                adTz(6) = dVecLengthOfdUdV
                adTz(7) = 0
                adTz(8) = dW


                adRz(0) = Cos(dAngleInRad)
                adRz(1) = -1 * Sin(dAngleInRad)
                adRz(2) = 0

                adRz(3) = Sin(dAngleInRad)
                adRz(4) = Cos(dAngleInRad)
                adRz(5) = 0

                adRz(6) = 0
                adRz(7) = 0
                adRz(8) = 1

                adInverseTxz = FnGetInvertMatrix(adTxz)
                adInverseTz = FnGetInvertMatrix(adTz)

                'Order in which multiplication was done in Varma's code
                'ITxz_ITz_RzAlpha_Tz_Txz = WorksheetFunction.MMult(WorksheetFunction.MMult(WorksheetFunction.MMult(ITxz, ITz), WorksheetFunction.MMult(Rz, Tz)), Txz)

                FnGetUFSession.Mtx3.Multiply(adInverseTz, adInverseTxz, adMultITxzITz)
                FnGetUFSession.Mtx3.Multiply(adTz, adRz, adMultRzTz)
                FnGetUFSession.Mtx3.Multiply(adMultRzTz, adMultITxzITz, adMultITxzITzRzTz)
                FnGetUFSession.Mtx3.Multiply(adTxz, adMultITxzITzRzTz, adMultITxzITzRzTzTxz)

                FnGetRotationMatrixAboutArbitraryVector = adMultITxzITzRzTzTxz
            End If
        End If
    End Function
    'Caode converted to vb.net based on Varma's code
    Function FnGetDCs(adVector() As Double) As Double()
        'to get direction cosines of a vector

        Dim dVecLength As Double = -1
        Dim adDCsVector(2) As Double

        dVecLength = FnComputeVectorLength(adVector(0), adVector(1), adVector(2))

        If (dVecLength > 10 ^ -1) Then
            adDCsVector(0) = adVector(0) / dVecLength
            adDCsVector(1) = adVector(1) / dVecLength
            adDCsVector(2) = adVector(2) / dVecLength
        Else
            adDCsVector(0) = adVector(0)
            adDCsVector(1) = adVector(1)
            adDCsVector(2) = adVector(2)
        End If
        FnGetDCs = adDCsVector
    End Function
    'Function to inverse 3X3 matrix.
    'This function is coded since UF function in NX doesnot have Inverse matrix for 3X3 matrix
    Function FnGetInvertMatrix(ad3X3Matrix() As Double) As Double()
        Dim ad4X4Matrix(15) As Double
        Dim adInverse4X4Matrix(15) As Double
        Dim adInverse3X3Matrix(8) As Double

        'Convert Matrix3X3 to Matrix4X4. This conversion is used, since NXAPI doesnot contain Inverse in 3X3 matrix
        FnGetUFSession.Mtx3.Mtx4(ad3X3Matrix, ad4X4Matrix)

        'Inverse the Matrix4X4
        FnGetUFSession.Mtx4.Invert(ad4X4Matrix, adInverse4X4Matrix)

        'Convert Matrix4X4 to Matrix3X3 by eliminating last row and last column data
        adInverse3X3Matrix(0) = adInverse4X4Matrix(0)
        adInverse3X3Matrix(1) = adInverse4X4Matrix(1)
        adInverse3X3Matrix(2) = adInverse4X4Matrix(2)

        adInverse3X3Matrix(3) = adInverse4X4Matrix(4)
        adInverse3X3Matrix(4) = adInverse4X4Matrix(5)
        adInverse3X3Matrix(5) = adInverse4X4Matrix(6)

        adInverse3X3Matrix(6) = adInverse4X4Matrix(8)
        adInverse3X3Matrix(7) = adInverse4X4Matrix(9)
        adInverse3X3Matrix(8) = adInverse4X4Matrix(10)

        FnGetInvertMatrix = adInverse3X3Matrix
    End Function

    Function FnGetRotMatOfAView(objMdlView As ModelingView) As Double()
        Dim adRotMat(8) As Double

        If Not objMdlView Is Nothing Then
            adRotMat(0) = objMdlView.Matrix.Xx
            adRotMat(1) = objMdlView.Matrix.Xy
            adRotMat(2) = objMdlView.Matrix.Xz
            adRotMat(3) = objMdlView.Matrix.Yx
            adRotMat(4) = objMdlView.Matrix.Yy
            adRotMat(5) = objMdlView.Matrix.Yz
            adRotMat(6) = objMdlView.Matrix.Zx
            adRotMat(7) = objMdlView.Matrix.Zy
            adRotMat(8) = objMdlView.Matrix.Zz
            FnGetRotMatOfAView = adRotMat
            Exit Function
        End If
        FnGetRotMatOfAView = Nothing
    End Function

    Function FnCheckIfPartIsValidMakeDetail(objPart As Part) As Boolean

        If Not objPart Is Nothing Then
            FnLoadPartFully(objPart)
            If (_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME) Then
                'For Chrysler, the part should not be a ALT PURCH component
                'We are skipping ALT PURCH component.
                'Email from Amitabh as directed by John dated Jun-19-2019 "Skip DB_ALT_PURCH Components
                'If FnGetStringUserAttribute(objPart, "DB_ALT_PURCH") = "" Then
                If FnGetStringUserAttribute(objPart, _ALTPURCH) = "" Then
                    FnCheckIfPartIsValidMakeDetail = True
                    Exit Function
                End If
            Else
                FnCheckIfPartIsValidMakeDetail = True
                Exit Function
            End If
        End If
        FnCheckIfPartIsValidMakeDetail = False
    End Function

    'cODE ADDEDJan-29-2020
    Function FnGetFeatNameMappingValueForSweepData(objFace As Face) As String
        Dim sFeatName As String = ""
        Dim sHoleSize As String = ""
        Dim bIsPreFabHole As Boolean = False
        Dim sMappedFeatName As String = ""

        If Not objFace Is Nothing Then
            sFeatName = FnGetStringUserAttribute(objFace, FEAT_NAME_FACE_ATTR)
            sHoleSize = FnGetStringUserAttribute(objFace, HOLE_SIZE)
            'Check if this hole is a PreFab Hole
            If FnGetStringUserAttribute(objFace, PRE_FAB_HOLE_ATTR_TITLE) = PRE_FAB_HOLE_ATTR_VALUE Then
                bIsPreFabHole = True
            Else
                bIsPreFabHole = False
            End If

            If sFeatName <> "" Then
                If sFeatName.ToUpper.Contains("DOWEL") Then
                    sMappedFeatName = "DOWEL"
                ElseIf sFeatName.ToUpper.Contains("PRECISION") Then
                    sMappedFeatName = "PRECISION"
                ElseIf sFeatName.ToUpper.Contains("COORD") Then
                    sMappedFeatName = "COORD"
                ElseIf (sFeatName.ToUpper.Contains("CLEAR_HOLE")) Or (sFeatName.ToUpper.Contains("CLR")) Then
                    If bIsPreFabHole Then
                        sMappedFeatName = "PRE-FAB"
                    Else
                        sMappedFeatName = "CLR"
                    End If

                ElseIf (sFeatName.ToUpper.Contains("CBORE")) Then
                    If sHoleSize <> "" Then
                        If sHoleSize.ToUpper.Contains("C'BORE") Then
                            sMappedFeatName = "CBORE"
                        ElseIf sHoleSize.ToUpper.Contains("S'FACE") Then
                            sMappedFeatName = "SFACE"
                        End If
                    End If
                ElseIf (sfeatname.ToUpper.Contains("C'BORE")) Then
                    sMappedFeatName = "CBORE"
                ElseIf (sfeatname.ToUpper.Contains("S'FACE")) Then
                    sMappedFeatName = "SFACE"
                ElseIf (sFeatName.ToUpper.Contains("CSINK")) Or (sFeatName.ToUpper.Contains("C'SINK")) Then
                    sMappedFeatName = "CSINK"
                ElseIf (sFeatName.ToUpper.Contains("ATTACH")) Or (sFeatName.ToUpper.Contains("TAP")) Then
                    sMappedFeatName = "TAP"
                ElseIf (sFeatName.ToUpper.Contains("CONST")) Then
                    sMappedFeatName = "CONST"
                ElseIf (sFeatName.ToUpper.Contains("VP")) Then
                    sMappedFeatName = "VP"
                Else
                    sMappedFeatName = sFeatName.ToUpper
                End If
            End If
        End If
        FnGetFeatNameMappingValueForSweepData = sMappedFeatName
    End Function

    'Code added Mar-09-2020
    'For a Given OEM - Supplier, get the correct name which needs to be added to the drawing sheet. 
    'This mapping name should be fetched from xml file
    Function FnGetMappingSupplierName(sOemName As String, sSuppNameInPSD As String) As String

        Dim sOemSuppDesignSourceFilePath As String = ""
        Dim xmlDoc As XDocument = Nothing
        Dim xmlElement As XElement = Nothing
        Dim aoxmlElements As XElement = Nothing
        Dim sSupplierName As String = ""
        Dim sSupplierValue As String = ""

        'Check if the DesignSource XML config file is present in the Config folder
        sOemSuppDesignSourceFilePath = Path.Combine(FnGetExecutionFolderPath(), CONFIG_FOLDER_NAME, DESIGN_SOURCE_CONFIG_FILE_NAME)
        If FnCheckFileExists(sOemSuppDesignSourceFilePath) Then
            sWriteToLogFile(DESIGN_SOURCE_CONFIG_FILE_NAME & " file is present inside the Config folder")
            xmlDoc = XDocument.Load(sOemSuppDesignSourceFilePath)
            If Not xmlDoc Is Nothing Then
                xmlElement = xmlDoc.Root
                If Not xmlElement Is Nothing Then
                    If (sOemName <> "") And (sSuppNameInPSD <> "") Then
                        If xmlElement.Element(sOemName).HasElements Then
                            For Each oxmlElement As XElement In xmlElement.Element(sOemName).Elements("Config")
                                sSupplierName = oxmlElement.Attribute("Name").Value
                                If sSupplierName <> "" Then
                                    If (sSuppNameInPSD.ToUpper) = (sSupplierName.ToUpper) Then
                                        sSupplierValue = oxmlElement.Attribute("Value").Value
                                        FnGetMappingSupplierName = sSupplierValue
                                        Exit Function
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        End If
        FnGetMappingSupplierName = sSuppNameInPSD
    End Function
    'List of Error Names for which the PSD file needs to be processed again and again will be in asErrorNamesInList
    Function FnCheckIfErrorNameIsInList(sExceptionErrorName As String, asErrorNamesInList() As String) As Boolean

        If Not asErrorNamesInList Is Nothing Then
            If sExceptionErrorName <> "" Then

                For Each sErrorNameInList As String In asErrorNamesInList
                    If sErrorNameInList <> "" Then
                        If sExceptionErrorName.ToUpper.Contains(sErrorNameInList) Then
                            FnCheckIfErrorNameIsInList = True
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
        FnCheckIfErrorNameIsInList = False
    End Function

    Function FnChannelABStockValuesForDaimler() As Dictionary(Of String, String)

        Dim dictOfChannelABValues As Dictionary(Of String, String) = New Dictionary(Of String, String)
        dictOfChannelABValues.Add(80, 45)
        dictOfChannelABValues.Add(100, 50)
        dictOfChannelABValues.Add(120, 55)
        dictOfChannelABValues.Add(140, 60)
        dictOfChannelABValues.Add(160, 65)
        dictOfChannelABValues.Add(180, 70)
        dictOfChannelABValues.Add(200, 75)
        dictOfChannelABValues.Add(220, 80)
        dictOfChannelABValues.Add(250, 85)
        dictOfChannelABValues.Add(300, 100)
        dictOfChannelABValues.Add(350, 100)
        dictOfChannelABValues.Add(400, 100)
        FnChannelABStockValuesForDaimler = dictOfChannelABValues
    End Function

    Function FnBeamABValuesForDaimler() As Dictionary(Of String, String)
        Dim dictOfBeamABValues As Dictionary(Of String, String) = New Dictionary(Of String, String)
        dictOfBeamABValues.Add(80, 40)
        dictOfBeamABValues.Add(100, 50)
        dictOfBeamABValues.Add(120, 60)
        dictOfBeamABValues.Add(140, 70)
        dictOfBeamABValues.Add(160, 80)
        dictOfBeamABValues.Add(180, 90)
        dictOfBeamABValues.Add(200, 100)
        dictOfBeamABValues.Add(220, 110)
        dictOfBeamABValues.Add(240, 120)
        dictOfBeamABValues.Add(250, 125)
        dictOfBeamABValues.Add(270, 125)
        dictOfBeamABValues.Add(300, 130)
        dictOfBeamABValues.Add(350, 140)
        dictOfBeamABValues.Add(400, 150)
        dictOfBeamABValues.Add(450, 160)
        dictOfBeamABValues.Add(500, 170)
        dictOfBeamABValues.Add(550, 180)
        dictOfBeamABValues.Add(600, 210)
        FnBeamABValuesForDaimler = dictOfBeamABValues
    End Function

    'Code added Sep-02-2020
    'New Logic to find the face center of a Conical face provided by Pradeep
    ' Use a In-built-In API To fetch "Face Axis direction" (A^) And "Face center on Axis" (C) For the Conical face. As Long these two are verified To be accurate
    '1. Compute geometric mean of edge centers of the two circular edges (C_m)
    '     C_m = (E1 + E2)/2
    '2. Compute the length D of the vector (C_m - C) . A^
    '     Retain the sign Of this distance, i.e., positive Or negative. 
    '3. Now, the New center which Is going to be within face bounds can be computed as: 
    '     C_new = C + (D times A^)
    '+, - , (.) above denote vector addition, subtraction And dot product.

    'This logic works even In Case split faces Of CATIA And Hirotec models. 

    Function FnGetFaceCenterOfConicalFace(objConicalFace As Face) As Double()
        Dim iFaceType As Integer = 0
        Dim adCenterPoint(2) As Double
        Dim adDir(2) As Double
        Dim adBox(5) As Double
        Dim dRadius As Double = 0.0
        Dim dRadData As Double = 0.0
        Dim iNormDir As Integer = 0
        Dim adFaceCenterOnAxis(2) As Double
        Dim adFaceAxisDir(2) As Double
        Dim aoAllEdgeInFace() As Edge = Nothing
        Dim aoAllCircularEdge() As Edge = Nothing
        Dim adGeoMeanEdgeCenter() As Double = Nothing
        Dim dLengthOfVec As Double = Nothing
        Dim dDiffInGeoMeanCenterAndFaceCenterONAxis(2) As Double
        Dim adNewFaceCenter(2) As Double




        If Not objConicalFace Is Nothing Then
            If objConicalFace.SolidFaceType = Face.FaceType.Conical Then
                FnGetUFSession.Modl.AskFaceData(objConicalFace.Tag, iFaceType, adCenterPoint, adDir, adBox, dRadius, dRadData, iNormDir)
                adFaceCenterOnAxis = adCenterPoint
                adFaceAxisDir = adDir
                aoAllEdgeInFace = objConicalFace.GetEdges
                If Not aoAllEdgeInFace Is Nothing Then
                    For Each objEDge As Edge In aoAllEdgeInFace
                        If objEDge.SolidEdgeType = Edge.EdgeType.Circular Then
                            If aoAllCircularEdge Is Nothing Then
                                ReDim Preserve aoAllCircularEdge(0)
                                aoAllCircularEdge(0) = objEDge
                            Else
                                ReDim Preserve aoAllCircularEdge(UBound(aoAllCircularEdge) + 1)
                                aoAllCircularEdge(UBound(aoAllCircularEdge)) = objEDge
                            End If
                        End If
                    Next
                    If Not aoAllCircularEdge Is Nothing Then
                        '1. Compute geometric mean of edge centers of the two circular edges (C_m)
                        '     C_m = (E1 + E2)/2
                        adGeoMeanEdgeCenter = FnGetGeometricMeanOfAllEdgeCenter(aoAllCircularEdge)
                        If Not adGeoMeanEdgeCenter Is Nothing Then
                            '2. Compute the length D of the vector (C_m - C) . A^
                            '     Retain the sign Of this distance, i.e., positive Or negative.
                            dDiffInGeoMeanCenterAndFaceCenterONAxis(0) = adGeoMeanEdgeCenter(0) - adFaceCenterOnAxis(0)
                            dDiffInGeoMeanCenterAndFaceCenterONAxis(1) = adGeoMeanEdgeCenter(1) - adFaceCenterOnAxis(1)
                            dDiffInGeoMeanCenterAndFaceCenterONAxis(2) = adGeoMeanEdgeCenter(2) - adFaceCenterOnAxis(2)

                            dLengthOfVec = FnGetDotProduct(dDiffInGeoMeanCenterAndFaceCenterONAxis, adFaceAxisDir)
                            '3. Now, the New center which Is going to be within face bounds can be computed as: 
                            '     C_new = C + (D times A^)

                            adNewFaceCenter(0) = adFaceCenterOnAxis(0) + (dLengthOfVec * adFaceAxisDir(0))
                            adNewFaceCenter(1) = adFaceCenterOnAxis(1) + (dLengthOfVec * adFaceAxisDir(1))
                            adNewFaceCenter(2) = adFaceCenterOnAxis(2) + (dLengthOfVec * adFaceAxisDir(2))
                            FnGetFaceCenterOfConicalFace = adNewFaceCenter
                            Exit Function
                        End If
                    End If
                End If

            End If
        End If
        FnGetFaceCenterOfConicalFace = Nothing
    End Function

    Function FnGetGeometricMeanOfAllEdgeCenter(aoAllCircularEdge() As Edge) As Double()
        Dim adEdgeCenter As Double() = Nothing
        Dim adSumOfEdgeCenter(2) As Double
        Dim iNumOfEdges As Integer = 0
        Dim adMeanEdgeCenter(2) As Double

        If Not aoAllCircularEdge Is Nothing Then
            For Each objEdge As Edge In aoAllCircularEdge
                adEdgeCenter = FnGetArcInfo(objEdge.Tag).center
                adSumOfEdgeCenter(0) = adSumOfEdgeCenter(0) + adEdgeCenter(0)
                adSumOfEdgeCenter(1) = adSumOfEdgeCenter(1) + adEdgeCenter(1)
                adSumOfEdgeCenter(2) = adSumOfEdgeCenter(2) + adEdgeCenter(2)
                iNumOfEdges = iNumOfEdges + 1
            Next
            adMeanEdgeCenter(0) = adSumOfEdgeCenter(0) / iNumOfEdges
            adMeanEdgeCenter(1) = adSumOfEdgeCenter(1) / iNumOfEdges
            adMeanEdgeCenter(2) = adSumOfEdgeCenter(2) / iNumOfEdges
            FnGetGeometricMeanOfAllEdgeCenter = adMeanEdgeCenter
            Exit Function
        End If
        FnGetGeometricMeanOfAllEdgeCenter = Nothing
    End Function
End Module
