Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Text.RegularExpressions
Imports NXOpen.Assemblies
Imports NXOpenUI
Imports NXOpen

Module OEM
    'Code added Mar-23-2020
    'Gestamp configuration
    'Public Const GESTAMP_TOOL_REGEX As String = "^(?'PROGNO'[a-zA-Z0-9]{3,4})-(?'CELLNO'[a-zA-Z0-9]{2,3})-(?'OPERATIONNO'[a-zA-Z0-9]{3})$"
    'Public Const GESTAMP_UNIT_REGEX As String = "^(?'PROGNO'[a-zA-Z0-9]{3,4})-(?'CELLNO'[a-zA-Z0-9]{2,3})-(?'OPERATIONNO'[a-zA-Z0-9]{3})?-(?'UNITNO'\d{3})$"
    'Public Const GESTAMP_MAKEDETAIL_REGEX As String = "^(?'PROGNO'[a-zA-Z0-9]{3,4})-(?'CELLNO'[a-zA-Z0-9]{2,3})-(?'OPERATIONNO'[a-zA-Z0-9]{3})?-(?'UNITNO'\d{3})-(?'DETAILNO'\d{3}?)$"
    'COde modified on May-19-2020
    'New Regex added with the Team Center Extension
    Public Const GESTAMP_TOOL_REGEX As String = "^(?'PROGNO'[a-zA-Z0-9]{3,4})-(?'CELLNO'[a-zA-Z0-9]{2,3})-(?'OPERATIONNO'[a-zA-Z0-9]{2,4})-(?'STATIONNO'[a-zA-Z0-9]{2,4})(?'TC'[_;a-zA-Z0-9]{0,15})$"
    Public Const GESTAMP_UNIT_REGEX As String = "^(?'PROGNO'[a-zA-Z0-9]{3,4})-(?'CELLNO'[a-zA-Z0-9]{2,3})-(?'OPERATIONNO'[a-zA-Z0-9]{2,4})-(?'STATIONNO'[a-zA-Z0-9]{2,4})-?(?'UNITNO'\d{3})(?'TC'[_;a-zA-Z0-9]{0,15})$"
    Public Const GESTAMP_MAKEDETAIL_REGEX As String = "^(?'PROGNO'[a-zA-Z0-9]{3,4})-(?'CELLNO'[a-zA-Z0-9]{2,3})-(?'OPERATIONNO'[a-zA-Z0-9]{2,4})?-(?'STATIONNO'[a-zA-Z0-9]{2,4})-?(?'UNITNO'\d{3})-(?'DETAILNO'\d{3}?)(?'TC'[_;a-zA-Z0-9]{0,15})$"


    'Check if this is a tool
    Public Function FnCheckIfThisIsATool(sFileName As String) As Boolean

        Dim sAssyPattern As String = ""
        Dim objRegex As Regex = Nothing
        Dim objRegexMatch As Match = Nothing
        If _sOemName = GM_OEM_NAME Then
            'sAssyPattern = "[a-z-[abcdefghijnopqstuvwxyz]]..\d\d\d\d\d[a-z-[abcdefghijkmnopqtuvwxyz]].f\d\d.\d\d\d\d"
            'sAssyPattern = "^[mlkrMLKR]+[a-zA-Z]+[a-zA-Z]+\d\d\d\d\d+[slrSLR]+.+[fF]+01.00+[1-9]+[0]$"
            'code modified on Jan-22-2020
            'Bug fixed in the regex. When we are checking for the dot, use the dot within square bracket [.]
            sAssyPattern = "^[mlkrMLKR]+[a-zA-Z]+[a-zA-Z]+\d\d\d\d\d+[slrSLR]+[.]+[fF]+01[.]00+[1-9]+[0]$"
        ElseIf _sOemName = CHRYSLER_OEM_NAME Then
            sAssyPattern = "^.*-[Aa][Ss][Ss][Yy]$"
        ElseIf _sOemName = DAIMLER_OEM_NAME Then
            If _sDivision = CAR_DIVISION Then
                '1. Tool Name should be 13Characters
                '2. First Character should be Alphabet
                '3. Character from 2 to 13 should be numeric
                'sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d$"
                sAssyPattern = "^F5\d{11}$"
            ElseIf _sDivision = TRUCK_DIVISION Then
                '1. Tool Name should be 13Characters
                '2. First Character should be Alphabet
                '3. Character from 2 to 13 should be numeric
                sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d$"
            End If
        ElseIf (_sOemName = GESTAMP_OEM_NAME) Then
            sAssyPattern = GESTAMP_TOOL_REGEX
        End If
        ' Instantiate the regular expression object.
        objRegex = New Regex(sAssyPattern, RegexOptions.IgnoreCase)

        ' Match the regular expression pattern against a text string.
        objRegexMatch = objRegex.Match(sFileName)

        If objRegexMatch.Success() Then
            FnCheckIfThisIsATool = True
            Exit Function
        End If
        FnCheckIfThisIsATool = False
    End Function

    'Check if this is a Unit
    Public Function FnCheckIfThisIsAUnit(sUnitName As String) As Boolean

        Dim sAssyPattern As String = ""
        Dim objRegex As Regex = Nothing
        Dim objRegexMatch As Match = Nothing

        If _sOemName = GM_OEM_NAME Then
            'sAssyPattern = "[a-z-[abcdefghijnopqstuvwxyz]]..\d\d\d\d\d[a-z-[abcdefghijkmnopqtuvwxyz]] \x5F\d\d\d.f\d\d.\d\d\d\d"
            'sAssyPattern = "^[mlkrMLKR]+[a-zA-Z]+[a-zA-Z]+\d\d\d\d\d+[slrSLR]+_+\d\d\d+[slrSLR].+[fF]+01.00+[1-9]+[09]$"
            'code modified on Jan-22-2020
            'Bug fixed in the regex. When we are checking for the dot, use the dot within square bracket [.]
            sAssyPattern = "^[mlkrMLKR]+[a-zA-Z]+[a-zA-Z]+\d\d\d\d\d+[slrSLR]+_+\d\d\d+[slrSLR][.]+[fF]+01[.]00+[1-9]+[09]$"
        ElseIf _sOemName = CHRYSLER_OEM_NAME Then
            'sAssyPattern = "WACUN\x2D\d\d\d\d\d\d\d\d\d\d\x2DS"
            If Not FnCheckIfThisIsAMakeDetailBasedOnName(sUnitName) Then
                sAssyPattern = "^.*-[Uu]\d{2}-.*$"
            Else
                FnCheckIfThisIsAUnit = False
                Exit Function
            End If
        ElseIf _sOemName = DAIMLER_OEM_NAME Then
            If _sDivision = CAR_DIVISION Then
                'sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d\d\d\d\d$"
                sAssyPattern = "^F5\d{15}$"
            ElseIf _sDivision = TRUCK_DIVISION Then
                '1. Unit Name will have 17 characters
                '2. Character 1 to 13 should be a tool name
                '3. Character 14 should be a numberic digit 6
                '4. Character 15 and 16 will always be a digit and will indicate unit number
                '5. Character 17 will be zero (0) always
                sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d6\d\d0$"
            End If
        ElseIf (_sOemName = GESTAMP_OEM_NAME) Then
            sAssyPattern = GESTAMP_UNIT_REGEX
        End If
        ' Instantiate the regular expression object.
        objRegex = New Regex(sAssyPattern, RegexOptions.IgnoreCase)

        ' Match the regular expression pattern against a text string.
        objRegexMatch = objRegex.Match(sUnitName)

        If objRegexMatch.Success() Then
            FnCheckIfThisIsAUnit = True
            Exit Function
        End If
        FnCheckIfThisIsAUnit = False
    End Function

    'Function to check if this is a make detail
    Function FnCheckIfThisIsAMakeDetailBasedOnName(sCompName As String) As Boolean
        Dim sAssyPattern As String = ""
        Dim objRegex As Regex = Nothing
        Dim objRegexMatch As Match = Nothing

        If sCompName <> "" Then
            If _sOemName = GM_OEM_NAME Then
                'sAssyPattern = "..\d\d\d\d\d\d\d[a-z-[abcdefghijkmnopqtuvwxyz]]f\d\d.\d\d\d0"
                'sAssyPattern = "^[mlkrMLKR]+[bBcC]+\d\d\d\d\d\d+[slrSLR]+.+[fF]+01.00+[1-9]+[09]$"
                'code modified on Jan-22-2020
                'Bug fixed in the regex. When we are checking for the dot, use the dot within square bracket [.]
                'sAssyPattern = "^[mlkrMLKR]+[bBcC]+\d\d\d\d\d\d+[slrSLR]+[.]+[fF]+01[.]00+[1-9]+[09]$"
                sAssyPattern = "^[mlkrMLKR]+[a-zA-Z]+\d\d\d\d\d\d+[slrSLR]+[.]+[fF]+01[.]00+[1-9]+[09]$"

            ElseIf _sOemName = CHRYSLER_OEM_NAME Then
                'sAssyPattern = "WACDU-\d\d\d\d\d\d\d\d\d\d\-"
                sAssyPattern = "^.*-\d{4,5}-\d{3}$"
            ElseIf _sOemName = DAIMLER_OEM_NAME Then
                'Check if it is a Weldment
                If FnCheckIfThisIsAWeldment(sCompName) Then
                    FnCheckIfThisIsAMakeDetailBasedOnName = True
                    Exit Function
                    'Check if it is a component
                Else
                    If _sDivision = CAR_DIVISION Then
                        '1. Make detail includes component and weldment
                        '2. Make Detail name will be 21 character
                        '3. Character 1 to 13 should be a tool name
                        '4. Character 18 should be 0
                        'sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d\d\d\d\d0\d\d\d$"
                        sAssyPattern = "^F5\d{15}0\d{3}$"
                    ElseIf _sDivision = TRUCK_DIVISION Then
                        '1. Make detail includes component and weldment
                        '2. Make Detail name will be 17 character
                        '3. Character 1 to 13 should be a tool name
                        '4. Character 14 should not be 6. (Because 6 indicates it as a Unit)
                        '5. Character 15 and 16 will always be a digit
                        '6. Character 17 should be an odd number
                        '7. Make Detail name should not have "_DWG"
                        sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d[0-9-[6]]\d\d\d$"
                    End If
                End If
            ElseIf (_sOemName = GESTAMP_OEM_NAME) Then
                sAssyPattern = GESTAMP_MAKEDETAIL_REGEX
            End If

            ' Instantiate the regular expression object.
            objRegex = New Regex(sAssyPattern, RegexOptions.IgnoreCase)

            ' Match the regular expression pattern against a text string.
            objRegexMatch = objRegex.Match(sCompName)

            If objRegexMatch.Success() Then
                FnCheckIfThisIsAMakeDetailBasedOnName = True
                Exit Function
            End If

        End If
        FnCheckIfThisIsAMakeDetailBasedOnName = False
    End Function

    'Check if this is a welded assembly
    Public Function FnCheckIfThisIsAWeldment(sWeldmentName As String) As Boolean


        Dim sAssyPattern As String = ""
        Dim objRegex As Regex = Nothing
        Dim objRegexMatch As Match = Nothing
        If _sOemName = DAIMLER_OEM_NAME Then
            If _sDivision = CAR_DIVISION Then
                '1. Weldment Name will have 21 characters
                '2. Character 1 to 13 should be a tool name
                '3. Character 14,15 should be sub tool
                '4. Character 16 and 17 will always be a digit - unit number
                '5. Character 18 will be 1

                'sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d\d\d\d\d1\d\d\d$"
                sAssyPattern = "^F5\d{15}1\d{3}$"
            ElseIf _sDivision = TRUCK_DIVISION Then
                '1. Weldment Name will have 17 characters
                '2. Character 1 to 13 should be a tool name
                '3. Character 14 should be a numberic digit 5
                '4. Character 15 and 16 will always be a digit 
                sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d5\d\d\d$"
            End If

            ' Instantiate the regular expression object.
            objRegex = New Regex(sAssyPattern, RegexOptions.IgnoreCase)

            ' Match the regular expression pattern against a text string.
            objRegexMatch = objRegex.Match(sWeldmentName)

            If objRegexMatch.Success() Then
                FnCheckIfThisIsAWeldment = True
                Exit Function
            End If
        End If
        FnCheckIfThisIsAWeldment = False
    End Function

    'Fucntion to check if the component is Child component container in weldment
    Public Function FnCheckIfThisIsAChildCompContainerInWeldment(sCompName As String) As Boolean

        Dim sAssyPattern As String = ""
        Dim objRegex As Regex = Nothing
        Dim objRegexMatch As Match = Nothing

        If _sDivision = CAR_DIVISION Then
            'sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d\d\d\d\d2\d\d\d$"
            sAssyPattern = "^F5\d{15}2\d{3}$"
        ElseIf _sDivision = TRUCK_DIVISION Then
            sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d\d\d\d\d$"
        End If

        ' Instantiate the regular expression object.
        objRegex = New Regex(sAssyPattern, RegexOptions.IgnoreCase)

        ' Match the regular expression pattern against a text string.
        objRegexMatch = objRegex.Match(sCompName)

        If objRegexMatch.Success() Then
            FnCheckIfThisIsAChildCompContainerInWeldment = True
            Exit Function
        End If
        FnCheckIfThisIsAChildCompContainerInWeldment = False
    End Function
    'Code modified on Nov-07-2018
    'Function modified to work for Daimler and FIAT OEM, where there can be Child of a welded assembly be a sub components.
    'Fucntion to check if the component is Child component in weldment
    Public Function FnCheckIfThisIsAChildCompInWeldment(objComp As Component, sOEMName As String) As Boolean

        Dim sAssyPattern As String = ""
        Dim objRegex As Regex = Nothing
        Dim objRegexMatch As Match = Nothing
        Dim sCompName As String = ""
        Dim objPart As Part = Nothing
        Dim sPartType As String = ""

        If sOEMName = DAIMLER_OEM_NAME Then
            sCompName = objComp.DisplayName.ToUpper
            If _sDivision = CAR_DIVISION Then
                ' sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d\d\d\d\d2\d\d\d_GEO"
                sAssyPattern = "^F5\d{15}2\d{3}_GEO$"
            ElseIf _sDivision = TRUCK_DIVISION Then
                sAssyPattern = "^F5\d\d\d\d\d\d\d\d\d\d\d\d\d\d\d$"
            End If

            ' Instantiate the regular expression object.
            objRegex = New Regex(sAssyPattern, RegexOptions.IgnoreCase)

            ' Match the regular expression pattern against a text string.
            objRegexMatch = objRegex.Match(sCompName)

            If objRegexMatch.Success() Then
                FnCheckIfThisIsAChildCompInWeldment = True
                Exit Function
            End If
        ElseIf sOEMName = FIAT_OEM_NAME Then
            objPart = FnGetPartFromComponent(objComp)
            If Not objPart Is Nothing Then
                FnLoadPartFully(objPart)
                sPartType = FnGetStringUserAttribute(objPart, B_PART_TYPE)
                If sPartType <> "" Then
                    If sPartType.ToUpper = WELDED_CHILD Then
                        FnCheckIfThisIsAChildCompInWeldment = True
                        Exit Function
                    End If
                End If
            End If
        End If

        FnCheckIfThisIsAChildCompInWeldment = False
    End Function
    'Function to identify feature group name
    Function FnGetFeatureGroupName(sDivision As String) As String
        Dim sFeatureGroupName As String = ""

        If sDivision = CAR_DIVISION Then
            sFeatureGroupName = CAR_DIVISION_FEATURE_GROUP
        ElseIf sDivision = TRUCK_DIVISION Then
            sFeatureGroupName = TRUCK_DIVISION_FEATURE_GROUP
        End If
        FnGetFeatureGroupName = sFeatureGroupName
    End Function

    'Code added Nov-07-2018
    'Function to check if the \given component is a make detail component.
    Function FnCheckIfThisIsAMakeDetailBasedOnAttribute(objPart As Part) As Boolean
        Dim sPartType As String = ""

        If Not objPart Is Nothing Then
            FnLoadPartFully(objPart)
            sPartType = FnGetStringUserAttribute(objPart, B_PART_TYPE)
            If sPartType <> "" Then
                If (sPartType.ToUpper = WELDED_ASS) Or (sPartType.ToUpper = NON_WELDED_ASS) Then
                    FnCheckIfThisIsAMakeDetailBasedOnAttribute = True
                    Exit Function
                End If
            End If
        End If
        FnCheckIfThisIsAMakeDetailBasedOnAttribute = False
    End Function

    'Code added Nov-08-2018
    Function FnCheckIfThisFileIsFIATMakeDetail(sPartName As String) As Boolean
        Dim sPartNumber As String = 0
        Dim asPartNumber() As String = Nothing

        If sPartName <> "" Then
            'Example 281702Z_M-020S16055ZR-020D16045ZR-001-A00
            'Code modified on Dec-13-2018
            'Nomenclature can have more than 41 character.
            'example 281702Z_M-050S30045AF-050D40045AF-001-A00-281
            If Split(sPartName.ToUpper, ".PRT")(0).Length >= 41 Then

                asPartNumber = Split(sPartName, "-")
                If Not asPartNumber Is Nothing Then
                    If asPartNumber.Length >= 5 Then
                        sPartNumber = asPartNumber(3)
                        If sPartNumber.ToUpper <> "SYM" And sPartNumber.ToUpper <> "ASS" Then
                            Try
                                'Part Number must be 001 to 999 and not 800
                                '800 will be Open Position, do not consider it
                                If CInt(sPartNumber) >= "001" And CInt(sPartNumber) <= 999 And CInt(sPartNumber) <> 800 Then
                                    FnCheckIfThisFileIsFIATMakeDetail = True
                                    Exit Function
                                End If
                            Catch ex As Exception
                                FnCheckIfThisFileIsFIATMakeDetail = False
                            End Try
                        End If
                    End If
                End If
            End If
        End If
        FnCheckIfThisFileIsFIATMakeDetail = False
    End Function
End Module
