Imports System.Math

Module MathAlgo

    '****************************************************************************************************************************************************
    'Description        : Given two vectors' DCs, determine if they are orthogonal.
    'Function Name      : fnCheckOrthogalityOfTwoVectors
    'Input Parameters   : adVector1 – DCs of Vector 1
    '                     adVector2 - DCs of Vector 2
    'Output Parameters  : Boolean Value - True if orthogonal.
    '****************************************************************************************************************************************************
    Function FnCheckOrthogalityOfTwoVectors(ByVal adVector1() As Double, ByVal adVector2() As Double) As Boolean

        Dim i As Integer = Nothing
        Dim dDotProdSum As Double = Nothing
        Dim dTolerance As Double = Nothing

        dDotProdSum = 0
        'dTolerance = 0.0349
        'dTolerance rounding it to 0.1degree
        '=COS(89.5*PI()/180)
        dTolerance = 0.001745
        'dTolerance = 0.0
        'Compute Dot Product.
        For i = LBound(adVector1) To UBound(adVector1)
            dDotProdSum = dDotProdSum + (adVector1(i) * adVector2(i))
        Next

        'Check if Dot Product is close to Zero. If yes, Orthogonal.
        If (Abs(dDotProdSum)) < dTolerance Then
            'If Abs(dDotProdSum) < dTolerance Then
            FnCheckOrthogalityOfTwoVectors = True
        Else
            FnCheckOrthogalityOfTwoVectors = False
        End If

    End Function

    '****************************************************************************************************************************************************
    'Description        : Given two orthogonal vectors' DCs, compute rotation matrix.
    'Function Name      : fnGetRotationMatrixGivenTwoOrthogonalVectors
    'Input Parameters   : adVector1 – DCs of Vector 1
    '                     adVector2 - DCs of Vector 2
    'Output Parameters  : Double Array containing 9 elements.
    '****************************************************************************************************************************************************
    Function FnGetRotationMatrixOfGivenTwoOrthogonalVectors(ByVal adVector1() As Double, ByVal adVector2() As Double) As Double()

        Dim i As Integer = Nothing
        Dim adRotnMatrix(8) As Double
        Dim adCrossProdVector() As Double = Nothing

        'Compute Cross Product Vector.
        adCrossProdVector = FnGetCrossProductVector(adVector1, adVector2)

        'Compse Rotation Matrix.
        For i = LBound(adVector1) To UBound(adVector1)
            adRotnMatrix(i) = adVector1(i)
            adRotnMatrix(i + 3) = adVector2(i)
            adRotnMatrix(i + 6) = adCrossProdVector(i)
        Next

        FnGetRotationMatrixOfGivenTwoOrthogonalVectors = adRotnMatrix

    End Function

    '****************************************************************************************************************************************************
    'Description        : Given two vectors' DCs, compute their cross product vector (orthogonal vector to both the vectors)
    'Function Name      : fnGetCrossProductVector
    'Input Parameters   : adVector1 – Vector 1
    '                   : adVector2 - Vector 2
    'Output Parameters  : Double Array containing cross product vector
    '****************************************************************************************************************************************************
    Function FnGetCrossProductVector(ByVal adVector1() As Double, ByVal adVector2() As Double) As Double()

        Dim adCrossProdVector(2) As Double

        'Cross Product Computation
        adCrossProdVector(0) = (adVector1(1) * adVector2(2)) - (adVector1(2) * adVector2(1))
        adCrossProdVector(1) = (adVector1(2) * adVector2(0)) - (adVector1(0) * adVector2(2))
        adCrossProdVector(2) = (adVector1(0) * adVector2(1)) - (adVector1(1) * adVector2(0))

        FnGetCrossProductVector = adCrossProdVector

    End Function

    '****************************************************************************************************************************************************
    'Description        : Compute Bounding Box Volume.
    'Function Name      : fnGetBoundingBoxVolume
    'Input Parameters   : dLength
    '                     dBreadth
    '                     dHeight
    'Output Parameters  : Double Value that containing Volume
    '****************************************************************************************************************************************************
    Function FnGetBoundingBoxVolume(ByVal dLength As Double, ByVal dBreadth As Double, ByVal dHeight As Double) As Double

        Dim dVolume As Double = Nothing
        If dVolume = Nothing Then
            dVolume = dLength * dBreadth * dHeight
            FnGetBoundingBoxVolume = dVolume
        End If

    End Function

    '****************************************************************************************************************************************************
    'Description        : Check if Rotation matrix is valid
    '                     Check 1: No direction must have zero dcs.
    '                     Check 2: Each direction-dcs must be a unit vector.
    '                     Check 3: Each direction-dcs must be orthogonal to all other direction-dcs.
    'Function Name      : fnCheckIfRotationMatrixIsConsistent
    'Input Parameters   : adRotationMatrix - Row vector of 9 elements
    'Output Parameters  : Boolean Value - True if Rotn Matrix is Valid.

    '****************************************************************************************************************************************************
    Function FnCheckIfRotationMatrixIsConsistent(ByVal adRotationMatrix() As Double) As Boolean

        Dim i As Integer = Nothing
        Dim j As Integer = Nothing

        Dim bPerpCheck1 As Boolean
        Dim bPerpCheck2 As Boolean
        Dim bPerpCheck3 As Boolean

        Dim dDotProdVal As Double
        Dim dTolerance As Double
        Dim dDotProdTolerance As Double

        Dim dMagnitude1 As Double
        Dim dMagnitude2 As Double
        Dim dMagnitude3 As Double

        Dim adXAxisDCs(2) As Double
        Dim adYAxisDCs(2) As Double
        Dim adZAxisDCs(2) As Double

        dTolerance = 0.001
        'dTolerance rounding it to 0.5degree
        '=COS(89.5*PI()/180)
        'dDotProdTolerance = 0.00873
        dDotProdTolerance = 0.0349
        'dDotProdTolerance = 0.001745


        'Check - 1: Direction DCs must NOT be Zeros.
        j = LBound(adRotationMatrix)
        For i = 1 To 3
            If ((Abs(adRotationMatrix(j)) < dTolerance) And _
                (Abs(adRotationMatrix(j + 1)) < dTolerance) And _
                (Abs(adRotationMatrix(j + 2)) < dTolerance)) Then
                fnCheckIfRotationMatrixIsConsistent = False
                Exit Function
            End If
            j = j + 3
        Next

        'Extract individual Direction DCs.
        For i = LBound(adXAxisDCs) To UBound(adXAxisDCs)
            adXAxisDCs(i) = adRotationMatrix(i)
            adYAxisDCs(i) = adRotationMatrix(i + 3)
            adZAxisDCs(i) = adRotationMatrix(i + 6)
        Next

        'Check - 2: Each Direction DCs must be a unit vector.
        dMagnitude1 = FnComputeVectorLength(adXAxisDCs(0), adXAxisDCs(1), adXAxisDCs(2))
        If (Not (Abs(dMagnitude1 - 1) < 0.001)) Then
            FnCheckIfRotationMatrixIsConsistent = False
            Exit Function
        End If
        dMagnitude2 = FnComputeVectorLength(adYAxisDCs(0), adYAxisDCs(1), adYAxisDCs(2))
        If (Not (Abs(dMagnitude2 - 1) < 0.001)) Then
            FnCheckIfRotationMatrixIsConsistent = False
            Exit Function
        End If
        dMagnitude3 = FnComputeVectorLength(adZAxisDCs(0), adZAxisDCs(1), adZAxisDCs(2))
        If (Not (Abs(dMagnitude3 - 1) < 0.001)) Then
            FnCheckIfRotationMatrixIsConsistent = False
            Exit Function
        End If

        'Check 3: Perpendicularity Checks, i.e, Each direction dcs must be orthogonal to every other direction dcs.
        dDotProdVal = (adXAxisDCs(0) * adYAxisDCs(0)) + (adXAxisDCs(1) * adYAxisDCs(1)) + (adXAxisDCs(2) * adYAxisDCs(2))
        If (Abs(dDotProdVal) < dDotProdTolerance) Then
            bPerpCheck1 = True
        End If
        dDotProdVal = (adYAxisDCs(0) * adZAxisDCs(0)) + (adYAxisDCs(1) * adZAxisDCs(1)) + (adYAxisDCs(2) * adZAxisDCs(2))
        If (Abs(dDotProdVal) < dDotProdTolerance) Then
            bPerpCheck2 = True
        End If
        dDotProdVal = (adZAxisDCs(0) * adXAxisDCs(0)) + (adZAxisDCs(1) * adXAxisDCs(1)) + (adZAxisDCs(2) * adXAxisDCs(2))
        If (Abs(dDotProdVal) < dDotProdTolerance) Then
            bPerpCheck3 = True
        End If

        If (bPerpCheck1 And bPerpCheck2 And bPerpCheck3) Then
            FnCheckIfRotationMatrixIsConsistent = True
            Exit Function
        End If

        FnCheckIfRotationMatrixIsConsistent = False

    End Function

    '****************************************************************************************************************************************************
    'Description        : Compute vector length.
    'Function Name      : fnComputeVectorLength
    'Input Parameters   : dVectComp_X - x component of vector
    '                     dVectComp_Y - y component of vector
    '                     dVectComp_Z - z component of vector
    'Output Parameters  : Double value that contains vector length
    '****************************************************************************************************************************************************
    Function FnComputeVectorLength(ByVal dVectCompX As Double, ByVal dVectCompY As Double, ByVal dVectCompZ As Double) As Double

        'Square root of Sum of squares of individual components of vector.
        'fnComputeVectorLength = WorksheetFunction.Power(WorksheetFunction.SumSq(dVectComp_X, dVectComp_Y, dVectComp_Z), 0.5)
        FnComputeVectorLength = Math.Sqrt((Math.Pow(dVectCompX, 2) + Math.Pow(dVectCompY, 2) + Math.Pow(dVectCompZ, 2)))

    End Function

    Public Function FnConvertDecimalToFraction(ByVal dDecimal As Double) As String
        Dim df As Double
        Dim lUpperPart As Long
        Dim lLowerPart As Long

        lUpperPart = 1
        lLowerPart = 1

        df = lUpperPart / lLowerPart
        While (df <> dDecimal)
            If (df < dDecimal) Then
                lUpperPart = lUpperPart + 1
            Else
                lLowerPart = lLowerPart + 1
                lUpperPart = dDecimal * lLowerPart
            End If
            df = lUpperPart / lLowerPart
        End While
        FnConvertDecimalToFraction = CStr(lUpperPart) & "/" & CStr(lLowerPart)
    End Function

    'Function GetFraction(ByVal d As Double) As String
    '    ' Get the initial denominator: 1 * (10 ^ decimal portion length)
    '    Dim Denom As Int32 = CInt(1 * (10 ^ tb1.Text.Split("."c)(1).Length))

    '    ' Get the initial numerator: integer portion of the number
    '    Dim Numer As Int32 = CInt(tb1.Text.Split("."c)(1))

    '    ' Use the Euclidean algorithm to find the gcd
    '    Dim a As Int32 = Numer
    '    Dim b As Int32 = Denom
    '    Dim t As Int32 = 0 ' t is a value holder

    '    ' Euclidean algorithm
    '    While b <> 0
    '        t = b
    '        b = a Mod b
    '        a = t
    '    End While

    '    ' Return our answer
    '    Return CInt(d) & " " & (Numer / a) & "/" & (Denom / a)
    'End Function

End Module
