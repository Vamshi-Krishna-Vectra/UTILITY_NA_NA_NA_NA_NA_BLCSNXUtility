Imports System
Imports NXOpen.Assemblies
Imports NXOpen.Features
Imports NXOpen
Imports NXOpen.Utilities
Imports NXOpenUI
Imports NXOpen.UF
Imports System.IO
Imports NXOpen.Drafting
Imports NXOpen.Drawings
Imports System.Math

Module NXWrapper
    Private _objOrientation As Matrix3x3
    Private _workPart As Part
    Private _displayPart As Part
    Private _objComponentPositioner As Positioning.ComponentPositioner
    Private _objComponentNetwork As Positioning.ComponentNetwork
    Private _objConstraintCollection As ConstraintCollection
    Private _markId2 As Session.UndoMarkId
    Private _NXPartfolderPath As String
    Private _NxComponentConstraints(100) As NXObject
    Private _ConstraintCounter As Integer = 0
    Private _iCountSheetMetalParts As Integer = 0
    Private _asSheetMetalParts() As String

    Public Function FnOpenNxPart(ByVal sFilePath As String) As BasePart
        Dim partLoadStatus As PartLoadStatus = Nothing
        'Dim iNumOfUnLoadedpart As Integer
        'Dim iGetStatus As Integer
        'Dim sDesc As String = ""
        FnOpenNxPart = FnGetPartCollectioninSession.OpenBaseDisplay(sFilePath, partLoadStatus)
        'iNumOfUnLoadedpart = partLoadStatus.NumberUnloadedParts
        'iGetStatus = partLoadStatus.GetStatus(iNumOfUnLoadedpart)
        'sDesc = partLoadStatus.GetStatusDescription(iNumOfUnLoadedpart)
        'sWriteToLogFile("iNumOfUnLoafpart " & iNumOfUnLoadedpart)
        'sWriteToLogFile("Status " & iGetStatus)
        'sWriteToLogFile("DEsc " & sDesc)
        sReportPartUnloadedStatus(partLoadStatus)
        partLoadStatus.Dispose()
    End Function
    Sub sReportPartUnloadedStatus(objPartLoadStatus As PartLoadStatus)
        Dim iNumOfUnLoadedpart As Integer = -1
        Dim iGetStatus As Integer = -1
        Dim sDesc As String = ""
        Dim ex As NXException = Nothing

        iNumOfUnLoadedpart = objPartLoadStatus.NumberUnloadedParts
        sWriteToLogFile("Nuber of UnLoaded part : " & iNumOfUnLoadedpart)
        If iNumOfUnLoadedpart = 0 Then
            sWriteToLogFile("Part Loaded Successfully")
        Else
            For iindex As Integer = 0 To iNumOfUnLoadedpart - 1
                iGetStatus = objPartLoadStatus.GetStatus(iindex)
                sDesc = objPartLoadStatus.GetStatusDescription(iindex)
                sWriteToLogFile("iNumOfUnLoafpart " & iindex)
                sWriteToLogFile("Status " & iGetStatus)
                sWriteToLogFile("DEsc " & sDesc)
                ex = NXException.Create(objPartLoadStatus.GetStatus(iindex))
                sWriteToLogFile("problem")
                sWriteToLogFile(objPartLoadStatus.GetPartName(iindex))
                sWriteToLogFile(ex.Message)
            Next
        End If

    End Sub
    Public Sub SetDisplayandWorkPart(ByVal objPart As Part)
        Dim partLoadStatus1 As PartLoadStatus = Nothing
        Dim status1 As PartCollection.SdpsStatus
        status1 = FnGetNxSession.Parts.SetDisplay(objPart, True, True, partLoadStatus1)

        _workPart = FnGetNxSession.Parts.Work
        _displayPart = FnGetNxSession.Parts.Display
        partLoadStatus1.Dispose()
    End Sub

    Sub SSetWorkComponent(ByVal objComp As Component)
        Dim partLoadStatus2 As PartLoadStatus = Nothing
        FnGetNxSession.Parts.SetWorkComponent(objComp, partLoadStatus2)

        _workPart = FnGetNxSession.Parts.Work
        partLoadStatus2.Dispose()
    End Sub
    Sub SetNXPartFolderPath(ByVal sPath As String)
        _NXPartfolderPath = sPath
    End Sub
    Public Function FnGetNxSession() As Session
        FnGetNxSession = NXOpen.Session.GetSession()
    End Function

    Public Function FnGetPartCollectioninSession() As PartCollection
        FnGetPartCollectioninSession = FnGetNxSession.Parts
    End Function

    Public Function FnGetCurrentWorkPartInSession() As Part
        FnGetCurrentWorkPartInSession = FnGetPartCollectioninSession.Work
    End Function

    Public Function FnGetPartObjectByName(ByVal sPartName As String) As Part
        Try
            FnGetPartObjectByName = CType(FnGetPartCollectioninSession.FindObject(sPartName), Part)
        Catch ex As Exception
            FnGetPartObjectByName = Nothing
        End Try
    End Function

    Public Sub SSetPartOrientation(ByVal xx As Double, ByVal xy As Double, ByVal xz As Double, _
        ByVal yx As Double, ByVal yy As Double, ByVal yz As Double, ByVal zx As Double, ByVal zy As Double, ByVal zz As Double)
        _objOrientation.Xx = xx
        _objOrientation.Xy = xy
        _objOrientation.Xz = xz
        _objOrientation.Yx = yx
        _objOrientation.Yy = yy
        _objOrientation.Yz = yz
        _objOrientation.Zx = zx
        _objOrientation.Zy = zy
        _objOrientation.Zz = zz
    End Sub
    Public Function FnGetPartOrientation() As Matrix3x3
        FnGetPartOrientation = _objOrientation
    End Function


    Public Sub SDisplay(ByVal sPartName As String)
        Dim objPartLoadStatus As PartLoadStatus = Nothing
        Dim objStatus As PartCollection.SdpsStatus = Nothing
        objStatus = FnGetPartCollectioninSession.SetDisplay(FnGetPartObjectByName(sPartName), True, True, objPartLoadStatus)
        objPartLoadStatus.Dispose()
    End Sub

    Public Sub SClosePart(ByVal sPartName As String)
        FnGetPartObjectByName(sPartName).Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.UseResponses, Nothing)
        'FnGetNxSession.Parts.CloseAll(BasePart.CloseModified.CloseModified, Nothing)
    End Sub

    Public Sub SCreateAssembly(ByVal sAssemblyName As String)
        Dim objfileNew As FileNew
        objfileNew = FnGetNxSession.Parts.FileNew()
        objfileNew.TemplateFileName = "assembly-mm-template.prt"
        objfileNew.Application = FileNewApplication.Assemblies
        objfileNew.Units = Part.Units.Millimeters
        objfileNew.NewFileName = My.Application.Info.DirectoryPath + "\" + sAssemblyName + ".prt"
        objfileNew.MakeDisplayedPart = True
        'If objfileNew.Validate() Then
        Dim objnX As NXObject
        objnX = objfileNew.Commit()
        'End If

        '_workPart = FnGetNxSession().Parts.Work
        '_displayPart = FnGetNxSession().Parts.Display
        objfileNew.Destroy()

    End Sub
    
    Public Function FnGetComponentInAssembly(ByVal sComponentName As String, ByVal objComp As Component) As Component
        FnGetComponentInAssembly = Nothing
        Try
            For Each subComp As Component In objComp.GetChildren()
                If subComp.Name.ToUpper() + ".PRT" = sComponentName.ToUpper() Then
                    FnGetComponentInAssembly = subComp
                    Exit Function
                End If
            Next
            FnGetComponentInAssembly = Nothing
        Catch ex As Exception
            FnGetComponentInAssembly = Nothing
        End Try
    End Function
    Public Function FnGetComponentsToBeFixedInSession(ByVal sName As String) As NXObject()
        Dim iCount As Integer
        iCount = _workPart.ComponentAssembly.RootComponent.GetChildren().Length()
        Dim objArrCompInSession(iCount - 2) As NXObject
        Dim index As Integer
        index = 0
        For Each comp As Component In _workPart.ComponentAssembly.RootComponent.GetChildren()
            If comp.Name.ToUpper() <> sName.ToUpper() Then
                objArrCompInSession(index) = comp
                index = index + 1
            End If
        Next
        FnGetComponentsToBeFixedInSession = objArrCompInSession
    End Function
    Public Function GetUnloadOption(ByVal dummy As String) As Integer

        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately

        '----Other unload options-------
        'Unloads the image when the NX session terminates
        'GetUnloadOption = NXOpen.Session.LibraryUnloadOption.AtTermination

        'Unloads the image explicitly, via an unload dialog
        'GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Explicitly
        '-------------------------------

    End Function

    Public Function FnGetFace(ByRef objPart As Part, ByVal faceName As String, ByVal sAttributeType As String) As String
        FnGetFace = Nothing
        'Dim sType As String
        'If sAttributeType = "Face" Then
        '    sType = "Face"
        'End If
        Dim wPart As Part = FnGetNxSession().Parts.Work
        For Each body As Body In objPart.Bodies()
            FnGetNxSession().Parts.SetWork(objPart)
            For Each face As Face In body.GetFaces()
                'On Error Resume Next
                If face.GetUserAttributes(NXObject.AttributeType.String).Length <> 0 Then
                    Try
                        If FnGetStringUserAttribute(face, sAttributeType) = faceName Then
                            'If face.Name = faceName Then
                            Dim sJournalIdName As String
                            sJournalIdName = body.JournalIdentifier + "|" + face.JournalIdentifier
                            FnGetFace = sJournalIdName
                            'FnGetFace = face

                            SSetWorkPart(wPart)
                            Exit Function
                            'End If
                        End If
                    Catch ex As Exception
                    End Try
                End If
            Next
        Next
        SSetWorkPart(wPart)
        Err.Clear()
    End Function
    'Code changed on Mar-14-2019
    Public Function FnGetPartFromComponent(ByRef instance As NXOpen.Assemblies.Component) As NXOpen.Part
        Dim objProtoTypePart As Part = Nothing

        If Not instance Is Nothing Then
            Try
                objProtoTypePart = instance.Prototype
                FnGetPartFromComponent = objProtoTypePart
            Catch ex As Exception
                FnGetPartFromComponent = Nothing
            End Try
        End If

        'Dim tag As NXOpen.Tag = -1
        'FnGetUFSession.Part.AskTagOfDispName(instance.DisplayName, tag)
        'Dim part As Part = FnGetNxSession().Parts.Work
        'For Each prt As NXOpen.Part In FnGetNxSession().Parts
        '    If prt.Tag = tag Then
        '        FnGetPartFromComponent = prt
        '        Exit Function
        '    End If
        'Next
        'FnGetPartFromComponent = Nothing
    End Function
    Public Sub SSetWorkPart(ByVal objPart As BasePart)
        FnGetNxSession().Parts.SetWork(objPart)
    End Sub
    Public Function FnGetFeature(ByVal featName As String, ByRef objPart As Part, ByVal sAttributeType As String) As NXObject
        FnGetFeature = Nothing
        'Dim sType As String
        'If sAttributeType = "Axis" Then
        '    sType = "Axis"
        'End If
        Dim wPart As Part = FnGetNxSession().Parts.Work
        For Each feature As Feature In objPart.Features
            'If feature.FeatureType.ToUpper().Contains(sType.ToUpper()) Then
            FnGetNxSession().Parts.SetWork(objPart)
            For Each entity As NXObject In feature.GetEntities()
                Try
                    If FnGetStringUserAttribute(entity, sAttributeType) = featName Then
                        feature.Highlight()
                        feature.Unhighlight()
                        FnGetFeature = entity
                        SSetWorkPart(wPart)
                        Exit Function
                    End If
                Catch ex As Exception
                End Try
            Next
            'End If
        Next
        SSetWorkPart(wPart)
    End Function
    Public Function FnGetUFSession() As UFSession
        FnGetUFSession = UFSession.GetUFSession()
    End Function

    'Public Sub SCreateAssemlyConstraints(ByVal workPart As Part)
    '    'Dim objComponentPositioner As Positioning.ComponentPositioner
    '    Dim objarrangement As Arrangement
    '    Dim objNetwork As Positioning.Network

    '    _objComponentPositioner = workPart.ComponentAssembly.Positioner
    '    _objComponentPositioner.ClearNetwork()

    '    objarrangement = workPart.ComponentAssembly.ActiveArrangement
    '    'objarrangement = CType(workPart.ComponentAssembly.Arrangements.FindObject("Arrangement 1"), Arrangement)
    '    _objComponentPositioner.PrimaryArrangement = objarrangement
    '    _objComponentPositioner.BeginAssemblyConstraints()

    '    objNetwork = _objComponentPositioner.EstablishNetwork()

    '    _objComponentNetwork = CType(objNetwork, Positioning.ComponentNetwork)
    '    _objComponentNetwork.MoveObjectsState = True
    '    Dim nullAssemblies_Component As Assemblies.Component = Nothing
    '    _objComponentNetwork.DisplayComponent = nullAssemblies_Component
    '    _objComponentNetwork.MoveObjectsState = True
    '    _objComponentNetwork.NetworkArrangementsMode = Positioning.ComponentNetwork.ArrangementsMode.Existing

    '    _markId2 = FnGetNxSession.SetUndoMark(Session.MarkVisibility.Invisible, "MyUndo")

    'End Sub

    'Public Sub CreateConstraint(ByVal sTypeofConstraint As String, ByVal objComp1ToMate As NXObject, Optional ByVal objComp2ToMate As NXObject = Nothing, _
    '    Optional ByVal objComponent1 As NXObject = Nothing, Optional ByVal objComponent2 As NXObject = Nothing)
    '    Dim objConstraint As Positioning.Constraint


    '    'Dim objComp1 As Object
    '    'Dim objComp2 As Object
    '    Dim movableobject(0) As NXObject
    '    Dim objComponentConstraint As Positioning.ComponentConstraint
    '    Dim objConstraintReference1 As Positioning.ConstraintReference
    '    Dim objConstraintReference2 As Positioning.ConstraintReference

    '    Dim markId8 As Session.UndoMarkId
    '    markId8 = FnGetNxSession().SetUndoMark(Session.MarkVisibility.Visible, "Create Constraint")

    '    objConstraint = _objComponentPositioner.CreateConstraint()

    '    objComponentConstraint = CType(objConstraint, Positioning.ComponentConstraint)

    '    'objComponentConstraint.Persistent = True
    '    'objComponentConstraint.Automatic = True

    '    If sTypeofConstraint = "Touch&Align-Axis" Then
    '        'objComponent1 = CType(objComponent1, DatumAxis)
    '        'objComponent2 = CType(objComponent2, DatumAxis)
    '        objComponentConstraint.ConstraintAlignment = Positioning.Constraint.Alignment.CoAlign
    '        objComponentConstraint.Persistent = True
    '        objComponentConstraint.ConstraintType = Positioning.Constraint.Type.Touch
    '        objConstraintReference2 = objComponentConstraint.CreateConstraintReference(objComp2ToMate, objComponent2, False, False, False)
    '        objConstraintReference1 = objComponentConstraint.CreateConstraintReference(objComp1ToMate, objComponent1, False, False, False)
    '        'objConstraintReference1.SetFixHint(True)
    '        'objComponentConstraint.SetAlignmentHint(Positioning.Constraint.Alignment.CoAlign)
    '        objConstraintReference1.SetFixHint(True)
    '        'objConstraintReference2.SetFixHintForUpdate(True)

    '    ElseIf sTypeofConstraint = "Touch&Align-Face" Then
    '        'objComponent1 = CType(objComponent1, Face)
    '        'objComponent2 = CType(objComponent2, Face)
    '        objComponentConstraint.ConstraintAlignment = Positioning.Constraint.Alignment.InferAlign
    '        objComponentConstraint.Persistent = True
    '        objComponentConstraint.ConstraintType = Positioning.Constraint.Type.Touch
    '        objConstraintReference2 = objComponentConstraint.CreateConstraintReference(objComp2ToMate, objComponent2, False, False, False)
    '        objConstraintReference1 = objComponentConstraint.CreateConstraintReference(objComp1ToMate, objComponent1, False, False, False)
    '        'objConstraintReference1.UsePortRotateFlag = True
    '        'objConstraintReference2.UsePortRotateFlag = True
    '        objConstraintReference1.SetFixHint(True)
    '        objComponentConstraint.SetAlignmentHint(Positioning.Constraint.Alignment.ContraAlign)

    '    ElseIf sTypeofConstraint = "Parallel" Then
    '        'objComponent1 = CType(objComponent1, Face)
    '        'objComponent2 = CType(objComponent2, Face)
    '        'objComponentConstraint.ConstraintAlignment = Positioning.Constraint.Alignment.InferAlign
    '        objComponentConstraint.Persistent = True
    '        objComponentConstraint.ConstraintType = Positioning.Constraint.Type.Parallel
    '        objConstraintReference2 = objComponentConstraint.CreateConstraintReference(objComp2ToMate, objComponent2, False, False, False)
    '        objConstraintReference1 = objComponentConstraint.CreateConstraintReference(objComp1ToMate, objComponent1, False, False, False)

    '        objConstraintReference1.SetFixHint(True)
    '        'objComponentConstraint.SetAlignmentHint(Positioning.Constraint.Alignment.CoAlign)
    '    ElseIf sTypeofConstraint = "Fix" Then
    '        'Dim workPart As Part = FnGetNxSession.Parts.Work
    '        'Dim component1 As Assemblies.Component = CType(workPart.ComponentAssembly.RootComponent.FindObject("COMPONENT Clamp 1"), Assemblies.Component)
    '        'objComponentConstraint.ConstraintAlignment = Positioning.Constraint.Alignment.InferAlign
    '        objComponentConstraint.Persistent = True
    '        objComponentConstraint.ConstraintType = Positioning.Constraint.Type.Fix
    '        objConstraintReference1 = objComponentConstraint.CreateConstraintReference(objComp1ToMate, objComp1ToMate, False, False, False)
    '        'objConstraintReference1.SetFixHint(True)
    '        'objComponentConstraint.SetAlignmentHint(Positioning.Constraint.Alignment.CoAlign)
    '        'ElseIf sTypeofConstraint = "Distance" Then
    '        '    objComponentConstraint.Persistent = True
    '        '    objComponentConstraint.ConstraintAlignment = Positioning.Constraint.Alignment.InferAlign
    '        '    objComponentConstraint.ConstraintType = Positioning.Constraint.Type.Distance
    '        '    objConstraintReference2 = objComponentConstraint.CreateConstraintReference(objComp2ToMate, objComponent2, False, False, False)
    '        '    objConstraintReference1 = objComponentConstraint.CreateConstraintReference(objComp1ToMate, objComponent1, False, False, False)

    '        '    objComponentConstraint.SetExpression("37")
    '        '    Dim expression1 As Expression
    '        '    expression1 = objComponentConstraint.Expression
    '        '    objConstraintReference1.SetFixHint(False)
    '        '    objConstraintReference2.SetFixHint(True)
    '        '    expression1.RightHandSide = "0"
    '    End If
    '    ReDim Preserve _NxComponentConstraints(_ConstraintCounter)
    '    _NxComponentConstraints(_ConstraintCounter) = objComponentConstraint
    '    _ConstraintCounter = _ConstraintCounter + 1

    '    _objComponentNetwork.AddConstraint(objConstraint)
    '    'movableobject(0) = objComp2ToMate
    '    'movableobject(1) = objComp1ToMate
    '    'movableobject(1) = objComp1ToMate
    '    'If Not sTypeofConstraint = "Fix" Then
    '    '    _objComponentNetwork.SetMovingGroup(FnGetComponentsToBeFixedInSession(objComp2ToMate.Name))
    '    '    _objComponentNetwork.NonMovingGroupGrounded = True
    '    'End If

    '    _objComponentNetwork.Solve()
    '    '_objComponentNetwork.Solve()

    '    'If objComponentConstraint.GetConstraintStatus() <> Positioning.Constraint.SolverStatus.Solved Then
    '    '    objComponentConstraint.ReverseDirection()
    '    'End If
    '    'If objComponentConstraint.GetConstraintStatus() <> Positioning.Constraint.SolverStatus.Solved Then
    '    '    objComponentConstraint.FlipAlignment()
    '    'End If
    '    '_objComponentNetwork.Solve()

    '    'objComponentConstraint.RememberOnComponent(objComp1ToMate)
    '    'objComponentConstraint.RememberOnComponent(objComp2ToMate)

    '    _objComponentNetwork.ResetDisplay()

    '    _objComponentNetwork.ApplyToModel()

    '    Dim nErr As Integer
    '    nErr = FnGetNxSession.UpdateManager.DoUpdate(markId8)
    '    FnGetNxSession.DeleteUndoMark(markId8, Nothing)

    '    Dim errList As ErrorList
    '    errList = FnGetNxSession().UpdateManager.ErrorList

    '    FnGetNxSession.UpdateManager.DoInterpartUpdate(markId8)
    '    '_objComponentNetwork.EmptyMovingGroup()

    '    '_objComponentNetwork.NonMovingGroupGrounded = False

    '    objConstraint = Nothing
    '    objComponentConstraint = Nothing
    '    objConstraintReference1 = Nothing
    '    objConstraintReference2 = Nothing
    '    movableobject = Nothing

    'End Sub

    'Public Sub SEndAssemblyConstraints()

    '    _objComponentNetwork.ResetDisplay()
    '    _objComponentNetwork.ApplyToModel()
    '    _objComponentPositioner.ClearNetwork()
    '    Dim nErrs1 As Integer
    '    nErrs1 = FnGetNxSession().UpdateManager.AddToDeleteList(_objComponentNetwork)
    '    Dim nErrs2 As Integer
    '    nErrs2 = FnGetNxSession.UpdateManager.DoUpdate(_markId2)
    '    _objComponentPositioner.DeleteNonPersistentConstraints()
    '    Dim nErrs3 As Integer
    '    nErrs3 = FnGetNxSession.UpdateManager.DoUpdate(_markId2)
    '    FnGetNxSession.UpdateManager.DoInterpartUpdate(_markId2)
    '    FnGetNxSession.DeleteUndoMark(_markId2, Nothing)

    '    Dim nullAssemblies_Arrangement As Assemblies.Arrangement = Nothing
    '    _objComponentPositioner.PrimaryArrangement = nullAssemblies_Arrangement

    '    _objComponentPositioner.EndAssemblyConstraints()
    '    _objComponentPositioner = Nothing
    '    _objComponentNetwork = Nothing

    'End Sub

    Function FnGetPartObjfromSession(ByVal sPartName As String) As Part
        FnGetPartObjfromSession = Nothing
        For Each part As Part In FnGetPartCollectioninSession().ToArray()
            If part.Leaf.ToString() + ".prt" = sPartName Then
                FnGetPartObjfromSession = part
                Exit Function
            End If
        Next
    End Function

    Public Sub FnSavePart(ByVal wrkPart As Part)
        'Dim partSaveStatus1 As PartSaveStatus
        'partSaveStatus1 = wrkPart.Save(BasePart.SaveComponents.True, BasePart.CloseAfterSave.False)
        'partSaveStatus1.Dispose()

        Dim objPartSaveStatus As PartSaveStatus = Nothing
        Dim RemUtil As RemoteUtilities = Nothing
        Dim iStatus As Integer = -1
        Dim iNumOfSavedParts As Integer = -1
        Dim iNumOfUnSavedParts As Integer = -1

        RemUtil = NXOpen.RemoteUtilities.GetRemoteUtilities()

        FnGetUFSession.Undo.DeleteAllMarks()
        sWriteToLogFile("Save Loging Starts for :" & wrkPart.Leaf)
        'Check if the part is modified
        'You can accidentally save an unmodified part as well
        If FnGetUFSession.Part.IsModified(wrkPart.Tag) Then

            If Not RemUtil.IsFileWritable(wrkPart.FullPath) Then
                sWriteToLogFile("File is read-only")
                Try
                    RemUtil.SetFileWritable(wrkPart.FullPath, True)
                    sWriteToLogFile(wrkPart.FullPath & " Write Access granted Programmatically")
                Catch ex As Exception
                    sWriteToLogFile("Write Access not allowed")
                End Try
            End If

            If RemUtil.IsFileWritable(wrkPart.FullPath) Then
                objPartSaveStatus = wrkPart.Save(BasePart.SaveComponents.True, BasePart.CloseAfterSave.False)
                sReportPartSaveStatus(objPartSaveStatus)
                objPartSaveStatus.Dispose()
            Else
                sWriteToLogFile(wrkPart.FullPath & " File didnot have write Access to save the part")
            End If
        Else
            sWriteToLogFile(wrkPart.Leaf & " File is not modified and not saved")
        End If
        sWriteToLogFile("Save Loging Ends for :" & wrkPart.Leaf)
    End Sub

    'Code added Dec-06-2018
    Sub sReportPartSaveStatus(objPartSaveStatus As PartSaveStatus)
        Dim iNumOfUnSavedParts As Integer = -1
        Dim ex As NXException = Nothing

        iNumOfUnSavedParts = objPartSaveStatus.NumberUnsavedParts
        sWriteToLogFile("Number of Unsaved part : " & iNumOfUnSavedParts)
        If iNumOfUnSavedParts = 0 Then
            sWriteToLogFile("Part Save Status is successfull")
        Else
            For iIndex As Integer = 0 To iNumOfUnSavedParts - 1
                ex = NXException.Create(objPartSaveStatus.GetStatus(iIndex))
                sWriteToLogFile("Problem with Part Save Status")
                sWriteToLogFile(objPartSaveStatus.GetPart(iIndex).Leaf)
                sWriteToLogFile(ex.Message)
            Next
        End If
    End Sub
    'Public Sub SDeleteConstraints(ByVal wrkPart As Part)
    '    FnGetNxSession().UpdateManager.ClearErrorList()
    '    Dim markId1 As Session.UndoMarkId
    '    markId1 = FnGetNxSession().SetUndoMark(Session.MarkVisibility.Visible, "Delete Constraints")
    '    Dim nErrs1 As Integer
    '    nErrs1 = FnGetNxSession().UpdateManager.AddToDeleteList(_NxComponentConstraints)

    '    Dim nErrs2 As Integer
    '    nErrs2 = FnGetNxSession().UpdateManager.DoUpdate(markId1)

    'End Sub

    'Delete the manually created features after the unit is built
    'Public Sub SDeleteFeatures(ByRef aFeatureList() As NXObject)

    '    Dim iFeatCount As Integer
    '    Dim asPopulatedList() As String
    '    Dim aoDeleteList() As NXObject
    '    Dim iLoopIndex As Integer
    '    Dim iLoopIndex2 As Integer
    '    Dim sPartName As String
    '    Dim iCounter As Integer = 0
    '    Dim iDeleteCounter As Integer = 0
    '    Dim bFound As Boolean = False
    '    Dim objParentComp As Component
    '    Dim objChildComp As Component

    '    iFeatCount = aFeatureList.Length

    '    For iLoopIndex = 0 To iFeatCount - 1
    '        bFound = False
    '        sPartName = aFeatureList(iLoopIndex).OwningPart.Leaf
    '        If asPopulatedList Is Nothing Then
    '            ReDim Preserve asPopulatedList(iCounter)
    '        End If
    '        For iLoopIndex2 = 0 To asPopulatedList.Length() - 1
    '            If asPopulatedList(iLoopIndex2) = sPartName Then
    '                bFound = True
    '                Exit For
    '            End If
    '        Next
    '        If Not bFound Then
    '            ReDim Preserve asPopulatedList(iCounter)
    '            asPopulatedList(iCounter) = sPartName
    '            iCounter = iCounter + 1
    '        End If
    '    Next

    '    For iLoopIndex2 = 0 To asPopulatedList.Length() - 1
    '        iDeleteCounter = 0
    '        For iLoopIndex = 0 To iFeatCount - 1
    '            sPartName = asPopulatedList(iLoopIndex2)
    '            Try
    '                If aFeatureList(iLoopIndex).OwningPart.Leaf = sPartName Then
    '                    ReDim Preserve aoDeleteList(iDeleteCounter)
    '                    aoDeleteList(iDeleteCounter) = aFeatureList(iLoopIndex)
    '                    iDeleteCounter = iDeleteCounter + 1
    '                End If
    '            Catch ex As Exception
    '            End Try
    '        Next

    '        objParentComp = FnGetWorkPart().ComponentAssembly.RootComponent
    '        objChildComp = FnGetComponentInAssembly(asPopulatedList(iLoopIndex2) + ".PRT", objParentComp)
    '        SSetWorkComponent(objChildComp)
    '        FnGetNxSession().UpdateManager.ClearErrorList()

    '        Dim markId2 As Session.UndoMarkId
    '        markId2 = FnGetNxSession().SetUndoMark(Session.MarkVisibility.Visible, "Delete Feature Planes")
    '        Dim nErrs1 As Integer
    '        nErrs1 = FnGetNxSession().UpdateManager.AddToDeleteList(aoDeleteList)

    '        Dim notifyOnDelete2 As Boolean
    '        notifyOnDelete2 = FnGetNxSession().Preferences.Modeling.NotifyOnDelete

    '        Dim nErrs2 As Integer
    '        nErrs2 = FnGetNxSession().UpdateManager.DoUpdate(markId2)

    '        aoDeleteList = Nothing

    '        Dim nullAssemblies_Component As Assemblies.Component = Nothing

    '        Dim partLoadStatus2 As PartLoadStatus
    '        FnGetNxSession.Parts.SetWorkComponent(nullAssemblies_Component, partLoadStatus2)

    '        _workPart = FnGetNxSession.Parts.Work
    '        partLoadStatus2.Dispose()

    '    Next
    'End Sub
    Public Function FnGetUI() As UI
        FnGetUI = UI.GetUI
    End Function
    Public Function FnSelectBody(ByRef selectedBody As NXObject, ByVal sTitle As String) As NXObject

        FnGetUI.LockAccess()
        Dim message As String = "Select Body"
        Dim title As String = sTitle
        Dim sel1 As Selection.Response

        Dim scope As Selection.SelectionScope = Selection.SelectionScope.AnyInAssembly
        Dim keepHighlighted As Boolean = False
        Dim includeFeatures As Boolean = True

        Dim selectionAction As Selection.SelectionAction = Selection.SelectionAction.ClearAndEnableSpecific

        Dim selectionMask_array(0) As Selection.MaskTriple
        selectionMask_array(0) = New Selection.MaskTriple(UFConstants.UF_solid_type, 0, UFConstants.UF_UI_SEL_FEATURE_BODY)

        Dim cursor As Point3d
        Do
            sel1 = FnGetUI.SelectionManager.SelectObject(message, title, scope, _
                 selectionAction, includeFeatures, _
                 keepHighlighted, selectionMask_array, selectedBody, cursor)

        Loop While sel1 = Selection.Response.ObjectSelected Or _
                   sel1 = Selection.Response.ObjectSelectedByName

        FnGetUI.UnlockAccess()

        Return selectedBody

    End Function
    Public Function FnSelectFace(ByRef selectedFace() As NXObject, ByVal sTitle As String) As Integer
        FnGetUI.LockAccess()
        'Dim ui As UI = NXOpen.UI.GetUI
        Dim message As String = "Select Face"
        Dim title As String = sTitle

        Dim scope As Selection.SelectionScope = Selection.SelectionScope.AnyInAssembly
        Dim keepHighlighted As Boolean = False
        Dim includeFeatures As Boolean = False
        Dim resp As Selection.Response

        Dim selectionAction As Selection.SelectionAction = Selection.SelectionAction.ClearAndEnableSpecific

        Dim selectionMask_array(0) As Selection.MaskTriple
        selectionMask_array(0) = New Selection.MaskTriple(UFConstants.UF_solid_type, 0, UFConstants.UF_UI_SEL_FEATURE_ANY_FACE)

        Do
            resp = FnGetUI.SelectionManager.SelectObjects(message, title, scope, selectionAction, includeFeatures, keepHighlighted, selectionMask_array, selectedFace)
        Loop While resp = Selection.Response.ObjectSelected Or _
                   resp = Selection.Response.ObjectSelectedByName

        FnGetUI.UnlockAccess()

        Return selectedFace.GetLength(0)

    End Function
    Public Function FnSelectComponents(ByRef comps() As NXObject, ByVal sTitle As String) As Integer

        'FnGetUI.LockAccess()

        Dim mask(0) As Selection.MaskTriple
        mask(0) = New Selection.MaskTriple(UFConstants.UF_component_type, 0, 0)

        Dim sel1 As Selection.Response

        Do
            sel1 = FnGetUI.SelectionManager.SelectObjects("Select Component", _
                sTitle, Selection.SelectionScope.AnyInAssembly, _
                Selection.SelectionAction.ClearAndEnableSpecific, _
                False, False, mask, comps)
        Loop While sel1 = Selection.Response.ObjectSelected Or _
                   sel1 = Selection.Response.ObjectSelectedByName
        'FnGetUI.UnlockAccess()

        Return comps.GetLength(0)

    End Function
    Public Sub FnLoadPartFully(ByVal objPart As Part)
        Dim objPartLoadStatus As PartLoadStatus
        Try
            objPartLoadStatus = objPart.LoadThisPartFully()
            objPartLoadStatus.Dispose()
            'sWriteToLogFile(objPart.Leaf.ToString & " loaded fully")
        Catch ex As Exception
            sWriteToLogFile(objPart.Leaf.ToString & " not loaded fully")
        End Try
     
    End Sub

    Public Sub SStartWaveGeomLinker(ByVal objWaveLinkerFace As Face)

        'Dim iCount As Integer
        Dim objWaveGeomLinkSelectedFace() As NXObject = Nothing
        Dim nullFeatures_Feature As Features.Feature = Nothing
        Dim objWaveLinkBuilder As Features.WaveLinkBuilder
        objWaveLinkBuilder = _workPart.BaseFeatures.CreateWaveLinkBuilder(nullFeatures_Feature)

        Dim objWaveDatumBuilder As Features.WaveDatumBuilder
        objWaveDatumBuilder = objWaveLinkBuilder.WaveDatumBuilder

        Dim objExtractFaceBuilder As Features.ExtractFaceBuilder
        objExtractFaceBuilder = objWaveLinkBuilder.ExtractFaceBuilder

        'FaceChain selection intent option  
        objExtractFaceBuilder.FaceOption = Features.ExtractFaceBuilder.FaceOptionType.FaceChain
        objWaveLinkBuilder.Type = Features.WaveLinkBuilder.Types.FaceLink
        'OtherPart inter part mode  
        objExtractFaceBuilder.ParentPart = Features.ExtractFaceBuilder.ParentPartType.OtherPart

        '************************************************************************
        'Default Values
        objExtractFaceBuilder.AngleTolerance = 45.0
        objWaveDatumBuilder.DisplayScale = 2.0
        objExtractFaceBuilder.ParentPart = Features.ExtractFaceBuilder.ParentPartType.OtherPart
        objExtractFaceBuilder.Associative = True
        objExtractFaceBuilder.FixAtCurrentTimestamp = False
        objExtractFaceBuilder.HideOriginal = False
        objExtractFaceBuilder.DeleteHoles = False
        objExtractFaceBuilder.InheritDisplayProperties = False
        '************************************************************************

        'Load all parts in the session fully.
        'FnLoadPartFully()

        Dim objFace(0) As Face
        'iCount = FnSelectFace(objWaveGeomLinkSelectedFace, "Select Wave Geometry Link Face")
        'ReDim Preserve arrTrimObjects(iTrimObjectCount)
        'arrTrimObjects(iTrimObjectCount) = objWaveGeomLinkSelectedFace(0)
        'iTrimObjectCount = iTrimObjectCount + 1
        'SWrite(objWaveGeomLinkSelectedFace(0).JournalIdentifier.ToString)
        'objFace(0) = objWaveGeomLinkSelectedFace(0)
        objFace(0) = objWaveLinkerFace
        Dim faceDumbRule1 As FaceDumbRule
        faceDumbRule1 = _workPart.ScRuleFactory.CreateRuleFaceDumb(objFace)

        Dim rules1(0) As SelectionIntentRule
        rules1(0) = faceDumbRule1
        objExtractFaceBuilder.FaceChain.ReplaceRules(rules1, False)

        Dim nXObject1 As NXObject
        nXObject1 = objWaveLinkBuilder.Commit()

        objWaveLinkBuilder.Destroy()

        Dim partLoadStatus3 As PartLoadStatus
        Dim status1 As PartCollection.SdpsStatus
        status1 = FnGetNxSession.Parts.SetDisplay(_workPart, True, True, partLoadStatus3)

        _workPart = FnGetNxSession.Parts.Work
        _displayPart = FnGetNxSession.Parts.Display
        partLoadStatus3.Dispose()


    End Sub

    Public Function FnGetWorkPart() As NXOpen.Part
        Return _workPart
    End Function

    Public Sub SReplaceFace(ByVal objReplaceFace As Face, ByVal objFaceToReplace As Face)
        'Dim iCount As Integer
        'Dim objSelectedReplaceFace() As NXObject = Nothing
        'Dim objSelectedFaceToReplace() As NXObject = Nothing
        Dim nullFeatures_Feature As Features.Feature = Nothing
        Dim nXObject1 As NXObject

        Dim objReplaceFaceBuilder As Features.ReplaceFaceBuilder
        objReplaceFaceBuilder = _workPart.Features.CreateReplaceFaceBuilder(nullFeatures_Feature)

        objReplaceFaceBuilder.OffsetDistance.RightHandSide = "0"

        'FnGetNxSession.SetUndoMarkName(markId6, "Replace Face Dialog")

        Dim faces2(0) As Face
        Dim extractFace1 As Features.ExtractFace = CType(nXObject1, Features.ExtractFace)

        'Dim face2 As Face = CType(extractFace1.FindObject("FACE 1 {(189.6348018946751,54.0999999999603,61.5642522169479) LINKED_FACE(55)}"), Face)
        'iCount = FnSelectFace(objSelectedReplaceFace, "Select Replacement Face")
        'ReDim Preserve arrTrimObjects(iTrimObjectCount)
        'arrTrimObjects(iTrimObjectCount) = objSelectedReplaceFace(0)
        'iTrimObjectCount = iTrimObjectCount + 1
        'SWrite(objSelectedReplaceFace(0).JournalIdentifier.ToString)

        'faces2(0) = objSelectedReplaceFace(0)
        faces2(0) = objReplaceFace
        Dim faceDumbRule2 As FaceDumbRule
        faceDumbRule2 = _workPart.ScRuleFactory.CreateRuleFaceDumb(faces2)

        Dim rules2(0) As SelectionIntentRule
        rules2(0) = faceDumbRule2
        objReplaceFaceBuilder.ReplacementFaces.ReplaceRules(rules2, False)

        Dim faces3(0) As Face
        'Dim trimBody1 As Features.TrimBody = CType(workPart.Features.FindObject("TRIM_BODY(31)"), Features.TrimBody)

        'Dim face3 As Face = CType(trimBody1.FindObject("FACE 1 {(-11.5,54.0999999999288,0) EXTRUDE(4)}"), Face)
        'iCount = FnSelectFace(objSelectedFaceToReplace, "Select Face To Replace")
        'ReDim Preserve arrTrimObjects(iTrimObjectCount)
        'arrTrimObjects(iTrimObjectCount) = objSelectedFaceToReplace(0)
        'iTrimObjectCount = iTrimObjectCount + 1
        'SWrite(objSelectedFaceToReplace(0).JournalIdentifier.ToString)

        'faces3(0) = objSelectedFaceToReplace(0)
        faces3(0) = objFaceToReplace
        Dim faceDumbRule3 As FaceDumbRule
        faceDumbRule3 = _workPart.ScRuleFactory.CreateRuleFaceDumb(faces3)

        Dim rules3(0) As SelectionIntentRule
        rules3(0) = faceDumbRule3
        objReplaceFaceBuilder.FaceToReplace.ReplaceRules(rules3, False)

        objReplaceFaceBuilder.ReverseDirection = True

        'Dim markId7 As Session.UndoMarkId
        'markId7 = FnGetNxSession.SetUndoMark(Session.MarkVisibility.Invisible, "Replace Face")

        Dim nXObject2 As NXObject
        nXObject2 = objReplaceFaceBuilder.Commit()

        'FnGetNxSession.DeleteUndoMark(markId7, Nothing)

        'FnGetNxSession.SetUndoMarkName(markId6, "Replace Face")

        Dim expression1 As Expression = objReplaceFaceBuilder.OffsetDistance

        objReplaceFaceBuilder.Destroy()

    End Sub

    'Public Sub SSetComponentAttribute(ByVal objComp As Component, ByVal sTitle As String, ByVal sValue As String)
    '    objComp.SetAttribute(sTitle, sValue)
    'End Sub

    'Public Sub SSetPartAttribute(ByVal objPart As Part, ByVal sTitle As String, ByVal sValue As String)
    '    Dim markId2 As Session.UndoMarkId
    '    markId2 = FnGetNxSession.SetUndoMark(Session.MarkVisibility.Visible, "Edit Properties")
    '    objPart.SetAttribute(sTitle, sValue)
    '    Dim nErrs1 As Integer
    '    nErrs1 = FnGetNxSession.UpdateManager.DoUpdate(markId2)
    'End Sub

    'Public Function FnGetComponentWithStringAttribute(ByVal objParentComp As Component, ByVal sTitle As String, ByVal sValue As String) As Component
    '    For Each subComp As Component In objParentComp.GetChildren()
    '        If subComp.GetStringAttribute(sTitle) = sValue Then
    '            FnGetComponentWithStringAttribute = subComp
    '            Exit Function
    '        End If
    '    Next
    '    FnGetComponentWithStringAttribute = Nothing
    'End Function
    Public Sub SSetCurrentWorkPart()
        _workPart = FnGetNxSession.Parts.Work
    End Sub

    'Public Sub SDeleteComponent(ByVal objCompToDelete As Component)
    '    Dim markId As Session.UndoMarkId
    '    FnGetNxSession().UpdateManager.ClearErrorList()
    '    markId = FnGetNxSession().SetUndoMark(Session.MarkVisibility.Visible, "Delete")
    '    FnGetNxSession().UpdateManager.AddToDeleteList(objCompToDelete)
    '    'FnGetNxSession.Preferences.Modeling.NotifyOnDelete = False
    '    Dim nErrs2 As Integer
    '    nErrs2 = FnGetNxSession.UpdateManager.DoUpdate(markId)
    'End Sub

    Public Sub ClosePartInSession(ByVal sPartName As String)
        Dim objPart As Part
        objPart = CType(FnGetNxSession.Parts.FindObject(sPartName), Part)
        objPart.Close(BasePart.CloseWholeTree.False, BasePart.CloseModified.UseResponses, Nothing)

    End Sub

    Public Function FnGetChildrenComponent(ByVal objComp As Component)
        Dim sPartName As String
        For Each comp As NXOpen.Assemblies.Component In objComp.GetChildren()
            ReDim Preserve _asSheetMetalParts(_iCountSheetMetalParts)
            sPartName = CType(comp.Prototype, NXOpen.Part).Leaf
            _asSheetMetalParts(_iCountSheetMetalParts) = sPartName
            _iCountSheetMetalParts = _iCountSheetMetalParts + 1

            If comp.GetChildren.Length > 0 Then
                FnGetChildrenComponent(comp)
            End If
        Next
    End Function

    Public Function FnGetComponentPosition(ByVal objComp As Component) As Double()
        Dim arrPosition(12) As Double
        Dim ObjPoint As Point3d
        Dim ObjOrient As Matrix3x3
        objComp.GetPosition(ObjPoint, ObjOrient)
        arrPosition(0) = ObjPoint.X
        arrPosition(1) = ObjPoint.Y
        arrPosition(2) = ObjPoint.Z
        arrPosition(3) = ObjOrient.Xx
        arrPosition(4) = ObjOrient.Xy
        arrPosition(5) = ObjOrient.Xz
        arrPosition(6) = ObjOrient.Yx
        arrPosition(7) = ObjOrient.Yy
        arrPosition(8) = ObjOrient.Yz
        arrPosition(9) = ObjOrient.Zx
        arrPosition(10) = ObjOrient.Zy
        arrPosition(11) = ObjOrient.Zz
        FnGetComponentPosition = arrPosition
    End Function

    'Public Sub SSetComponentPosition(ByVal objComp As Component, ByVal X As Double, ByVal Y As Double, ByVal Z As Double, _
    '                                    ByVal Xx As Double, ByVal Xy As Double, ByVal Xz As Double, ByVal Yx As Double, ByVal Yy As Double, ByVal Yz As Double, _
    '                                    ByVal Zx As Double, ByVal Zy As Double, ByVal Zz As Double)

    '    'SSetWorkComponent(objComp)
    '    Dim objComponentPositioner As Positioning.ComponentPositioner
    '    objComponentPositioner = FnGetWorkPart.ComponentAssembly.Positioner
    '    objComponentPositioner.ClearNetwork()
    '    'Dim arrangement1 As Assemblies.Arrangement = CType(FnGetWorkPart.ComponentAssembly.Arrangements.FindObject("Arrangement 1"), Assemblies.Arrangement)
    '    'objComponentPositioner.PrimaryArrangement = arrangement1
    '    objComponentPositioner.BeginMoveComponent()
    '    Dim objNetwork As Positioning.Network
    '    objNetwork = objComponentPositioner.EstablishNetwork()

    '    Dim objComponentNetwork1 As Positioning.ComponentNetwork = CType(objNetwork, Positioning.ComponentNetwork)
    '    'objComponentNetwork1.MoveObjectsState = True
    '    'objComponentNetwork1.DisplayComponent = objComp
    '    'objComponentNetwork1.NetworkArrangementsMode = Positioning.ComponentNetwork.ArrangementsMode.Existing
    '    'objComponentNetwork1.RemoveAllConstraints()

    '    Dim movableObjects1(0) As NXObject
    '    'Dim component2 As Assemblies.Component = CType(component1.FindObject("COMPONENT agg75206.f01.0049 1"), Assemblies.Component)

    '    movableObjects1(0) = objComp
    '    objComponentNetwork1.SetMovingGroup(movableObjects1)

    '    'objComponentNetwork1.Solve()

    '    'theSession.SetUndoMarkName(markId2, "Move Component Dialog")

    '    objComponentNetwork1.MoveObjectsState = True

    '    'objComponentNetwork1.NetworkArrangementsMode = Positioning.ComponentNetwork.ArrangementsMode.Existing

    '    Dim markId3 As Session.UndoMarkId
    '    markId3 = FnGetNxSession.SetUndoMark(Session.MarkVisibility.Invisible, "Move Component Update")

    '    'Dim markId4 As Session.UndoMarkId
    '    'markId4 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Transform Origin")

    '    'Dim loaded1 As Boolean
    '    'loaded1 = objComponentNetwork1.IsReferencedGeometryLoaded()

    '    'objComponentNetwork1.BeginDrag()

    '    'Dim translation1 As Vector3d = New Vector3d(X, Y, Z)
    '    'objComponentNetwork1.DragByTranslation(translation1)

    '    'objComponentNetwork1.EndDrag()

    '    'objComponentNetwork1.ResetDisplay()

    '    'objComponentNetwork1.ApplyToModel()

    '    objComponentNetwork1.BeginDrag()

    '    Dim translation2 As Vector3d = New Vector3d(X, Y, Z)
    '    Dim rotation1 As Matrix3x3
    '    rotation1.Xx = Xx
    '    rotation1.Xy = Xy
    '    rotation1.Xz = Xz
    '    rotation1.Yx = Yx
    '    rotation1.Yy = Yy
    '    rotation1.Yz = Yz
    '    rotation1.Zx = Zx
    '    rotation1.Zy = Zy
    '    rotation1.Zz = Zz
    '    objComponentNetwork1.DragByTransform(translation2, rotation1)

    '    objComponentNetwork1.EndDrag()

    '    'objComponentNetwork1.ResetDisplay()

    '    'objComponentNetwork1.ApplyToModel()

    '    objComponentNetwork1.Solve()

    '    'objComponentNetwork1.ResetDisplay()

    '    'objComponentNetwork1.ApplyToModel()

    '    'objComponentPositioner.ClearNetwork()

    '    Dim nErrs1 As Integer
    '    nErrs1 = FnGetNxSession.UpdateManager.AddToDeleteList(objComponentNetwork1)

    '    Dim nErrs2 As Integer
    '    nErrs2 = FnGetNxSession.UpdateManager.DoUpdate(markId3)

    '    'objComponentPositioner.DeleteNonPersistentConstraints()

    '    Dim nErrs3 As Integer
    '    nErrs3 = FnGetNxSession.UpdateManager.DoUpdate(markId3)

    '    FnGetNxSession.DeleteUndoMark(markId3, Nothing)

    '    'Dim nullAssemblies_Component As Assemblies.Component = Nothing

    '    'Dim partLoadStatus2 As PartLoadStatus
    '    'FnGetNxSession.Parts.SetWorkComponent(nullAssemblies_Component, partLoadStatus2)

    '    '_workPart = FnGetNxSession.Parts.Work
    '    'partLoadStatus2.Dispose()

    'End Sub

    Public Sub SSetComponentPosition(ByVal objComp As Component, ByVal X As Double, ByVal Y As Double, ByVal Z As Double, _
                                        ByVal Xx As Double, ByVal Xy As Double, ByVal Xz As Double, ByVal Yx As Double, ByVal Yy As Double, ByVal Yz As Double, _
                                        ByVal Zx As Double, ByVal Zy As Double, ByVal Zz As Double)

        Dim ac As Assemblies.ComponentAssembly = FnGetNxSession.Parts.Work.ComponentAssembly()
        Dim translation As Vector3d = New Vector3d(X, Y, Z)
        Dim rotation As Matrix3x3
        rotation.Xx = Xx
        rotation.Xy = Xy
        rotation.Xz = Xz
        rotation.Yx = Yx
        rotation.Yy = Yy
        rotation.Yz = Yz
        rotation.Zx = Zx
        rotation.Zy = Zy
        rotation.Zz = Zz
        ac.MoveComponent(objComp, translation, rotation)
    End Sub

    'Public Function SEditExpression(ByVal sExpressionName As String, ByVal sNewValue As String, ByVal sUnitName As String, ByVal objPart As Part)
    '    Dim markId As Session.UndoMarkId
    '    markId = FnGetNxSession.SetUndoMark(Session.MarkVisibility.Visible, "Expression")

    '    Dim objExpression As Expression
    '    For Each exp As Expression In objPart.Expressions
    '        If exp.Name = sExpressionName Then
    '            objExpression = exp
    '            Exit For
    '        End If
    '    Next

    '    Dim objUnit As Unit
    '    For Each Unit As Unit In objPart.UnitCollection
    '        If Unit.Name = sUnitName Then
    '            objUnit = Unit
    '            Exit For
    '        End If
    '    Next

    '    'Edit the expression
    '    If Not objExpression Is Nothing And Not objUnit Is Nothing Then
    '        objPart.Expressions.EditWithUnits(objExpression, objUnit, sNewValue)
    '    End If

    '    Dim nErrs1 As Integer
    '    nErrs1 = FnGetNxSession.UpdateManager.DoUpdate(markId)

    'End Function



    Public Sub SSetLayer(ByVal ilayerNos As Integer, ByVal objComp As Component)
        Dim objectArray(0) As DisplayableObject
        objectArray(0) = objComp
        _workPart = FnGetNxSession.Parts.Work
        _workPart.Layers.MoveDisplayableObjects(ilayerNos, objectArray)
    End Sub

    Public Function FnGetAllComponentsInSession() As Component()
        Dim listcomp() As Component = Nothing
        ReDim Preserve listcomp(0)
        If Not FnGetNxSession.Parts.Work.ComponentAssembly.RootComponent Is Nothing Then
            listcomp(0) = FnGetNxSession.Parts.Work.ComponentAssembly.RootComponent
            For Each objComp As Component In FnGetNxSession.Parts.Work.ComponentAssembly.RootComponent.GetChildren()
                ReDim Preserve listcomp(UBound(listcomp) + 1)
                listcomp(UBound(listcomp)) = objComp
                sGetAllChildren(objComp, listcomp)
            Next
            FnGetAllComponentsInSession = listcomp
        Else
            FnGetAllComponentsInSession = Nothing
        End If
    End Function

    Public Sub sGetAllChildren(ByVal objParentComp As Component, ByRef listcomp() As Component)
        For Each objComp As Component In objParentComp.GetChildren()
            ReDim Preserve listcomp(UBound(listcomp) + 1)
            listcomp(UBound(listcomp)) = objComp
            For Each objChildComp As Component In objComp.GetChildren()
                sGetAllChildren(objChildComp, listcomp)
            Next
        Next
    End Sub
    'Set the Assembly load search option from folder
    Public Sub sLoadAssemblySearchFromFolder(asSearchPaths() As String)

        FnGetNxSession.Parts.LoadOptions.LoadLatest = False
        FnGetNxSession.Parts.LoadOptions.ComponentLoadMethod = LoadOptions.LoadMethod.SearchDirectories
        'Dim searchDirectories1(0) As String
        'searchDirectories1(0) = "C:\Vectra\Input Files\NXParts\"
        Dim abSearchSubDirs() As Boolean = Nothing
        ReDim Preserve abSearchSubDirs(UBound(asSearchPaths))
        For iIndex As Integer = 0 To UBound(abSearchSubDirs)
            abSearchSubDirs(iIndex) = True
        Next
        'Dim searchSubDirs1(0) As Boolean
        'searchSubDirs1(0) = True
        FnGetNxSession.Parts.LoadOptions.SetSearchDirectories(asSearchPaths, abSearchSubDirs)
        FnGetNxSession.Parts.LoadOptions.ComponentsToLoad = LoadOptions.LoadComponents.All
        FnGetNxSession.Parts.LoadOptions.UsePartialLoading = True
        FnGetNxSession.Parts.LoadOptions.UseLightweightRepresentations = False
        FnGetNxSession.Parts.LoadOptions.SetInterpartData(False, LoadOptions.Parent.Partial)
        FnGetNxSession.Parts.LoadOptions.AllowSubstitution = False
        FnGetNxSession.Parts.LoadOptions.GenerateMissingPartFamilyMembers = True
        FnGetNxSession.Parts.LoadOptions.AbortOnFailure = False

        Dim referenceSets1(3) As String
        referenceSets1(0) = "As Saved"
        referenceSets1(1) = "Entire Part"
        referenceSets1(2) = "Empty"
        referenceSets1(3) = "Use Model"
        FnGetNxSession.Parts.LoadOptions.SetDefaultReferenceSets(referenceSets1)
        FnGetNxSession.Parts.LoadOptions.ReferenceSetOverride = False
        FnGetNxSession.Parts.LoadOptions.SetBookmarkComponentsToLoad(True, False, LoadOptions.BookmarkComponents.LoadVisible)
        FnGetNxSession.Parts.LoadOptions.BookmarkRefsetLoadBehavior = LoadOptions.BookmarkRefsets.ImportData

    End Sub
    '3/23/16 - Added
    'Set the Assembly load search option from folder
    Public Sub sLoadAssemblyAsSaved()
        FnGetNxSession.Parts.LoadOptions.LoadLatest = False
        FnGetNxSession.Parts.LoadOptions.ComponentLoadMethod = LoadOptions.LoadMethod.AsSaved
        FnGetNxSession.Parts.LoadOptions.ComponentsToLoad = LoadOptions.LoadComponents.All
        FnGetNxSession.Parts.LoadOptions.UsePartialLoading = False
        FnGetNxSession.Parts.LoadOptions.UseLightweightRepresentations = False
        FnGetNxSession.Parts.LoadOptions.SetInterpartData(False, LoadOptions.Parent.Partial)
        FnGetNxSession.Parts.LoadOptions.AllowSubstitution = False
        FnGetNxSession.Parts.LoadOptions.GenerateMissingPartFamilyMembers = True
        FnGetNxSession.Parts.LoadOptions.AbortOnFailure = False

        Dim referenceSets1(3) As String
        referenceSets1(0) = "As Saved"
        referenceSets1(1) = "Entire Part"
        referenceSets1(2) = "Empty"
        referenceSets1(3) = "Use Model"
        FnGetNxSession.Parts.LoadOptions.SetDefaultReferenceSets(referenceSets1)
        FnGetNxSession.Parts.LoadOptions.ReferenceSetOverride = False
        FnGetNxSession.Parts.LoadOptions.SetBookmarkComponentsToLoad(True, False, LoadOptions.BookmarkComponents.LoadVisible)
        FnGetNxSession.Parts.LoadOptions.BookmarkRefsetLoadBehavior = LoadOptions.BookmarkRefsets.ImportData
    End Sub
    'Set the Assembly load search option from folder
    Public Sub sLoadAssemblyFromFolder()

        FnGetNxSession.Parts.LoadOptions.LoadLatest = False
        FnGetNxSession.Parts.LoadOptions.ComponentLoadMethod = LoadOptions.LoadMethod.FromDirectory
        FnGetNxSession.Parts.LoadOptions.ComponentsToLoad = LoadOptions.LoadComponents.All
        FnGetNxSession.Parts.LoadOptions.UsePartialLoading = False
        FnGetNxSession.Parts.LoadOptions.UseLightweightRepresentations = False
        FnGetNxSession.Parts.LoadOptions.SetInterpartData(False, LoadOptions.Parent.Partial)
        FnGetNxSession.Parts.LoadOptions.AllowSubstitution = True
        FnGetNxSession.Parts.LoadOptions.GenerateMissingPartFamilyMembers = True
        FnGetNxSession.Parts.LoadOptions.AbortOnFailure = False
        FnGetNxSession.Parts.LoadOptions.ReferenceSetOverride = True
        FnGetNxSession.Parts.LoadOptions.SetBookmarkComponentsToLoad(True, False, LoadOptions.BookmarkComponents.LoadVisible)
        FnGetNxSession.Parts.LoadOptions.BookmarkRefsetLoadBehavior = LoadOptions.BookmarkRefsets.ImportData

    End Sub
    Public Function FnCalculateMassProperties(ByVal objPart As Part, ByVal objBody As Body, Optional ByVal iIndexProp As Integer = 2) As String

        Dim ufs As UFSession = UFSession.GetUFSession()
        Dim propToCalculate As Double
        Dim units As Integer
        Dim accuracy As Integer = 1
        Dim accValue(10) As Double
        Dim density As Double
        Dim massProps(46) As Double
        Dim statistics(12) As Double
        Dim bods(0) As NXOpen.Tag
        accValue(0) = 0.99
        '1 = Solid Bodies 
        '2 = Thin Shell - Sheet Bodies 
        '3 = Bounded by Sheet Bodies
        Dim solidBodyType As Integer = 1
        Dim partUnits As Integer

        ufs.Part.AskUnits(FnGetNxSession.Parts.Work.Tag, partUnits)
        If partUnits = UFConstants.UF_PART_ENGLISH Then
            units = 1       ' Pounds
        ElseIf partUnits = UFConstants.UF_PART_METRIC Then
            units = 4       'Kilograms
        End If

        bods(0) = objBody.Tag
        Try
            ufs.Modl.AskMassProps3d(bods, 1, solidBodyType, units, _
                                    density, accuracy, accValue, _
                                              massProps, statistics)
            propToCalculate = massProps(iIndexProp)
            FnCalculateMassProperties = propToCalculate.ToString()
        Catch ex As NXException
            'Unable to compute mass
            FnCalculateMassProperties = ""
        End Try

    End Function
    'Delete a body attribute
    Public Sub sDeleteBodyAttribute(ByVal objBody As Body, ByVal sType As String, ByVal sAttrTitle As String)
        Try
            If sType = "Integer" Then
                sDeleteIntegerUserAttribute(objBody, sAttrTitle)
            ElseIf sType = "String" Then
                sDeleteStringUserAttribute(objBody, sAttrTitle)
            ElseIf sType = "Real" Then
                sDeleteRealUserAttribute(objBody, sAttrTitle)
            End If
        Catch ex As Exception
        End Try
        sUpdateAttributesInModel()
    End Sub

    'NX 9.0.0 Functions
    'Set a String User Attribute
    Sub sSetStringUserAttribute(objNX As NXObject, sTitle As String, sValue As String, Optional sCategory As String = "VECTRA")
        'Dim objAttrInfo As NXObject.AttributeInformation = Nothing
        'objAttrInfo.Category = "VECTRA"
        'objAttrInfo.Title = sTitle
        'objAttrInfo.StringValue = sValue
        'objAttrInfo.Type = NXObject.AttributeType.String
        'objNX.SetUserAttribute(objAttrInfo, Update.Option.Later)

        Dim attPropBuilder As NXOpen.AttributePropertiesBuilder = _
  FnGetNxSession.AttributeManager.CreateAttributePropertiesBuilder(FnGetWorkPart, New NXObject() {objNX}, AttributePropertiesBuilder.OperationType.Create)

        attPropBuilder.Title = sTitle
        attPropBuilder.StringValue = sValue
        attPropBuilder.Category = sCategory
        attPropBuilder.CreateAttribute()

        attPropBuilder.Category = ""
        attPropBuilder.Title = ""
        attPropBuilder.StringValue = ""

        attPropBuilder.Commit()
        attPropBuilder.Destroy()
    End Sub
    'Set a Integer User Attribute
    Sub sSetIntegerUserAttribute(objNX As NXObject, sTitle As String, iValue As Integer, Optional sCategory As String = "VECTRA")
        Dim objAttrInfo As NXObject.AttributeInformation = Nothing
        objAttrInfo.Category = sCategory
        objAttrInfo.Title = sTitle
        objAttrInfo.IntegerValue = iValue
        objAttrInfo.Type = NXObject.AttributeType.Integer
        objNX.SetUserAttribute(objAttrInfo, Update.Option.Later)
    End Sub
    'Set a Real User Attribute
    'Sub sSetRealUserAttribute(objNX As NXObject, sTitle As String, dValue As Double)
    '    Dim objAttrInfo As NXObject.AttributeInformation = Nothing
    '    objAttrInfo.Title = sTitle
    '    objAttrInfo.RealValue = dValue
    '    objAttrInfo.Type = NXObject.AttributeType.Real
    '    objNX.SetUserAttribute(objAttrInfo, Update.Option.Later)
    'End Sub

    Sub sSetRealUserAttribute(objPart As Part, objNX As NXObject, sTitle As String, dValue As Double)
        Dim attributePropertiesBuilder1 As AttributePropertiesBuilder
        attributePropertiesBuilder1 = FnGetNxSession.AttributeManager.CreateAttributePropertiesBuilder(objPart, {objNX}, _
                                                                                                   AttributePropertiesBuilder.OperationType.None)
        attributePropertiesBuilder1.IsArray = False
        attributePropertiesBuilder1.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.Number
        attributePropertiesBuilder1.Title = sTitle
        attributePropertiesBuilder1.Units = "Kilogram"
        attributePropertiesBuilder1.NumberValue = dValue
        Dim nXObject1 As NXObject
        nXObject1 = attributePropertiesBuilder1.Commit()

    End Sub

    'Get a String user attribute
    Function FnGetStringUserAttribute(objNX As NXObject, sTitle As String) As String
        Dim sAttribute As String = ""
        Try
            sAttribute = objNX.GetUserAttributeAsString(sTitle, NXObject.AttributeType.String, -1)
            'Code added July-05-2019
            If (_sSupplierName <> "" And _sOemName <> "") Then
                If (_sSupplierName = VALIANT_NAME) And ((_sOemName = CHRYSLER_OEM_NAME) Or (_sOemName = GESTAMP_OEM_NAME)) Then
                    If sAttribute <> "" Then
                        If sAttribute.ToUpper = "X" Then
                            sAttribute = ""
                        End If
                    End If
                End If
            End If
            FnGetStringUserAttribute = sAttribute
        Catch ex As NXException
            FnGetStringUserAttribute = ""
        End Try
    End Function
    'Get a integer user attribute
    Function FnGetIntegerUserAttribute(objNX As NXObject, sTitle As String) As String
        Try
            FnGetIntegerUserAttribute = objNX.GetUserAttributeAsString(sTitle, NXObject.AttributeType.Integer, -1)
        Catch ex As NXException
            FnGetIntegerUserAttribute = ""
        End Try
    End Function
    'Get a real user attribute
    Function FnGetRealUserAttribute(objNX As NXObject, sTitle As String) As String
        Try
            FnGetRealUserAttribute = objNX.GetUserAttributeAsString(sTitle, NXObject.AttributeType.Real, -1)
        Catch ex As NXException
            FnGetRealUserAttribute = ""
        End Try
    End Function
    'Delete String attribute
    Sub sDeleteStringUserAttribute(objNX As NXObject, sTitle As String)
        objNX.DeleteUserAttribute(NXObject.AttributeType.String, sTitle, True, Update.Option.Later)
    End Sub
    'Delete integer attribute
    Sub sDeleteIntegerUserAttribute(objNX As NXObject, sTitle As String)
        objNX.DeleteUserAttribute(NXObject.AttributeType.Integer, sTitle, True, Update.Option.Later)
    End Sub
    'Delete real attribute
    Sub sDeleteRealUserAttribute(objNX As NXObject, sTitle As String)
        objNX.DeleteUserAttribute(NXObject.AttributeType.Real, sTitle, True, Update.Option.Later)
    End Sub
    'Update the part
    Sub sUpdateAttributesInModel()
        Dim markID As Session.UndoMarkId = Nothing
        markID = FnGetNxSession.SetUndoMark(Session.MarkVisibility.Visible, "Start Attributes Update")
        Dim iErr As Integer = 0
        iErr = FnGetNxSession.UpdateManager().DoUpdate(markID)
    End Sub
    Public Function FnComputeMinDistance(ByVal objPart As Part, ByVal obj1 As DisplayableObject, ByVal obj2 As DisplayableObject) As Double
        'FnComputeMinDistance = objPart.MeasureManager.NewDistance(Nothing, MeasureManager.MeasureType.Minimum, obj1, obj2).Value

        Dim adClosePoint1 As Point3d
        Dim adClosePoint2 As Point3d
        Dim dAccuracy As Double
        FnComputeMinDistance = FnGetNxSession.Measurement.GetMinimumDistance(obj1, obj2, adClosePoint1, adClosePoint2, dAccuracy)
      
    End Function

    'Compute if face is mating
    'Default tolerance of 0.1
    Public Function FnIsFaceMating(ByVal objPart As Part, ByVal objFace1 As Face, ByVal objFace2 As Face, _
                                   Optional ByVal dTolerance As Double = 0.1, _
                                   Optional ByVal sMatingParentPartFaceType As String = PLANAR_FACE) As Boolean

        Dim dirFace1 As Direction = Nothing
        Dim dirFace2 As Direction = Nothing
        Dim dDotProductValue As Double = 0.0
        Dim objFace1RefPoint1 As Point3d = Nothing
        Dim objFace1RefPoint2 As Point3d = Nothing
        Dim objFace2RefPoint1 As Point3d = Nothing
        Dim objFace2RefPoint2 As Point3d = Nothing
        Dim objGeomProp As GeometricAnalysis.GeometricProperties = Nothing

        Dim objGeomPropFace1 As GeometricAnalysis.GeometricProperties.Face = Nothing
        Dim objGeomPropFace2 As GeometricAnalysis.GeometricProperties.Face = Nothing
        Dim dDistance As Double = -1

       
        dDistance = FnComputeMinDistance(objPart, objFace1, objFace2)

        If dDistance <= dTolerance Then
            If sMatingParentPartFaceType = PLANAR_FACE Or sMatingParentPartFaceType = CYLINDRICAL_FACE Then
                'Get the Dot product of the face vectors to check antiparallelity
                Try
                    dirFace1 = objPart.Directions.CreateDirection(objFace1, Sense.Forward, SmartObject.UpdateOption.WithinModeling)
                    dirFace2 = objPart.Directions.CreateDirection(objFace2, Sense.Forward, SmartObject.UpdateOption.WithinModeling)
                    dDotProductValue = (((dirFace1.Vector.X) * (dirFace2.Vector.X)) + ((dirFace1.Vector.Y) * (dirFace2.Vector.Y)) + _
                                        ((dirFace1.Vector.Z) * (dirFace2.Vector.Z)))
                    If sMatingParentPartFaceType = PLANAR_FACE Then
                        'Check for anti-parallel
                        If Abs(dDotProductValue + 1) < 0.01 Then
                            FnIsFaceMating = True
                            Exit Function
                        End If
                    ElseIf sMatingParentPartFaceType = CYLINDRICAL_FACE Then
                        'Check for both parallel or anti-parallel
                        If Abs(dDotProductValue + 1) < 0.01 Or Abs(dDotProductValue - 1) < 0.01 Then
                            FnIsFaceMating = True
                            Exit Function
                        End If
                    End If

                Catch ex As Exception
                    objGeomProp = objPart.AnalysisManager.CreateGeometricPropertiesObject()
                    objFace1.GetEdges(0).GetVertices(objFace1RefPoint1, objFace1RefPoint2)
                    objGeomProp.GetFaceProperties(objFace1, objFace1RefPoint1, objGeomPropFace1)

                    objFace2.GetEdges(0).GetVertices(objFace2RefPoint1, objFace2RefPoint2)
                    objGeomProp.GetFaceProperties(objFace2, objFace2RefPoint1, objGeomPropFace2)

                    dDotProductValue = ((objGeomPropFace1.NormalInWcs.X * objGeomPropFace2.NormalInWcs.X) + _
                                    (objGeomPropFace1.NormalInWcs.Y * objGeomPropFace2.NormalInWcs.Y) + _
                                    (objGeomPropFace1.NormalInWcs.Z * objGeomPropFace2.NormalInWcs.Z))
                    If sMatingParentPartFaceType = PLANAR_FACE Then
                        'Check for anti-parallel
                        If Abs(dDotProductValue + 1) < 0.01 Then
                            FnIsFaceMating = True
                            objGeomProp.Destroy()
                            Exit Function
                        End If
                    ElseIf sMatingParentPartFaceType = CYLINDRICAL_FACE Then
                        'Check for both parallel or anti-parallel
                        If Abs(dDotProductValue + 1) < 0.01 Or Abs(dDotProductValue - 1) < 0.01 Then
                            FnIsFaceMating = True
                            objGeomProp.Destroy()
                            Exit Function
                        End If
                    End If
                    objGeomProp.Destroy()
                End Try
            Else
                FnIsFaceMating = True
                Exit Function
            End If
        End If

        FnIsFaceMating = False
    End Function

    'Get the Face Center
    Function FnGetFaceCenter(objFace As Face) As Double()
        Dim uvMinMax(3) As Double
        Dim adCenter(2) As Double
        Dim unitNormal(2) As Double
        Dim u1(2) As Double
        Dim u2(2) As Double
        Dim v1(2) As Double
        Dim v2(2) As Double
        Dim radii(1) As Double
        Dim param(1) As Double
        'Computes the u,v parameter space min, max of a face.
        FnGetUFSession.Modl.AskFaceUvMinmax(objFace.Tag, uvMinMax)
        param(0) = ((uvMinMax(0) + uvMinMax(1)) / 2)
        param(1) = ((uvMinMax(2) + uvMinMax(3)) / 2)
        'tag_t face_id Input Face identifier. 
        'double param [ 2 ]  Input Parameter (u,v) on face (param[2]). 
        'double point [ 3 ]  Output Point at parameter (point[3]). 
        'double u1 [ 3 ]  Output First derivative in U (u1[3]). 
        'double v1 [ 3 ]  Output First derivative in V (v1[3]). 
        'double u2 [ 3 ]  Output Second derivative in U (u2[3]). 
        'double v2 [ 3 ]  Output Second derivative in V (v2[3]). 
        'double unit_norm [ 3 ]  Output Unit face normal (unit_norm[3]). 
        'double radii [ 2 ]  Output Principal radii of curvature (radii[2]). 

        FnGetUFSession.Modl.AskFaceProps(objFace.Tag, param, adCenter, u1, _
                              v1, u2, v2, unitNormal, radii)
        FnGetFaceCenter = adCenter
    End Function

    'Function to Change the View Type (Either Modeling or Drafting View)
    Public Sub FnViewType(ByVal Obj As Part)
        Dim viewtype As Integer
        '1 = Modeling VIew
        '2 = DrawingView
        FnGetUFSession.Draw.AskDisplayState(viewtype)
        If viewtype = 2 Then
            FnGetUFSession.Draw.SetDisplayState(1)
        End If
    End Sub

    Public Sub sOpenDraftingApplication(ByVal objPart As Part)
        For Each sheet As DrawingSheet In objPart.DrawingSheets
            sheet.Open()
            Exit For
        Next
    End Sub

    'Function to Change View Stettings
    Public Sub sChangeViewSettingsForIsoVw(ByVal objview As DraftingView)
        objview.Style.SmoothEdges.SmoothEdge = True
        objview.Style.HiddenLines.HiddenlineFont = Preferences.Font.Invisible
        objview.Style.General.Silhouettes = False
        objview.Commit()
        FnGetUFSession.Disp.SetDisplay(UFConstants.UF_DISP_SUPPRESS_DISPLAY)
    End Sub
    'Function to Get Displayed Part
    Public Function FnGetDisplayedPart() As Part
        FnGetDisplayedPart = FnGetNXSession.Parts.Display
    End Function

    'Delete a modelling view
    Public Sub sDeleteModellingView(ByVal objPart As Part, ByVal sViewName As String)
        Dim viewType As Integer = 0
        Dim objMdlViewTobeDeleted As ModelingView = Nothing
        '1 = modeling view
        '2 = drawing view
        'other = error
        FnGetUFSession.Draw.AskDisplayState(viewType)

        'if drawing sheet shown, change to modeling view
        If viewType = 2 Then
            FnGetUFSession.Draw.SetDisplayState(1)
        End If

        For Each objModView As ModelingView In objPart.ModelingViews
            If objModView.Name.ToUpper = sViewName.ToUpper Then
                objMdlViewTobeDeleted = objModView
                Exit For
            End If
        Next
        'Check if the model view to be deleted is the current work view
        If objPart.ModelingViews.WorkView Is objMdlViewTobeDeleted Then
            'Change another modeling view as the work view
            For Each objModView As ModelingView In objPart.ModelingViews
                'CODE CHANGED - 4/20/16 - If the model view to be deleted is the work view then make one of the canned
                'views as the work view and then delete
                If Not objPart.ModelingViews.WorkView Is Nothing Then
                    If objPart.ModelingViews.WorkView.Name.ToUpper <> objMdlViewTobeDeleted.Name.ToUpper Then
                        objPart.ModelingViews.WorkView.Fit()
                        Exit For
                    Else
                        'Make the Top view as the work view
                        If objModView.Name.ToUpper = "TOP" Then
                            'If objModView.Name.ToUpper <> sViewName.ToUpper Then
                            sReplaceViewInLayout(objPart, objModView)
                            Exit For
                        End If
                    End If
                Else
                    'Make the Top view as the work view
                    If objModView.Name.ToUpper = "TOP" Then
                        'If objModView.Name.ToUpper <> sViewName.ToUpper Then
                        sReplaceViewInLayout(objPart, objModView)
                        Exit For
                    End If
                End If
            Next
        End If
        SDeleteObjects({objMdlViewTobeDeleted})
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
    Public Function FnGetFaceArea(ByVal objPart As Part, ByVal objFace As Face) As Double
        Dim nullUnit As Unit = Nothing
        Dim areaUnit As Unit = Nothing
        Dim lengthUnit As Unit = Nothing
        Dim accuracy As Double = 0.99

        ' Dim measureDistance1 As MeasureFaceBuilder = objPart.MeasureManager.CreateMeasureFaceBuilder(objFace)
        FnGetFaceArea = objPart.MeasureManager.NewFaceProperties(areaUnit, lengthUnit, accuracy, {objFace}).Area
    End Function
    
    'Check for parallel 
    Public Function FnParallelAntiParallelCheck(ByVal adVector1() As Double, ByVal adVector2() As Double, Optional bConcentricCheck As Boolean = False) As Boolean
        Dim dDotProductValue As Double = 0.0
        dDotProductValue = (adVector1(0) * adVector2(0)) + (adVector1(1) * adVector2(1)) + (adVector1(2) * adVector2(2))
        'Code modified on Nov-16-2017
        'Only for concentric check tolerance should be 10degree. for all others it can be 1 degree (Confirmed by Pradeep)
        'Code Changed on Mar-14-2017
        'Tolerance of 1 degree changed to 10 degree
        'This change was made based on the reference part mb444961 in tool maa53846
        'Anti-Parallel Check
        If bConcentricCheck Then
            If Abs(dDotProductValue + 1) < ANGULAR_TOLERANCE_FOR_CONCENTRIC_CHECK Then
                FnParallelAntiParallelCheck = True
                Exit Function
                'Parallel Check
            ElseIf Abs(dDotProductValue - 1) < ANGULAR_TOLERANCE_FOR_CONCENTRIC_CHECK Then
                FnParallelAntiParallelCheck = True
                Exit Function
            End If

        Else
            If Abs(dDotProductValue + 1) < ONE_DEG Then
                FnParallelAntiParallelCheck = True
                Exit Function
                'Parallel Check
            ElseIf Abs(dDotProductValue - 1) < ONE_DEG Then
                FnParallelAntiParallelCheck = True
                Exit Function
            End If
        End If

        FnParallelAntiParallelCheck = False
    End Function

    'Check for parallel 
    Public Function FnParallelCheck(ByVal adVector1() As Double, ByVal adVector2() As Double, Optional bConcentricCheck As Boolean = False) As Boolean
        Dim dDotProductValue As Double = 0.0
        dDotProductValue = (adVector1(0) * adVector2(0)) + (adVector1(1) * adVector2(1)) + (adVector1(2) * adVector2(2))
       
        If bConcentricCheck Then
            'Parallel Check
            If Abs(dDotProductValue - 1) < ANGULAR_TOLERANCE_FOR_CONCENTRIC_CHECK Then
                FnParallelCheck = True
                Exit Function
            End If

        Else
            'Parallel Check
            If Abs(dDotProductValue - 1) < ONE_DEG Then
                FnParallelCheck = True
                Exit Function
            End If
        End If

        FnParallelCheck = False
    End Function

    'Code commented on Feb-01-2019
    'For Parallel Anti Parallel check apart from concentricity check also check with tolerance 1 degree tolerance
    ' ''Code added Oct-12-2018
    ' ''This function was added to compute LCS computation. While obtaining the rotational matrix check for the exact parallel anti parallel check.
    ' ''Ref part WACDU-0010144689-S in tool WACUN-0000144689 has the Slot Face with 0.5 degree tolerance, which lead to wrong orientation.
    ' ''Check for parallel 
    ''Public Function FnParallelAntiParallelCheckWithoutTolerance(ByVal adVector1() As Double, ByVal adVector2() As Double) As Boolean
    ''    Dim dDotProductValue As Double = 0.0
    ''    dDotProductValue = (adVector1(0) * adVector2(0)) + (adVector1(1) * adVector2(1)) + (adVector1(2) * adVector2(2))

    ''    'Anti-Parallel Check
    ''    If Abs(dDotProductValue + 1) < ANGULAR_TOLERANCE_FOR_CONCENTRIC_CHECK Then
    ''        FnParallelAntiParallelCheckWithoutTolerance = True
    ''        Exit Function
    ''        'Parallel Check
    ''    ElseIf Abs(dDotProductValue - 1) < ANGULAR_TOLERANCE_FOR_CONCENTRIC_CHECK Then
    ''        FnParallelAntiParallelCheckWithoutTolerance = True
    ''        Exit Function
    ''    End If

    ''    FnParallelAntiParallelCheckWithoutTolerance = False
    ''End Function

    '****************************************************************************************************************************************************
    'function to Get the Axis Vector of a cylindrical face
    'Description        : Get Axis Vector of Cylindrical face
    'Function Name      : FnGetAxisVecCylFace
    'Input Parameter    : objCylFace - Cylinderical face object
    'Output Parameter   : axis vector(double()) of cylindericalface
    '****************************************************************************************************************************************************
    Public Function FnGetAxisVecCylFace(ByVal objPart As Part, ByVal objCylFace As Face) As Double()
        Dim dir As Direction = Nothing
        dir = objPart.Directions.CreateDirection(objCylFace, Sense.Forward, SmartObject.UpdateOption.WithinModeling)
        FnGetAxisVecCylFace = {dir.Vector.X, dir.Vector.Y, dir.Vector.Z}
    End Function
    'Function to check if the given two face curvatures are opposite to each other.
    '(i.e) if objface1 is concave, then objface2 should be convex and similarly viceversa.
    Function FnCheckIfFaceAreOppositeCurvature(objFace1 As Face, objFace2 As Face) As Boolean

        If (Not objFace1 Is Nothing) And (Not objFace2 Is Nothing) Then
            If FnCheckIfFaceIsConcave(objFace1) Then
                If Not FnCheckIfFaceIsConcave(objFace2) Then
                    FnCheckIfFaceAreOppositeCurvature = True
                    Exit Function
                End If
            Else
                If FnCheckIfFaceIsConcave(objFace2) Then
                    FnCheckIfFaceAreOppositeCurvature = True
                    Exit Function
                End If
            End If
        End If
        FnCheckIfFaceAreOppositeCurvature = False
    End Function
    Public Function FnCheckIfFaceIsConcave(ByVal objFace As Face) As Boolean
        Dim num_Radii As Integer = 0
        Dim radii(1) As Double
        Dim position(5) As Double
        Dim params(3) As Double
        'The magnitude of the radius has a positive sign if the surface is 
        'concave with respect to its normal and a negative sign if the surface is 
        'convex with respect to its normal. 

        FnGetUFSession.Modl.AskFaceMinRadii(objFace.Tag, num_Radii, radii, position, params)
        If radii(0) > 0 Then
            FnCheckIfFaceIsConcave = True
            Exit Function
        Else
            FnCheckIfFaceIsConcave = False
            Exit Function
        End If
        If radii(1) > 0 Then
            FnCheckIfFaceIsConcave = True
            Exit Function
        Else
            FnCheckIfFaceIsConcave = False
            Exit Function
        End If
        FnCheckIfFaceIsConcave = False
    End Function

    'Measure distance between objects
    Function FnGetDistanceBetweenObjects(ByVal objPart As Part, ByVal obj1 As NXObject, _
                                                  ByVal obj2 As NXObject) As Double
        'FnGetDistanceBetweenObjects = objPart.MeasureManager.NewDistance(Nothing, MeasureManager.MeasureType.Minimum, False, obj1, obj2).Value
        Dim adClosePoint1 As Point3d
        Dim adClosePoint2 As Point3d
        Dim dAccuracy As Double

        FnGetDistanceBetweenObjects = FnGetNxSession.Measurement.GetMinimumDistance(obj1, obj2, adClosePoint1, adClosePoint2, dAccuracy)
    End Function
    'Function to get the minimum radius of a face
    Public Function FnGetCylFaceRadius(ByVal objFace As Face) As Double
        'Dim num_Radii As Integer = 0
        'Dim radii(1) As Double
        'Dim position(5) As Double
        'Dim params(3) As Double

        Dim iFaceType As Integer = 0
        Dim adCenterPoint(2) As Double
        Dim adDir(2) As Double
        Dim adBox(5) As Double
        Dim dRadius As Double = 0.0
        Dim dRadData As Double = 0.0
        Dim iNormDir As Integer = 0
        'Dim adVector As Vector3d

        FnGetUFSession.Modl.AskFaceData(objFace.Tag, iFaceType, adCenterPoint, adDir, adBox, dRadius, dRadData, iNormDir)

        'FnGetUFSession.Modl.AskFaceMinRadii(objFace.Tag, num_Radii, radii, position, params)
        FnGetCylFaceRadius = Math.Abs(dRadius)
    End Function
    Function FnGetCylFaceCenter(objFace As Face) As Double()
        Dim iFaceType As Integer = 0
        Dim adCenterPoint(2) As Double
        Dim adDir(2) As Double
        Dim adBox(5) As Double
        Dim dRadius As Double = 0.0
        Dim dRadData As Double = 0.0
        Dim iNormDir As Integer = 0
        FnGetUFSession.Modl.AskFaceData(objFace.Tag, iFaceType, adCenterPoint, adDir, adBox, dRadius, dRadData, iNormDir)
        FnGetCylFaceCenter = adCenterPoint
    End Function
    Public Function FnGetVectorDirCosByTwoPoints(ByVal objPart As Part, ByVal objStartPoint As NXOpen.Point, ByVal objEndPoint As NXOpen.Point) As Double()
        Dim dir As Direction = Nothing
        dir = objPart.Directions.CreateDirection(objStartPoint, objEndPoint, SmartObject.UpdateOption.WithinModeling)
        FnGetVectorDirCosByTwoPoints = {dir.Vector.X, dir.Vector.Y, dir.Vector.Z}
    End Function
    'Measure Angle Using vectors
    Public Function FnGetAngleByVectors(ByVal objPart As Part, ByVal objVector1 As NXObject, _
                                                     ByVal objVector2 As NXObject, ByVal bMinorAngle As Boolean) As Double

        FnGetAngleByVectors = objPart.MeasureManager.NewAngle(Nothing, objVector1, MeasureManager.EndpointType.EndPoint, _
                                                                            objVector2, MeasureManager.EndpointType.EndPoint, bMinorAngle).Value

    End Function
    'Get the direction vector of a face
    Public Function FnGetDirOfFace(ByVal objPart As Part, ByVal ObjFace As Face) As Direction
        Dim dir As Direction = Nothing
        dir = objPart.Directions.CreateDirection(ObjFace, Sense.Forward, SmartObject.UpdateOption.WithinModeling)
        FnGetDirOfFace = dir
    End Function

    Function FnGetDirVecOfALinearEdge(objPart As Part, objEdge As Edge) As Double()
        Dim objVertex1 As Point3d = Nothing
        Dim objVertex2 As Point3d = Nothing
        Dim objCreatedVrtPt1 As NXOpen.Point = Nothing
        Dim objCreatedVrtPt2 As NXOpen.Point = Nothing

        If Not objPart Is Nothing Then
            If Not objEdge Is Nothing Then

                objEdge.GetVertices(objVertex1, objVertex2)
                objCreatedVrtPt1 = objPart.Points.CreatePoint(objVertex1)
                objCreatedVrtPt2 = objPart.Points.CreatePoint(objVertex2)
                FnGetDirVecOfALinearEdge = FnGetVectorDirCosByTwoPoints(objPart, objCreatedVrtPt1, objCreatedVrtPt2)
                Exit Function
            End If
        End If
        FnGetDirVecOfALinearEdge = Nothing
    End Function
End Module
