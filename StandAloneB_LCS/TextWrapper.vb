Imports System.IO

Module TextWrapper
    Public sPartPath As String = ""

    Public Sub SWrite(ByVal sText As String, ByVal sFileNamePath As String)
        If Not File.Exists(sFileNamePath) = True Then
            Dim theFile As FileStream = File.Create(sFileNamePath)
            theFile.Close()
        End If
        'Append text to the file
        Dim objWriter As New StreamWriter(sFileNamePath, True)
        objWriter.WriteLine(sText)
        objWriter.Close()
        objWriter = Nothing
    End Sub
    Public Function FnCheckFileExists(ByVal sFileNamePath As String) As Boolean
        If File.Exists(sFileNamePath) = True Then
            FnCheckFileExists = True
        Else
            FnCheckFileExists = False
        End If
    End Function
    Public Sub SClearContentsOfFile(ByVal sFilePath As String)
        Dim stream As New StreamWriter(sFilePath, False)
        stream.WriteLine("")
        stream.Close()
    End Sub
    Public Sub SDeleteFile(ByVal sFilePath As String)
        If File.Exists(sFilePath) = True Then
            File.Delete(sFilePath)
        End If
    End Sub
    Public Function FnReadFile(ByVal sFileNamePath As String, ByVal iLineNumberToRead As Integer) As String
        Dim oRead As StreamReader
        Dim LineIn As String
        Dim iLineCounter As Integer
        oRead = File.OpenText(sFileNamePath)
        iLineCounter = 0
        While oRead.Peek <> -1
            iLineCounter = iLineCounter + 1
            LineIn = oRead.ReadLine()
            If iLineCounter = iLineNumberToRead Then
                oRead.Close()
                Return LineIn
            End If
        End While
        oRead.Close()
    End Function
    Public Function FnGetLineNumberInFile(ByVal sFileNamePath As String, ByVal sLineToFind As String) As Integer
        Dim oRead As StreamReader
        Dim LineIn As String
        Dim iLineCounter As Integer
        oRead = File.OpenText(sFileNamePath)
        iLineCounter = 0
        While oRead.Peek <> -1
            iLineCounter = iLineCounter + 1
            LineIn = oRead.ReadLine()
            If LineIn = sLineToFind Then
                oRead.Close()
                Return iLineCounter
            End If
        End While
        Return 0
        oRead.Close()
    End Function
    Public Function FnReadFullFile(ByVal sFileNamePath As String) As String()
        Dim oRead As StreamReader
        Dim sLines() As String = Nothing
        Dim iLineCounter As Integer
        oRead = File.OpenText(sFileNamePath)
        iLineCounter = 0
        While oRead.Peek <> -1
            ReDim Preserve sLines(iLineCounter)
            sLines(iLineCounter) = oRead.ReadLine()
            iLineCounter = iLineCounter + 1
        End While
        oRead.Close()
        FnReadFullFile = sLines
    End Function
    Public Function SCreateDirectory(ByVal sPath As String) As Boolean
        Try
            If Not (Directory.Exists(sPath)) Then
                Directory.CreateDirectory(sPath)
                Return True
            Else
                Return True
            End If
        Catch E As Exception
            sWriteToLogFile("Error creating directory")
            Return False
        End Try
    End Function

    Public Function FnCheckFolderExists(ByVal sFolderPath As String) As Boolean
        If Not (Directory.Exists(sFolderPath)) Then
            Return False
        Else
            Return True
        End If
    End Function


    Public Function FnGetFilePathWithinFolderRecursively(ByVal sFolderPath As String, ByVal fileExt As String) As FileInfo()
        sPartPath = ""
        Dim dir As New DirectoryInfo(sFolderPath)
        Dim listofFiles() As FileInfo = Nothing
        If System.IO.Directory.GetDirectories(sFolderPath).Length > 0 Then
            For Each Directory As DirectoryInfo In dir.GetDirectories()
                If Directory.GetDirectories().Length > 0 Then
                    FnFindFolderWithinFolder(Directory.FullName, fileExt, listofFiles)
                End If
                For Each file As FileInfo In Directory.GetFiles()
                    If file.Extension = fileExt Then
                        If listofFiles Is Nothing Then
                            ReDim Preserve listofFiles(0)
                            listofFiles(0) = file
                        Else
                            Dim iCount As Integer = UBound(listofFiles)
                            ReDim Preserve listofFiles(iCount + 1)
                            listofFiles(iCount + 1) = file
                        End If
                    End If
                Next
            Next
        End If
        For Each file As FileInfo In dir.GetFiles()
            If file.Extension = fileExt Then
                If listofFiles Is Nothing Then
                    ReDim Preserve listofFiles(0)
                    listofFiles(0) = file
                Else
                    Dim iCount As Integer = UBound(listofFiles)
                    ReDim Preserve listofFiles(iCount + 1)
                    listofFiles(iCount + 1) = file
                End If
            End If
        Next
        FnGetFilePathWithinFolderRecursively = listofFiles
    End Function
    Public Function FnGetPartPath() As String
        If sPartPath <> "" Then
            Return sPartPath
        End If
    End Function
    Public Sub FnFindFolderWithinFolder(ByVal sFolderPath As String, ByVal fileExt As String, ByRef listofFiles() As FileInfo)
        Dim dir As New DirectoryInfo(sFolderPath)
        If System.IO.Directory.GetDirectories(sFolderPath).Length > 0 Then
            For Each Directory As DirectoryInfo In dir.GetDirectories()
                For Each file As FileInfo In Directory.GetFiles()
                    If file.Extension = fileExt Then
                        If listofFiles Is Nothing Then
                            ReDim Preserve listofFiles(0)
                            listofFiles(0) = file
                        Else
                            Dim iCount As Integer = UBound(listofFiles)
                            ReDim Preserve listofFiles(iCount + 1)
                            listofFiles(iCount + 1) = file
                        End If
                    End If
                Next
                If Directory.GetDirectories().Length > 0 Then
                    Call FnFindFolderWithinFolder(Directory.FullName, fileExt, listofFiles)
                End If
            Next
        End If
    End Sub

    Public Sub SDeleteDirectory(ByVal sPath As String, Optional ByVal rec As Boolean = True)
        Directory.Delete(sPath, rec)
    End Sub

    Public Sub SCopyFiles(ByVal DestinationPath As String, Optional ByVal sourcePath As String = "", Optional ByVal sFilePath As String = "")

        If Not sFilePath = "" Then
            If File.Exists(sFilePath) Then
                Dim dFile As String = String.Empty
                dFile = Path.GetFileName(sFilePath)
                Dim dFilePath As String = String.Empty
                dFilePath = DestinationPath + "\" + dFile
                File.Copy(sFilePath, dFilePath, True)
            End If
        Else
            If (Directory.Exists(sourcePath)) Then
                For Each fName As String In Directory.GetFiles(sourcePath)
                    If File.Exists(fName) Then
                        Dim dFile As String = String.Empty
                        dFile = Path.GetFileName(fName)
                        Dim dFilePath As String = String.Empty
                        dFilePath = DestinationPath + "\" + dFile
                        File.Copy(fName, dFilePath, True)
                    End If
                Next
            End If
        End If

    End Sub

    Public Sub SRenameFile(ByVal sFilePath As String, ByVal sNewFileName As String)
        If FnCheckFileExists(sFilePath) Then
            'Sample - My.Computer.FileSystem.RenameFile("C:\Test.txt", "SecondTest.txt")
            My.Computer.FileSystem.RenameFile(sFilePath, sNewFileName)
        End If
    End Sub
End Module
