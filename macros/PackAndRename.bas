'----------------------------------------------
' PackAndRename.bas
' VBA macro for SOLIDWORKS to perform Pack & Go
' with automated renaming of parts, subassemblies,
' and associated drawings.
'----------------------------------------------

Option Explicit

' Entry point for the macro
Public Sub Main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No active document. Please open an assembly.", vbCritical, "Pack & Rename"
        Exit Sub
    End If

    If swModel.GetType <> swDocASSEMBLY Then
        MsgBox "Active document is not an assembly. Open the top-level assembly and try again.", vbCritical, "Pack & Rename"
        Exit Sub
    End If

    If Len(swModel.GetPathName) = 0 Then
        MsgBox "Please save the assembly before running Pack & Rename so the output location can be resolved.", vbExclamation, "Pack & Rename"
        Exit Sub
    End If

    ' Show the configuration form explicitly on a new instance to avoid
    ' issues where the default form instance is not created by SOLIDWORKS
    ' (e.g. when macros are imported in oudere builds). Using a dedicated
    ' instance guarantees the pop-up is shown.
    Dim dialog As PackAndRenameForm
    Set dialog = New PackAndRenameForm
    dialog.InitializeWithDefaults swModel
    dialog.Show vbModal
End Sub

' Perform pack and go with renaming based on user input
Public Sub ExecutePackAndRename(ByVal swModel As SldWorks.ModelDoc2, _
                                ByVal prefix As String, _
                                ByVal outputFolder As String, _
                                ByVal includeDrawings As Boolean)

    Dim swPack As SldWorks.PackAndGo
    Dim docNames As Variant
    Dim modelPath As String
    Dim i As Long
    Dim docCount As Long

    modelPath = swModel.GetPathName
    If Len(modelPath) = 0 Then
        MsgBox "Please save the assembly before running Pack & Rename so the output location can be resolved.", vbExclamation, "Pack & Rename"
        Exit Sub
    End If

    If Len(Dir$(outputFolder, vbDirectory)) = 0 Then
        MsgBox "Output folder does not exist: " & outputFolder, vbCritical, "Pack & Rename"
        Exit Sub
    End If

    Set swPack = swModel.Extension.GetPackAndGo
    swPack.IncludeDrawings = includeDrawings
    swPack.FlattenToSingleFolder = True

    ' Define target folder
    swPack.SetSaveToName True, outputFolder

    ' Get current document list
    docCount = swPack.GetDocumentNamesCount
    docNames = swPack.GetDocumentNames

    ' Map new names
    Dim partMap As Object
    Dim asmMap As Object
    Dim drawingMap As Object
    Dim usedDrawingNames As Object
    Dim partCounter As Long
    Dim asmCounter As Long
    Dim drawingCounter As Long
    Dim baseModelName As String
    Dim currentPath As String
    Dim newBase As String
    Dim ext As String

    Set partMap = CreateObject("Scripting.Dictionary")
    Set asmMap = CreateObject("Scripting.Dictionary")
    Set drawingMap = CreateObject("Scripting.Dictionary")
    Set usedDrawingNames = CreateObject("Scripting.Dictionary")

    partCounter = 1
    asmCounter = 1
    drawingCounter = 1

    For i = 0 To docCount - 1
        currentPath = docNames(i)
        ext = LCase$(GetExtension(currentPath))

        Select Case ext
            Case "sldprt"
                If Not partMap.Exists(currentPath) Then
                    newBase = BuildNumber(prefix, "P", partCounter)
                    partMap.Add currentPath, newBase
                    partCounter = partCounter + 1
                End If
            Case "sldasm"
                If StrComp(currentPath, modelPath, vbTextCompare) = 0 Then
                    newBase = prefix & "-A00"
                    asmMap.Add currentPath, newBase
                ElseIf Not asmMap.Exists(currentPath) Then
                    newBase = BuildNumber(prefix, "A", asmCounter)
                    asmMap.Add currentPath, newBase
                    asmCounter = asmCounter + 1
                End If
            Case "slddrw"
                ' Drawing names align with their referenced model when possible
                baseModelName = GetBaseName(currentPath)
                drawingMap(currentPath) = baseModelName
        End Select
    Next i

    ' Resolve drawing names using corresponding model map when possible
    Dim key As Variant
    For Each key In drawingMap.Keys
        baseModelName = drawingMap(key)
        newBase = ResolveDrawingBase(baseModelName, partMap, asmMap)
        If Len(newBase) = 0 Then
            newBase = BuildNumber(prefix, "D", drawingCounter)
            drawingCounter = drawingCounter + 1
        End If
        newBase = EnsureUniqueDrawingName(newBase, usedDrawingNames, drawingCounter)
        drawingMap(key) = newBase
    Next key

    ' Build final save paths
    Dim newPaths() As String
    ReDim newPaths(0 To docCount - 1)

    For i = 0 To docCount - 1
        currentPath = docNames(i)
        ext = LCase$(GetExtension(currentPath))

        Select Case ext
            Case "sldprt"
                newBase = partMap(currentPath)
            Case "sldasm"
                newBase = asmMap(currentPath)
            Case "slddrw"
                newBase = drawingMap(currentPath)
            Case Else
                newBase = GetBaseName(currentPath)
        End Select

        newPaths(i) = BuildTargetPath(outputFolder, newBase, ext)
    Next i

    ' Apply new names and save
    swPack.SetDocumentSaveToNames newPaths

    Dim result As Boolean
    result = swModel.Extension.SavePackAndGo(swPack)

    If result Then
        MsgBox "Pack & Go completed successfully to: " & outputFolder, vbInformation, "Pack & Rename"
    Else
        MsgBox "Pack & Go failed. Verify file access and try again.", vbCritical, "Pack & Rename"
    End If
End Sub

Private Function BuildNumber(ByVal prefix As String, ByVal designator As String, ByVal number As Long) As String
    BuildNumber = prefix & "-" & designator & Format$(number, "00")
End Function

Private Function BuildTargetPath(ByVal folder As String, ByVal baseName As String, ByVal ext As String) As String
    If Right$(folder, 1) = "\" Or Right$(folder, 1) = "/" Then
        BuildTargetPath = folder & baseName & "." & ext
    Else
        BuildTargetPath = folder & "\" & baseName & "." & ext
    End If
End Function

Private Function GetExtension(ByVal path As String) As String
    Dim parts As Variant
    parts = Split(path, ".")
    If UBound(parts) >= 1 Then
        GetExtension = parts(UBound(parts))
    Else
        GetExtension = ""
    End If
End Function

Public Function GetBaseName(ByVal path As String) As String
    Dim namePart As String
    namePart = Mid$(path, InStrRev(path, "\") + 1)
    If InStr(namePart, ".") > 0 Then
        namePart = Left$(namePart, InStrRev(namePart, ".") - 1)
    End If
    GetBaseName = namePart
End Function

Private Function ResolveDrawingBase(ByVal baseModelName As String, _
                                    ByVal partMap As Object, _
                                    ByVal asmMap As Object) As String
    Dim key As Variant

    For Each key In partMap.Keys
        If StrComp(GetBaseName(CStr(key)), baseModelName, vbTextCompare) = 0 Then
            ResolveDrawingBase = partMap(key)
            Exit Function
        End If
    Next key

    For Each key In asmMap.Keys
        If StrComp(GetBaseName(CStr(key)), baseModelName, vbTextCompare) = 0 Then
            ResolveDrawingBase = asmMap(key)
            Exit Function
        End If
    Next key

    ResolveDrawingBase = ""
End Function

Private Function EnsureUniqueDrawingName(ByVal candidate As String, _
                                          ByVal usedNames As Object, _
                                          ByRef drawingCounter As Long) As String
    Dim uniqueName As String

    uniqueName = candidate

    If usedNames.Exists(LCase$(uniqueName)) Then
        Do
            uniqueName = candidate & "-DWG" & Format$(drawingCounter, "00")
            drawingCounter = drawingCounter + 1
        Loop While usedNames.Exists(LCase$(uniqueName))
    End If

    usedNames.Add LCase$(uniqueName), True
    EnsureUniqueDrawingName = uniqueName
End Function
