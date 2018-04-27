Attribute VB_Name = "Main"
Sub Main()
Call clearDebugConsole
Debug.Print ("Begining")

Dim oDoc As AssemblyDocument
Dim oApp As Inventor.Application
Dim compDef As AssemblyComponentDefinition

Set oApp = ThisApplication
Set oDoc = oApp.ActiveDocument
Set compDef = oDoc.ComponentDefinition

' Get all of the referenced documents.
Dim oRefDocs As DocumentsEnumerator
Set oRefDocs = oDoc.AllReferencedDocuments

Dim exportPath As String
Dim fileName As String
Dim outputFile As String
exportPath = ThisApplication.DesignProjectManager.ActiveDesignProject.workspacePath & "\EXPORT\"
outputFile = exportPath & "output.cmd"

' ------------------------------------ '
'First erase previous tmp files
Open outputFile For Output As #1
    Print #1, "!created by ossip"
Close #1
fileName = exportPath & "Materials" & ".txt"
Open fileName For Output As #1
Close #1
fileName = exportPath & "RigidBodies" & ".txt"
Open fileName For Output As #1
Close #1
'--------------------------------------'

' Go through the list of documents and export each stp in EXPORT folder
' Also populate Material File (asume that each ipt reference has the same material)
Dim oRefDoc As PartDocument
Dim materials() As String
ReDim materials(oRefDocs.Count)
Dim numMaterials As Integer
numMaterials = 0
Debug.Print ("Looking for referenced files and materials")
For Each oRefDoc In oRefDocs
    Debug.Print ("File: " & oRefDoc.DisplayName)
    Dim materialName As String
    materialName = oRefDoc.ActiveMaterial.DisplayName
    Debug.Print ("Material: " & materialName)

    Dim oldMaterial As Boolean
    oldMaterial = False
    For Index = 0 To numMaterials
        If materials(Index) = materialName Then
            oldMaterial = True
        End If
    Next Index
    If oldMaterial = False Then
        'Write material to file
        Call AppendMaterial(materialName, numMaterials + 1, 207000, 0.29, 0.000007801) 'TODO get properties from Inventor
        materials(numMaterials) = materialName
        numMaterials = numMaterials + 1
    End If
    
    Dim libAsset As Asset
    Set libAsset = ThisApplication.ActiveMaterialLibrary.MaterialAssets.Item(materialName)
    Call ExportToSTEP(oRefDoc)
Next

' Search for the leaves
''Show number of parts
Debug.Print ("")
Debug.Print ("Saving rigid bodies' configuration")
Dim oLeafOccs As ComponentOccurrencesEnumerator
    Set oLeafOccs = compDef.Occurrences.AllLeafOccurrences

    ' Iterate through the occurrences and print the name.
    Dim oOcc As ComponentOccurrence

    Dim id As Integer
    id = 2
    For Each oOcc In oLeafOccs
        Dim componentName As String
        componentName = Replace(oOcc.name, ":", "_")
        
        Debug.Print ("Export: " & componentName)
        '_Pose_
        'Position:
        Dim trans As Vector
        Set trans = oOcc.Transformation.Translation
        trans.ScaleBy (10)
        Debug.Print (trans.x & " " & trans.y & " " & trans.Z)
        
        
        'Orientation:
        Dim angles(2) As Double
        Call CalculateRotationAngles(oOcc.Transformation, angles)
        
        Call AppendRigidBody(componentName, id, trans, angles)
        'Append markers ' TO DO TO DO TO DO
        'Append mass properties
        Call AppendMassProperties(componentName, "Generic")
        'Append Geometry Properties
        Dim geoFile As String
        Dim geoName As String
          
        Dim part As PartDocument
        Set part = oOcc.Definition.Document
        geoName = Left(part.DisplayName, Len(part.DisplayName) - 4)
        geoFile = geoName & ".stp"
        Debug.Print ("Geometry file: " & geoFile)
        
        
        Call AppendGeometryProperties(componentName, geoFile, "SOLID" & (id - 1))
        
        id = id + 1
        

    Next
    

'Join files to output.cmd
fileName = exportPath & "Base1" & ".txt"
Call AppendFiles(outputFile, fileName)
fileName = exportPath & "Materials" & ".txt"
Call AppendFiles(outputFile, fileName)
fileName = exportPath & "Base2" & ".txt"
Call AppendFiles(outputFile, fileName)
fileName = exportPath & "RigidBodies" & ".txt"
Call AppendFiles(outputFile, fileName)



End Sub

