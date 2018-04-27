Attribute VB_Name = "Export"
Public Const vbDoubleQuote As String = """"
 
Public Sub ExportToSTEP(Doc As Document)
    ' Get the STEP translator Add-In.
    Dim oSTEPTranslator As TranslatorAddIn
    Set oSTEPTranslator = ThisApplication.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

    If oSTEPTranslator Is Nothing Then
        MsgBox "Could not access STEP translator."
        Exit Sub
    End If

    Dim oContext As TranslationContext
    Set oContext = ThisApplication.TransientObjects.CreateTranslationContext
    Dim oOptions As NameValueMap
    Set oOptions = ThisApplication.TransientObjects.CreateNameValueMap
    If oSTEPTranslator.HasSaveCopyAsOptions(Doc, oContext, oOptions) Then
        ' Set application protocol.
        ' 2 = AP 203 - Configuration Controlled Design
        ' 3 = AP 214 - Automotive Design
        oOptions.value("ApplicationProtocolType") = 3

        ' Other options...
        'oOptions.Value("Author") = ""
        'oOptions.Value("Authorization") = ""
        'oOptions.Value("Description") = ""
        'oOptions.Value("Organization") = ""

        oContext.Type = kFileBrowseIOMechanism

        Dim oData As DataMedium
        Set oData = ThisApplication.TransientObjects.CreateDataMedium
        
        Dim fileName As String
        fileName = Left(Doc.DisplayName(), (Len(Doc.DisplayName()) - 4))
        oData.fileName = ThisApplication.DesignProjectManager.ActiveDesignProject.workspacePath & "\EXPORT\" & fileName & ".stp"
        Debug.Print ("Export to: " & oData.fileName)
        Call oSTEPTranslator.SaveCopyAs(Doc, oContext, oOptions, oData)
    End If
End Sub

Public Sub ExportActiveToSTEP()
    ' Get the STEP translator Add-In.
    Dim oSTEPTranslator As TranslatorAddIn
    Set oSTEPTranslator = ThisApplication.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

    If oSTEPTranslator Is Nothing Then
        MsgBox "Could not access STEP translator."
        Exit Sub
    End If

    Dim oContext As TranslationContext
    Set oContext = ThisApplication.TransientObjects.CreateTranslationContext
    Dim oOptions As NameValueMap
    Set oOptions = ThisApplication.TransientObjects.CreateNameValueMap
    If oSTEPTranslator.HasSaveCopyAsOptions(ThisApplication.ActiveDocument, oContext, oOptions) Then
        ' Set application protocol.
        ' 2 = AP 203 - Configuration Controlled Design
        ' 3 = AP 214 - Automotive Design
        oOptions.value("ApplicationProtocolType") = 3

        ' Other options...
        'oOptions.Value("Author") = ""
        'oOptions.Value("Authorization") = ""
        'oOptions.Value("Description") = ""
        'oOptions.Value("Organization") = ""

        oContext.Type = kFileBrowseIOMechanism

        Dim oData As DataMedium
        Set oData = ThisApplication.TransientObjects.CreateDataMedium
        oData.fileName = "D:\tmp\temptest.stp"

        Call oSTEPTranslator.SaveCopyAs(ThisApplication.ActiveDocument, oContext, oOptions, oData)
    End If
End Sub

Public Sub AppendMaterial(name As String, id As Integer, young As Double, poisson As Double, density As Double)
    Dim fileName As String
    fileName = ThisApplication.DesignProjectManager.ActiveDesignProject.workspacePath & "\EXPORT\" & "Materials" & ".txt"
    Open fileName For Append As #1
    Print #1, "material create &"
    Print #1, vbTab & "material_name = .model_1." & name & " &"
    Print #1, vbTab & "adams_id = " & id & " &"
    Print #1, vbTab & "youngs_modulus = " & young & " &"
    Print #1, vbTab & "poissons_ratio = " & poisson & " &"
    Print #1, vbTab & "density = " & density
    Print #1, "!"
    Close #1
End Sub

Public Sub AppendRigidBody(name As String, id As Integer, location As Vector, orientation() As Double)
    Dim fileName As String
    fileName = ThisApplication.DesignProjectManager.ActiveDesignProject.workspacePath & "\EXPORT\" & "RigidBodies" & ".txt"
    Open fileName For Append As #1
    'Default coordinate system to ground
    Print #1, "defaults coordinate_system  &"
    Print #1, "default_coordinate_system = .model_1.ground"
    Print #1, "!"
    Print #1, "part create rigid_body name_and_position &"
    Print #1, vbTab & "part_name = .model_1." & name & " &"
    Print #1, vbTab & "adams_id = " & id & " &"
    Print #1, vbTab & "location = " & location.x & ", " & location.y & ", " & location.Z & " &"
    Print #1, vbTab & "orientation = " & orientation(0) & "d, " & orientation(1) & "d, " & orientation(2) & "d"
    Print #1, "!"
    Print #1, "defaults coordinate_system &"
    Print #1, vbTab & "default_coordinate_system = .model_1." & name
    Print #1, "!"
    
    Close #1
    
    
End Sub

Public Sub AppendMassProperties(name As String, materialName As String)
    Dim fileName As String
    fileName = ThisApplication.DesignProjectManager.ActiveDesignProject.workspacePath & "\EXPORT\" & "RigidBodies" & ".txt"
    Open fileName For Append As #1
    Print #1, "!"
    Print #1, "part create rigid_body mass_properties &"
    Print #1, vbTab & "part_name = .model_1." & name & " &"
    Print #1, vbTab & "material_type = .model_1." & materialName
    Close #1
End Sub
Public Sub AppendGeometryProperties(name As String, geometryFile As String, geoName As String)
    Dim fileName As String
    fileName = ThisApplication.DesignProjectManager.ActiveDesignProject.workspacePath & "\EXPORT\" & "RigidBodies" & ".txt"
    Open fileName For Append As #1
    Print #1, "!"
    Print #1, "file geometry read file_name = " & vbDoubleQuote & geometryFile & vbDoubleQuote & "  &"
    Print #1, vbTab & "part_name = .model_1." & name & " &"
    Print #1, vbTab & "single_shell = no &"
    Print #1, vbTab & "create_geometry = solid &"
    Print #1, vbTab & "type_of_geometry = stp"
    Print #1, "!"
    Print #1, "geometry attributes &"
    Print #1, vbTab & "geometry_name = .model_1." & name & "." & geoName; " &"
    Print #1, vbTab & "color = BLUE"
    Print #1, "!"
    Close #1
End Sub


