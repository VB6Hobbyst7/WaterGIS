Imports ESRI.ArcGIS.Geodatabase

Public Class CreateTablesScenarios




    Public Function CreateTableScenariosWithPipeOrJunc(ByVal workspace As IWorkspace, ByVal className As String) As IObjectClass
        ' Cast the workspace to IFeatureWorkspace and IWorkspace2. An explicit cast is not used
        ' for IWorkspace2 because not all workspaces implement it.
        Dim featureWorkspace As IFeatureWorkspace = CType(workspace, IFeatureWorkspace)
        Dim workspace2 As IWorkspace2 = TryCast(workspace, IWorkspace2)

        ' See if a class with the provided name already exists. If so, return it.
        ' Note that this will not work with file-based workspaces, such as shapefile workspaces.
        If Not workspace2 Is Nothing Then
            If workspace2.NameExists(esriDatasetType.esriDTTable, className) Then
                Dim existingTable As ITable = featureWorkspace.OpenTable(className)
                Return CType(existingTable, IObjectClass)
            End If
        End If

        ' Use IFieldChecker.ValidateTableName to validate the name of the class. The tableNameErrorType
        ' variable is not used in this example, but it indicates the reasons the table name was invalid,
        ' if any.
        Dim fieldChecker As IFieldChecker = New FieldCheckerClass()
        Dim validatedClassName As String = Nothing
        fieldChecker.ValidateWorkspace = workspace
        Dim tableNameErrorType As Integer = fieldChecker.ValidateTableName(className, validatedClassName)

        ' Create an object class description and get the required fields from it.
        Dim ocDescription As IObjectClassDescription = New ObjectClassDescriptionClass()
        Dim fields As IFields = ocDescription.RequiredFields
        Dim fieldsEdit As IFieldsEdit = CType(fields, IFieldsEdit)

        ' Add the Name field to the required fields.
        Dim NumOfScenario As IField = New FieldClass()
        Dim NumOfScenarioEdit As IFieldEdit = CType(NumOfScenario, IFieldEdit)
        NumOfScenarioEdit.Name_2 = "NumScenario"
        NumOfScenarioEdit.AliasName_2 = "IDScenario"
        NumOfScenarioEdit.Required_2 = True
        NumOfScenarioEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        fieldsEdit.AddField(NumOfScenario)

        ' Add the Name field to the required fields.
        Dim NumOfPipe As IField = New FieldClass()
        Dim NumOfPipeEdit As IFieldEdit = CType(NumOfPipe, IFieldEdit)
        NumOfPipeEdit.Name_2 = "NumPipe"
        NumOfPipeEdit.AliasName_2 = "IDPipe"
        NumOfPipeEdit.Required_2 = True
        NumOfPipeEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        fieldsEdit.AddField(NumOfPipe)

        ' Add the Diameter field to the required fields.
        Dim DiaField As IField = New FieldClass()
        Dim DiaFieldEdit As IFieldEdit = CType(DiaField, IFieldEdit)
        DiaFieldEdit.Name_2 = "DiameterPipe"
        DiaFieldEdit.AliasName_2 = "DiameterPipe"
        DiaFieldEdit.Required_2 = True
        DiaFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(DiaField)

        ' Add the Length field to the required fields.
        Dim LengthField As IField = New FieldClass()
        Dim LengthFieldEdit As IFieldEdit = CType(LengthField, IFieldEdit)
        LengthFieldEdit.Name_2 = "LengthPipe"
        LengthFieldEdit.AliasName_2 = "LengthPipe"
        LengthFieldEdit.Required_2 = True
        LengthFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(LengthField)

        ' Add the Hazen field to the required fields.
        Dim HazenCField As IField = New FieldClass()
        Dim HazenCFieldEdit As IFieldEdit = CType(HazenCField, IFieldEdit)
        HazenCFieldEdit.Name_2 = "HazenCPipe"
        HazenCFieldEdit.AliasName_2 = "HazenCPipe"
        HazenCFieldEdit.Required_2 = True
        HazenCFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        fieldsEdit.AddField(HazenCField)

        ' Add the Material field to the required fields.
        Dim MaterialField As IField = New FieldClass()
        Dim MaterialFieldEdit As IFieldEdit = CType(MaterialField, IFieldEdit)
        MaterialFieldEdit.Name_2 = "MaterialPipe"
        MaterialFieldEdit.AliasName_2 = "MaterialPipe"
        MaterialFieldEdit.Required_2 = True
        MaterialFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        fieldsEdit.AddField(MaterialField)

        ' Add the NumJunc field to the required fields.
        Dim NumOfJunc As IField = New FieldClass()
        Dim NumOfJuncEdit As IFieldEdit = CType(NumOfJunc, IFieldEdit)
        NumOfJuncEdit.Name_2 = "NumJunc"
        NumOfJuncEdit.AliasName_2 = "IDJunc"
        NumOfJuncEdit.Required_2 = True
        NumOfJuncEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        fieldsEdit.AddField(NumOfJunc)

        ' Add the NumJunc field to the required fields.
        Dim ElevationField As IField = New FieldClass()
        Dim ElevationFieldEdit As IFieldEdit = CType(ElevationField, IFieldEdit)
        ElevationFieldEdit.Name_2 = "ElevationJunc"
        ElevationFieldEdit.AliasName_2 = "ElevationJunc"
        ElevationFieldEdit.Required_2 = True
        ElevationFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(ElevationField)

        ' Add the Name field to the required fields.
        Dim NumOfUser As IField = New FieldClass()
        Dim NumOfUserEdit As IFieldEdit = CType(NumOfUser, IFieldEdit)
        NumOfUserEdit.Name_2 = "NumUser"
        NumOfUserEdit.AliasName_2 = "IDUser"
        NumOfUserEdit.Required_2 = True
        NumOfUserEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        fieldsEdit.AddField(NumOfUser)

        ' Add the Masraf field to the required fields.
        Dim MasrafOfUser As IField = New FieldClass()
        Dim MasrafUserEdit As IFieldEdit = CType(MasrafOfUser, IFieldEdit)
        MasrafUserEdit.Name_2 = "MasrafUser"
        MasrafUserEdit.AliasName_2 = "MasrafUser"
        MasrafUserEdit.Required_2 = True
        MasrafUserEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(MasrafOfUser)

        ' Use the field checker created earlier to validate the fields.
        Dim enumFieldError As IEnumFieldError = Nothing
        Dim validatedFields As IFields = Nothing
        fieldChecker.Validate(fields, enumFieldError, validatedFields)

        ' Create and return the object class.
        Dim table As ITable = featureWorkspace.CreateTable(validatedClassName, validatedFields, ocDescription.InstanceCLSID, Nothing, "")
        Return CType(table, IObjectClass)
    End Function


    Public Function CreateTableCompareScenarios(ByVal workspace As IWorkspace, ByVal className As String) As IObjectClass
        ' Cast the workspace to IFeatureWorkspace and IWorkspace2. An explicit cast is not used
        ' for IWorkspace2 because not all workspaces implement it.
        Dim featureWorkspace As IFeatureWorkspace = CType(workspace, IFeatureWorkspace)
        Dim workspace2 As IWorkspace2 = TryCast(workspace, IWorkspace2)

        ' See if a class with the provided name already exists. If so, return it.
        ' Note that this will not work with file-based workspaces, such as shapefile workspaces.
        If Not workspace2 Is Nothing Then
            If workspace2.NameExists(esriDatasetType.esriDTTable, className) Then
                Dim existingTable As ITable = featureWorkspace.OpenTable(className)
                Return CType(existingTable, IObjectClass)
            End If
        End If

        ' Use IFieldChecker.ValidateTableName to validate the name of the class. The tableNameErrorType
        ' variable is not used in this example, but it indicates the reasons the table name was invalid,
        ' if any.
        Dim fieldChecker As IFieldChecker = New FieldCheckerClass()
        Dim validatedClassName As String = Nothing
        fieldChecker.ValidateWorkspace = workspace
        Dim tableNameErrorType As Integer = fieldChecker.ValidateTableName(className, validatedClassName)

        ' Create an object class description and get the required fields from it.
        Dim ocDescription As IObjectClassDescription = New ObjectClassDescriptionClass()
        Dim fields As IFields = ocDescription.RequiredFields
        Dim fieldsEdit As IFieldsEdit = CType(fields, IFieldsEdit)

        ' Add the Name field to the required fields.
        Dim Num1OfPipe As IField = New FieldClass()
        Dim Num1OfPipeEdit As IFieldEdit = CType(Num1OfPipe, IFieldEdit)
        Num1OfPipeEdit.Name_2 = "NumPipe1"
        Num1OfPipeEdit.AliasName_2 = "IDPipe1"
        Num1OfPipeEdit.Required_2 = True
        Num1OfPipeEdit.Type_2 = esriFieldType.esriFieldTypeString
        fieldsEdit.AddField(Num1OfPipe)


        ' Add the Velocity field to the required fields.
        Dim Velocity1Field As IField = New FieldClass()
        Dim Velocity1FieldEdit As IFieldEdit = CType(Velocity1Field, IFieldEdit)
        Velocity1FieldEdit.Name_2 = "VelocityPipe1"
        Velocity1FieldEdit.AliasName_2 = "VelocityPipe1"
        Velocity1FieldEdit.Required_2 = True
        Velocity1FieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Velocity1Field)

        ' Add the Headloss field to the required fields.
        Dim Headloss1Field As IField = New FieldClass()
        Dim Headloss1FieldEdit As IFieldEdit = CType(Headloss1Field, IFieldEdit)
        Headloss1FieldEdit.Name_2 = "Headloss1"
        Headloss1FieldEdit.AliasName_2 = "Headloss1"
        Headloss1FieldEdit.Required_2 = True
        Headloss1FieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Headloss1Field)

        ' Add the Name field to the required fields.
        Dim Num2OfPipe As IField = New FieldClass()
        Dim Num2OfPipeEdit As IFieldEdit = CType(Num2OfPipe, IFieldEdit)
        Num2OfPipeEdit.Name_2 = "NumPipe2"
        Num2OfPipeEdit.AliasName_2 = "IDPipe2"
        Num2OfPipeEdit.Required_2 = True
        Num2OfPipeEdit.Type_2 = esriFieldType.esriFieldTypeString
        fieldsEdit.AddField(Num2OfPipe)

        ' Add the Velocity field to the required fields.
        Dim Velocity2Field As IField = New FieldClass()
        Dim Velocity2FieldEdit As IFieldEdit = CType(Velocity2Field, IFieldEdit)
        Velocity2FieldEdit.Name_2 = "VelocityPipe2"
        Velocity2FieldEdit.AliasName_2 = "VelocityPipe2"
        Velocity2FieldEdit.Required_2 = True
        Velocity2FieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Velocity2Field)


        ' Add the Headloss field to the required fields.
        Dim Headloss2Field As IField = New FieldClass()
        Dim Headloss2FieldEdit As IFieldEdit = CType(Headloss2Field, IFieldEdit)
        Headloss2FieldEdit.Name_2 = "Headloss2"
        Headloss2FieldEdit.AliasName_2 = "Headloss2"
        Headloss2FieldEdit.Required_2 = True
        Headloss2FieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Headloss2Field)

        ' Use the field checker created earlier to validate the fields.
        Dim enumFieldError As IEnumFieldError = Nothing
        Dim validatedFields As IFields = Nothing
        fieldChecker.Validate(fields, enumFieldError, validatedFields)

        ' Create and return the object class.
        Dim table As ITable = featureWorkspace.CreateTable(validatedClassName, validatedFields, ocDescription.InstanceCLSID, Nothing, "")
        Return CType(table, IObjectClass)
    End Function


    Public Function CreateTableScenarios(ByVal workspace As IWorkspace, ByVal className As String) As IObjectClass
        ' Cast the workspace to IFeatureWorkspace and IWorkspace2. An explicit cast is not used
        ' for IWorkspace2 because not all workspaces implement it.
        Dim featureWorkspace As IFeatureWorkspace = CType(workspace, IFeatureWorkspace)
        Dim workspace2 As IWorkspace2 = TryCast(workspace, IWorkspace2)

        ' See if a class with the provided name already exists. If so, return it.
        ' Note that this will not work with file-based workspaces, such as shapefile workspaces.
        If Not workspace2 Is Nothing Then
            If workspace2.NameExists(esriDatasetType.esriDTTable, className) Then
                Dim existingTable As ITable = featureWorkspace.OpenTable(className)
                Return CType(existingTable, IObjectClass)
            End If
        End If

        ' Use IFieldChecker.ValidateTableName to validate the name of the class. The tableNameErrorType
        ' variable is not used in this example, but it indicates the reasons the table name was invalid,
        ' if any.
        Dim fieldChecker As IFieldChecker = New FieldCheckerClass()
        Dim validatedClassName As String = Nothing
        fieldChecker.ValidateWorkspace = workspace
        Dim tableNameErrorType As Integer = fieldChecker.ValidateTableName(className, validatedClassName)

        ' Create an object class description and get the required fields from it.
        Dim ocDescription As IObjectClassDescription = New ObjectClassDescriptionClass()
        Dim fields As IFields = ocDescription.RequiredFields
        Dim fieldsEdit As IFieldsEdit = CType(fields, IFieldsEdit)



        ' Add the NameScenario field to the required fields.
        Dim NameScenario As IField = New FieldClass()
        Dim NameScenarioFieldEdit As IFieldEdit = CType(NameScenario, IFieldEdit)
        NameScenarioFieldEdit.Name_2 = "NameScenario"
        NameScenarioFieldEdit.AliasName_2 = "NameScenario"
        'NameScenarioFieldEdit.Editable_2 = False
        NameScenarioFieldEdit.Required_2 = True
        NameScenarioFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        fieldsEdit.AddField(NameScenario)

        ' Add the TypeScenario field to the required fields.
        Dim TypeScenario As IField = New FieldClass()
        Dim TypeScenarioFieldEdit As IFieldEdit = CType(TypeScenario, IFieldEdit)
        TypeScenarioFieldEdit.Name_2 = "TypeScenario"
        TypeScenarioFieldEdit.AliasName_2 = "TypeScenario"
        TypeScenarioFieldEdit.Required_2 = True
        TypeScenarioFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        fieldsEdit.AddField(TypeScenario)

        ' Add the DateScenario field to the required fields.
        Dim DateScenario As IField = New FieldClass()
        Dim DateScenarioFieldEdit As IFieldEdit = CType(DateScenario, IFieldEdit)
        DateScenarioFieldEdit.Name_2 = "DateScenario"
        DateScenarioFieldEdit.AliasName_2 = "DateScenario"
        DateScenarioFieldEdit.Required_2 = True
        DateScenarioFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        fieldsEdit.AddField(DateScenario)

        ' Use the field checker created earlier to validate the fields.
        Dim enumFieldError As IEnumFieldError = Nothing
        Dim validatedFields As IFields = Nothing
        fieldChecker.Validate(fields, enumFieldError, validatedFields)

        ' Create and return the object class.
        Dim table As ITable = featureWorkspace.CreateTable(validatedClassName, validatedFields, ocDescription.InstanceCLSID, Nothing, "")
        Return CType(table, IObjectClass)
    End Function


    Public Function CreateTableEdgeCompareScenarios(ByVal workspace As IWorkspace, ByVal className As String) As IObjectClass
        ' Cast the workspace to IFeatureWorkspace and IWorkspace2. An explicit cast is not used
        ' for IWorkspace2 because not all workspaces implement it.
        Dim featureWorkspace As IFeatureWorkspace = CType(workspace, IFeatureWorkspace)
        Dim workspace2 As IWorkspace2 = TryCast(workspace, IWorkspace2)

        ' See if a class with the provided name already exists. If so, return it.
        ' Note that this will not work with file-based workspaces, such as shapefile workspaces.
        If Not workspace2 Is Nothing Then
            If workspace2.NameExists(esriDatasetType.esriDTTable, className) Then
                Dim existingTable As ITable = featureWorkspace.OpenTable(className)
                Return CType(existingTable, IObjectClass)
            End If
        End If

        ' Use IFieldChecker.ValidateTableName to validate the name of the class. The tableNameErrorType
        ' variable is not used in this example, but it indicates the reasons the table name was invalid,
        ' if any.
        Dim fieldChecker As IFieldChecker = New FieldCheckerClass()
        Dim validatedClassName As String = Nothing
        fieldChecker.ValidateWorkspace = workspace
        Dim tableNameErrorType As Integer = fieldChecker.ValidateTableName(className, validatedClassName)

        ' Create an object class description and get the required fields from it.
        Dim ocDescription As IObjectClassDescription = New ObjectClassDescriptionClass()
        Dim fields As IFields = ocDescription.RequiredFields
        Dim fieldsEdit As IFieldsEdit = CType(fields, IFieldsEdit)

        ' Add the Name field to the required fields.
        Dim NumOfPipe As IField = New FieldClass()
        Dim NumOfPipeEdit As IFieldEdit = CType(NumOfPipe, IFieldEdit)
        NumOfPipeEdit.Name_2 = "NumPipe"
        NumOfPipeEdit.AliasName_2 = "IDPipe"
        NumOfPipeEdit.Required_2 = True
        NumOfPipeEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
        fieldsEdit.AddField(NumOfPipe)

        ' Add the Length field to the required fields.
        Dim LengthField1 As IField = New FieldClass()
        Dim LengthFieldEdit1 As IFieldEdit = CType(LengthField1, IFieldEdit)
        LengthFieldEdit1.Name_2 = "LengthPipe1"
        LengthFieldEdit1.AliasName_2 = "LengthPipe1"
        LengthFieldEdit1.Required_2 = True
        LengthFieldEdit1.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(LengthField1)

        ' Add the Diameter field to the required fields.
        Dim DiaField1 As IField = New FieldClass()
        Dim DiaFieldEdit1 As IFieldEdit = CType(DiaField1, IFieldEdit)
        DiaFieldEdit1.Name_2 = "DiameterPipe1"
        DiaFieldEdit1.AliasName_2 = "DiameterPipe1"
        DiaFieldEdit1.Required_2 = True
        DiaFieldEdit1.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(DiaField1)

        ' Add the Type field to the required fields.
        Dim TypeField1 As IField = New FieldClass()
        Dim TypeFieldEdit1 As IFieldEdit = CType(TypeField1, IFieldEdit)
        TypeFieldEdit1.Name_2 = "TypePipe1"
        TypeFieldEdit1.AliasName_2 = "TypePipe1"
        TypeFieldEdit1.Required_2 = True
        TypeFieldEdit1.Type_2 = esriFieldType.esriFieldTypeString
        fieldsEdit.AddField(TypeField1)

        ' Add the C_Hazen field to the required fields.
        Dim C_HazenField1 As IField = New FieldClass()
        Dim C_HazenFieldEdit1 As IFieldEdit = CType(C_HazenField1, IFieldEdit)
        C_HazenFieldEdit1.Name_2 = "C_HazenPipe1"
        C_HazenFieldEdit1.AliasName_2 = "C_HazenPipe1"
        C_HazenFieldEdit1.Required_2 = True
        C_HazenFieldEdit1.Type_2 = esriFieldType.esriFieldTypeSmallInteger
        fieldsEdit.AddField(C_HazenField1)

        ' Add the Discharge field to the required fields.
        Dim DischargeField1 As IField = New FieldClass()
        Dim DischargeFieldEdit1 As IFieldEdit = CType(DischargeField1, IFieldEdit)
        DischargeFieldEdit1.Name_2 = "DischargePipe1"
        DischargeFieldEdit1.AliasName_2 = "DischargePipe1"
        DischargeFieldEdit1.Required_2 = True
        DischargeFieldEdit1.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(DischargeField1)

        ' Add the Velocity field to the required fields.
        Dim VelocityField1 As IField = New FieldClass()
        Dim VelocityFieldEdit1 As IFieldEdit = CType(VelocityField1, IFieldEdit)
        VelocityFieldEdit1.Name_2 = "VelocityPipe1"
        VelocityFieldEdit1.AliasName_2 = "VelocityPipe1"
        VelocityFieldEdit1.Required_2 = True
        VelocityFieldEdit1.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(VelocityField1)

        ' Add the Headloss field to the required fields.
        Dim HeadlossALLField1 As IField = New FieldClass()
        Dim HeadlossALLFieldEdit1 As IFieldEdit = CType(HeadlossALLField1, IFieldEdit)
        HeadlossALLFieldEdit1.Name_2 = "HeadlossALLPipe1"
        HeadlossALLFieldEdit1.AliasName_2 = "HeadlossALLPipe1"
        HeadlossALLFieldEdit1.Required_2 = True
        HeadlossALLFieldEdit1.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(HeadlossALLField1)



        ' Add the Length field to the required fields.
        Dim LengthField2 As IField = New FieldClass()
        Dim LengthFieldEdit2 As IFieldEdit = CType(LengthField2, IFieldEdit)
        LengthFieldEdit2.Name_2 = "LengthPipe2"
        LengthFieldEdit2.AliasName_2 = "LengthPipe2"
        LengthFieldEdit2.Required_2 = True
        LengthFieldEdit2.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(LengthField2)

        ' Add the Diameter field to the required fields.
        Dim DiaField2 As IField = New FieldClass()
        Dim DiaFieldEdit2 As IFieldEdit = CType(DiaField2, IFieldEdit)
        DiaFieldEdit2.Name_2 = "DiameterPipe2"
        DiaFieldEdit2.AliasName_2 = "DiameterPipe2"
        DiaFieldEdit2.Required_2 = True
        DiaFieldEdit2.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(DiaField2)

        ' Add the Type field to the required fields.
        Dim TypeField2 As IField = New FieldClass()
        Dim TypeFieldEdit2 As IFieldEdit = CType(TypeField2, IFieldEdit)
        TypeFieldEdit2.Name_2 = "TypePipe2"
        TypeFieldEdit2.AliasName_2 = "TypePipe2"
        TypeFieldEdit2.Required_2 = True
        TypeFieldEdit2.Type_2 = esriFieldType.esriFieldTypeString
        fieldsEdit.AddField(TypeField2)

        ' Add the C_Hazen field to the required fields.
        Dim C_HazenField2 As IField = New FieldClass()
        Dim C_HazenFieldEdit2 As IFieldEdit = CType(C_HazenField2, IFieldEdit)
        C_HazenFieldEdit2.Name_2 = "C_HazenPipe2"
        C_HazenFieldEdit2.AliasName_2 = "C_HazenPipe2"
        C_HazenFieldEdit2.Required_2 = True
        C_HazenFieldEdit2.Type_2 = esriFieldType.esriFieldTypeSmallInteger
        fieldsEdit.AddField(C_HazenField2)

        ' Add the Discharge field to the required fields.
        Dim DischargeField2 As IField = New FieldClass()
        Dim DischargeFieldEdit2 As IFieldEdit = CType(DischargeField2, IFieldEdit)
        DischargeFieldEdit2.Name_2 = "DischargePipe2"
        DischargeFieldEdit2.AliasName_2 = "DischargePipe2"
        DischargeFieldEdit2.Required_2 = True
        DischargeFieldEdit2.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(DischargeField2)

        ' Add the Velocity field to the required fields.
        Dim VelocityField2 As IField = New FieldClass()
        Dim VelocityFieldEdit2 As IFieldEdit = CType(VelocityField2, IFieldEdit)
        VelocityFieldEdit2.Name_2 = "VelocityPipe2"
        VelocityFieldEdit2.AliasName_2 = "VelocityPipe2"
        VelocityFieldEdit2.Required_2 = True
        VelocityFieldEdit2.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(VelocityField2)

        ' Add the Headloss field to the required fields.
        Dim HeadlossALLField2 As IField = New FieldClass()
        Dim HeadlossALLFieldEdit2 As IFieldEdit = CType(HeadlossALLField2, IFieldEdit)
        HeadlossALLFieldEdit2.Name_2 = "HeadlossALLPipe2"
        HeadlossALLFieldEdit2.AliasName_2 = "HeadlossALLPipe2"
        HeadlossALLFieldEdit2.Required_2 = True
        HeadlossALLFieldEdit2.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(HeadlossALLField2)

        ' Use the field checker created earlier to validate the fields.
        Dim enumFieldError As IEnumFieldError = Nothing
        Dim validatedFields As IFields = Nothing
        fieldChecker.Validate(fields, enumFieldError, validatedFields)

        ' Create and return the object class.
        Dim table As ITable = featureWorkspace.CreateTable(validatedClassName, validatedFields, ocDescription.InstanceCLSID, Nothing, "")
        Return CType(table, IObjectClass)
    End Function



    Public Function CreateTableJuncCompareScenarios(ByVal workspace As IWorkspace, ByVal className As String) As IObjectClass
        ' Cast the workspace to IFeatureWorkspace and IWorkspace2. An explicit cast is not used
        ' for IWorkspace2 because not all workspaces implement it.
        Dim featureWorkspace As IFeatureWorkspace = CType(workspace, IFeatureWorkspace)
        Dim workspace2 As IWorkspace2 = TryCast(workspace, IWorkspace2)

        ' See if a class with the provided name already exists. If so, return it.
        ' Note that this will not work with file-based workspaces, such as shapefile workspaces.
        If Not workspace2 Is Nothing Then
            If workspace2.NameExists(esriDatasetType.esriDTTable, className) Then
                Dim existingTable As ITable = featureWorkspace.OpenTable(className)
                Return CType(existingTable, IObjectClass)
            End If
        End If

        ' Use IFieldChecker.ValidateTableName to validate the name of the class. The tableNameErrorType
        ' variable is not used in this example, but it indicates the reasons the table name was invalid,
        ' if any.
        Dim fieldChecker As IFieldChecker = New FieldCheckerClass()
        Dim validatedClassName As String = Nothing
        fieldChecker.ValidateWorkspace = workspace
        Dim tableNameErrorType As Integer = fieldChecker.ValidateTableName(className, validatedClassName)

        ' Create an object class description and get the required fields from it.
        Dim ocDescription As IObjectClassDescription = New ObjectClassDescriptionClass()
        Dim fields As IFields = ocDescription.RequiredFields
        Dim fieldsEdit As IFieldsEdit = CType(fields, IFieldsEdit)

        ' Add the Name field to the required fields.
        Dim NumOfJunc As IField = New FieldClass()
        Dim NumOfJuncEdit As IFieldEdit = CType(NumOfJunc, IFieldEdit)
        NumOfJuncEdit.Name_2 = "NumJunc"
        NumOfJuncEdit.AliasName_2 = "IDJunc"
        NumOfJuncEdit.Required_2 = True
        'NumOfJuncEdit.Editable_2 = False
        NumOfJuncEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        fieldsEdit.AddField(NumOfJunc)

        ' Add the Elevation field to the required fields.
        Dim ElevationField1 As IField = New FieldClass()
        Dim ElevationFieldEdit1 As IFieldEdit = CType(ElevationField1, IFieldEdit)
        ElevationFieldEdit1.Name_2 = "ElevationJunc1"
        ElevationFieldEdit1.AliasName_2 = "ElevationJunc1"
        ElevationFieldEdit1.Required_2 = True
        'ElevationFieldEdit.Editable_2 = False
        ElevationFieldEdit1.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(ElevationField1)

        ' Add the Demand field to the required fields.
        Dim DemandField1 As IField = New FieldClass()
        Dim DemandFieldEdit1 As IFieldEdit = CType(DemandField1, IFieldEdit)
        DemandFieldEdit1.Name_2 = "DemandJunc1"
        DemandFieldEdit1.AliasName_2 = "DemandJunc1"
        DemandFieldEdit1.Required_2 = True
        'DemandFieldEdit.Editable_2 = False
        DemandFieldEdit1.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(DemandField1)

        ' Add the Headloss field to the required fields.
        Dim HeadLoosField1 As IField = New FieldClass()
        Dim HeadLoosFieldEdit1 As IFieldEdit = CType(HeadLoosField1, IFieldEdit)
        HeadLoosFieldEdit1.Name_2 = "HeadLossJunc1"
        HeadLoosFieldEdit1.AliasName_2 = "HeadLossJunc1"
        HeadLoosFieldEdit1.Required_2 = True
        ' HeadLoosFieldEdit.Editable_2 = False
        HeadLoosFieldEdit1.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(HeadLoosField1)


        ' Add the Elevation field to the required fields.
        Dim ElevationField2 As IField = New FieldClass()
        Dim ElevationFieldEdit2 As IFieldEdit = CType(ElevationField2, IFieldEdit)
        ElevationFieldEdit2.Name_2 = "ElevationJunc2"
        ElevationFieldEdit2.AliasName_2 = "ElevationJunc2"
        ElevationFieldEdit2.Required_2 = True
        'ElevationFieldEdit.Editable_2 = False
        ElevationFieldEdit2.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(ElevationField2)


        ' Add the Demand field to the required fields.
        Dim DemandField2 As IField = New FieldClass()
        Dim DemandFieldEdit2 As IFieldEdit = CType(DemandField2, IFieldEdit)
        DemandFieldEdit2.Name_2 = "DemandJunc2"
        DemandFieldEdit2.AliasName_2 = "DemandJunc2"
        DemandFieldEdit2.Required_2 = True
        'DemandFieldEdit.Editable_2 = False
        DemandFieldEdit2.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(DemandField2)

        ' Add the Headloss field to the required fields.
        Dim HeadLoosField2 As IField = New FieldClass()
        Dim HeadLoosFieldEdit2 As IFieldEdit = CType(HeadLoosField2, IFieldEdit)
        HeadLoosFieldEdit2.Name_2 = "HeadLossJunc2"
        HeadLoosFieldEdit2.AliasName_2 = "HeadLossJunc2"
        HeadLoosFieldEdit2.Required_2 = True
        ' HeadLoosFieldEdit.Editable_2 = False
        HeadLoosFieldEdit2.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(HeadLoosField2)


        ' Use the field checker created earlier to validate the fields.
        Dim enumFieldError As IEnumFieldError = Nothing
        Dim validatedFields As IFields = Nothing
        fieldChecker.Validate(fields, enumFieldError, validatedFields)

        ' Create and return the object class.
        Dim table As ITable = featureWorkspace.CreateTable(validatedClassName, validatedFields, ocDescription.InstanceCLSID, Nothing, "")
        Return CType(table, IObjectClass)
    End Function


    Public Function CreateTableDemandUserCompareScenarios(ByVal workspace As IWorkspace, ByVal className As String) As IObjectClass
        ' Cast the workspace to IFeatureWorkspace and IWorkspace2. An explicit cast is not used
        ' for IWorkspace2 because not all workspaces implement it.
        Dim featureWorkspace As IFeatureWorkspace = CType(workspace, IFeatureWorkspace)
        Dim workspace2 As IWorkspace2 = TryCast(workspace, IWorkspace2)

        ' See if a class with the provided name already exists. If so, return it.
        ' Note that this will not work with file-based workspaces, such as shapefile workspaces.
        If Not workspace2 Is Nothing Then
            If workspace2.NameExists(esriDatasetType.esriDTTable, className) Then
                Dim existingTable As ITable = featureWorkspace.OpenTable(className)
                Return CType(existingTable, IObjectClass)
            End If
        End If

        ' Use IFieldChecker.ValidateTableName to validate the name of the class. The tableNameErrorType
        ' variable is not used in this example, but it indicates the reasons the table name was invalid,
        ' if any.
        Dim fieldChecker As IFieldChecker = New FieldCheckerClass()
        Dim validatedClassName As String = Nothing
        fieldChecker.ValidateWorkspace = workspace
        Dim tableNameErrorType As Integer = fieldChecker.ValidateTableName(className, validatedClassName)

        ' Create an object class description and get the required fields from it.
        Dim ocDescription As IObjectClassDescription = New ObjectClassDescriptionClass()
        Dim fields As IFields = ocDescription.RequiredFields
        Dim fieldsEdit As IFieldsEdit = CType(fields, IFieldsEdit)


        ' Add the Name field to the required fields.
        Dim NumOfUser As IField = New FieldClass()
        Dim NumOfUserEdit As IFieldEdit = CType(NumOfUser, IFieldEdit)
        NumOfUserEdit.Name_2 = "NumUser"
        NumOfUserEdit.AliasName_2 = "IDUser"
        NumOfUserEdit.Required_2 = True
        'NumOfJuncEdit.Editable_2 = False
        NumOfUserEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        fieldsEdit.AddField(NumOfUser)

        ' Add the Name field to the required fields.
        Dim MasrafUser1 As IField = New FieldClass()
        Dim MasrafUserEdit1 As IFieldEdit = CType(MasrafUser1, IFieldEdit)
        MasrafUserEdit1.Name_2 = "MasrafUser1"
        MasrafUserEdit1.AliasName_2 = "MasrafUser1"
        MasrafUserEdit1.Required_2 = True
        'NumOfJuncEdit.Editable_2 = False
        MasrafUserEdit1.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(MasrafUser1)


        ' Add the Name field to the required fields.
        Dim MasrafUser2 As IField = New FieldClass()
        Dim MasrafUserEdit2 As IFieldEdit = CType(MasrafUser2, IFieldEdit)
        MasrafUserEdit2.Name_2 = "MasrafUser2"
        MasrafUserEdit2.AliasName_2 = "MasrafUser2"
        MasrafUserEdit2.Required_2 = True
        'NumOfJuncEdit.Editable_2 = False
        MasrafUserEdit2.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(MasrafUser2)

        ' Use the field checker created earlier to validate the fields.
        Dim enumFieldError As IEnumFieldError = Nothing
        Dim validatedFields As IFields = Nothing
        fieldChecker.Validate(fields, enumFieldError, validatedFields)

        ' Create and return the object class.
        Dim table As ITable = featureWorkspace.CreateTable(validatedClassName, validatedFields, ocDescription.InstanceCLSID, Nothing, "")
        Return CType(table, IObjectClass)
    End Function

    Public Function CreateTablePattern(ByVal workspace As IWorkspace, ByVal className As String) As IObjectClass
        ' Cast the workspace to IFeatureWorkspace and IWorkspace2. An explicit cast is not used
        ' for IWorkspace2 because not all workspaces implement it.
        Dim featureWorkspace As IFeatureWorkspace = CType(workspace, IFeatureWorkspace)
        Dim workspace2 As IWorkspace2 = TryCast(workspace, IWorkspace2)

        ' See if a class with the provided name already exists. If so, return it.
        ' Note that this will not work with file-based workspaces, such as shapefile workspaces.
        If Not workspace2 Is Nothing Then
            If workspace2.NameExists(esriDatasetType.esriDTTable, className) Then
                Dim existingTable As ITable = featureWorkspace.OpenTable(className)
                Return CType(existingTable, IObjectClass)
            End If
        End If

        ' Use IFieldChecker.ValidateTableName to validate the name of the class. The tableNameErrorType
        ' variable is not used in this example, but it indicates the reasons the table name was invalid,
        ' if any.
        Dim fieldChecker As IFieldChecker = New FieldCheckerClass()
        Dim validatedClassName As String = Nothing
        fieldChecker.ValidateWorkspace = workspace
        Dim tableNameErrorType As Integer = fieldChecker.ValidateTableName(className, validatedClassName)

        ' Create an object class description and get the required fields from it.
        Dim ocDescription As IObjectClassDescription = New ObjectClassDescriptionClass()
        Dim fields As IFields = ocDescription.RequiredFields
        Dim fieldsEdit As IFieldsEdit = CType(fields, IFieldsEdit)

        ' Add the Name field to the required fields.
        Dim NumePattern As IField = New FieldClass()
        Dim NumePatternEdit As IFieldEdit = CType(NumePattern, IFieldEdit)
        NumePatternEdit.Name_2 = "NumePattern"
        NumePatternEdit.AliasName_2 = "NumePattern"
        NumePatternEdit.Required_2 = True
        NumePatternEdit.Type_2 = esriFieldType.esriFieldTypeString
        fieldsEdit.AddField(NumePattern)

        ' Add the Name field to the required fields.
        Dim Hour1 As IField = New FieldClass()
        Dim Hour1Edit As IFieldEdit = CType(Hour1, IFieldEdit)
        Hour1Edit.Name_2 = "Hour1"
        Hour1Edit.AliasName_2 = "Hour1"
        Hour1Edit.Required_2 = True
        Hour1Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour1)
        ' Add the Name field to the required fields.
        Dim Hour2 As IField = New FieldClass()
        Dim Hour2Edit As IFieldEdit = CType(Hour2, IFieldEdit)
        Hour2Edit.Name_2 = "Hour2"
        Hour2Edit.AliasName_2 = "Hour2"
        Hour2Edit.Required_2 = True
        Hour2Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour2)
        ' Add the Name field to the required fields.
        Dim Hour3 As IField = New FieldClass()
        Dim Hour3Edit As IFieldEdit = CType(Hour3, IFieldEdit)
        Hour3Edit.Name_2 = "Hour3"
        Hour3Edit.AliasName_2 = "Hour3"
        Hour3Edit.Required_2 = True
        Hour3Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour3)
        ' Add the Name field to the required fields.
        Dim Hour4 As IField = New FieldClass()
        Dim Hour4Edit As IFieldEdit = CType(Hour4, IFieldEdit)
        Hour4Edit.Name_2 = "Hour4"
        ' Hour4Edit.AliasName_2 = "ساعت 4"
        Hour4Edit.Required_2 = True
        Hour4Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour4)
        ' Add the Name field to the required fields.
        Dim Hour5 As IField = New FieldClass()
        Dim Hour5Edit As IFieldEdit = CType(Hour5, IFieldEdit)
        Hour5Edit.Name_2 = "Hour5"
        ' Hour5Edit.AliasName_2 = "ساعت 5"
        Hour5Edit.Required_2 = True
        Hour5Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour5)
        ' Add the Name field to the required fields.
        Dim Hour6 As IField = New FieldClass()
        Dim Hour6Edit As IFieldEdit = CType(Hour6, IFieldEdit)
        Hour6Edit.Name_2 = "Hour6"
        ' Hour6Edit.AliasName_2 = "ساعت 6"
        Hour6Edit.Required_2 = True
        Hour6Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour6)
        ' Add the Name field to the required fields.
        Dim Hour7 As IField = New FieldClass()
        Dim Hour7Edit As IFieldEdit = CType(Hour7, IFieldEdit)
        Hour7Edit.Name_2 = "Hour7"
        ' Hour7Edit.AliasName_2 = "ساعت 7"
        Hour7Edit.Required_2 = True
        Hour7Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour7)
        ' Add the Name field to the required fields.
        Dim Hour8 As IField = New FieldClass()
        Dim Hour8Edit As IFieldEdit = CType(Hour8, IFieldEdit)
        Hour8Edit.Name_2 = "Hour8"
        ' Hour8Edit.AliasName_2 = "ساعت 8"
        Hour8Edit.Required_2 = True
        Hour8Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour8)
        ' Add the Name field to the required fields.
        Dim Hour9 As IField = New FieldClass()
        Dim Hour9Edit As IFieldEdit = CType(Hour9, IFieldEdit)
        Hour9Edit.Name_2 = "Hour9"
        ' Hour9Edit.AliasName_2 = "ساعت 9"
        Hour9Edit.Required_2 = True
        Hour9Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour9)
        ' Add the Name field to the required fields.
        Dim Hour10 As IField = New FieldClass()
        Dim Hour10Edit As IFieldEdit = CType(Hour10, IFieldEdit)
        Hour10Edit.Name_2 = "Hour10"
        ' Hour10Edit.AliasName_2 = "ساعت 10"
        Hour10Edit.Required_2 = True
        Hour10Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour10)

        ' Add the Name field to the required fields.
        Dim Hour11 As IField = New FieldClass()
        Dim Hour11Edit As IFieldEdit = CType(Hour11, IFieldEdit)
        Hour11Edit.Name_2 = "Hour11"
        ' Hour11Edit.AliasName_2 = "ساعت 11"
        Hour11Edit.Required_2 = True
        Hour11Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour11)
        ' Add the Name field to the required fields.
        Dim Hour12 As IField = New FieldClass()
        Dim Hour12Edit As IFieldEdit = CType(Hour12, IFieldEdit)
        Hour12Edit.Name_2 = "Hour12"
        ' Hour12Edit.AliasName_2 = "ساعت 12"
        Hour12Edit.Required_2 = True
        Hour12Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour12)
        ' Add the Name field to the required fields.
        Dim Hour13 As IField = New FieldClass()
        Dim Hour13Edit As IFieldEdit = CType(Hour13, IFieldEdit)
        Hour13Edit.Name_2 = "Hour13"
        ' Hour13Edit.AliasName_2 = "ساعت 13"
        Hour13Edit.Required_2 = True
        Hour13Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour13)
        ' Add the Name field to the required fields.
        Dim Hour14 As IField = New FieldClass()
        Dim Hour14Edit As IFieldEdit = CType(Hour14, IFieldEdit)
        Hour14Edit.Name_2 = "Hour14"
        ' Hour14Edit.AliasName_2 = "ساعت 14"
        Hour14Edit.Required_2 = True
        Hour14Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour14)
        ' Add the Name field to the required fields.
        Dim Hour15 As IField = New FieldClass()
        Dim Hour15Edit As IFieldEdit = CType(Hour15, IFieldEdit)
        Hour15Edit.Name_2 = "Hour15"
        ' Hour15Edit.AliasName_2 = "ساعت 15"
        Hour15Edit.Required_2 = True
        Hour15Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour15)
        ' Add the Name field to the required fields.
        Dim Hour16 As IField = New FieldClass()
        Dim Hour16Edit As IFieldEdit = CType(Hour16, IFieldEdit)
        Hour16Edit.Name_2 = "Hour16"
        ' Hour16Edit.AliasName_2 = "ساعت 16"
        Hour16Edit.Required_2 = True
        Hour16Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour16)
        ' Add the Name field to the required fields.
        Dim Hour17 As IField = New FieldClass()
        Dim Hour17Edit As IFieldEdit = CType(Hour17, IFieldEdit)
        Hour17Edit.Name_2 = "Hour17"
        ' Hour17Edit.AliasName_2 = "ساعت 17"
        Hour17Edit.Required_2 = True
        Hour17Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour17)
        ' Add the Name field to the required fields.
        Dim Hour18 As IField = New FieldClass()
        Dim Hour18Edit As IFieldEdit = CType(Hour18, IFieldEdit)
        Hour18Edit.Name_2 = "Hour18"
        ' Hour18Edit.AliasName_2 = "ساعت 18"
        Hour18Edit.Required_2 = True
        Hour18Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour18)
        ' Add the Name field to the required fields.
        Dim Hour19 As IField = New FieldClass()
        Dim Hour19Edit As IFieldEdit = CType(Hour19, IFieldEdit)
        Hour19Edit.Name_2 = "Hour19"
        ' Hour19Edit.AliasName_2 = "ساعت 19"
        Hour19Edit.Required_2 = True
        Hour19Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour19)
        ' Add the Name field to the required fields.
        Dim Hour20 As IField = New FieldClass()
        Dim Hour20Edit As IFieldEdit = CType(Hour20, IFieldEdit)
        Hour20Edit.Name_2 = "Hour20"
        ' Hour20Edit.AliasName_2 = "ساعت 20"
        Hour20Edit.Required_2 = True
        Hour20Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour20)
        ' Add the Name field to the required fields.
        Dim Hour21 As IField = New FieldClass()
        Dim Hour21Edit As IFieldEdit = CType(Hour21, IFieldEdit)
        Hour21Edit.Name_2 = "Hour21"
        ' Hour21Edit.AliasName_2 = "ساعت 21"
        Hour21Edit.Required_2 = True
        Hour21Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour21)

        Dim Hour22 As IField = New FieldClass()
        Dim Hour22Edit As IFieldEdit = CType(Hour22, IFieldEdit)
        Hour22Edit.Name_2 = "Hour22"
        ' Hour22Edit.AliasName_2 = "ساعت 22"
        Hour22Edit.Required_2 = True
        Hour22Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour22)

        Dim Hour23 As IField = New FieldClass()
        Dim Hour23Edit As IFieldEdit = CType(Hour23, IFieldEdit)
        Hour23Edit.Name_2 = "Hour23"
        ' Hour23Edit.AliasName_2 = "ساعت 23"
        Hour23Edit.Required_2 = True
        Hour23Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour23)

        Dim Hour24 As IField = New FieldClass()
        Dim Hour24Edit As IFieldEdit = CType(Hour24, IFieldEdit)
        Hour24Edit.Name_2 = "Hour24"
        ' Hour24Edit.AliasName_2 = "ساعت 24"
        Hour24Edit.Required_2 = True
        Hour24Edit.Type_2 = esriFieldType.esriFieldTypeDouble
        fieldsEdit.AddField(Hour24)


        Dim SetDefault As IField = New FieldClass()
        Dim SetDefaultEdit As IFieldEdit = CType(SetDefault, IFieldEdit)
        SetDefaultEdit.Name_2 = "SetDefault"
        '  SetDefaultEdit.AliasName_2 = "الگوی پیش فرض"
        SetDefaultEdit.Required_2 = True
        SetDefaultEdit.DefaultValue_2 = 0
        SetDefaultEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        fieldsEdit.AddField(SetDefault)

        Dim HourMabna As IField = New FieldClass()
        Dim HourMabnaEdit As IFieldEdit = CType(HourMabna, IFieldEdit)
        HourMabnaEdit.Name_2 = "HourMabnaEdit"
        ' HourMabnaEdit.AliasName_2 = "ساعت مبنا در الگوی پیش فرض"
        HourMabnaEdit.Required_2 = True
        HourMabnaEdit.DefaultValue_2 = 0
        HourMabnaEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        fieldsEdit.AddField(HourMabna)

        ' Use the field checker created earlier to validate the fields.
        Dim enumFieldError As IEnumFieldError = Nothing
        Dim validatedFields As IFields = Nothing
        fieldChecker.Validate(fields, enumFieldError, validatedFields)

        ' Create and return the object class.
        Dim table As ITable = featureWorkspace.CreateTable(validatedClassName, validatedFields, ocDescription.InstanceCLSID, Nothing, "")
        Return CType(table, IObjectClass)
    End Function
End Class
