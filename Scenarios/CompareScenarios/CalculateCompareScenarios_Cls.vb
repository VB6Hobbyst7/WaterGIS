Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports System.Runtime.InteropServices

Public Class CalculateCompareScenarios_Cls

    Private m_hookHelper As IHookHelper = Nothing
    Private IDScenario1, IDScenario2, IDScenarioBase As Integer
    Private Map As IMap
    Private m_FCPipe As IFeatureClass
    Private CollScenarioPipe1, CollScenarioPipe2 As Collection
    Private CollScenarioJunc1, CollScenarioJunc2 As Collection
    Private TableIntermediate, TableEdgeCompare, TableJuncCompare As ITable
    Private CalculateHeadloss As CalculateHeadlossJunctionScenario_Cls
    Private IndexNumPipe, IndexNumJunc, IndexNumScenario As Integer
    Private GeometricNetwork As IGeometricNetwork
    Private SourceFC As IFeatureClass

    Public Sub New(ByVal hookHelper As IHookHelper)
        m_hookHelper = hookHelper
    End Sub

    Public WriteOnly Property Get_IDScenario1() As Integer
        Set(ByVal value As Integer)
            IDScenario1 = value
        End Set
    End Property

    Public WriteOnly Property Get_IDScenarioBase() As Integer
        Set(ByVal value As Integer)
            IDScenarioBase = value
        End Set
    End Property

    Public WriteOnly Property Get_IDScenario2() As Integer
        Set(ByVal value As Integer)
            IDScenario2 = value
        End Set
    End Property

    Public Function CalculationCompare() As Boolean

        If m_hookHelper Is Nothing Then Exit Function

        Map = TryCast(m_hookHelper.FocusMap, IMap)
        If Map Is Nothing Then Exit Function

        CalculationCompare = False
        m_FCPipe = m_FCMainPipe
        SourceFC = m_FCSource

        TableJuncCompare = m_TableJuncScenarios
        TableEdgeCompare = m_TableEdgeScenarios
        TableIntermediate = m_TableIntermediateScenario
        GeometricNetwork = m_GeometricNetwork

        InitializeCollections()

        Fill_TableEdgeCompare()
        Fill_TableJuncCompare()

        If CalculateHeadloss Is Nothing Then
            CalculateHeadloss = New CalculateHeadlossJunctionScenario_Cls(m_hookHelper)
        End If

        CalculateHeadloss.CalculateCompare()
        CalculationCompare = True

    End Function


    Private Sub InitializeCollections()

        Dim Cursor As ICursor
        Dim RowCount As Integer
        Dim Row As IRow

        CollScenarioJunc1 = New Collection
        CollScenarioJunc2 = New Collection
        CollScenarioPipe1 = New Collection
        CollScenarioPipe2 = New Collection

        IndexNumPipe = TableIntermediate.FindField("NumPipe")
        IndexNumScenario = TableIntermediate.FindField("NumScenario")
        IndexNumJunc = TableIntermediate.FindField("NumJunc")
        If IndexNumPipe = -1 Or IndexNumScenario = -1 Or IndexNumJunc = -1 Then
            Exit Sub
        End If

        RowCount = TableIntermediate.RowCount(Nothing)
        Cursor = TableIntermediate.Search(Nothing, False)
        For i = 0 To RowCount - 1
            Row = Cursor.NextRow
            If Row.HasOID Then
                If Row.Value(IndexNumScenario) = IDScenario1 Then
                    If Row.Value(IndexNumPipe) <> -1 Then
                        CollScenarioPipe1.Add(Row.Value(IndexNumPipe))
                    End If
                    If Row.Value(IndexNumJunc) <> -1 Then
                        CollScenarioJunc1.Add(Row.Value(IndexNumJunc))
                    End If
                End If
                If Row.Value(IndexNumScenario) = IDScenario2 Then
                    If Row.Value(IndexNumPipe) <> -1 Then
                        CollScenarioPipe2.Add(Row.Value(IndexNumPipe))
                    End If
                    If Row.Value(IndexNumJunc) <> -1 Then
                        CollScenarioJunc2.Add(Row.Value(IndexNumJunc))
                    End If
                End If
            End If
        Next

        Marshal.ReleaseComObject(Cursor)
    End Sub

    Private Function SearchCollection(ByVal CollectionID As Collection, ByVal OID As Integer) As Boolean
        SearchCollection = False
        For i = CollectionID.Count - 1 To 0 Step -1
            If CollectionID(i + 1) = OID Then
                SearchCollection = True
                Exit For
            End If
        Next
    End Function

    Private Sub Fill_TableJuncCompare()
        On Error GoTo ErrorHandler

        Dim WorkspaceEdit As IWorkspaceEdit
        Dim FCJunction As IFeatureClass
        Dim Row As IRow

        Dim NameFieldHeadlssFC As String = "Headloss"
        Dim NameFieldElevationFC As String = "Elevation"
        Dim NameFieldDemandFC As String = "MasrafCol"

        Dim NameFieldHeadlss1Table As String = "HeadLossJunc1"
        Dim NameFieldElevation1Table As String = "ElevationJunc1"
        Dim NameFieldDemand1Table As String = "DemandJunc1"
        Dim NameFieldNumJuncTable As String = "NumJunc"

        Dim NameFieldHeadlss2Table As String = "HeadLossJunc2"
        Dim NameFieldElevation2Table As String = "ElevationJunc2"
        Dim NameFieldDemand2Table As String = "DemandJunc2"


        Dim IndexHeadlossFC As Integer
        Dim IndexElevationFC As Integer
        Dim IndexDemandFC As Integer


        Dim IndexHeadloss1Table As Integer
        Dim IndexElevation1Table As Integer
        Dim IndexDemand1Table As Integer
        Dim IndexNumJuncTable As Integer

        Dim IndexHeadloss2Table As Integer
        Dim IndexElevation2Table As Integer
        Dim IndexDemand2Table As Integer


        FCJunction = m_FCJunction

        IndexDemandFC = FCJunction.FindField(NameFieldDemandFC)
        IndexElevationFC = FCJunction.FindField(NameFieldElevationFC)
        IndexHeadlossFC = FCJunction.FindField(NameFieldHeadlssFC)

        IndexDemand1Table = TableJuncCompare.FindField(NameFieldDemand1Table)
        IndexElevation1Table = TableJuncCompare.FindField(NameFieldElevation1Table)
        IndexHeadloss1Table = TableJuncCompare.FindField(NameFieldHeadlss1Table)
        IndexNumJuncTable = TableJuncCompare.FindField(NameFieldNumJuncTable)

        IndexDemand2Table = TableJuncCompare.FindField(NameFieldDemand2Table)
        IndexElevation2Table = TableJuncCompare.FindField(NameFieldElevation2Table)
        IndexHeadloss2Table = TableJuncCompare.FindField(NameFieldHeadlss2Table)


        WorkspaceEdit = CType(FCJunction.FeatureDataset.Workspace, IWorkspaceEdit)
        WorkspaceEdit.StartEditing(True)
        WorkspaceEdit.StartEditOperation()

        Dim Cursor As ICursor = TableJuncCompare.Search(Nothing, False)
        Dim RowCount As Integer = TableJuncCompare.RowCount(Nothing)

        For i = 0 To RowCount - 1
            Row = Cursor.NextRow
            Row.Delete()
        Next


        Dim FeatureCount As Integer = FCJunction.FeatureCount(Nothing)
        Dim FCursor As IFeatureCursor = FCJunction.Search(Nothing, False)
        Dim Feature As IFeature
        For i = 0 To FeatureCount - 1
            Feature = FCursor.NextFeature
            If Feature.HasOID Then

                Row = TableJuncCompare.CreateRow
                If IDScenario1 = 0 OrElse Not SearchCollection(CollScenarioJunc1, Feature.OID) Then

                    'Scenario1
                    Row.Value(IndexNumJuncTable) = Feature.OID
                    Row.Value(IndexElevation1Table) = Math.Round(Feature.Value(IndexElevationFC), 2)
                    Row.Value(IndexDemand1Table) = Math.Round(Feature.Value(IndexDemandFC), 2)
                    ' Row.Value(IndexHeadloss1Table) = Feature.Value(IndexHeadlossFC)
                    Row.Store()

                Else
                    Dim RowTableIntermediate As IRow
                    RowTableIntermediate = Find_RowInTableIntermediate(IDScenario1, -1, Feature.OID)
                    If Not RowTableIntermediate Is Nothing Then

                        'Scenario1

                        Row.Value(IndexNumJuncTable) = Feature.OID
                        Row.Value(IndexDemand1Table) = Math.Round(Feature.Value(IndexDemandFC), 2)
                        Row.Value(IndexElevation1Table) = Math.Round(RowTableIntermediate.Value(TableIntermediate.FindField("ElevationJunc")), 2)
                        Row.Store()

                    End If

                End If
                If IDScenario2 = 0 OrElse Not SearchCollection(CollScenarioJunc2, Feature.OID) Then

                    'Scenario2
                    Row.Value(IndexElevation2Table) = Math.Round(Feature.Value(IndexElevationFC), 2)
                    Row.Value(IndexDemand2Table) = Math.Round(Feature.Value(IndexDemandFC), 2)
                    Row.Store()

                Else
                    Dim RowTableIntermediate As IRow
                    RowTableIntermediate = Find_RowInTableIntermediate(IDScenario2, -1, Feature.OID)
                    If Not RowTableIntermediate Is Nothing Then

                        'Scenario2

                        Row.Value(IndexDemand2Table) = Math.Round(Feature.Value(IndexDemandFC), 2)
                        Row.Value(IndexElevation2Table) = Math.Round(RowTableIntermediate.Value(TableIntermediate.FindField("ElevationJunc")), 2)
                        Row.Store()

                    End If

                End If

            End If


        Next


        WorkspaceEdit.StopEditOperation()
        WorkspaceEdit.StopEditing(True)
        Marshal.ReleaseComObject(FCursor)

        Exit Sub
ErrorHandler:
        WorkspaceEdit.AbortEditOperation()

    End Sub

    Private Sub Fill_TableEdgeCompare()
        On Error GoTo ErrorHandler

        Dim WorkspaceEdit As IWorkspaceEdit

        Dim NameFieldDia1Table As String = "DiameterPipe1"
        Dim NameFieldC1Table As String = "C_HazenPipe1"
        Dim NameFieldType1Table As String = "TypePipe1"
        Dim NameFieldDebi1Table As String = "DischargePipe1"
        Dim NameFieldVelocity1Table As String = "VelocityPipe1"
        Dim NameFieldNumPipeTable As String = "NumPipe"
        Dim NameFieldLength1Table As String = "LengthPipe1"
        Dim NameFieldHeadloss1Table As String = "HeadlossALLPipe1"


        Dim NameFieldDia2Table As String = "DiameterPipe2"
        Dim NameFieldC2Table As String = "C_HazenPipe2"
        Dim NameFieldType2Table As String = "TypePipe2"
        Dim NameFieldDebi2Table As String = "DischargePipe2"
        Dim NameFieldVelocity2Table As String = "VelocityPipe2"
        Dim NameFieldLength2Table As String = "LengthPipe2"
        Dim NameFieldHeadloss2Table As String = "HeadlossALLPipe2"

        Dim IndexDebi1FieldTable As Integer
        Dim IndexVelocity1FieldTable As Integer
        Dim IndexDia1FieldTable As Integer
        Dim IndexC1FieldTable As Integer
        Dim IndexType1FieldTable As Integer
        Dim IndexNumFieldTable As Integer
        Dim IndexLength1FieldTable As Integer
        Dim IndexHeadloss1FieldTable As Integer

        Dim IndexDebi2FieldTable As Integer
        Dim IndexVelocity2FieldTable As Integer
        Dim IndexDia2FieldTable As Integer
        Dim IndexC2FieldTable As Integer
        Dim IndexType2FieldTable As Integer
        Dim IndexLength2FieldTable As Integer
        Dim IndexHeadloss2FieldTable As Integer

        Dim FCPipe As IFeatureClass
        Dim Row As IRow


        FCPipe = m_FCPipe

        IndexC1FieldTable = TableEdgeCompare.FindField(NameFieldC1Table)
        IndexDebi1FieldTable = TableEdgeCompare.FindField(NameFieldDebi1Table)
        IndexDia1FieldTable = TableEdgeCompare.FindField(NameFieldDia1Table)
        IndexVelocity1FieldTable = TableEdgeCompare.FindField(NameFieldVelocity1Table)
        IndexType1FieldTable = TableEdgeCompare.FindField(NameFieldType1Table)
        IndexNumFieldTable = TableEdgeCompare.FindField(NameFieldNumPipeTable)
        IndexLength1FieldTable = TableEdgeCompare.FindField(NameFieldLength1Table)
        IndexHeadloss1FieldTable = TableEdgeCompare.FindField(NameFieldHeadloss1Table)

        IndexC2FieldTable = TableEdgeCompare.FindField(NameFieldC2Table)
        IndexDebi2FieldTable = TableEdgeCompare.FindField(NameFieldDebi2Table)
        IndexDia2FieldTable = TableEdgeCompare.FindField(NameFieldDia2Table)
        IndexVelocity2FieldTable = TableEdgeCompare.FindField(NameFieldVelocity2Table)
        IndexType2FieldTable = TableEdgeCompare.FindField(NameFieldType2Table)
        IndexLength2FieldTable = TableEdgeCompare.FindField(NameFieldLength2Table)
        IndexHeadloss2FieldTable = TableEdgeCompare.FindField(NameFieldHeadloss2Table)

        If m_IndexMainPipeDaimeter = -1 Then Exit Sub

        WorkspaceEdit = m_Workspace
        WorkspaceEdit.StartEditing(True)
        WorkspaceEdit.StartEditOperation()

        Dim Cursor As ICursor = TableEdgeCompare.Search(Nothing, False)
        Dim RowCount As Integer = TableEdgeCompare.RowCount(Nothing)

        For i = 0 To RowCount - 1
            Row = Cursor.NextRow
            Row.Delete()
        Next


        Dim FeatureCount As Integer = FCPipe.FeatureCount(Nothing)
        Dim FCursor As IFeatureCursor = FCPipe.Search(Nothing, False)
        Dim Feature As IFeature
        For i = 0 To FeatureCount - 1
            Feature = FCursor.NextFeature
            If Feature.HasOID Then

                Row = TableEdgeCompare.CreateRow
                If IDScenario1 = 0 OrElse Not SearchCollection(CollScenarioPipe1, Feature.OID) Then

                    'Scenario1
                    Row.Value(IndexNumFieldTable) = Feature.OID
                    Row.Value(IndexLength1FieldTable) = Math.Round(Feature.Value(FCPipe.FindField(FCPipe.LengthField.Name)), 2)
                    Row.Value(IndexDia1FieldTable) = Feature.Value(m_IndexMainPipeDaimeter)

                    Select Case Feature.Value(m_IndexMainPipeMaterial)

                        Case 1
                            Row.Value(IndexType1FieldTable) = "Cement"
                        Case 2
                            Row.Value(IndexType1FieldTable) = "Asbestos"
                        Case 3
                            Row.Value(IndexType1FieldTable) = "Steel"
                        Case 4
                            Row.Value(IndexType1FieldTable) = "Ductile"
                        Case 5
                            Row.Value(IndexType1FieldTable) = "Galvanized"
                        Case 6
                            Row.Value(IndexType1FieldTable) = "Plyka"
                        Case 7
                            Row.Value(IndexType1FieldTable) = "Cast-iron"
                        Case 8
                            Row.Value(IndexType1FieldTable) = "Copper"
                        Case 9
                            Row.Value(IndexType1FieldTable) = "PE"
                        Case 10
                            Row.Value(IndexType1FieldTable) = "Other"
                    End Select

                    Row.Value(IndexC1FieldTable) = Feature.Value(m_IndexMainPipeCHaizen)
                    Row.Value(IndexDebi1FieldTable) = Math.Round(Feature.Value(m_IndexMainPipeDebi), 2)
                    Row.Value(IndexVelocity1FieldTable) = Math.Round(Feature.Value(m_IndexMainPipeVelocity), 3)
                    Row.Value(IndexHeadloss1FieldTable) = Math.Round(Feature.Value(m_IndexMainPipeHeadloss), 3)
                    Row.Store()
                Else
                    Dim RowTableIntermediate As IRow
                    RowTableIntermediate = Find_RowInTableIntermediate(IDScenario1, Feature.OID, -1)
                    If Not RowTableIntermediate Is Nothing Then

                        'Scenario1

                        Row.Value(IndexNumFieldTable) = Feature.OID
                        Row.Value(IndexLength1FieldTable) = Math.Round(RowTableIntermediate.Value(TableIntermediate.FindField("LengthPipe")), 2)
                        Row.Value(IndexDia1FieldTable) = RowTableIntermediate.Value(TableIntermediate.FindField("DiameterPipe"))
                        Row.Value(IndexType1FieldTable) = RowTableIntermediate.Value(TableIntermediate.FindField("MaterialPipe"))
                        Row.Value(IndexC1FieldTable) = RowTableIntermediate.Value(TableIntermediate.FindField("HazenCPipe"))
                        Row.Value(IndexDebi1FieldTable) = Math.Round(Feature.Value(m_IndexMainPipeDebi), 2)
                        Row.Store()

                        Row.Value(IndexVelocity1FieldTable) = Math.Round((Math.Abs(Row.Value(IndexDebi1FieldTable)) * 0.001) / ((Math.Pow((Row.Value(IndexDia1FieldTable) * 0.001), 2) * Math.PI) / 4), 3)
                        Row.Store()
                        Row.Value(IndexHeadloss1FieldTable) = Math.Round((6.78 / Math.Round(Math.Pow((Row.Value(IndexDia1FieldTable) * 0.001), 1.165), 4)) * (Math.Pow(Math.Round((Row.Value(IndexVelocity1FieldTable) / Row.Value(IndexC1FieldTable)), 4), 1.85)) * Row.Value(IndexLength1FieldTable), 4)
                        Row.Store()

                    End If

                End If
                If IDScenario2 = 0 OrElse Not SearchCollection(CollScenarioPipe2, Feature.OID) Then

                    'Scenario2
                    Row.Value(IndexLength2FieldTable) = Feature.Value(FCPipe.FindField(FCPipe.LengthField.Name))
                    Row.Value(IndexDia2FieldTable) = Feature.Value(m_IndexMainPipeDaimeter)

                    Select Case Feature.Value(m_IndexMainPipeMaterial)
                        Case 1
                            Row.Value(IndexType2FieldTable) = "Cement"
                        Case 2
                            Row.Value(IndexType2FieldTable) = "Asbestos"
                        Case 3
                            Row.Value(IndexType2FieldTable) = "Steel"
                        Case 4
                            Row.Value(IndexType2FieldTable) = "Ductile"
                        Case 5
                            Row.Value(IndexType2FieldTable) = "Galvanized"
                        Case 6
                            Row.Value(IndexType2FieldTable) = "Plyka"
                        Case 7
                            Row.Value(IndexType2FieldTable) = "Cast-iron"
                        Case 8
                            Row.Value(IndexType2FieldTable) = "Copper"
                        Case 9
                            Row.Value(IndexType2FieldTable) = "PE"
                        Case 10
                            Row.Value(IndexType2FieldTable) = "Other"
                    End Select

                    Row.Value(IndexC2FieldTable) = Feature.Value(m_IndexMainPipeCHaizen)
                    Row.Value(IndexDebi2FieldTable) = Math.Round(Feature.Value(m_IndexMainPipeDebi), 2)
                    Row.Value(IndexVelocity2FieldTable) = Math.Round(Feature.Value(m_IndexMainPipeVelocity), 3)
                    Row.Value(IndexHeadloss2FieldTable) = Math.Round(Feature.Value(m_IndexMainPipeHeadloss), 3)
                    Row.Store()

                Else
                    Dim RowTableIntermediate As IRow
                    RowTableIntermediate = Find_RowInTableIntermediate(IDScenario2, Feature.OID, -1)
                    If Not RowTableIntermediate Is Nothing Then

                        ' Scenario2
                        Row.Value(IndexLength2FieldTable) = RowTableIntermediate.Value(TableIntermediate.FindField("LengthPipe"))
                        Row.Value(IndexDia2FieldTable) = RowTableIntermediate.Value(TableIntermediate.FindField("DiameterPipe"))
                        Row.Value(IndexType2FieldTable) = RowTableIntermediate.Value(TableIntermediate.FindField("MaterialPipe"))
                        Row.Value(IndexC2FieldTable) = RowTableIntermediate.Value(TableIntermediate.FindField("HazenCPipe"))
                        Row.Value(IndexDebi2FieldTable) = Math.Round(Feature.Value(m_IndexMainPipeDebi), 2)
                        Row.Store()
                        Row.Value(IndexVelocity2FieldTable) = Math.Round((Math.Abs(Row.Value(IndexDebi2FieldTable)) * 0.001) / ((Math.Pow((Row.Value(IndexDia2FieldTable) * 0.001), 2) * Math.PI) / 4), 3)
                        Row.Store()
                        Row.Value(IndexHeadloss2FieldTable) = Math.Round((6.78 / Math.Round(Math.Pow((Row.Value(IndexDia2FieldTable) * 0.001), 1.165), 4)) * (Math.Pow(Math.Round((Row.Value(IndexVelocity2FieldTable) / Row.Value(IndexC2FieldTable)), 4), 1.85)) * Row.Value(IndexLength2FieldTable), 4)
                        Row.Store()

                    End If

                End If

            End If
        Next

        WorkspaceEdit.StopEditOperation()
        WorkspaceEdit.StopEditing(True)
        Marshal.ReleaseComObject(FCursor)
        Exit Sub
ErrorHandler:
        WorkspaceEdit.AbortEditOperation()

    End Sub

    Private Function Find_RowInTableIntermediate(ByVal IDScenario As Integer, ByVal NumPipe As Integer, ByVal NumJunc As Integer) As IRow
        Dim Cursor As ICursor
        Dim RowCount As Integer
        Dim Row As IRow

        RowCount = TableIntermediate.RowCount(Nothing)
        Cursor = TableIntermediate.Search(Nothing, False)
        Find_RowInTableIntermediate = Nothing

        For i = 0 To RowCount - 1
            Row = Cursor.NextRow
            If Row.Value(IndexNumScenario) = IDScenario Then
                If NumPipe <> -1 Then
                    If Row.Value(IndexNumPipe) = NumPipe Then
                        Find_RowInTableIntermediate = Row
                        Exit For
                    End If
                End If
                If NumJunc <> -1 Then
                    If Row.Value(IndexNumJunc) = NumJunc Then
                        Find_RowInTableIntermediate = Row
                        Exit For
                    End If
                End If
            End If
        Next



    End Function
End Class
