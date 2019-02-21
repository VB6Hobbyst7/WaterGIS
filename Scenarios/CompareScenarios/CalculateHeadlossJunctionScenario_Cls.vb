Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.Geodatabase
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Carto

Public Class CalculateHeadlossJunctionScenario_Cls

    Private Map As IMap
    Private UtilityNetwork As IUtilityNetwork
    Private Network As INetwork
    Private TableJuncCompare, TableEdgeCompare As ITable
    Public Shared CorrectCalculation As Boolean
    Private IndexHeadField1, IndexHeadField2 As Integer
    Private IndexNumJunc, IndexNumEdge As Integer
    Private m_hookHelper As IHookHelper = Nothing
    Private ResultMessage As String

    Public Sub New(ByVal hookHelper As IHookHelper)
        m_hookHelper = hookHelper
    End Sub

    Public Sub CalculateCompare()

        CorrectCalculation = False
        Dim ForwardStar As IForwardStar
        Dim NetElement As INetElements

        ResultMessage = ""
        If m_hookHelper Is Nothing Then Exit Sub

        Map = TryCast(m_hookHelper.FocusMap, IMap)
        If Map Is Nothing Then Exit Sub

        TableEdgeCompare = m_TableEdgeScenarios
        TableJuncCompare = m_TableJuncScenarios

        Network = m_GeometricNetwork.Network
        NetElement = Network
        UtilityNetwork = Network
        ForwardStar = Network.CreateForwardStar(True, Nothing, Nothing, Nothing, Nothing)

        IndexNumJunc = TableJuncCompare.FindField("NumJunc")
        IndexNumEdge = TableEdgeCompare.FindField("NumPipe")
        IndexHeadField1 = TableJuncCompare.FindField("HeadLossJunc1")
        IndexHeadField2 = TableJuncCompare.FindField("HeadLossJunc2")

        If m_CollNetworkFeaturesComplexEF Is Nothing OrElse m_CollNetworkFeaturesComplexEF.Count <> m_FCMainPipe.FeatureCount(Nothing) Then
            ' ResultMessage = "اعمال تغییرات را بر روی شبکه تکرار نمایید"
            ResultMessage = "Please , ClickValidation again"
            Exit Sub
        End If

        Calculate_AllHeadLoos(ForwardStar, UtilityNetwork, IndexHeadField1)

        Calculate_AllHeadLoos(ForwardStar, UtilityNetwork, IndexHeadField2)

        CorrectCalculation = True
    End Sub


    Private Sub Calculate_AllHeadLoos(ByVal ForwardStar As IForwardStar, ByVal UtilityNetwork As IUtilityNetwork, ByVal IndexHead As Integer)
        On Error GoTo ErrorMessage

        Dim WorkspaceEdit As IWorkspaceEdit
        Dim Dataset As IDataset

        Dim Pipe, InPipe, MainPipe As IComplexEdgeFeature
        Dim PipeFeat, InPipeFeat As IFeature
        Dim FromJunc, ToJunc As ISimpleJunctionFeature
        Dim FromJuncFeat, ToJuncFeat As IFeature

        Dim NameElevationField As String = "Elevation"
        Dim NameHeadLossField As String = "Headloss"
        Dim NameVelocityField As String = "Velocity"
        Dim IndexHeadFieldJunc As Integer
        Dim IndexHeadFieldPipe As Integer
        Dim IndexEleField As Integer
        Dim IndexVelocityField As Integer
        Dim IndexLengthField As Integer

        Dim FCSource As IFeatureClass
        ' Dim FCJunction As IFeatureClass
        Dim HeadLoss As Double
        Dim EdgeCount As Integer
        Dim Orient As Boolean
        Dim AdjEdgeEID As Integer
        Dim Weight As Double
        Dim RowTableFrom As IRow
        Dim RowTableTo As IRow
        Dim RowTableInEdge, RowTableEdge As IRow
        Dim lngUserClassId, lngUserID, lngSubID As Long
        Dim NetElement As INetElements
        Dim MainPipeFeature As IFeature
        Dim Factor As Integer

        Dim EleFromJunc, EleToJunc As Double
        Dim VelocityInPipe, VelocityPipe As Double
        Dim HeadlossFromJunc, HeadlossToJunc, HeadlossPipe, LengthPipe As Double
        Const Ground As Double = 9.8


        Dataset = CType(TableJuncCompare, IDataset)
        NetElement = m_GeometricNetwork.Network
        If IndexHead = IndexHeadField1 Then
            IndexHeadFieldJunc = TableJuncCompare.FindField("HeadLossJunc1")
            IndexEleField = TableJuncCompare.FindField("ElevationJunc1")
            IndexVelocityField = TableEdgeCompare.FindField("VelocityPipe1")
            IndexHeadFieldPipe = TableEdgeCompare.FindField("HeadlossALLPipe1")
            IndexLengthField = TableEdgeCompare.FindField("LengthPipe1")

        ElseIf IndexHead = IndexHeadField2 Then
            IndexHeadFieldJunc = TableJuncCompare.FindField("HeadLossJunc2")
            IndexEleField = TableJuncCompare.FindField("ElevationJunc2")
            IndexVelocityField = TableEdgeCompare.FindField("VelocityPipe2")
            IndexHeadFieldPipe = TableEdgeCompare.FindField("HeadlossALLPipe2")
            IndexLengthField = TableEdgeCompare.FindField("LengthPipe2")
        End If

        If IndexEleField = -1 Then Exit Sub
        If IndexHeadFieldJunc = -1 Then Exit Sub
        If IndexVelocityField = -1 Then Exit Sub
        If IndexHeadFieldPipe = -1 Then Exit Sub
        If IndexHeadField1 = -1 Then Exit Sub
        If IndexHeadField2 = -1 Then Exit Sub
        If IndexLengthField = -1 Then Exit Sub

        WorkspaceEdit = CType(Dataset.Workspace, IWorkspaceEdit)

        WorkspaceEdit.StartEditing(True)
        WorkspaceEdit.StartEditOperation()
        MainPipeFeature = m_CollNetworkFeaturesComplexEF.Item(1)
        If MainPipeFeature Is Nothing Then Exit Sub

        MainPipe = CType(MainPipeFeature, IComplexEdgeFeature)
        FromJunc = CType(MainPipe.JunctionFeature(0), ISimpleJunctionFeature)
        FromJuncFeat = CType(FromJunc, IFeature)
        EleFromJunc = FromJuncFeat.Value(m_IndexSourceElevation)

        ToJunc = CType(MainPipe.JunctionFeature(MainPipe.JunctionFeatureCount - 1), ISimpleJunctionFeature)
        ToJuncFeat = CType(ToJunc, IFeature)
        RowTableTo = SearchInTableJuncCompare(ToJuncFeat.OID)
        RowTableEdge = SearchInTableEdgeCompare(MainPipeFeature.OID)

        EleToJunc = RowTableTo.Value(IndexEleField)

        VelocityPipe = RowTableEdge.Value(IndexVelocityField)

        Dim MainHeadloss, Mainlength As Double
        MainHeadloss = RowTableEdge.Value(IndexHeadFieldPipe)
        Mainlength = RowTableEdge.Value(IndexLengthField)

        HeadLoss = EleFromJunc - EleToJunc - (MainHeadloss) - ((VelocityPipe ^ 2) / (2 * Ground))


        RowTableTo.Value(IndexHead) = Math.Abs(Math.Round(HeadLoss, 3))
        RowTableTo.Store()

        For i = 1 To m_CollNetworkFeaturesComplexEF.Count - 1
            PipeFeat = m_CollNetworkFeaturesComplexEF.Item(i + 1)
            Pipe = CType(PipeFeat, IComplexEdgeFeature)
            RowTableEdge = SearchInTableEdgeCompare(PipeFeat.OID)

            FromJunc = CType(Pipe.JunctionFeature(0), ISimpleJunctionFeature)
            FromJuncFeat = CType(FromJunc, IFeature)
            RowTableFrom = SearchInTableJuncCompare(FromJuncFeat.OID)

            ToJunc = CType(Pipe.JunctionFeature(Pipe.JunctionFeatureCount - 1), ISimpleJunctionFeature)
            ToJuncFeat = CType(ToJunc, IFeature)
            RowTableTo = SearchInTableJuncCompare(ToJuncFeat.OID)

            ForwardStar.FindAdjacent(0, FromJunc.EID, EdgeCount)
            For j = 0 To EdgeCount - 1
                ForwardStar.QueryAdjacentEdge(j, AdjEdgeEID, Orient, Weight)
                If Orient Then
                    Dim Feature As IFeature
                    NetElement.QueryIDs(AdjEdgeEID, esriElementType.esriETEdge, lngUserClassId, lngUserID, lngSubID)
                    Feature = m_FCMainPipe.GetFeature(lngUserID)

                    InPipe = CType(Feature, IComplexEdgeFeature)
                    InPipeFeat = Feature
                    RowTableInEdge = SearchInTableEdgeCompare(InPipeFeat.OID)

                    Select Case UtilityNetwork.GetFlowDirection(NetElement.GetEIDs(PipeFeat.Class.ObjectClassID, PipeFeat.OID, esriElementType.esriETEdge).Next)
                        Case esriFlowDirection.esriFDWithFlow
                            Factor = -1
                        Case esriFlowDirection.esriFDAgainstFlow
                            Factor = 1
                    End Select

                    EleFromJunc = Math.Round(RowTableFrom.Value(IndexEleField), 3)
                    EleToJunc = Math.Round(RowTableTo.Value(IndexEleField), 3)
                    VelocityInPipe = Math.Round(RowTableInEdge.Value(IndexVelocityField), 3)
                    VelocityPipe = Math.Round(RowTableEdge.Value(IndexVelocityField), 3)
                    RowTableFrom = SearchInTableJuncCompare(FromJuncFeat.OID)
                    HeadlossFromJunc = Math.Round(RowTableFrom.Value(IndexHeadFieldJunc), 3)
                    HeadlossPipe = Math.Round(RowTableEdge.Value(IndexHeadFieldPipe), 3)
                    LengthPipe = Math.Round(RowTableEdge.Value(IndexLengthField), 3)

                    HeadlossToJunc = (EleFromJunc - EleToJunc) + (((VelocityInPipe ^ 2) / (2 * Ground)) - ((VelocityPipe ^ 2) / (2 * Ground))) + HeadlossFromJunc + (Factor * HeadlossPipe)
                    RowTableTo = SearchInTableJuncCompare(ToJuncFeat.OID)
                    RowTableTo.Value(IndexHead) = Math.Abs(Math.Round(HeadlossToJunc, 3))
                    RowTableTo.Store()

                    Exit For
                End If
            Next

        Next

        WorkspaceEdit.StopEditOperation()
        WorkspaceEdit.StopEditing(True)

        Exit Sub

ErrorMessage:
        WorkspaceEdit.AbortOperation()

    End Sub


    Private Function SearchInTableJuncCompare(ByVal pOID As Integer) As IRow
        Dim Cursor As ICursor
        Dim Row As IRow
        Dim RowCount As Integer

        SearchInTableJuncCompare = Nothing

        RowCount = TableJuncCompare.RowCount(Nothing)
        Cursor = TableJuncCompare.Search(Nothing, False)

        For i = 0 To RowCount - 1
            Row = Cursor.NextRow
            If Row.Value(IndexNumJunc) = pOID Then
                SearchInTableJuncCompare = Row
                Exit For
            End If
        Next

        Marshal.ReleaseComObject(Cursor)
    End Function

    Private Function SearchInTableEdgeCompare(ByVal pOID As Integer) As IRow
        Dim Cursor As ICursor
        Dim Row As IRow
        Dim RowCount As Integer

        SearchInTableEdgeCompare = Nothing

        RowCount = TableEdgeCompare.RowCount(Nothing)
        Cursor = TableEdgeCompare.Search(Nothing, False)

        For i = 0 To RowCount - 1
            Row = Cursor.NextRow
            If Row.Value(IndexNumEdge) = pOID Then
                SearchInTableEdgeCompare = Row
                Exit For
            End If
        Next

        Marshal.ReleaseComObject(Cursor)
    End Function
End Class
