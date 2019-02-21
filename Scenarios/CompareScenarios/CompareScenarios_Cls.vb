Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Controls
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing

Public Class CompareScenarios_Cls

    Private m_hookHelper As IHookHelper = Nothing
    Private WithEvents FormCompare As CompareTablesScenarios_Form
    Private FirstScenarioName, SecondScenarioName As String
    Private CalculateCompare As CalculateCompareScenarios_Cls
    Private CompareTablesScenarios_Cmd As CompareTablesScenarios_Command
    Private CompareGraphicalScenarios_Cmd As CompareGraphicalScenarios_Command
    Private ResultMessage As String
    Private ScreenRectangle As Rectangle
    Private ProgressBar_MainForm As ProgressBar

    Public Sub New(ByVal hookHelper As IHookHelper)
        m_hookHelper = hookHelper
    End Sub

    Public WriteOnly Property Get_FirstScenarioName() As String
        Set(ByVal value As String)
            FirstScenarioName = value
        End Set
    End Property

    Public WriteOnly Property Get_SecondScenarioName() As String
        Set(ByVal value As String)
            SecondScenarioName = value
        End Set
    End Property

    Public WriteOnly Property Set_ScreenRectangle() As Rectangle
        Set(ByVal value As Rectangle)
            ScreenRectangle = value
        End Set
    End Property

    Public WriteOnly Property Set_Prog() As ProgressBar
        Set(ByVal value As ProgressBar)
            ProgressBar_MainForm = value
        End Set
    End Property

    Public Function CompareScenarios() As Boolean
        On Error GoTo Err

        Dim IDScenario1, IDScenario2 As Integer
        Dim Row As IRow

        CompareScenarios = False
        ResultMessage = ""
        If FirstScenarioName = SecondScenarioName Then
            ResultMessage = "Please first choose two different scenarios"
            GoTo Err
        End If
        ProgressBar(10, 30)

        If FirstScenarioName = "Base Scenario" Then
            IDScenario1 = 0
        Else
            Row = Find_IDScenario(FirstScenarioName)
            If Row Is Nothing Then
                ' ResultMessage = "اطلاعات مربوط به سناریوی اول یافت نشد"
                ResultMessage = "There is not first scenario"
                GoTo Err
            End If
            IDScenario1 = Row.Value(m_TableScenarios.FindField(m_TableScenarios.OIDFieldName))
        End If

        If SecondScenarioName = "Base Scenario" Then
            IDScenario2 = 0
        Else
            Row = Find_IDScenario(SecondScenarioName)
            If Row Is Nothing Then
                'ResultMessage = "اطلاعات مربوط به سناریوی دوم یافت نشد"
                ResultMessage = "There is not second scenario"
                GoTo Err
            End If
            IDScenario2 = Row.Value(m_TableScenarios.FindField(m_TableScenarios.OIDFieldName))
        End If

        If CalculateCompare Is Nothing Then
            CalculateCompare = New CalculateCompareScenarios_Cls(m_hookHelper)
        End If

        CalculateCompare.Get_IDScenario1 = IDScenario1
        CalculateCompare.Get_IDScenario2 = IDScenario2
        CalculateCompare.Get_IDScenarioBase = 0

        If Not CalculateCompare.CalculationCompare() Then GoTo Err

        ' System.Threading.Thread.Sleep(500)
        ProgressBar(40, 30)

        If Not CalculateHeadlossJunctionScenario_Cls.CorrectCalculation Then
            CalculateCompare = Nothing
            GoTo Err
        End If
        ' System.Threading.Thread.Sleep(500)
        ProgressBar(70, 20)

        CompareScenarios = True
        Exit Function
Err:
        If ResultMessage = "" Then
            ' ResultMessage = "اشکال در مقایسه سناریوها"
            ResultMessage = "Failure in compare two scenarios"
        End If
        ProgressBar_MainForm.Value = 0
        Show_AlertCustom(ResultMessage)
    End Function

    Private Sub ProgressBar(ByVal Value As Integer, ByVal StepProg As Integer)
        ProgressBar_MainForm.Value = Value
        ProgressBar_MainForm.Step = StepProg
        ProgressBar_MainForm.PerformStep()
        Application.DoEvents()
       
    End Sub


    Private Function Find_IDScenario(ByVal NameScenario As String) As IRow
        Dim Row As IRow
        Dim Cursor As ICursor
        Dim RowCount As Integer
        Dim IndexNameScenario As Integer

        Cursor = m_TableScenarios.Search(Nothing, False)
        RowCount = m_TableScenarios.RowCount(Nothing)
        IndexNameScenario = m_TableScenarios.FindField("NameScenario")

        Find_IDScenario = Nothing
        Row = Cursor.NextRow
        Do Until Row Is Nothing
            If Row.Value(IndexNameScenario) = NameScenario Then
                Find_IDScenario = Row
                Exit Function
            End If
            Row = Cursor.NextRow
        Loop

        Marshal.ReleaseComObject(Cursor)
    End Function

    Private Sub Show_AlertCustom(ByVal MessageError As String)
        Dim AlertCustom As New AlertCustom_Cls
        AlertCustom.Set_ScreenRectangle = ScreenRectangle
        AlertCustom.ShowLoadAlert(Nothing, MessageError)
    End Sub
End Class
