Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Controls
Imports System.Windows.Forms

<ComClass(CompareTablesScenarios_Command.ClassId, CompareTablesScenarios_Command.InterfaceId, CompareTablesScenarios_Command.EventsId), _
 ProgId("WaterEngine_AnalysisNetwork.CompareTablesScenarios_Command")> _
Public NotInheritable Class CompareTablesScenarios_Command
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "631083c5-b09f-4d3f-b4fb-bd6c9791f988"
    Public Const InterfaceId As String = "6dcfb590-3823-45f4-b53c-ca8078fc43af"
    Public Const EventsId As String = "aaed7864-841a-43f6-8e0c-9df093c43f91"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        ControlsCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        ControlsCommands.Unregister(regKey)

    End Sub

#End Region
#End Region


    Private m_hookHelper As IHookHelper
    Private WithEvents FrmCompareTables As CompareTablesScenarios_Form
    Private FirstScenarioName, SecondScenarioName As String
    Private ProgressBar_MainForm As ProgressBar

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = ""  'localizable text 
        MyBase.m_caption = ""   'localizable text 
        MyBase.m_message = ""   'localizable text 
        MyBase.m_toolTip = "" 'localizable text 
        MyBase.m_name = ""  'unique id, non-localizable (e.g. "MyCategory_MyCommand")

        Try
            'TODO: change bitmap name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try
    End Sub


    Public Overrides Sub OnCreate(ByVal hook As Object)
        If m_hookHelper Is Nothing Then m_hookHelper = New HookHelperClass

        If Not hook Is Nothing Then
            m_hookHelper.Hook = hook
        End If

        If m_hookHelper.ActiveView Is Nothing Then m_hookHelper = Nothing
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

    Public WriteOnly Property Set_Prog() As ProgressBar
        Set(ByVal value As ProgressBar)
            ProgressBar_MainForm = value
        End Set
    End Property

    Public Overrides Sub OnClick()
        On Error GoTo Err
        Dim Win32 As IWin32Window

        If m_hookHelper Is Nothing Then Return
        If FrmCompareTables Is Nothing Then
            FrmCompareTables = New CompareTablesScenarios_Form(m_hookHelper)
            FrmCompareTables.Get_FirstScenarioName = FirstScenarioName
            FrmCompareTables.Get_SecondScenarioName = SecondScenarioName
            FrmCompareTables.Set_Prog = ProgressBar_MainForm
            FrmCompareTables.Show(Win32)
        End If
        Exit Sub
Err:

    End Sub

    Public Sub Close_Form()
        If FrmCompareTables IsNot Nothing Then FrmCompareTables.Close()
    End Sub

    Private Sub FrmCompareTables_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles FrmCompareTables.FormClosing
        ProgressBar_MainForm = Nothing
        FrmCompareTables = Nothing
    End Sub
End Class



