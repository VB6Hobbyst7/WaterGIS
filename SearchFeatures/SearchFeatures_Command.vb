Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Controls
Imports System.Windows.Forms

<ComClass(SearchFeatures_Command.ClassId, SearchFeatures_Command.InterfaceId, SearchFeatures_Command.EventsId), _
 ProgId("WaterEngine_AnalysisNetwork.SearchFeatures_Command")> _
Public NotInheritable Class SearchFeatures_Command
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "2fcec261-9d37-4e0e-bbb8-71a21166b958"
    Public Const InterfaceId As String = "0cd5f75e-c2ee-44e9-abf2-de013f11f81c"
    Public Const EventsId As String = "f52fe71f-a837-4461-a117-decde45b3404"
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

    Public Overrides Sub OnClick()
        Dim Win32 As IWin32Window
        If m_hookHelper Is Nothing Then Exit Sub
        Dim Search_Form As New SearchFeatures_Form(m_hookHelper)
        Search_Form.Show(Win32)
    End Sub
End Class



