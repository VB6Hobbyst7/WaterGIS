Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Controls
Imports System.Windows.Forms

<ComClass(EditAttributeFeatures_Command.ClassId, EditAttributeFeatures_Command.InterfaceId, EditAttributeFeatures_Command.EventsId), _
 ProgId("WaterEngine_DesignNetwork.EditAttributeFeatures_Command")> _
Public NotInheritable Class EditAttributeFeatures_Command
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "93e0677d-2870-482e-8d87-ea81acf383e3"
    Public Const InterfaceId As String = "b063a566-0d1d-49de-835d-8a9033c55e71"
    Public Const EventsId As String = "a71b2361-4eef-47b5-b97a-07a646eebd28"
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
    Private WithEvents EditAttribute_Frm As EditAttributeForm

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
        If EditAttribute_Frm Is Nothing Then
            EditAttribute_Frm = New EditAttributeForm(m_hookHelper)
            EditAttribute_Frm.Show(Win32)
        End If
    End Sub

    Private Sub EditAttribute_Frm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles EditAttribute_Frm.FormClosing
        EditAttribute_Frm = Nothing
    End Sub

    Public Sub Close_Form()
        If EditAttribute_Frm IsNot Nothing Then EditAttribute_Frm.Close()
    End Sub
End Class



