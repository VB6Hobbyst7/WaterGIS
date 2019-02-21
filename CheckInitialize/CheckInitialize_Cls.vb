Imports ESRI.ArcGIS.esriSystem

Public Class CheckInitialize_Cls

    Private WithEvents CheckInitialize_Frm As CheckInitialize_Form

    Public Function CheckInitialize(ByVal LicenseStatus As esriLicenseStatus) As Boolean
        CheckInitialize = False
        'If license initialization failed,
        If LicenseStatus <> esriLicenseStatus.esriLicenseCheckedOut Then
            CheckInitialize = True
            MsgBox(String.Format("لطفاً قبل از اجرای نرم افزار پیش نیازهای استفاده از نرم افزار  " & vbNewLine & "را بطور کامل بر روی سیستم خود نصب نمایید"))
        End If
    End Function

    Private Sub CheckInitialize_Frm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles CheckInitialize_Frm.FormClosing
        CheckInitialize_Frm = Nothing
    End Sub
End Class
