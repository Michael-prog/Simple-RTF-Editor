
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.WindowsRuntime
Imports Windows.ApplicationModel.DataTransfer
Imports Windows.Foundation

<ComImport, Guid("3A3DCD6C-3EAB-43DC-BCDE-45671CE800C8")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IDataTransferManagerInterOp
    <PreserveSig>
    Function GetForWindow(
    <[In]> ByVal appWindow As IntPtr,
    <[In]> ByRef riid As Guid,
    <Out> ByRef pDataTransferManager As DataTransferManager) As UInteger
    <PreserveSig>
    Function ShowShareUIForWindow(ByVal appWindow As IntPtr) As UInteger
End Interface


Public Class Windows10Share

    Dim DataTransferManager As IDataTransferManagerInterOp = CType(WindowsRuntimeMarshal.GetActivationFactory(GetType(DataTransferManager)), IDataTransferManagerInterOp)

    Public ReadOnly Property BaseForm As Form

    Public Sub New(BaseForm As Form, CallBack As TypedEventHandler(Of DataTransferManager, DataRequestedEventArgs))
        Me.BaseForm = BaseForm
        Dim x As DataTransferManager = Nothing
        DataTransferManager.GetForWindow(BaseForm.Handle, New Guid("a5caee9b-8708-49d1-8d36-67d25a8da00c"), x)
        AddHandler x.DataRequested, CallBack
    End Sub

    Public Sub ShowShareUI()
        DataTransferManager.ShowShareUIForWindow(BaseForm.Handle)
    End Sub

End Class
