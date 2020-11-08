Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Module Extension
    <Extension()>
    Public Sub SetInnerMargins(ByVal textBox As TextBoxBase, margin As Integer)
        Dim rect = New RECT()
        SendMessage(textBox.Handle, EmGetrect, IntPtr.Zero, rect)
        rect.Left += margin
        rect.Top += margin
        rect.Bottom -= margin
        rect.Right -= margin
        SendMessage(textBox.Handle, EmSetrect, IntPtr.Zero, rect)
    End Sub

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer

        Private Sub New(ByVal left As Integer, ByVal top As Integer, ByVal right As Integer, ByVal bottom As Integer)
            left = left
            top = top
            right = right
            bottom = bottom
        End Sub

        Public Sub New(ByVal r As Rectangle)
            Me.New(r.Left, r.Top, r.Right, r.Bottom)
        End Sub
    End Structure

    <DllImport("user32.dll", EntryPoint:="SendMessage", CharSet:=CharSet.Auto)>
    Private Function SendMessage(ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As IntPtr, ByRef lParam As RECT) As Integer : End Function
    Private Const EmGetrect As Integer = &HB2
    Private Const EmSetrect As Integer = &HB3
End Module
