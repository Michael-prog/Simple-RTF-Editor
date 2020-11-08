Imports System.ComponentModel

Public Class NotSelectableButton
    Inherits Button
    Public Sub New()
        SetStyle(ControlStyles.Selectable, False)
    End Sub

    <Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible), DefaultValue(False)>
    Public Property ClickNotify As Boolean = False

    Private Sub NotSelectableButton_Click(sender As Object, e As EventArgs) Handles Me.Click
        If Not ClickNotify Then Exit Sub
        If Parent Is Nothing Then Exit Sub
        If GetType(IClickNotify).IsAssignableFrom(Parent.GetType()) Then CType(Parent, IClickNotify).ClickNotify(Me)
    End Sub
End Class
