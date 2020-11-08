Imports System.Runtime.InteropServices
Imports Windows.ApplicationModel.DataTransfer
Imports Windows.Foundation
Imports Windows.Storage

Public Class Form1
    Implements IClickNotify

#Region "Variables"
    Public ReadOnly Property Path As String
    Public ReadOnly Property Saved As Boolean = True
#End Region

#Region "General"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.SetInnerMargins(10)
        SetSavedIndicator()
    End Sub

    Private Sub RichTextBox1_KeyDown_1(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyData = (Keys.Control Or Keys.S) Then
            Save()
            Exit Sub
        End If
        If e.KeyData = (Keys.Control Or Keys.O) Then
            Open()
            Exit Sub
        End If
    End Sub

    Private Sub Form1_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles MyBase.PreviewKeyDown
        If e.KeyData = (Keys.Control Or Keys.S) Then
            Save()
            Exit Sub
        End If
        If e.KeyData = (Keys.Control Or Keys.O) Then
            Open()
            Exit Sub
        End If
        If e.KeyData = (Keys.Control Or Keys.Z) Then
            Undo()
            Exit Sub
        End If
        If e.KeyData = (Keys.Control Or Keys.Y) Then
            Redo()
            Exit Sub
        End If
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        MadeChanges(sender)
    End Sub

    Public Sub MadeChanges(sender As Object) Implements IClickNotify.ClickNotify
        SavedLabel.Text = "Unsaved Changes"
        If Not Me.Text.EndsWith("*") Then Me.Text += "*"
        _Saved = False
    End Sub

    Public Sub SetSavedIndicator()
        SavedLabel.Text = "Saved"
        If Me.Text.EndsWith("*") Then Me.Text = Me.Text.Substring(0, Me.Text.Length - 1)
        _Saved = True
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Not Saved Then
            Dim result = MessageBox.Show("You have Unsaved Changes!" + vbNewLine + "Do you want to Save before quitting the application?!", "Unsaved Changes", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation)
            If result = DialogResult.Cancel Then e.Cancel = True
            If result = DialogResult.Yes Then Save()
        End If
    End Sub

#End Region

#Region "MenuStrip"
    Private Sub Save() Handles Button10.Click
        If Path = "" Then
            Dim dialog As New SaveFileDialog()
            dialog.Filter = "Meine coole TextDatei|*.rtf"
            If dialog.ShowDialog() = DialogResult.OK Then
                RichTextBox1.SaveFile(dialog.FileName)
                _Path = dialog.FileName
                SetSavedIndicator()
            End If
        Else
            RichTextBox1.SaveFile(Path)
            SetSavedIndicator()
        End If
    End Sub

    Private Sub Open() Handles Button11.Click
        Dim dialog As New OpenFileDialog()
        dialog.Filter = "Meine coole TextDatei|*.rtf"
        If dialog.ShowDialog() = DialogResult.OK Then
            RichTextBox1.LoadFile(dialog.FileName)
            _Path = dialog.FileName
        End If
        SetSavedIndicator()
    End Sub

    Private Sub Undo() Handles Button12.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub Redo() Handles Button13.Click
        RichTextBox1.Redo()
    End Sub
#End Region

#Region "ToolBar"

#Region "#1"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NotSelectableButton34.Click, NotSelectableButton13.Click, NotSelectableButton1.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, NotSelectableButton33.Click, NotSelectableButton2.Click, NotSelectableButton14.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NotSelectableButton31.Click, NotSelectableButton3.Click, NotSelectableButton15.Click
        RichTextBox1.Copy()
    End Sub
#End Region

#Region "#2"
    Sub ToggleFontStyle(style As FontStyle)
        Dim font = RichTextBox1.SelectionFont
        If font.Style.HasFlag(style) Then
            RichTextBox1.SelectionFont = New Font(font, font.Style And Not style)
        Else
            RichTextBox1.SelectionFont = New Font(font, font.Style Or style)
        End If
        RichTextBox1_SelectionChanged()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, NotSelectableButton4.Click, NotSelectableButton30.Click, NotSelectableButton16.Click
        ToggleFontStyle(FontStyle.Bold)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, NotSelectableButton5.Click, NotSelectableButton29.Click, NotSelectableButton17.Click
        ToggleFontStyle(FontStyle.Italic)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NotSelectableButton6.Click, NotSelectableButton27.Click, NotSelectableButton18.Click
        ToggleFontStyle(FontStyle.Strikeout)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, NotSelectableButton7.Click, NotSelectableButton26.Click, NotSelectableButton19.Click
        ToggleFontStyle(FontStyle.Underline)
    End Sub

    Private Sub RichTextBox1_SelectionChanged() Handles RichTextBox1.SelectionChanged
        If RichTextBox1.SelectionFont Is Nothing Then Exit Sub
        Dim style = RichTextBox1.SelectionFont.Style
        If style.HasFlag(FontStyle.Bold) Then
            Button4.FlatAppearance.BorderSize = 1
        Else
            Button4.FlatAppearance.BorderSize = 0
        End If
        If style.HasFlag(FontStyle.Italic) Then
            Button5.FlatAppearance.BorderSize = 1
        Else
            Button5.FlatAppearance.BorderSize = 0
        End If
        If style.HasFlag(FontStyle.Strikeout) Then
            Button6.FlatAppearance.BorderSize = 1
        Else
            Button6.FlatAppearance.BorderSize = 0
        End If
        If style.HasFlag(FontStyle.Underline) Then
            Button7.FlatAppearance.BorderSize = 1
        Else
            Button7.FlatAppearance.BorderSize = 0
        End If
    End Sub
#End Region

#Region "#3"
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click, NotSelectableButton8.Click, NotSelectableButton25.Click, NotSelectableButton20.Click
        Dim dialog As New ColorDialog()
        dialog.Color = RichTextBox1.SelectionColor
        If dialog.ShowDialog() = DialogResult.OK Then
            RichTextBox1.SelectionColor = dialog.Color
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click, NotSelectableButton9.Click, NotSelectableButton28.Click, NotSelectableButton21.Click
        Dim dialog As New FontDialog()
        dialog.Font = RichTextBox1.SelectionFont
        If dialog.ShowDialog() = DialogResult.OK Then
            RichTextBox1.SelectionFont = dialog.Font
        End If
    End Sub
#End Region

#Region "#4"
    Private Sub Standart() Handles Button14.Click
        RichTextBox1.SelectionFont = New Font("Calibri", 11)
        RichTextBox1.SelectionColor = Color.Black
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click, NotSelectableButton35.Click, NotSelectableButton23.Click, NotSelectableButton11.Click
        RichTextBox1.SelectionFont = New Font("Calibri", 28)
        RichTextBox1.SelectionColor = Color.Black
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click, NotSelectableButton36.Click, NotSelectableButton24.Click, NotSelectableButton12.Click
        RichTextBox1.SelectionFont = New Font("Calibri", 16)
        RichTextBox1.SelectionColor = Color.FromArgb(47, 84, 150)
    End Sub

    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyData = Keys.Enter Then
            Standart()
        End If
    End Sub
#End Region

#End Region

#Region "Share"
    Public ReadOnly Property ShareManager As Windows10Share

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim ref As DataTransferManager = Nothing
        _ShareManager = New Windows10Share(Me, Sub(s As DataTransferManager, ev As DataRequestedEventArgs)
                                                   Dim data = ev.Request.Data
                                                   If Path = "" Then
                                                       ev.Request.FailWithDisplayText("Keine Datei geöffnet / gespeichert!")
                                                   End If
                                                   data.Properties.Title = IO.Path.GetFileName(Path)
                                                   data.Properties.Description = "RTF Editor"
                                                   Dim x = StorageFile.GetFileFromPathAsync(Path)
                                                   While x.Status = AsyncStatus.Started : End While
                                                   data.SetStorageItems({x.GetResults()})
                                               End Sub)
    End Sub

    Private Sub Share(sender As Object, e As EventArgs) Handles Button17.Click
        ShareManager.ShowShareUI()
    End Sub

    Private Sub Standart(sender As Object, e As EventArgs) Handles Button14.Click, NotSelectableButton32.Click, NotSelectableButton22.Click, NotSelectableButton10.Click

    End Sub

#End Region

End Class
