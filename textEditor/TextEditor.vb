'Luke richards
'NETD 2202
'2020/4/3
'text editor


Option Strict On
Imports System.IO

Public Class frmTextEditor
    Dim saved As Boolean = True 'for the confirm prompt
    Dim filePath As String = "" ' effectively what file we are editing
    ' this opens the dialog to get a filepath then calls the save function
    Private Sub saveAs()
        Dim accessedFile As SaveFileDialog = New SaveFileDialog
        accessedFile.Filter = "txt files|*.txt"
        If accessedFile.ShowDialog() = 1 Then
            filePath = accessedFile.FileName
            Me.Text = filePath
            save()
        End If
    End Sub
    ' writes the textbox contents to the file at specified location, creates the file if it does not exist
    Private Sub save()
        If filePath = "" Then
            saveAs()
        Else
            Dim openfile As FileStream = New FileStream(filePath, FileMode.Create, FileAccess.Write)
            Dim writer As StreamWriter = New StreamWriter(openfile)
            writer.Write(txtMain.Text)
            writer.Close()
            openfile.Close()
            saved = True
        End If
    End Sub
    ' calls save function if there unsaved work about to be lost
    Private Sub confirmClose()
        If Not saved Then
            If MsgBox("unsaved work would be lost. would you like to save?", vbYesNo) = vbYes Then
                save()
            End If
        End If
    End Sub
    ' opens and reads files into the textbox, also changes filepath variable
    Private Sub mnuFileOpen_Click(sender As Object, e As EventArgs) Handles mnuFileOpen.Click
        confirmClose()
        Dim reader As StreamReader
        Dim accessedFile As OpenFileDialog = New OpenFileDialog
        Dim openfile As FileStream
        accessedFile.Filter = "txt files|*.txt"
        If accessedFile.ShowDialog() = 1 Then
            filePath = accessedFile.FileName
            Me.Text = filePath
            openfile = New FileStream(filePath, FileMode.Open, FileAccess.Read)
            reader = New StreamReader(openfile)
            txtMain.Text = reader.ReadToEnd
            openfile.Close()
            saved = True
        End If
    End Sub
    ' resets the form
    Private Sub mnuFileNew_Click(sender As Object, e As EventArgs) Handles mnuFileNew.Click
        confirmClose()
        filePath = ""
        Me.Text = "New File"
        txtMain.Text = ""
        saved = True
    End Sub
    'calls the save function
    Private Sub mnuFileSave_Click(sender As Object, e As EventArgs) Handles mnuFileSave.Click
        save()
    End Sub
    ' calls the save as function
    Private Sub mnuFileSaveAs_Click(sender As Object, e As EventArgs) Handles mnuFileSaveAs.Click
        saveAs()
    End Sub
    ' closes program, see formclosing for prompt
    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        Me.Close()
    End Sub
    ' for heathens who dont use shortcuts for copying
    Private Sub mnuEditCopy_Click(sender As Object, e As EventArgs) Handles mnuEditCopy.Click
        Clipboard.SetText(txtMain.SelectedText)
    End Sub
    ' for heathens who dont use shortcuts for cutting
    Private Sub mnuEditCut_Click(sender As Object, e As EventArgs) Handles mnuEditCut.Click
        Clipboard.SetText(txtMain.SelectedText)
        txtMain.SelectedText = ""
    End Sub
    ' for heathens who dont use shortcuts for pasting
    Private Sub mnuEditPaste_Click(sender As Object, e As EventArgs) Handles mnuEditPaste.Click
        txtMain.SelectedText = Clipboard.GetText()
    End Sub
    ' help dialog
    Private Sub mnuHelpAbout_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        MsgBox("NetD 2202" +
               Environment.NewLine +
               Environment.NewLine +
               "Lab 5 " +
               Environment.NewLine +
               Environment.NewLine +
               "Luke Richards")
    End Sub
    ' for checking if the file has been saved
    Private Sub txtMain_TextChanged(sender As Object, e As EventArgs) Handles txtMain.TextChanged
        If saved Then
            Me.Text = Me.Text + "*"
        End If
        saved = False
    End Sub
    ' calls the confirm close when the form would close, even gets alt+f4
    Private Sub frmTextEditor_FormClosing_1(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        confirmClose()
    End Sub
End Class
