Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class MainForm
    Inherits Form

    Private label As Label

    Public Sub New()
        label = New Label() With {.Text = "Initial Text", .Dock = DockStyle.Top}
        Me.Controls.Add(label)
        Me.Load += Async Sub(sender, e) Await UpdateLabelAsync()
    End Sub

    Private Async Function UpdateLabelAsync() As Task
        ' Simulate a background task
        Await Task.Run(Sub()
                           Thread.Sleep(2000) ' Simulate work
                       End Sub)

        ' Update UI on the main thread
        label.Invoke(Sub() label.Text = "Updated Text")
    End Function
End Class

Module Program
    Sub Main()
        Application.Run(New MainForm())
    End Sub
End Module
