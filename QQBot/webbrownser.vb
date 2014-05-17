Imports System.Threading

Public Class webbrownser

    Private isLoadOver As Boolean
    Private Sub webbrownser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.Visible = False
    End Sub

    Public Function getParamFromWebpage(ByVal url As String, ByVal reqparam As String) As String

    End Function

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        isLoadOver = True
    End Sub

    Private Sub WebBrowser1_Navigated(sender As Object, e As WebBrowserNavigatedEventArgs) Handles WebBrowser1.Navigated
        isLoadOver = True
    End Sub
    Private Sub watchthread()
        While 1
            If isLoadOver = True Then Exit Sub
            Thread.Sleep(200)
        End While

    End Sub
End Class