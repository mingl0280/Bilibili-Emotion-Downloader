Imports System.Net
Imports System.Text.RegularExpressions
Imports System.IO

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate("https://www.bilibili.com/video/av7/")

    End Sub

    Private DownList As New List(Of DownloadFileAttr)

    Private Class DownloadFileAttr

        Public Property Uri As Uri
        Public Property Size As Size

        Public Property Name As String
        Public Property FileType As String

        Public Property FileName As String

        Public Property Category As String

        Public Property Path As String

        Public Sub New(src As String, title As String)
            If src.StartsWith("//") Then
                src = "https:" + src
            End If
            Uri = New Uri(src)
            Size = getSize(src)
            Name = title.Replace("[", "").Replace("]", "")
            FileType = getFileType(src)
            FileName = Name + "." + FileType
            Category = getCategory(title)
        End Sub

        Private Function getSize(src As String) As Size
            Dim R As New Regex("\d\dx\d\d")
            Dim Rsul = R.Match(src).ToString().Split("x")
            Return New Size(Val(Rsul(0)), Val(Rsul(1)))
        End Function

        Private Function getFileType(src As String) As String
            Return src.Substring(src.LastIndexOf(".") + 1, src.Length - src.LastIndexOf(".") - 1)
        End Function

        Private Function getCategory(title As String) As String
            Return title.Substring(1, title.IndexOf("_") - 1)
        End Function

    End Class

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CurAddList As New List(Of String)
        Dim o = WebBrowser1.Document.GetElementsByTagName("img")
        For Each Img As HtmlElement In o
            Debug.WriteLine(Img.Parent.GetAttribute("Class"))
            Debug.WriteLine(Img.Parent.InnerHtml)
            If Img.Parent.GetAttribute("className") = "emoji-list emoji-icon" Then
                DownList.Add(New DownloadFileAttr(Img.GetAttribute("src"), Img.GetAttribute("title")))
                CurAddList.Add(Img.GetAttribute("title"))
            End If
        Next
        Dim str As String = ""
        For Each item In CurAddList
            str += item + vbCrLf
        Next
        MessageBox.Show("已缓存：" + vbCrLf + str, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        WebBrowser1.Navigate(TextBox1.Text)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim foldbox As New FolderBrowserDialog()
        Dim Ret = foldbox.ShowDialog()
        If Ret = DialogResult.OK Then
            For Each Item In DownList
                Item.Path = foldbox.SelectedPath
                Dim dl As New WebClient
                AddHandler dl.DownloadDataCompleted, AddressOf onDlEventFinished

                dl.DownloadDataAsync(Item.Uri, Item)
            Next
        End If
    End Sub
    Private Sub onDlEventFinished(sender As Object, e As DownloadDataCompletedEventArgs)
        Dim datas() As Byte = e.Result
        Dim Item As DownloadFileAttr = e.UserState
        Dim folderpath As String = Item.Path + "\" + Item.Category + "\"
        If Not Directory.Exists(folderpath) Then
            Directory.CreateDirectory(folderpath)
        End If
        Dim FileFullPath As String = folderpath + Item.FileName
        Dim fsw As FileStream = File.Create(FileFullPath)
        fsw.Write(datas, 0, datas.LongLength)
        fsw.Close()
        Label2.Text = Item.Name + "Download Complete."
    End Sub
End Class
