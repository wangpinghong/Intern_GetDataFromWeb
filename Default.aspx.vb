Imports System.Web
Imports System.Net.Sockets 'HttpWebRequest、HttpWebResponse類別
Imports System.Net
Imports System.IO 'StreamReader類
Partial Class _Default

    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
     
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Label1.Text = ""
        '要抓的網址
        Dim request As HttpWebRequest = WebRequest.Create("http://cathlifefund.moneydj.com/w/wb/wb02_SHZ96-BRU029.djhtm")
        Dim mResponse As HttpWebResponse = request.GetResponse()
        Dim sr As New StreamReader(mResponse.GetResponseStream, Encoding.GetEncoding("big5"))
        '寫出成TXT檔
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("C:\Users\品閎\Desktop\test.txt", True)
        '一行一行讀
        Do Until sr.EndOfStream
            Dim strContent = sr.ReadLine()
            If strContent.StartsWith("2014") Then
                file.WriteLine(strContent.ToString)
                strContent = sr.ReadLine
                file.WriteLine("匯率為" & strContent.ToCharArray)
            End If
            file.Flush()
        Loop
        file.Close()
        sr.Close()
        request = Nothing
        mResponse = Nothing
    End Sub
End Class
