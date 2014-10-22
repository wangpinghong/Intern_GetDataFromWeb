Imports System.Web
Imports System.Net.Sockets 'HttpWebRequest、HttpWebResponse類別
Imports System.Net
Imports System.IO 'StreamReader類
Partial Class _Default

    Inherits System.Web.UI.Page

    Private Property a As Integer
    Private Property b As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
     
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Label1.Text = ""
        '資料網址
        Dim request As HttpWebRequest = WebRequest.Create("http://cathlifefund.moneydj.com/w/wb/wb02_SHZ96-BRU029.djhtm")
        Dim mResponse As HttpWebResponse = request.GetResponse()
        Dim sr As New StreamReader(mResponse.GetResponseStream, Encoding.GetEncoding("big5"))
        '寫成TXT檔
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("C:\Users\品閎\Desktop\test.txt", True)
        '迴圈讀data
        Do Until sr.EndOfStream
            Dim strContent = sr.ReadLine()
            '判斷是否為表格內資料
            If strContent.Contains("2014") Then
                'Dim ans As String
                'ans = GetData(strContent.ToString(), ">", "<") '得到日期
                'file.WriteLine(ans)
                'strContent = sr.ReadLine
                ''相同方法取得匯率
                'ans = GetData(strContent.ToString(), ">", "<") '得到匯率
                'file.WriteLine("匯率為" & ans)
            End If
            file.Flush()
        Loop
        file.Close()
        sr.Close()
        request = Nothing
        mResponse = Nothing
    End Sub
    Public Function GetData(ByVal str, ByVal startindex, ByVal endindex)
        a = (InStr(str, startindex)) '取得開頭index
        str = Mid(str, a + 1) '刪除前半
        b = (InStr(str, endindex)) '取得結尾index
        If b > 0 Then
            str = Mid(str, b - 1) '刪除後半
        End If
        GetData = str '完成
    End Function
End Class
