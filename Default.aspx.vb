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
        '資料網址
        Dim request As HttpWebRequest = WebRequest.Create("http://cathlifefund.moneydj.com/w/wb/wb02_SHZ96-BRU029.djhtm")
        Dim mResponse As HttpWebResponse = request.GetResponse()
        Dim sr As New StreamReader(mResponse.GetResponseStream, Encoding.GetEncoding("big5"))
        '寫成TXT檔
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("C:\Users\品閎\Desktop\test.txt", True)
        '宣告日期和匯率re
        Dim dateRe As New Regex("\d{4}/\d{2}/\d{2}")
        Dim rateRe As New Regex("\d{2}.\d{4}")
        '迴圈讀data
        Do Until sr.EndOfStream
            Dim strContent = sr.ReadLine().ToString
            '用dateRe判斷是否為日期
            If dateRe.IsMatch(strContent.ToString) Then
                Dim ans As String
                ans = deHtml(strContent.ToString(), ">", "<")
                '得到日期
                file.WriteLine("日期" & ans)
            End If
            '用rateRe判斷是否為匯率
            If rateRe.IsMatch(strContent.ToString) Then
                Dim ans As String
                ans = deHtml(strContent.ToString(), ">", "<")
                '得到匯率
                file.WriteLine("匯率" & ans)
            End If
            'End If

            file.Flush()
        Loop
        file.Close()
        sr.Close()
        request = Nothing
        mResponse = Nothing
    End Sub
    '刪除html格式的function
    Public Function deHtml(ByVal str, ByVal startindex, ByVal endindex)
        Dim s = str
        '取得開頭index
        a = (InStr(s, startindex))
        '刪除前半
        s = Mid(s, a + 1)
        '取得結尾index
        b = (InStr(s, endindex))
        '刪除後半
        If b > 0 Then
            s = Mid(s, 1, b - 1)
        End If

        deHtml = s '完成
    End Function

End Class
