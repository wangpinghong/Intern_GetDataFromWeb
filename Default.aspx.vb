Imports System.Web
Imports System.Net.Sockets 'HttpWebRequest、HttpWebResponse類別
Imports System.Net
Imports System.IO 'StreamReader類
Partial Class _Default

    Inherits System.Web.UI.Page

    Private Property a As Integer
    Private Property b As Integer
    '寫成TXT檔
    Private Property url As String
    Dim file As System.IO.StreamWriter
    Dim itemIndex As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        For l_index As Integer = 1 To ListBox1.Items.Count - 1
            '得到現在所以選擇的網址
            getWeb(ListBox1.SelectedItem.ToString())
            '把得到的資料轉成txt
            outputFile(l_index, "C:\Users\品閎\Desktop\")
        Next
    End Sub
    '抓取名字
    Private Function getWeb(ByVal str As String) As String
        Try
            Dim curItem As String = str
            Return curItem
        Catch ex As Exception
            MsgBox("請選擇要開始的網址")
        End Try
    End Function
    '將得到的資料轉成txt
    Public Sub outputFile(ByVal index As Integer, ByVal cur_path As String)
        Dim request As HttpWebRequest = WebRequest.Create(getWeb(ListBox1.SelectedItem.ToString()))
        Dim mResponse As HttpWebResponse = request.GetResponse()
        Dim sr As New StreamReader(mResponse.GetResponseStream, Encoding.GetEncoding("big5"))
        file = My.Computer.FileSystem.OpenTextFileWriter(cur_path & index & ".txt", True)
        Try
            '迴圈讀data
            Do Until sr.EndOfStream
                Dim strContent = sr.ReadLine().ToString
                '得到日期
                getDate(strContent)
                '得到匯率
                getRate(strContent)
            Loop
            file.Flush()
        Catch ex As Exception

        End Try
        file.Close()
        sr.Close()
        request = Nothing
        mResponse = Nothing
    End Sub
    '刪除html格式的function
    Public Function deHtml(ByVal str As String, ByVal startindex As String, ByVal endindex As String) As String
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
    '抓取日期
    Public Sub getDate(ByVal str As String)
        Dim dateRe As New Regex("\d{4}/\d{2}/\d{2}")
        Dim dateRe2 As New Regex("\d{2}/\d{2}")
        If dateRe.IsMatch(str.ToString) Then
            Dim ans As String
            ans = deHtml(str.ToString(), ">", "<")
            '得到日期
            file.WriteLine(ans)
        ElseIf dateRe2.IsMatch(str.ToString) Then
            Dim ans As String
            ans = deHtml(str.ToString(), ">", "<")
            '得到日期
            file.WriteLine(ans)
        End If
    End Sub
    '抓取匯率
    Public Sub getRate(ByVal str As String)
        Dim rateRe As New Regex("\d{2}.\d{4}")
        Dim rateRe2 As New Regex("\d{1}.\d*")
        If rateRe.IsMatch(str.ToString) Then
            Dim ans As String
            ans = deHtml(str.ToString(), ">", "<")
            '得到匯率
            file.WriteLine(ans)
        End If
    End Sub
    '讀取網頁位置
    Public Sub readFile(ByVal l_box As ListBox, ByVal web_path As String)
        Dim str As IO.StreamReader = New IO.StreamReader(web_path, System.Text.Encoding.Default)
        Do Until str.EndOfStream
            '將讀到的資料存成listboxitems
            l_box.Items.Add(str.ReadLine)
        Loop
        str.Close()
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        readFile(ListBox1, "C:\Users\品閎\Desktop\web.txt")
    End Sub
End Class
