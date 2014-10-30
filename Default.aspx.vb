Imports System.Web
Imports System.Net.Sockets 'HttpWebRequest、HttpWebResponse類別
Imports System.Net
Imports System.IO 'StreamReader類
Imports System.Text.RegularExpressions
Partial Class _Default

    Inherits System.Web.UI.Page
    Private Property a As Integer
    Private Property b As Integer
    '寫成TXT檔
    Private Property url As String
    Dim file As System.IO.StreamWriter
    Dim itemIndex As Integer = 0
    Dim curItem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub
    '全部output的button
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        For l_index As Integer = 0 To ListBox1.Items.Count - 1
            '得到網址
            getWeb(ListBox1.Items(l_index).ToString)
            '把得到的資料轉成txt
            outputFile(l_index, "C:\Users\品閎\Desktop\")
        Next
    End Sub
    '選取output的button
    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
    End Sub
    '讀取網站資料
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        readFile(ListBox1, "C:\Users\品閎\Desktop\web")
    End Sub
    '抓取名字
    Private Function getWeb(ByVal str As String) As String
        Try
            curItem = str
            Return curItem
        Catch ex As Exception
            MsgBox("請選擇要開始的網址")
        End Try
    End Function
    '將得到的資料輸出成txt檔
    Public Sub outputFile(ByVal index As Integer, ByVal cur_path As String)
        Dim request As HttpWebRequest = WebRequest.Create(curItem)
        Dim mResponse As HttpWebResponse = request.GetResponse()

        Dim sr As New StreamReader(mResponse.GetResponseStream, Encoding.GetEncoding("big5"))
        file = My.Computer.FileSystem.OpenTextFileWriter(cur_path & index & ".txt", True)

        Try
            '判斷是否有符合格式
            match(sr)
        Catch ex As Exception

        End Try
        file.Close()
        sr.Close()
        request = Nothing
        mResponse = Nothing
    End Sub

    '讀取要抓取的網站的位置
    Public Sub readFile(ByVal l_box As ListBox, ByVal web_path As String)
        Dim str As IO.StreamReader = New IO.StreamReader(web_path & ".txt", System.Text.Encoding.Default)
        Do Until str.EndOfStream
            '將讀到的資料存成listboxitems
            l_box.Items.Add(str.ReadLine)
        Loop
        str.Close()
    End Sub
    '判斷是否有符合要求格式
    Private Sub match(ByVal sr As StreamReader)
        Try
            Dim patternDate As String = "\d{4}/\d{1,2}/\d{1,2}"
            Dim patternRate As String = "\d{1,3}.\d{1,4}"
            ' 定義RE
            Dim rDate As Regex = New Regex(patternDate, RegexOptions.IgnoreCase)
            Dim rRate As Regex = New Regex(patternRate, RegexOptions.IgnoreCase)
            '各別讀取兩行
            Dim fir = sr.ReadLine
            Dim sec = sr.ReadLine
            '讀取網頁至結尾
            Do Until sr.EndOfStream
                Dim mDate As Match = rDate.Match(fir)
                Dim mRate As Match = rRate.Match(sec)
                '當兩者皆符合的時候，便抓取兩行的資料
                Do While mDate.Success And mRate.Success
                    file.WriteLine(mDate)
                    file.WriteLine(mRate)
                    mDate = mDate.NextMatch()
                    mRate = mRate.NextMatch()
                Loop
                '讀取格行，並非單雙單雙
                mDate = mDate.NextMatch()
                fir = sec
                sec = sr.ReadLine
            Loop
        Catch ex As Exception

        End Try
    End Sub

End Class
