Imports System.Web
Imports System.Net.Sockets 'HttpWebRequest、HttpWebResponse類別
Imports System.Net
Imports System.IO 'StreamReader類
Imports System.Text.RegularExpressions
Partial Class _Default
    Inherits System.Web.UI.Page
    Dim file As System.IO.StreamWriter
    Dim itemIndex As Integer = 0
    Dim cur_Item As String
    Dim patternType As String = "\d{4}/\d{1,2}/\d{1,2}.{0,}\n.{0,}\d{1,3}.\d{1,7}" 'type1 date:2014/01/01,rate:~.~
    Dim patternType2 As String = "\d{1,2}/\d{1,2}.{0,}\n.{0,}\d{1,3}.\d{1,7}" 'type2 date:01/01,rate:~.~
    Dim patternDate As String = "\d{4}/\d{1,2}/\d{1,2}" 'date  2014/01/01
    Dim patternDate2 As String = "\d{1,2}/\d{1,2}" ' date  01/01
    Dim patternRate As String = "\d{1,3}\.\d{1,7}" 'rate 
    Dim mDate, mRate, mdate2 As Match
    Dim rDate As Regex = New Regex(patternDate, RegexOptions.IgnoreCase)
    Dim rDate2 As Regex = New Regex(patternDate2, RegexOptions.IgnoreCase)
    Dim rRate As Regex = New Regex(patternRate, RegexOptions.IgnoreCase)
    Dim rType As Regex = New Regex(patternType, RegexOptions.IgnoreCase)
    Dim rType2 As Regex = New Regex(patternType2, RegexOptions.IgnoreCase)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub
    '全部output的button
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        'output檔案
        For l_index As Integer = 0 To ListBox1.Items.Count - 1
            getWeb(ListBox1.Items(l_index).ToString) '按序得到網址
            newFile(l_index, "C:\Users\品閎\Desktop\") '創造新的檔案
            getWebContent(l_index, "C:\Users\品閎\Desktop\") '得到內容
            file.Close()
        Next
        GC.Collect()
    End Sub
    '讀取網站資料
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        readFile(ListBox1, "C:\Users\品閎\Desktop\web")
    End Sub
    '得到網址名稱
    Private Sub getWeb(ByVal str As String)
        Try
            cur_Item = str
        Catch ex As Exception
            MsgBox("請選擇要開始的網址")
        End Try
    End Sub
    Private Sub newFile(ByVal index As Integer, ByVal cur_path As String)
        file = My.Computer.FileSystem.OpenTextFileWriter(cur_path & index & ".txt", True)
    End Sub
    '將得到的資料輸出成txt檔
    Public Sub getWebContent(ByVal index As Integer, ByVal cur_path As String)
        '讀取網頁內容
        Dim request As HttpWebRequest = WebRequest.Create(cur_Item)
        Dim mResponse As HttpWebResponse = request.GetResponse()
        Dim sr As New StreamReader(mResponse.GetResponseStream, Encoding.GetEncoding("big5"))
        Try
            '判斷是否有符合格式
            match(sr)
        Catch ex As Exception
        End Try
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
            Dim type As Integer
            '先將資料讀進來
            Dim text = sr.ReadToEnd.ToString
            '判斷讀進來的網頁是否為tpye1
            Dim mcDate As MatchCollection = rType.Matches(text)
            '判斷讀進來的網頁是否為type2
            If mcDate.Count = 0 Then
                mcDate = rType2.Matches(text)
                type = 2
            Else
                type = 1
            End If
            '將這兩行中個別的資料取出來
            For Each m As Match In mcDate
                getDate(m, type)
                getRate(m)
            Next
        Catch ex As Exception
        End Try
    End Sub
    '篩選日期格式
    Private Sub getDate(ByVal m As Match, ByVal int As Integer)
        If int = 1 Then
            mDate = rDate.Match(m.ToString)
            file.WriteLine(mDate)
        ElseIf int = 2 Then
            mDate2 = rDate2.Match(m.ToString)
            file.WriteLine(mDate2)
        End If
    End Sub
    '篩選匯率格式
    Private Sub getRate(ByVal m As Match)
        mRate = rRate.Match(m.ToString)
        file.WriteLine(mRate)
    End Sub
End Class
