﻿Imports System.Web
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
    Dim patternDate As String = "\d{4}/\d{1,2}/\d{1,2}"
    Dim patternDate2 As String = "\d{1,2}/\d{1,2}"
    Dim patternRate As String = "\d{1,3}.\d{1,7}"
    Dim dbCheck As Boolean

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
        GC.Collect()
    End Sub
    '選取output的button
    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
    End Sub
    '讀取網站資料
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        readFile(ListBox1, "C:\Users\品閎\Desktop\web")
    End Sub
    '抓取名字
    Private Sub getWeb(ByVal str As String)
        Try
            curItem = str
        Catch ex As Exception
            MsgBox("請選擇要開始的網址")
        End Try
    End Sub
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
            ' 定義RE
            Dim rDate As Regex = New Regex(patternDate, RegexOptions.IgnoreCase)
            Dim rDate2 As Regex = New Regex(patternDate2, RegexOptions.IgnoreCase)
            Dim rRate As Regex = New Regex(patternRate, RegexOptions.IgnoreCase)
            ''定義判斷參數
            Dim fir = sr.ReadLine, sec = sr.ReadLine
            '讀取網頁至結尾
            Dim mDate, mDate2, mRate As Match
            Do
                '讓他隔行讀取，並非單雙行讀取
                mDate = rDate.Match(fir)
                mDate2 = rDate2.Match(fir)
                mRate = rRate.Match(sec)
                '判斷日期格式
                If mDate.Success And mRate.Success Then
                    typeDate1(mDate, mRate)
                ElseIf mDate2.Success And mRate.Success Then
                    typeDate2(mDate2, mRate)
                End If
                ''判斷日期格式
                'typeDate1(mDate, mRate)
                'typeDate2(mDate2, mRate)
                '抓取下一個符合格式的
                mDate = mDate.NextMatch()
                mDate2 = mDate2.NextMatch()
                fir = sec
                sec = sr.ReadLine
            Loop Until sr.EndOfStream
        Catch ex As Exception
        End Try
    End Sub

    Private Sub typeDate1(ByVal mDate As Match, ByVal mRate As Match)
        If mDate.Success And mRate.Success Then
            file.WriteLine(mDate)
            file.WriteLine(mRate)
        End If
    End Sub
    Private Sub typeDate2(ByVal mDate As Match, ByVal mRate As Match)
        If mDate.Success And mRate.Success Then
            file.WriteLine(mDate)
            file.WriteLine(mRate)
        End If
    End Sub
    Private Function checkRate(ByVal sr As StreamReader,dbCheck As Boolean)
        Dim rRate As Regex = New Regex(patternRate, RegexOptions.IgnoreCase)
        Dim mRate As Match
        mRate = rRate.Match(sr.ReadLine)
        If dbCheck Then
            Return rRate
        Else
            Return mRate.Success
        End If
    End Function
End Class
