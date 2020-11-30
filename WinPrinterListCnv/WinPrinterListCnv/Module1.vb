
'実行サンプル1
'   cscript C:\Windows\System32\Printing_Admin_Scripts\ja-JP\prnmngr.vbs -l | D:\GitHub\WinPrinterListCnv\WinPrinterListCnv\bin\Debug\WinPrinterListCnv.exe | clip

'実行サンプル2
'   cscript C:\Windows\System32\Printing_Admin_Scripts\ja-JP\prnmngr.vbs -l > プリンタ一覧.txt
'   WinPrinterListCnv.exe < プリンタ一覧.txt | clip

Module Module1

    Structure PrintInfo_Str
        Dim PrinterName As String
        Dim ShareName As String
        Dim DriverName As String
        Dim PortName As String
        Dim Comment As String
        Dim Location As String
    End Structure

    Sub Main()

        Dim PrintInfoList = New List(Of PrintInfo_Str)()

        Dim WrkPrintInfo As New PrintInfo_Str
        While True
            Dim WrkLine As String
            Dim WrkStr As String

            WrkLine = Console.ReadLine()
            If WrkLine Is Nothing Then
                If WrkPrintInfo.PrinterName <> "" Then
                    PrintInfoList.Add(WrkPrintInfo)
                End If
                Exit While
            End If

            WrkStr = GetInfoMsg(WrkLine, "プリンター名 ")
            If WrkStr <> "" Then
                If WrkPrintInfo.PrinterName <> "" Then
                    PrintInfoList.Add(WrkPrintInfo)
                End If
                WrkPrintInfo.PrinterName = WrkStr
                'Console.WriteLine(WrkStr)
            End If

            WrkStr = GetInfoMsg(WrkLine, "共有名 ")
            If WrkStr <> "" Then
                WrkPrintInfo.ShareName = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "ドライバー名 ")
            If WrkStr <> "" Then
                WrkPrintInfo.DriverName = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "ポート名 ")
            If WrkStr <> "" Then
                WrkPrintInfo.PortName = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "コメント ")
            If WrkStr <> "" Then
                WrkPrintInfo.Comment = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "場所 ")
            If WrkStr <> "" Then
                WrkPrintInfo.Location = WrkStr
            End If

        End While

        Dim FirstFLG As Boolean = True
        For Each EachPrintInfo In PrintInfoList
            If FirstFLG = True Then
                Console.Write("プリンター名" & vbTab)
                Console.Write("共有名" & vbTab)
                Console.Write("ドライバー名" & vbTab)
                Console.Write("ポート名" & vbTab)
                Console.Write("コメント" & vbTab)
                Console.Write("場所" & vbCrLf)
                FirstFLG = False
            End If
            Console.Write(EachPrintInfo.PrinterName & vbTab)
            Console.Write(EachPrintInfo.ShareName & vbTab)
            Console.Write(EachPrintInfo.DriverName & vbTab)
            Console.Write(EachPrintInfo.PortName & vbTab)
            Console.Write(EachPrintInfo.Comment & vbTab)
            Console.Write(EachPrintInfo.Location & vbCrLf)
        Next

        ''デバッグの場合、画面が消えるので対策
        'Console.WriteLine(vbCrLf + vbCrLf + "続行するには Ctrl+C を押してください．．．")
        'While True
        '    Console.Write("")
        'End While

    End Sub

    Private Function GetInfoMsg(ByVal BaseStr As String, ByVal CheckStr As String) As String

        Dim StrPos = BaseStr.IndexOf(CheckStr)

        If StrPos >= 0 Then
            StrPos = StrPos + CheckStr.Length
            Return BaseStr.Substring(StrPos)
        End If

        Return ""
    End Function

End Module
