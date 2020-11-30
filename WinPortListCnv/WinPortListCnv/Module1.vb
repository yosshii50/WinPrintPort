
'実行サンプル1
'   cscript C:\Windows\System32\Printing_Admin_Scripts\ja-JP\prnport.vbs -l | D:\GitHub\WinPortListCnv\WinPortListCnv\bin\Debug\WinPortListCnv.exe | clip

'実行サンプル2
'   cscript C:\Windows\System32\Printing_Admin_Scripts\ja-JP\prnport.vbs -l > ポート一覧.txt
'   WinPortListCnv.exe < ポート一覧.txt | clip

Module Module1

    Structure PortInfo_Str
        Dim PortName As String
        Dim Address As String
        Dim Protocol As String
        Dim PortNo As String
        Dim Queue As String
        Dim ByteCount As String
        Dim SNMP As String
        Dim Community As String
        Dim DeviceIndex As String
    End Structure

    Sub Main()

        Dim IsListOut As Boolean = False

        For Each WrkCommandLine In System.Environment.GetCommandLineArgs()

            WrkCommandLine = WrkCommandLine.Replace("-", "/")
            WrkCommandLine = WrkCommandLine.ToUpper()

            If WrkCommandLine = "/L" Then
                IsListOut = True
            End If
        Next

        Dim PortInfoList = New List(Of PortInfo_Str)()

        Dim WrkPortInfo As New PortInfo_Str
        While True
            Dim WrkLine As String
            Dim WrkStr As String

            WrkLine = Console.ReadLine()
            If WrkLine Is Nothing Then
                If WrkPortInfo.PortName <> "" Then
                    PortInfoList.Add(WrkPortInfo)
                End If
                Exit While
            End If

            WrkStr = GetInfoMsg(WrkLine, "ポート名 ")
            If WrkStr <> "" Then
                If WrkPortInfo.PortName <> "" Then
                    PortInfoList.Add(WrkPortInfo)
                End If
                WrkPortInfo.PortName = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "ホスト アドレス ")
            If WrkStr <> "" Then
                WrkPortInfo.Address = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "プロトコル ")
            If WrkStr <> "" Then
                WrkPortInfo.Protocol = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "ポート番号 ")
            If WrkStr <> "" Then
                WrkPortInfo.PortNo = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "キュー ")
            If WrkStr <> "" Then
                WrkPortInfo.Queue = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "バイト カウント")
            If WrkStr <> "" Then
                WrkPortInfo.ByteCount = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "SNMP ")
            If WrkStr <> "" Then
                WrkPortInfo.SNMP = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "コミュニティ ")
            If WrkStr <> "" Then
                WrkPortInfo.Community = WrkStr
            End If

            WrkStr = GetInfoMsg(WrkLine, "デバイス インデックス ")
            If WrkStr <> "" Then
                WrkPortInfo.DeviceIndex = WrkStr
            End If

        End While

        If IsListOut = True Then

            '一覧表出力
            Dim FirstFLG As Boolean = True
            For Each EachPortInfo In PortInfoList
                If FirstFLG = True Then
                    Console.Write("ポート名" & vbTab)
                    Console.Write("ホストアドレス" & vbTab)
                    Console.Write("プロトコル" & vbTab)
                    Console.Write("ポート番号" & vbTab)
                    Console.Write("キュー" & vbTab)
                    Console.Write("バイトカウント" & vbTab)
                    Console.Write("SNMP" & vbTab)
                    Console.Write("コミュニティ" & vbTab)
                    Console.Write("デバイスインデックス" & vbCrLf)
                    FirstFLG = False
                End If
                Console.Write(EachPortInfo.PortName & vbTab)
                Console.Write(EachPortInfo.Address & vbTab)
                Console.Write(EachPortInfo.Protocol & vbTab)
                Console.Write(EachPortInfo.PortNo & vbTab)
                Console.Write(EachPortInfo.Queue & vbTab)
                Console.Write(EachPortInfo.ByteCount & vbTab)
                Console.Write(EachPortInfo.SNMP & vbTab)
                Console.Write(EachPortInfo.Community & vbTab)
                Console.Write(EachPortInfo.DeviceIndex & vbCrLf)
            Next

        Else

            'Port生成用コマンド
            Console.WriteLine("管理者として cmd 実行")
            Console.WriteLine("cd C:\Windows\System32\Printing_Admin_Scripts\ja-JP")
            For Each EachPortInfo In PortInfoList
                Console.Write("cscript prnport.vbs -a -r " & EachPortInfo.PortName)
                Console.Write(" -h " & EachPortInfo.Address)
                If EachPortInfo.Protocol = "RAW" Then
                    'RAW
                    Console.Write(" -o raw -n " & EachPortInfo.PortNo)
                Else
                    'LPR
                    Console.Write(" -o lpr -q " & EachPortInfo.Queue)
                    If EachPortInfo.ByteCount = "有効" Then
                        '有効
                        Console.Write(" -2e")
                    Else
                        '無効
                        Console.Write(" -2d")
                    End If
                End If
                If EachPortInfo.SNMP = "無効" Then
                    '無効
                    Console.Write(" -md")
                Else
                    '有効
                    Console.Write(" -me -y " & EachPortInfo.Community & " -i " & EachPortInfo.DeviceIndex)
                End If
                Console.Write(vbCrLf)
            Next

        End If

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
