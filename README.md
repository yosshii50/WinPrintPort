
WinPrinterListCnv 

    cscript C:\Windows\System32\Printing_Admin_Scripts\ja-JP\prnmngr.vbs -l
    で出力された文字列を表形式に変換する
    項目間は TAB 区切りなので、そのままエクセルに貼り付けて使う

    下記コマンドでクリップボードに貼り付けられる
    cscript C:\Windows\System32\Printing_Admin_Scripts\ja-JP\prnmngr.vbs -l | WinPrinterListCnv.exe | clip


WinPortListCnv

    cscript C:\Windows\System32\Printing_Admin_Scripts\ja-JP\prnport.vbs -l
    で出力された文字列から、ポート作成用スクリプトを生成する

    下記のコマンドでの使用を想定
    cscript C:\Windows\System32\Printing_Admin_Scripts\ja-JP\prnport.vbs -l | WinPortListCnv.exe
    
    以下の場所で実行するとポートが作成される
    cd /d C:\Windows\System32\Printing_Admin_Scripts\ja-JP

コマンドプロンプトは管理者で実行する事
