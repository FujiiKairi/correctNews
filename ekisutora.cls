VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub ekisutoraNEWS()
    
'---コード1｜インターネットに接続してブラウザを開く---
    Dim objIE As InternetExplorer                               '操作するIEを入れるオブジェクトを準備
    Set objIE = CreateObject("InternetExplorer.Application")    '上のオブジェクトにIEを入れる
    objIE.Visible = False                                   '見えるように
    
'---コード2｜インターネットの特定のページを開く---
    objIE.navigate "https://www.sankei.com/politics/newslist/politics-n1.html#12"               'このページを開く
    Call IEWait(objIE)   'IEを待機                              IEWaitを呼ぶ
    Call WaitFor(3) '3秒停止                                    WaitForをよぶ

    Dim anchor As HTMLAnchorElement
    Dim numNews As Integer
    Dim i As Integer
    Dim pageIndex As Integer
'    i = 0
'    'url情報をエクセルに書き込む
'    For Each elmListItem In objIE.document.getElementsByClassName("entry tertitary vertical")
'            Set anchor = elmListItem.getElementsByTagName("a")(0)
'            Worksheets("Sheet2").Cells(i + 2, 1).Value = anchor.href 'url情報
'            i = i + 1
'    Next
   
    For pageIndex = 244 To 251
        objIE.navigate (Cells(pageIndex, 1).Value)
        'Sleep (10000)
        Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
            DoEvents
        Loop
        Worksheets("Sheet2").Cells(pageIndex, 2).Value = objIE.document.getElementById("__r_article_title__").innerHTML '記事タイトル
        Worksheets("Sheet2").Cells(pageIndex, 3).Value = objIE.document.getElementById("__r_publish_date__").innerHTML '記事の日付
        Sleep (1000)
        
        For Each elmListItem In objIE.document.getElementsByClassName("post_content")
            For Each objLink In elmListItem.getElementsByTagName("p")
                Worksheets("Sheet2").Cells(pageIndex, 4).Value = Worksheets("Sheet2").Cells(pageIndex, 4).Value + objLink.innerHTML
            Next
        Next
        
        'Sleep (10000)
    Next
        
        
'    For pageIndex = 101 To 150
'        objIE.navigate (Cells(pageIndex, 1).Value)
'        Sleep 5000
'        'ページのコード番号の取得
'        tmp = Split(Cells(pageIndex, 1).Value, "/")
'        Worksheets("imagedata").Cells(pageIndex, 3).Value = tmp(4)
'        'フォルダ作成
'        FolderName = "C:\Users\mnuus\Downloads\二次ブログ\画像\" & Worksheets("imagedata").Cells(pageIndex, 3).Value
'        If Dir(FolderName, vbDirectory) = "" Then   '同名のフォルダがない場合フォルダを作成'
'            MkDir FolderName
'        End If
'        'ここからタグの取得
'        tagString = objIE.document.getElementsByClassName("tag-block")(0).getElementsByClassName("ga-event")(0).innerHTML
'        tmp = Split(tagString, "<")
'        tagString = tmp(0)
'        numTag = objIE.document.getElementsByClassName("tag-block")(0).getElementsByClassName("ga-event").Length
'        For i = 1 To numTag - 1
'            tmp = Split(objIE.document.getElementsByClassName("tag-block")(0).getElementsByClassName("ga-event")(i).innerHTML, "<")
'            tagString = tagString & "," & tmp(0)
'        Next
'        Worksheets("imagedata").Cells(pageIndex, 4).Value = tagString
'        'ここから画像保存
'        For Each objLink In objIE.document.getElementsByClassName("gallery-block")
'            numPicture = objLink.getElementsByTagName("a").Length - 1
'        Next
'            'urlを保存する
'        For numPicIndex = 1 To numPicture
'            For Each objLink In objIE.document.getElementsByClassName("gallery-block")
'                Set anchor = objLink.getElementsByTagName("a")(numPicIndex - 1)
'                Worksheets("imagedata").Cells(pageIndex, 5 + numPicIndex - 1).Value = anchor.href 'url情報
'
'            Next
'        Next
'        For numPicIndex = 0 To numPicture - 1
'              imgURL = Cells(pageIndex, 5 + numPicIndex).Value
'              '画像ファイル名
'              Filename = Mid(imgURL, InStrRev(imgURL, "/") + 1)
'
'              '画像保存先(+画像ファイル名）
'              savePath = FolderName & "\" & numPicIndex & ".jpg"
'
'              'キャッシュクリア
'              cacheDel = DeleteUrlCacheEntry(imgURL)
'
'              '画像ダウンロード
'              result = URLDownloadToFile(0, imgURL, savePath, 0, 0)
'        Next
'
'    Next
'---コード6｜IEを閉じる---
    MsgBox ("終わりました")
    objIE.Quit
    Set objIE = Nothing
ErrLabel:
    'msg = msg & "エラーが発生しました"
    Resume Next
End Sub

Function OpenPage(ByVal url As String, ByRef objIE As Object)
    objIE.navigate (url)               'このページを開く
    Call WaitFor(3) '3秒停止                                    WaitForをよぶ
End Function

'---コード2-1｜IEを待機する関数---
Function IEWait(ByRef objIE As Object)                      'オブジェクトを参照渡し
    Do While objIE.Busy = True Or objIE.readyState <> 4     'busyプロパティがtrueもしくはreadystateが4（IEオブジェクトの全データ読み込み完了状態）
        DoEvents
    Loop
End Function

'---コード2-2｜指定した秒だけ停止する関数---
Function WaitFor(ByVal second As Integer)
    Dim futureTime As Date
 
    futureTime = DateAdd("s", second, Now)                  'functimeを今の時間+second(ここでは3秒)にする
 
    While Now < futureTime
        DoEvents                                            'キャンセルボタンなどのイベントが起こった時にその処理をOSにわたす
    Wend                                                    'while do loop と同じ
End Function













