'===============================================================
' modStakeScraper   PedigreeQuery 1?3着馬リスト収集
'===============================================================
Option Explicit

Sub ScrapePedigreeQuery()

'    Const LIST_URL As String = _
'        "http://www.pedigreequery.com/index.php?query_type=stakes&search_bar=stakes&field=name&h="

'    Const LIST_URL As String = _
'        "http://www.pedigreequery.com/index.php?query_type=stakes&search_bar=stakes&field=name&h=japan+cup"
        
        
    Const BASE_URL As String = "https://www.pedigreequery.com"
    
        '―― 1. 一覧ページ HTML ――――――――――――――――――――――
    Dim LIST_URL As String
    LIST_URL = ThisWorkbook.Worksheets("MAIN").Range("A1")
    ' ① 一覧ページを取得（ここまで従来どおり）
    Dim docList As Object: Set docList = GetHTMLDoc(LIST_URL)
    
    ' ② body の “生 HTML” を取り出し
    Dim rawHtml As String
    rawHtml = docList.body.innerHTML      ' ← ここにテーブル全部入っている
    
    ' ③ 新しい htmlfile に書き込んで再パース
    Dim docTbl As Object: Set docTbl = CreateObject("htmlfile")
    docTbl.write rawHtml: docTbl.Close

    
    '―― 2. レース URL 収集 ―――――――――――――――――――――――――
    Dim raceDict As Object: Set raceDict = CreateObject("Scripting.Dictionary")
    Dim td As Object, aTag As Object, href$, title$
    
    For Each td In docTbl.getElementsByTagName("td")
    If LCase$(Trim$(td.className)) = "w2" Then          ' 1列目
        If td.getElementsByTagName("a").Length > 0 Then
            Set aTag = td.getElementsByTagName("a")(0)
            
            href = aTag.getAttribute("href")
            If Left$(href, 1) = "/" Then href = BASE_URL & href
            href = Split(href, "#")(0)                  ' #アンカー除去
            
            title = Trim$(Replace(aTag.innerText, vbCrLf, ""))
            If title = "" Then title = "(no title)"
            
            If Not raceDict.Exists(href) Then raceDict.Add href, title
        End If
    End If
    Next td
    
    Debug.Print "レース URL 件数: "; raceDict.count
    
    
    Debug.Print "raw len=" & Len(rawHtml)

    If raceDict.count = 0 Then
        MsgBox "レース URL を取得できませんでした。", vbExclamation
        Exit Sub
    End If
    
        
    '―― 3. 各レース詳細 → 1?3着馬収集 ――――――――――――――――
    Dim horseDict As Object: Set horseDict = CreateObject("Scripting.Dictionary")
    
    Dim raceURL As Variant
    For Each raceURL In raceDict.Keys
        ParseRacePage CStr(raceURL), CStr(raceDict(raceURL)), horseDict
    Next raceURL
    
    '―― 4. 出力 ―――――――――――――――――――――――――――――――
    DumpResult horseDict, BASE_URL
    
    MsgBox "完了！ レース " & raceDict.count & " 件 ／ 馬 " & horseDict.count & " 頭", vbInformation
End Sub


'===============================================================
' レース詳細ページ → 年・着順ごとに馬名取得
'===============================================================
Private Sub ParseRacePage(raceURL As String, raceName As String, horseDict As Object)

    Dim doc As Object: Set doc = GetHTMLDoc(raceURL)
    If doc Is Nothing Then Exit Sub
    
    Dim tr As Object, td As Object
    Dim yr$, win$, sec$, thr$
    
    For Each tr In doc.getElementsByTagName("tr")
        If tr.getElementsByTagName("td").Length >= 4 Then
            Set td = tr.getElementsByTagName("td")
            yr = Trim$(td(0).innerText)
            win = LCase$(Trim$(td(1).innerText))
            sec = LCase$(Trim$(td(10).innerText))
            thr = LCase$(Trim$(td(11).innerText))
            
            If Len(win) > 0 Then AddNote horseDict, Encode(win), yr & " " & raceName & " 1着"
            If Len(sec) > 0 Then AddNote horseDict, Encode(sec), yr & " " & raceName & " 2着"
            If Len(thr) > 0 Then AddNote horseDict, Encode(thr), yr & " " & raceName & " 3着"
        End If
    Next tr
End Sub


'===============================================================
' horseDict へ備考を追記
'===============================================================
Private Sub AddNote(dict As Object, key As String, note As String)
    Dim arr As Variant
    If dict.Exists(key) Then
        arr = dict(key)
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr)) = note
        dict(key) = arr
    Else
        dict.Add key, Array(note)
    End If
End Sub


'===============================================================
' Result シートへ書き出し
'===============================================================
Private Sub DumpResult(hDict As Object, baseURL As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Result")
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add: ws.Name = "Result"
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.[A1:C1].Value = Array("horse", "profile_url", "notes")
    
    Dim r As Long: r = 2
    Dim k As Variant
    Dim count As Long: count = 1
    For Each k In hDict.Keys
        If count > 3 Then
            ws.Cells(r, 1).Value = k
            ws.Cells(r, 2).Value = baseURL & "/" & k
            ws.Cells(r, 3).Value = Join(hDict(k), ", ")
            r = r + 1
        End If
        count = count + 1
    Next k
End Sub


'===============================================================
' 馬名を URL 用キーへエンコード
'===============================================================
Private Function Encode(s As String) As String
    s = Replace(LCase$(Trim$(s)), " ", "+")
    s = Replace(s, "'", "")
    Encode = s
End Function


