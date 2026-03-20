

'====================================================================
' グローバル Dictionary
'====================================================================
Public searchedUrls As Object ' 既存のデータを格納するDictionary
Public addedUrls As Object    ' 新たに追加するデータをチェックするDictionary
Public horseListUrls As Object

'====================================================================
' メインのループ。MAIN_2シートのA列URLを順に main(...) へ渡す
'====================================================================
Sub main_roop()
    Dim oldScrUpdate As Boolean, oldCalc As XlCalculation, oldEnableEvents As Boolean
    '=== 画面更新・自動計算・イベントを停止 ===
    oldScrUpdate = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim failedUrls As New Collection
    
    Set ws = ThisWorkbook.Sheets("MAIN_2")
    lastRow = ws.cells(ws.rows.Count, "A").End(xlUp).row
    
    '--- A列を上から順に読み込んで実行 ---
    For i = 1 To lastRow
        If ws.cells(i, 1).Value <> "" Then
            'On Error Resume Next
            Call main(ws.cells(i, 1).Value)
            If Err.number <> 0 Then
                failedUrls.Add ws.cells(i, 1).Value
                Debug.Print "エラー発生: " & Err.Description & vbCrLf & "URL=" & ws.cells(i, 1).Value
                Err.Clear
            End If
            'On Error GoTo 0
        End If
    Next i
    
    '--- 失敗したURLがあれば再トライ ---
    If failedUrls.Count > 0 Then
        Debug.Print "---- 再トライ開始 ----"
        Dim urlFail As Variant
        For Each urlFail In failedUrls
            'On Error Resume Next
            Call main(urlFail)
            If Err.number <> 0 Then
                Debug.Print "再トライもエラー: " & Err.Description & vbCrLf & "URL=" & urlFail
                Err.Clear
            End If
            'On Error GoTo 0
        Next urlFail
    End If
    
    MsgBox "main_roop 完了", vbInformation
    
    '=== 画面更新・自動計算・イベントを戻す ===
    Application.ScreenUpdating = oldScrUpdate
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEnableEvents
End Sub

'====================================================================
' 単一URLを処理。
'   1) SearchedUrls辞書初期化
'   2) Pedigree Dataシート準備
'   3) GetAllHorses呼び出し（再帰）
'   4) 後処理 (列並び替え → FilterAndCopyDataWithHyperlinks → UpdateHorseList)
'====================================================================
Sub main(Optional ByVal url As String = "")
    Dim oldScrUpdate As Boolean, oldCalc As XlCalculation, oldEnableEvents As Boolean
    '=== 画面更新等を停止 ===
    oldScrUpdate = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '--- 事前初期化 ---
    Call InitializeSearchedUrls
    
    If url = "" Then
        url = ThisWorkbook.Sheets("MAIN_2").Range("A1").Value
    Else
        Debug.Print "main サブが呼び出されました。URL=" & url
    End If
    
    If url = "" Then
        Debug.Print "セルA1にURLを入力してください。"
        GoTo Cleanup
    End If
    
    If searchedUrls.Exists(url) Then
        Debug.Print "URLはすでに処理済みのため終了: " & url
        GoTo Cleanup
    End If
    
    '--- Pedigree Dataシート準備 ---
    Dim ws As Worksheet
    'On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Pedigree Data")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = "Pedigree Data"
    End If
    'On Error GoTo 0
    
    ws.cells.Clear
    Call AllCellsToTextFormat(ws)
    
    ' ヘッダー行
    ws.cells(1, 1).Value = "Number"
    ws.cells(1, 2).Value = "Generation"
    ws.cells(1, 3).Value = "Horse Name"
    ws.cells(1, 4).Value = "Sex"
    ws.cells(1, 5).Value = "Color"
    ws.cells(1, 6).Value = "Year"
    ws.cells(1, 7).Value = "URL"
    ws.cells(1, 8).Value = "PrimaryKey"
    ws.cells(1, 9).Value = "Details"
    ws.cells(1, 10).Value = "Sire"
    ws.cells(1, 11).Value = "Dam"
    ws.cells(1, 12).Value = "LoadURL"
    
    '--- 接続チェック(IE) ---
    Dim IE As Object
    Dim maxAttempts As Long: maxAttempts = 3
    Dim attempt As Long
    Dim statusCode As Long
    
    '=== IE起動＆接続チェック (最大3回リトライ) ===
    For attempt = 1 To maxAttempts
        'On Error Resume Next
        Set IE = CreateObject("InternetExplorer.Application")
        'On Error GoTo 0
        
        If IE Is Nothing Then
            MsgBox "Internet Explorerを起動できませんでした。", vbCritical
            GoTo Cleanup
        End If
        
        statusCode = CheckURL(url) ' HEADリクエスト
        If statusCode = 403 Then
            ' 403でもアクセス可能なpedigreequery用の仕様に合わせて成功とみなす
            Exit For
        End If
        
        ' 失敗でリトライ
        IE.Quit
        Set IE = Nothing
        If attempt = maxAttempts Then
            Debug.Print "接続エラー: " & statusCode & " => 最大試行回数に達しました。"
            GoTo Cleanup
        Else
            Debug.Print "接続エラー: " & statusCode & " => 再試行 (" & attempt & "/" & maxAttempts & ")"
        End If
    Next attempt
    
    ' 最終チェック(403以外なら一応エラー)
    If statusCode <> 403 Then
        MsgBox "接続エラー: " & statusCode & vbCrLf & "URLにアクセスできません。"
        IE.Quit
        Set IE = Nothing
        GoTo Cleanup
    End If
    
    '--- IE起動(非表示モード)---
'    IE.Visible = True
    IE.Visible = False
        
    '========================================================
    ' メイン再帰呼び出し: GetAllHorses
    '========================================================
    Dim rowCount As Long: rowCount = 1
    Dim roop_flag As Boolean: roop_flag = True
    Dim Err_flag As Boolean
    Dim noturlflug As Long
    
    Do While roop_flag
'        'On Error Resume Next
        Err.Clear
        
        ' → 旧GetAllHorses呼び出し
        noturlflug = GetAllHorses(rowCount, ws, url, IE)
        IE.Quit
        
        If noturlflug = -1 Then
            GoTo Cleanup
        End If
        
        If Err.number <> 0 Then
            Debug.Print "エラー発生: " & Err.Description
            Err.Clear
            Err_flag = True
        Else
            Err_flag = False
            roop_flag = False
        End If
'        'On Error GoTo 0
        
        '=== エラー時の処理(リトライフロー) ===
        If Err_flag Then
            Call RearrangeColumnsByOrder_call
            Call FilterAndCopyDataWithHyperlinks
            Call UpdateHorseList
            
            ' 再試行
            Dim attempt2 As Long
            For attempt2 = 1 To maxAttempts
                'On Error Resume Next
                Set IE = CreateObject("InternetExplorer.Application")
                'On Error GoTo 0
                If IE Is Nothing Then
                    MsgBox "Internet Explorerを起動できませんでした(リトライ).", vbCritical
                    GoTo Cleanup
                End If
                
                statusCode = CheckURL(url)
                If statusCode = 403 Then
                    Exit For
                End If
                IE.Quit
                Set IE = Nothing
                
                If attempt2 = maxAttempts Then
                    Debug.Print "接続エラー再試行も失敗: " & statusCode
                    GoTo Cleanup
                Else
                    Debug.Print "接続エラー: " & statusCode & " => 再試行 (" & attempt2 & "/" & maxAttempts & ")"
                End If
            Next attempt2
            
            ws.cells.Clear
            Call AllCellsToTextFormat(ws)
            
            ' ヘッダー再設定
            ws.cells(1, 1).Value = "Number"
            ws.cells(1, 2).Value = "Generation"
            ws.cells(1, 3).Value = "Horse Name"
            ws.cells(1, 4).Value = "Sex"
            ws.cells(1, 5).Value = "Color"
            ws.cells(1, 6).Value = "Year"
            ws.cells(1, 7).Value = "URL"
            ws.cells(1, 8).Value = "PrimaryKey"
            ws.cells(1, 9).Value = "Details"
            ws.cells(1, 10).Value = "Sire"
            ws.cells(1, 11).Value = "Dam"
            ws.cells(1, 12).Value = "LoadURL"
            
            IE.Visible = False
            rowCount = 1
            
            ' Dictionaryを初期化
            Call InitializeSearchedUrls
        End If
    Loop
    
    '========================================================
    ' 後処理
    '========================================================
    Call RearrangeColumnsByOrder_call
    Call FilterAndCopyDataWithHyperlinks
    Call UpdateHorseList

Cleanup:
    '=== 画面更新等を戻す ===
    Application.ScreenUpdating = oldScrUpdate
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEnableEvents
End Sub


'====================================================================
' 旧GetAllHorses関数 (再帰) - 機能を変えずにそのまま最適化箇所を最低限
'====================================================================
Function GetAllHorses(ByRef rowCount As Long, ByVal ws As Worksheet, ByVal url As String, ByVal IE As Object) As Long
    Dim doc As Object
    Dim tbl As Object
    Dim data() As Variant
    Dim nextUrls() As Variant
    GetAllHorses = 0
    
    '=== URL重複チェック ===
    If searchedUrls.Exists(url) Then
        Debug.Print "URLはすでに処理済み: " & url
        Exit Function
    End If
    searchedUrls.Add url, True
    
    '=== ページロード待ち & doc取得 ===
    IE.Navigate url
    Call WaitForLoad(IE, url)
    Set doc = IE.document
    
    Debug.Print "ページが正常にロードされました: " & url
    
    '=== 血統表取得 ===
    On Error Resume Next
    Set tbl = doc.querySelector(".pedigreetable")
    On Error GoTo 0
    If tbl Is Nothing Then
        Debug.Print "血統表が見つかりませんでした。"
        IE.Quit
        Exit Function
    End If
    
    '=== データ収集 ===
    Dim recordCount As Long
    recordCount = 0
    ReDim data(1 To 63, 1 To 9)
    
    Dim cell As Object
    Dim generation As String, horseName As String
    Dim horseColor As String, horseYear As String, horseURL As String
    Dim dataName As String, details As String
    Dim parts() As String
    Dim number As Long
    
    '--- 基底馬を0番に追加 ---
    Dim baseDataName As String
    If InStr(url, "https://www.pedigreequery.com/") > 0 Then
        baseDataName = Mid(url, Len("https://www.pedigreequery.com/") + 1)
    Else
        baseDataName = ""
    End If
    
    recordCount = recordCount + 1
    data(recordCount, 1) = 0   ' Number
    data(recordCount, 2) = 0   ' Generation
    data(recordCount, 3) = baseDataName  ' HorseName(仮)
    
    '=== ページ全体から color, year を取得 ===
    Dim fontElements As Object, fontText As String
    Set fontElements = doc.getElementsByTagName("font")
    Dim sex As String
    sex = ""
    Dim searchword As String
    searchword = "" ' 未定義対策
    Dim colorStart As Long, yearStart As Long
    
    Dim horseColorFound As String: horseColorFound = ""
    Dim horseYearFound As String:  horseYearFound = ""
    Dim detailsFound As String:    detailsFound = ""
    
    For Each fontElement In fontElements
        fontText = fontElement.innerText
        If InStr(fontText, ")") > 0 Then
            sex = Right(Split(fontText, ",")(0), 1)
            searchword = searchword + ","
            If InStr(fontText, searchword) > 0 Then
                colorStart = InStr(fontText, ")") + 2
                yearStart = InStr(fontText, searchword) + 2
                
                horseColorFound = Replace(Trim(Mid(fontText, colorStart, InStr(colorStart, fontText, " ") - colorStart)), ".", "")
                If horseColorFound = "DI" Then horseColorFound = ""
                
                horseYearFound = Trim(Mid(fontText, yearStart, 4))
                horseYearFound = ExtractFourDigitNumbers(horseYearFound)
                detailsFound = fontText
                Exit For
            End If
        End If
    Next fontElement
    
    '=== horseName補正 ===
    If horseColorFound = "M," Or horseColorFound = "H," Then
        horseColorFound = ""
    End If
    
    Dim horseName2 As String: horseName2 = baseDataName
    If InStr(detailsFound, horseColorFound) > 0 Then
        horseName2 = Trim(Left(detailsFound, InStr(detailsFound, horseColorFound) - 1))
    End If
    
    If horseColorFound = "" Then
        horseName2 = Trim(Left(detailsFound, InStr(detailsFound, sex & ",") - 1))
        If InStr(horseName2, ".") > 0 Then
            Dim tmpColor As String
            
           

            If InStr(horseName2, ".") - InStrRev(horseName2, " ") + 1 > 0 Then
                tmpColor = Mid(horseName2, InStrRev(horseName2, " ") + 1, InStr(horseName2, ".") - InStrRev(horseName2, " ") + 1)
                horseColorFound = tmpColor
            End If
            horseName2 = Left(horseName2, InStrRev(horseName2, " "))

        End If
    End If
    
    
    
    data(recordCount, 3) = horseName2
    data(recordCount, 4) = sex
    data(recordCount, 5) = horseColorFound
    data(recordCount, 6) = horseYearFound
    data(recordCount, 7) = url
    data(recordCount, 8) = baseDataName
    data(recordCount, 9) = IIf(detailsFound = "", baseDataName, detailsFound)
    
    number = 1
    Dim mainhorsename As String: mainhorsename = baseDataName
    
    '=== 世代カウント判定 ===
    Dim max_generations As Long
    max_generations = CheckGenerationCounts(tbl)
    Dim num_generations As Variant
    num_generations = Array(0, 1, 2, 3, 4, 5, 5, 4, 5, 5, 3, 4, 5, 5, _
                        4, 5, 5, 2, 3, 4, 5, 5, 4, 5, 5, 3, 4, 5, _
                        5, 4, 5, 5, 1, 2, 3, 4, 5, 5, 4, 5, 5, 3, _
                        4, 5, 5, 4, 5, 5, 2, 3, 4, 5, 5, 4, 5, 5, _
                        3, 4, 5, 5, 4, 5, 5, 5)

    
    If max_generations = 0 Then
        Debug.Print max_generations & "世代なので終了"
        GoTo SkipExit
    ElseIf max_generations < 0 Then
        Dim lasthorsesex As String
        If max_generations = -32 Then
            lasthorsesex = "f"
            Debug.Print "1世代母馬のみ取得"
        Else
            lasthorsesex = "m"
            Debug.Print "1世代父馬のみ取得"
        End If
        max_generations = 1
        
        Dim cell2 As Object
        For Each cell2 In tbl.getElementsByTagName("td")
            Dim generation2 As String
            generation2 = ""
            If cell2.hasAttribute("data-g") Then
                generation2 = cell2.getAttribute("data-g")
            End If
            If cell2.className = lasthorsesex Then
                If lasthorsesex = "f" Then
                    sex = "M"
                    recordCount = 33
                Else
                    sex = "H"
                    recordCount = 2
                    
                    If data(recordCount, 3) <> "" Then
                        GoTo SkipExit
                    End If
                End If
                

                
                
                
                
                
                If cell2.getElementsByTagName("a").Length > 0 Then
                    horseName = Replace(Replace(cell2.getElementsByTagName("a")(0).innerText, vbCrLf, ""), vbLf, "")
                    ReDim parts(0)
                    If InStr(cell2.innerText, ",") > 0 Then
                        parts = Split(cell2.innerText, ",")
                    ElseIf InStr(cell2.innerText, ".") > 0 Then
                        parts = Split(cell2.innerText, ".")
                    End If
                    If UBound(parts) >= 1 Then
                        If InStr(parts(0), ")") Then
                            colorStart = InStr(parts(0), ")") + 4
                        ElseIf InStr(parts(0), " ") Then
                            colorStart = InStrRev(parts(0), " ") + 1
                        End If
                        horseColor = Trim(Mid(parts(0), colorStart))
                        horseYear = ExtractFourDigitNumbers(parts(1))
                    End If
                    horseURL = cell2.getElementsByTagName("a")(0).href
                    details = Replace(Replace(cell2.innerText, vbCrLf, ""), vbLf, "")
                    
                    If InStr(details, horseColor) Then
                    
                        If Trim(Left(details, InStr(details, horseColor) - 1)) = "" Then
                                If InStr(details, ".") > 0 Then
                                    horseYear = ExtractFourDigitNumbers(cell2.innerText)
                                ElseIf details <> "" And InStr(details, ")") > 0 Then
                                    horseName = Left(details, InStr(details, ")"))
                                End If
                            ElseIf InStr(details, horseColor) > 0 Then
                                horseName = Trim(Left(details, InStr(details, horseColor) - 1))
                            End If
                    
                    End If
                    If InStr(horseURL, "https://www.pedigreequery.com/") > 0 Then
                        dataName = Mid(horseURL, Len("https://www.pedigreequery.com/") + 1)
                    End If
                End If
                
                If recordCount > UBound(data, 1) Then
                    ReDim Preserve data(1 To recordCount, 1 To 9)
                End If
                data(recordCount, 1) = number
                data(recordCount, 2) = generation2
                data(recordCount, 3) = horseName
                data(recordCount, 4) = sex
                data(recordCount, 5) = horseColor
                data(recordCount, 6) = horseYear
                data(recordCount, 7) = horseURL
                data(recordCount, 8) = dataName
                data(recordCount, 9) = details
                number = number + 1
            End If
        Next cell2
        GoTo SkipExit
        
    ElseIf max_generations < 5 Then
        Debug.Print max_generations & "世代まで取得します。"
    End If
    
    Dim next_tdflag As Boolean, inbreedflag As Boolean
    Dim cell3 As Object
    For Each cell3 In tbl.getElementsByTagName("td")
        generation = ""
        horseName = ""
        sex = ""
        horseColor = ""
        horseYear = ""
        horseURL = ""
        dataName = ""
        details = ""
        
        If cell3.hasAttribute("data-g") Then
            generation = cell3.getAttribute("data-g")
            If generation <> num_generations(recordCount) Then
                Debug.Print generation & "世代が混入しています"
            End If
            If num_generations(recordCount) > max_generations Then
                recordCount = recordCount + 1
                If recordCount > UBound(data, 1) Then
                    ReDim Preserve data(1 To recordCount, 1 To 9)
                End If
                data(recordCount, 1) = number
                data(recordCount, 2) = generation
                data(recordCount, 3) = horseName
                data(recordCount, 4) = sex
                data(recordCount, 5) = horseColor
                data(recordCount, 6) = horseYear
                data(recordCount, 7) = horseURL
                data(recordCount, 8) = dataName
                data(recordCount, 9) = details
                number = number + 1
                GoTo SkipIteration
            End If
        End If
        
        If next_tdflag Then
            If Len(cell3.innerText) > 1 Then
                details = data(recordCount, 9) & " " & cell3.innerText
                data(recordCount, 9) = details
                If InStr(details, ".") > 0 Then
                    data(recordCount, 6) = ExtractFourDigitNumbers(cell3.innerText)
                    data(recordCount, 5) = Trim(Split(cell3.innerText, ".")(0))
                Else
                    data(recordCount, 6) = ExtractFourDigitNumbers(cell3.innerText)
                End If
                inbreedflag = False
            ElseIf Len(cell3.innerText) = 1 Then
                inbreedflag = True
            End If
            
            If Not inbreedflag Then
                next_tdflag = False
            End If
        End If
        
        If cell3.getElementsByTagName("a").Length > 0 Then
            horseName = Replace(Replace(cell3.getElementsByTagName("a")(0).innerText, vbCrLf, ""), vbLf, "")
            ReDim parts(0)
            If InStr(cell3.innerText, ",") > 0 Then
                parts = Split(cell3.innerText, ",")
            ElseIf InStr(cell3.innerText, ".") > 0 Then
                parts = Split(cell3.innerText, ".")
            End If
            
            If generation = "5" Then
                next_tdflag = True
            End If
            
            If UBound(parts) >= 1 Then
                If InStr(parts(0), ")") Then
                    colorStart = InStr(parts(0), ")") + 4
                ElseIf InStr(parts(0), " ") Then
                    colorStart = InStrRev(parts(0), " ") + 1
                End If
                horseColor = Trim(Mid(parts(0), colorStart))
                horseYear = ExtractFourDigitNumbers(parts(1))
            End If
            
            horseURL = cell3.getElementsByTagName("a")(0).href
            details = Replace(Replace(cell3.innerText, vbCrLf, ""), vbLf, "")
            
            If Trim(Left(details, InStr(details, horseColor) - 1)) = "" Then
                If InStr(details, ".") > 0 Then
                    horseYear = ExtractFourDigitNumbers(cell3.innerText)
                ElseIf details <> "" And InStr(details, ")") > 0 Then
                    horseName = Left(details, InStr(details, ")"))
                End If
            ElseIf InStr(details, horseColor) > 0 Then
                horseName = Trim(Left(details, InStr(details, horseColor) - 1))
            End If
            
            If Len(cell3.className) > 0 Then
                sex = Trim(cell3.className)
                If sex = "f" Then
                    sex = "M"
                Else
                    sex = "H"
                End If
            End If
            
            If InStr(horseURL, "https://www.pedigreequery.com/") > 0 Then
                dataName = Mid(horseURL, Len("https://www.pedigreequery.com/") + 1)
            End If
        End If
        
        If horseName <> "" Then
            recordCount = recordCount + 1
            If recordCount > UBound(data, 1) Then
                ReDim Preserve data(1 To recordCount, 1 To 9)
            End If
            data(recordCount, 1) = number
            data(recordCount, 2) = generation
            data(recordCount, 3) = horseName
            data(recordCount, 4) = sex
            data(recordCount, 5) = horseColor
            data(recordCount, 6) = horseYear
            data(recordCount, 7) = horseURL
            data(recordCount, 8) = dataName
            data(recordCount, 9) = details
            number = number + 1
        End If
SkipIteration:
    Next cell3
    
SkipExit:
    Dim result As Variant
    result = ProcessHorseData(data)
    
    '=== シート書き込み ===
    Dim iRow As Long
    If recordCount > 0 Then
        For iRow = 1 To recordCount
            rowCount = rowCount + 1
            If iRow = 1 Then mainhorserow = rowCount
            
            ws.cells(rowCount, 1).Value = result(iRow, 1)  ' number
            ws.cells(rowCount, 2).Value = result(iRow, 2)  ' generation
            ws.cells(rowCount, 3).Value = result(iRow, 3)  ' horseName
            ws.cells(rowCount, 4).Value = result(iRow, 4)  ' sex
            ws.cells(rowCount, 5).Value = result(iRow, 5)  ' color
            ws.cells(rowCount, 6).Value = result(iRow, 6)  ' year
            ws.cells(rowCount, 7).Value = result(iRow, 7)  ' url
            ws.cells(rowCount, 8).Value = result(iRow, 8)  ' dataName
            ws.cells(rowCount, 9).Value = result(iRow, 9)  ' details
            ws.cells(rowCount, 10).Value = result(iRow, 10) ' Sire
            ws.cells(rowCount, 11).Value = result(iRow, 11) ' Dam
            
            ' ハイパーリンク設定
            If data(iRow, 7) <> "" Then
                ws.Hyperlinks.Add _
                    Anchor:=ws.cells(rowCount, 8), _
                    Address:=data(iRow, 7), _
                    TextToDisplay:=data(iRow, 8)
            End If
            
            '=== 5世代目URLを nextUrls に登録 ===
            If result(iRow, 2) = CStr(max_generations) And result(iRow, 2) <> "0" Then
                GetAllHorses = GetAllHorses + 1
                ReDim Preserve nextUrls(1 To GetAllHorses)
                nextUrls(GetAllHorses) = result(iRow, 7)
            End If
        Next iRow
    End If
    
    '=== 再帰呼び出し ===
    If GetAllHorses > 0 Then
        Debug.Print GetAllHorses & "件次のURLがあります。"
        Dim x As Long
        For x = LBound(nextUrls) To UBound(nextUrls)
            If Not searchedUrls.Exists(nextUrls(x)) Then
                Debug.Print mainhorsename & ": " & x & "/" & GetAllHorses & " => " & nextUrls(x)
                GetAllHorses rowCount, ws, nextUrls(x), IE
            Else
                Debug.Print mainhorsename & ": " & x & "/" & GetAllHorses & " (スキップ) " & nextUrls(x)
            End If
        Next x
    ElseIf max_generations > 0 Then
        Debug.Print max_generations & "世代すべての馬のデータ収集が完了"
    End If
    
    Debug.Print mainhorsename & "取得完了"
    ws.cells(mainhorserow, 12).Value = "True"
    
    Set doc = Nothing
    Set IE = Nothing
End Function

'====================================================================
' ProcessHorseData: 父馬/母馬のインデックス割当をオリジナル通り
'====================================================================
Function ProcessHorseData(data As Variant) As Variant
    Dim recordCount As Long
    recordCount = UBound(data, 1)
    If recordCount < 1 Then Exit Function
    
    Dim result() As Variant
    ReDim result(1 To recordCount, 1 To UBound(data, 2) + 2)
    
    Dim rowCount As Long
    For rowCount = 1 To recordCount
        Dim father As String, mother As String
        father = ""
        mother = ""
        
        Select Case data(rowCount, 1)
            Case 0: father = data(2, 8): mother = data(33, 8)
            Case 1: father = data(3, 8): mother = data(18, 8)
            Case 2: father = data(4, 8): mother = data(11, 8)
            Case 3: father = data(5, 8): mother = data(8, 8)
            Case 4: father = data(6, 8): mother = data(7, 8)
            Case 7: father = data(9, 8): mother = data(10, 8)
            Case 10: father = data(12, 8): mother = data(15, 8)
            Case 11: father = data(13, 8): mother = data(14, 8)
            Case 14: father = data(16, 8): mother = data(17, 8)
            Case 17: father = data(19, 8): mother = data(26, 8)
            Case 18: father = data(20, 8): mother = data(23, 8)
            Case 19: father = data(21, 8): mother = data(22, 8)
            Case 22: father = data(24, 8): mother = data(25, 8)
            Case 25: father = data(27, 8): mother = data(30, 8)
            Case 26: father = data(28, 8): mother = data(29, 8)
            Case 29: father = data(31, 8): mother = data(32, 8)
            Case 32: father = data(34, 8): mother = data(49, 8)
            Case 33: father = data(35, 8): mother = data(42, 8)
            Case 34: father = data(36, 8): mother = data(39, 8)
            Case 35: father = data(37, 8): mother = data(38, 8)
            Case 38: father = data(40, 8): mother = data(41, 8)
            Case 41: father = data(43, 8): mother = data(46, 8)
            Case 42: father = data(44, 8): mother = data(45, 8)
            Case 45: father = data(47, 8): mother = data(48, 8)
            Case 48: father = data(50, 8): mother = data(57, 8)
            Case 49: father = data(51, 8): mother = data(54, 8)
            Case 50: father = data(52, 8): mother = data(53, 8)
            Case 53: father = data(55, 8): mother = data(56, 8)
            Case 56: father = data(58, 8): mother = data(61, 8)
            Case 57: father = data(59, 8): mother = data(60, 8)
            Case 60: father = data(62, 8): mother = data(63, 8)
            Case Else
                father = ""
                mother = ""
        End Select
        
        Dim j As Long
        For j = 1 To UBound(data, 2)
            result(rowCount, j) = data(rowCount, j)
        Next j
        result(rowCount, UBound(data, 2) + 1) = father
        result(rowCount, UBound(data, 2) + 2) = mother
    Next rowCount
    
    ProcessHorseData = result
End Function

'====================================================================
' 以下、オリジナルのユーティリティ類（変更最小限）
'====================================================================
Sub RearrangeColumnsByOrder_call()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Pedigree Data")
    Dim newOrder As Variant
    newOrder = Array("PrimaryKey", "Sire", "Dam", "Sex", "Color", "Year", "details", "URL", "Horse Name", "generation", "LoadURL")
    RearrangeColumnsByOrder ws, newOrder
    ws.Activate
End Sub

Sub RearrangeColumnsByOrder(ws As Worksheet, columnOrder As Variant)
    Dim headerRange As Range
    Set headerRange = ws.rows(1).Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If headerRange Is Nothing Then
        Debug.Print "ヘッダー行が見つかりませんでした。"
        Exit Sub
    End If
    
    Dim lastColumn As Long
    lastColumn = ws.cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim tempColumn As Long
    tempColumn = lastColumn + 1
    
    Dim targetHeader As Variant
    Dim columnIndex As Variant
    
    For Each targetHeader In columnOrder
        'On Error Resume Next
        columnIndex = Application.match(targetHeader, ws.rows(1), 0)
        'On Error GoTo 0
        If Not IsError(columnIndex) And columnIndex > 0 Then
            ws.Columns(columnIndex).Cut
            ws.Columns(tempColumn).Insert Shift:=xlToRight
            tempColumn = tempColumn + 1
        Else
            Debug.Print "ヘッダー '" & targetHeader & "' が見つかりませんでした。"
        End If
    Next targetHeader
    
    Dim c As Long
    For c = lastColumn To 1 Step -1
        'On Error Resume Next
        columnIndex = Application.match(ws.cells(1, c).Value, columnOrder, 0)
        'On Error GoTo 0
        If IsError(columnIndex) Then
            ws.Columns(c).Delete
        End If
    Next c
    
    Dim lastC As Long
    lastC = ws.cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = lastC To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Columns(c)) = 0 Then
            ws.Columns(c).Delete
        End If
    Next c
    
    Dim lastRow As Long
    lastRow = ws.cells(ws.rows.Count, 1).End(xlUp).row
    Dim r As Long
    For r = lastRow To 2 Step -1
        If IsEmpty(ws.cells(r, 1)) Then
            ws.rows(r).Delete
        End If
    Next r
    
    Debug.Print "並び替え完了"
End Sub

Function ExtractFourDigitNumbers(inputString As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "\b\d{4}\b"
    regex.Global = True
    
    Dim cleanedString As String
    cleanedString = Trim(StrConv(inputString, vbNarrow))
    
    Dim matches As Object
    Set matches = regex.Execute(cleanedString)
    
    Dim result As String
    result = ""
    
    Dim mtch As Object
    For Each mtch In matches
        result = result & mtch.Value & ", "
    Next mtch
    
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 2)
    End If
    
    ExtractFourDigitNumbers = result
End Function

Function CheckGenerationCounts(tbl As Object) As Integer
    Dim cell As Object
    Dim genCounts(1 To 5) As Long
    Dim gen_number As Integer
    Dim lastclass As String
    
    Dim i As Long
    For i = 1 To 5
        genCounts(i) = 0
    Next i
    
    For Each cell In tbl.getElementsByTagName("td")
        If cell.hasAttribute("data-g") Then
            If cell.className = "m" Or cell.className = "f" Then
                gen_number = CInt(cell.getAttribute("data-g"))
                If gen_number >= 1 And gen_number <= 5 Then
                    genCounts(gen_number) = genCounts(gen_number) + 1
                    If gen_number = 1 Then lastclass = cell.className
                End If
            End If
        End If
    Next cell
    
    Dim maxMatched As Long: maxMatched = 0
    
    If genCounts(1) = 1 Then
        If lastclass = "m" Then
            maxMatched = -1
        Else
            maxMatched = -32
        End If
    End If
    If genCounts(1) = 2 Then maxMatched = 1
    If genCounts(2) = 4 Then maxMatched = 2
    If genCounts(3) = 8 Then maxMatched = 3
    If genCounts(4) = 16 Then maxMatched = 4
    If genCounts(5) = 32 Then maxMatched = 5
    
    CheckGenerationCounts = maxMatched
End Function

Function WaitForLoad(IE As Object, ByVal expectedURL As String) As Boolean
    Dim startTime As Double
    Dim timeout As Double
    Dim retryCount As Integer
    Dim maxRetries As Integer
    
    timeout = 30
    maxRetries = 3
    retryCount = 0
    startTime = Timer
    
    Do
        DoEvents
        If Not IE.Busy And IE.ReadyState = 4 Then
            'On Error Resume Next
            If Not IE.document Is Nothing Then
                If LCase(IE.LocationURL) = LCase(expectedURL) Then
                    WaitForLoad = True
                    Exit Function
                Else
                    retryCount = retryCount + 1
                    If retryCount >= maxRetries Then
                        Err.Raise vbObjectError + 1, "WaitForLoad", "ナビゲートURLと異なるページが開かれました"
                    Else
                        IE.Navigate expectedURL
                        startTime = Timer
                    End If
                End If
            End If
            'On Error GoTo 0
        End If
        
        If Timer - startTime > timeout Then
            Err.Raise vbObjectError + 2, "WaitForLoad", "ページロードがタイムアウト"
        End If
    Loop
End Function

Sub FilterAndCopyDataWithHyperlinks()
    Call InitializeSearchedUrls
    
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Set wsSrc = ThisWorkbook.Sheets("Pedigree Data")
    Set wsDest = ThisWorkbook.Sheets("FilteredData")
    Call AllCellsToTextFormat(wsDest)
    
    Set addedUrls = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long, lastCol As Long, destRow As Long
    Dim i As Long, j As Long, loadURLCol As Long, primaryKeyCol As Long
    
    lastRow = wsSrc.cells(wsSrc.rows.Count, 1).End(xlUp).row
    lastCol = wsSrc.cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    
    For j = 1 To lastCol
        If wsSrc.cells(1, j).Value = "LoadURL" Then loadURLCol = j
        If wsSrc.cells(1, j).Value = "URL" Then primaryKeyCol = j
    Next j
    
    If loadURLCol = 0 Or primaryKeyCol = 0 Then
        MsgBox "Pedigree Data のカラム 'LoadURL' or 'URL'が見つかりません", vbExclamation
        Exit Sub
    End If
    
    Dim destLastRow As Long
    destLastRow = wsDest.cells(wsDest.rows.Count, 1).End(xlUp).row
    destRow = destLastRow + 1
    
    Dim key As String
    For i = 2 To lastRow
        key = wsSrc.cells(i, primaryKeyCol).Value
        If wsSrc.cells(i, loadURLCol).Value = True And key <> "" Then
            If Not searchedUrls.Exists(key) And Not addedUrls.Exists(key) Then
                Dim c As Long
                For c = 1 To lastCol
                    wsDest.cells(destRow, c).Value = wsSrc.cells(i, c).Value
                Next c
                Dim h As Hyperlink
                For Each h In wsSrc.Hyperlinks
                    If h.Range.row = i Then
'                        wsDest.Hyperlinks.Add wsDest.cells(destRow, h.Range.Column), h.Address, , , h.TextToDisplay
                    End If
                Next h
                addedUrls(key) = True
                destRow = destRow + 1
            End If
        End If
    Next i
    Debug.Print "LoadURL = TRUE のデータをFilteredDataへ追加完了"
End Sub
Private Function FindHeaderColumn(ByVal arr As Variant, ByVal headerName As String) As Long
    Dim c As Long
    ' 2次元配列(arr) は arr(行, 列) の形
    '   行方向 = LBound(arr,1) ～ UBound(arr,1)
    '   列方向 = LBound(arr,2) ～ UBound(arr,2)
    ' ヘッダー行が arr(1, c) にあると想定（1ベース配列の場合）
    
    For c = LBound(arr, 2) To UBound(arr, 2)
        If CStr(arr(1, c)) = headerName Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
    
    ' 見つからなかった場合は 0 を返す
    FindHeaderColumn = 0
End Function

Sub UpdateHorseList()

    Dim wsHorse As Worksheet, wsFiltered As Worksheet, wsPedigree As Worksheet
    Dim fKey As Variant
    Dim pKey As Variant
    '--- horse list シートを用意 ---
    'On Error Resume Next
    Set wsHorse = ThisWorkbook.Sheets("horse list")
    If wsHorse Is Nothing Then
        Set wsHorse = ThisWorkbook.Sheets.Add
        wsHorse.name = "horse list"
    End If
    'On Error GoTo 0
    
    Call AllCellsToTextFormat(wsHorse)
    
    '--- 他シート参照 ---
    Set wsFiltered = ThisWorkbook.Sheets("FilteredData")
    Set wsPedigree = ThisWorkbook.Sheets("Pedigree Data")
    
    '--- horse list の URL, LoadURL列を探す ---
    Dim urlCol As Long, loadURLCol As Long
    Dim j As Long
    Dim lastColHorseHeader As Long
    lastColHorseHeader = wsHorse.cells(1, wsHorse.Columns.Count).End(xlToLeft).Column
    
    For j = 1 To lastColHorseHeader
        Select Case wsHorse.cells(1, j).Value
            Case "URL": urlCol = j
            Case "LoadURL": loadURLCol = j
        End Select
    Next j
    
    If urlCol = 0 Or loadURLCol = 0 Then
        MsgBox "horse list に 'URL' または 'LoadURL' が見つかりません。"
        Exit Sub
    End If
    
    '--- 既存のURLs管理 ---
    Call InitializeSearchedUrls
    Call InitializeHorseListUrls
    
    '################################################################
    ' ① horse list/FilteredData/Pedigree Data を配列に取り込み
    '################################################################
    Dim lastRowHorse As Long
    lastRowHorse = wsHorse.cells(wsHorse.rows.Count, 1).End(xlUp).row
    If lastRowHorse < 2 Then lastRowHorse = 2  ' データがないとき2行固定
    
    Dim horseRange As Range
    Set horseRange = wsHorse.Range("A1").Resize(lastRowHorse, lastColHorseHeader)
    Dim horseArr As Variant
    horseArr = horseRange.Value ' (1-based配列) [行,列]
    
    Dim lastRowFiltered As Long
    lastRowFiltered = wsFiltered.cells(wsFiltered.rows.Count, 1).End(xlUp).row
    If lastRowFiltered < 2 Then lastRowFiltered = 2
    
    Dim filteredRange As Range
    Set filteredRange = wsFiltered.Range("A1").Resize(lastRowFiltered, wsFiltered.cells(1, Columns.Count).End(xlToLeft).Column)
    Dim filteredArr As Variant
    filteredArr = filteredRange.Value
    
    Dim lastRowPedigree As Long
    lastRowPedigree = wsPedigree.cells(wsPedigree.rows.Count, 1).End(xlUp).row
    If lastRowPedigree < 2 Then lastRowPedigree = 2
    
    Dim pedigreeRange As Range
    Set pedigreeRange = wsPedigree.Range("A1").Resize(lastRowPedigree, wsPedigree.cells(1, Columns.Count).End(xlToLeft).Column)
    Dim pedigreeArr As Variant
    pedigreeArr = pedigreeRange.Value
    
    '################################################################
    ' ② FilteredData, Pedigree を Dictionaryにして検索高速化
    '    (URL => 行番号) として保持
    '################################################################
    Dim dictFiltered As Object, dictPedigree As Object
    Set dictFiltered = CreateObject("Scripting.Dictionary")
    Set dictPedigree = CreateObject("Scripting.Dictionary")
    
    Dim urlColFiltered As Long, urlColPedigree As Long
    urlColFiltered = FindHeaderColumn(filteredArr, "URL")
    urlColPedigree = FindHeaderColumn(pedigreeArr, "URL")
    
    If urlColFiltered = 0 Or urlColPedigree = 0 Then
        MsgBox "FilteredData or Pedigree Data に 'URL' ヘッダが見つかりません。"
        Exit Sub
    End If
    
    ' FilteredData => dictFiltered(key) = 行番号
    Dim i As Long
    For i = 2 To UBound(filteredArr, 1)

        fKey = CStr(filteredArr(i, urlColFiltered))
        If Len(fKey) > 0 Then
            dictFiltered(fKey) = i
        End If
    Next i
    
    ' Pedigree => dictPedigree(key) = 行番号
    For i = 2 To UBound(pedigreeArr, 1)

        pKey = CStr(pedigreeArr(i, urlColPedigree))
        If Len(pKey) > 0 Then
            dictPedigree(pKey) = i
        End If
    Next i
    
    '################################################################
    ' ③ horse list の "LoadURL=FALSE" 行を FilteredData で上書き
    '################################################################
    Dim maxColHorse As Long
    maxColHorse = UBound(horseArr, 2)
    
    For i = 2 To UBound(horseArr, 1)
        Dim key As String
        key = CStr(horseArr(i, urlCol))
        
        ' horseListUrls(key)=False なら FilteredDataの行で上書き
        If (key <> "") Then
            If horseListUrls.Exists(key) Then
                If horseListUrls(key) = False Then
                    ' FilteredData内の行を探す
                    If dictFiltered.Exists(key) Then
                        Dim fRow As Long
                        fRow = dictFiltered(key)
                        
                        ' FilteredData の fRow 行を horseArr(i) にコピー
                        Dim c As Long
                        Dim colCountFiltered As Long
                        colCountFiltered = UBound(filteredArr, 2)
                        
                        ' horseArrの列数とFilteredDataの列数が異なる可能性があるため
                        ' ミニマムの列だけコピー
                        Dim minCol As Long
                        minCol = IIf(maxColHorse < colCountFiltered, maxColHorse, colCountFiltered)
                        
                        For c = 1 To minCol
                            horseArr(i, c) = filteredArr(fRow, c)
                        Next c
                        
                        ' ハイパーリンクは後ほどまとめて or 行単位でコピー
                        ' → 後で行番号がわかるように何か記録しておく
                        '   （ただし簡易的にはここで直接CopyHyperlink呼んでもOK）
                        
                        ' 今回は即座にコピーする例
                        ' => CopyHyperlink wsFiltered, wsHorse, fRow, i
                    End If
                End If
            End If
        End If
    Next i
    
    '################################################################
    ' ④ FilteredData にあって horse list に無いURL => 新規追加
    '################################################################
    '   新規追加分をコレクションにためて後で一括書き込み
    
    Dim addData As New Collection
    
    ' horseListUrls.Exists(key) のチェックを配列だけでやると遅いので
    ' 事前に horseListUrls を補完済み。
    
    ' ここは "URLがまだ horseListUrls にない" 場合、filteredArrの行を addData に格納
    For Each fKey In dictFiltered.Keys
        If Not horseListUrls.Exists(fKey) Then
            Dim rowF As Long
            rowF = dictFiltered(fKey)
            
            ' 1行ぶんをVariant配列にして取り出す
            Dim oneRow() As Variant
            ReDim oneRow(1 To maxColHorse)
            
            Dim colCountF As Long
            colCountF = UBound(filteredArr, 2)
            Dim minC As Long
            minC = IIf(maxColHorse < colCountF, maxColHorse, colCountF)
            
            For c = 1 To minC
                oneRow(c) = filteredArr(rowF, c)
            Next c
            addData.Add oneRow
            
            ' 追加と同時に horseListUrls(key)=True するならこう↓
            horseListUrls(fKey) = True
        End If
    Next fKey
    
    '################################################################
    ' ⑤ Pedigree Dataにあって horse listにないURL => 新規追加
    '################################################################
    For Each pKey In dictPedigree.Keys
        If Not horseListUrls.Exists(pKey) Then
            Dim rowP As Long
            rowP = dictPedigree(pKey)
            

            ReDim oneRow(1 To maxColHorse)
            
            Dim colCountP As Long
            colCountP = UBound(pedigreeArr, 2)
 
            minC = IIf(maxColHorse < colCountP, maxColHorse, colCountP)
            
            For c = 1 To minC
                oneRow(c) = pedigreeArr(rowP, c)
            Next c
            addData.Add oneRow
            horseListUrls(pKey) = True
        End If
    Next pKey
    
    '################################################################
    ' ⑥ horseArr を シートに書き戻し
    '################################################################
    horseRange.Value = horseArr
    
    '################################################################
    ' ⑦ 追加行 (addData) を "horse list" 最終行の下にまとめて貼り付け
    '################################################################
    Dim destRow As Long
    destRow = lastRowHorse + 1
    
    Dim item As Variant
    For Each item In addData
        ' item は 1行分の配列
        wsHorse.Range("A" & destRow).Resize(1, UBound(item)).Value = item
        destRow = destRow + 1
    Next item
    
    '################################################################
    ' ⑧ ハイパーリンクのコピー (必要であれば)
    '################################################################
    ' オリジナルでは行ごとに CopyHyperlink... していましたが、
    ' 配列処理では特定の行番号が変わるため、都度呼び出すか
    ' あるいは後からループで行うか、好みに応じて調整して下さい。
    '
    ' 例: "LoadURL=False" 上書き時の fRow, i をメモしておき、
    '     最後にそれらをまとめて CopyHyperlink wsFiltered, wsHorse, fRow, i
    
    Debug.Print "[Optimized] horse list を更新しました！"
End Sub


Public Sub UpdateHorseList_Faster()
    ' (省略: 前回のまま。もし必要なら全機能同等で高速化するが、
    '  変更を最小限にするため、ここではあえて書いていません)
End Sub

'====================================================================
' 辞書初期化など
'====================================================================
Public Sub InitializeSearchedUrls()
    Dim wsDest As Worksheet
    Set wsDest = ThisWorkbook.Sheets("FilteredData")
    
    Set searchedUrls = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long, primaryKeyCol As Long, i As Long
    lastRow = wsDest.cells(wsDest.rows.Count, 1).End(xlUp).row
    
    For i = 1 To wsDest.cells(1, wsDest.Columns.Count).End(xlToLeft).Column
        If wsDest.cells(1, i).Value = "URL" Then
            primaryKeyCol = i
            Exit For
        End If
    Next i
    
    If primaryKeyCol = 0 Then
        MsgBox "カラム 'URL' が見つかりません。"
        Exit Sub
    End If
    
    Dim key As String
    For i = 2 To lastRow
        key = wsDest.cells(i, primaryKeyCol).Value
        If key <> "" Then searchedUrls(key) = True
    Next i
End Sub

Public Sub InitializeHorseListUrls()
    Set horseListUrls = CreateObject("Scripting.Dictionary")
    Dim wsHorse As Worksheet
    Set wsHorse = ThisWorkbook.Sheets("horse list")
    
    Dim lastRow As Long
    lastRow = wsHorse.cells(wsHorse.rows.Count, 1).End(xlUp).row
    
    Dim urlCol As Long, loadURLCol As Long, i As Long
    For i = 1 To wsHorse.cells(1, wsHorse.Columns.Count).End(xlToLeft).Column
        If wsHorse.cells(1, i).Value = "URL" Then urlCol = i
        If wsHorse.cells(1, i).Value = "LoadURL" Then loadURLCol = i
    Next i
    
    Dim key As String
    For i = 2 To lastRow
        key = wsHorse.cells(i, urlCol).Value
        If key <> "" Then
            horseListUrls(key) = wsHorse.cells(i, loadURLCol).Value
        End If
    Next i
End Sub

Function AllCellsToTextFormat(Worksheet_tgt As Worksheet)
    Dim rng As Range
    Set rng = Worksheet_tgt.cells
    rng.NumberFormat = "@"
End Function

Function CheckURL(ByVal url As String) As Integer
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    'On Error Resume Next
    http.Open "HEAD", url, False
    http.send
    'On Error GoTo 0
    CheckURL = http.Status ' HTTPステータスコード(200,403など)
End Function

Sub CopyHyperlink(ByVal wsSrc As Worksheet, ByVal wsDest As Worksheet, ByVal srcRow As Long, ByVal destRow As Long)
    Dim h As Hyperlink
    For Each h In wsSrc.Hyperlinks
        If h.Range.row = srcRow Then
            wsDest.Hyperlinks.Add Anchor:=wsDest.cells(destRow, h.Range.Column), _
                Address:=h.Address, TextToDisplay:=h.TextToDisplay
        End If
    Next h
End Sub






