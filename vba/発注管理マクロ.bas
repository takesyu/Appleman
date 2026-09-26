Attribute VB_Name = "発注管理マクロ"
Option Explicit

'==============================================================
'  発注シート印刷.xlsm  自動化マクロ
'
'   予約と入荷待ちを振り分け : 「発注割れ」F列の入力を処理し
'                              予約発注リスト／入荷待ちリストへ転記
'                              （発注点割れデータの行は削除しない。
'                                予約中／入荷待ちは色分け＋バーコード
'                                非表示で発注割れシートに残す）
'   予約分を復活             : 予約発注リストの予定日到来分を発注可能に戻す
'   発注点割れデータを取込   : レセコン出力CSVを発注点割れデータへ全入替
'==============================================================

Private Const SH_MAIN   As String = "発注割れ"
Private Const SH_YOYAKU  As String = "予約発注リスト"
Private Const SH_NYUKA   As String = "入荷待ちリスト"
Private Const SH_MASTER  As String = "マスタ"
Private Const SH_DATA    As String = "発注点割れデータ"
Private Const FIRST_ROW  As Long = 3
Private Const LAST_ROW   As Long = 200

'--------------------------------------------------------------
' ① 予約と入荷待ちを振り分け
'    発注点割れデータの行は削除しない（#REF!防止・一覧から消さない）。
'--------------------------------------------------------------
Public Sub 予約と入荷待ちを振り分け()
    Dim wsM As Worksheet, wsY As Worksheet, wsN As Worksheet
    Dim i As Long
    Dim savedCalc As XlCalculation

    On Error GoTo EH
    Set wsM = ThisWorkbook.Worksheets(SH_MAIN)
    Set wsY = ThisWorkbook.Worksheets(SH_YOYAKU)
    Set wsN = ThisWorkbook.Worksheets(SH_NYUKA)

    savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.Calculate

    '--- 1. 入荷待ちリストのクリーンアップ --------------------
    '   今日の「発注割れ」C列(JAN)に無い＝入庫して発注点超え
    Dim todayCodes As Object
    Set todayCodes = CreateObject("Scripting.Dictionary")
    todayCodes.CompareMode = vbTextCompare

    Dim code As String
    For i = FIRST_ROW To LAST_ROW
        code = NormCode(wsM.Cells(i, 3).Value)
        If Len(code) > 0 And code <> "JAN未登録" Then
            If Not todayCodes.Exists(code) Then todayCodes.Add code, True
        End If
    Next i

    Dim nLast As Long
    nLast = LastDataRow(wsN, 1)
    For i = nLast To 2 Step -1
        code = NormCode(wsN.Cells(i, 1).Value)
        If Len(code) > 0 Then
            If Not todayCodes.Exists(code) Then wsN.Rows(i).Delete
        End If
    Next i

    '--- 2. 既存の入荷待ちコード集合を作成 -------------------
    Dim nyukaSet As Object
    Set nyukaSet = CreateObject("Scripting.Dictionary")
    nyukaSet.CompareMode = vbTextCompare
    nLast = LastDataRow(wsN, 1)
    For i = 2 To nLast
        code = NormCode(wsN.Cells(i, 1).Value)
        If Len(code) > 0 And Not nyukaSet.Exists(code) Then nyukaSet.Add code, True
    Next i

    '--- 3. F列を走査して振り分け（発注点割れデータは削除しない） ---
    Dim moved As Long
    Dim kbn As String, nm As String, rawCode As String
    Dim qty As Variant, gd As Variant
    Dim ry As Long, rn As Long

    For i = FIRST_ROW To LAST_ROW
        kbn = Trim$(CStr(wsM.Cells(i, 6).Value))
        If kbn = "予約" Or kbn = "入荷待ち" Then
            nm = CStr(wsM.Cells(i, 1).Value)
            rawCode = CStr(wsM.Cells(i, 3).Value)
            qty = wsM.Cells(i, 2).Value
            gd = wsM.Cells(i, 7).Value

            If Len(Trim$(nm)) > 0 Then
                If kbn = "予約" Then
                    ry = LastDataRow(wsY, 1) + 1
                    If ry < 2 Then ry = 2
                    wsY.Cells(ry, 1).Value = rawCode
                    wsY.Cells(ry, 2).Value = nm
                    wsY.Cells(ry, 3).Value = qty
                    If IsDate(gd) Then
                        wsY.Cells(ry, 4).Value = CDate(gd)
                    Else
                        wsY.Cells(ry, 4).Value = gd
                    End If
                    wsY.Cells(ry, 5).Value = "予約中"
                    moved = moved + 1

                Else    ' 入荷待ち
                    If Not nyukaSet.Exists(NormCode(rawCode)) Then
                        rn = LastDataRow(wsN, 1) + 1
                        If rn < 2 Then rn = 2
                        wsN.Cells(rn, 1).Value = rawCode
                        wsN.Cells(rn, 2).Value = nm
                        wsN.Cells(rn, 3).Value = qty
                        wsN.Cells(rn, 4).Value = Date
                        wsN.Cells(rn, 5).Value = "入荷待ち"
                        If Len(NormCode(rawCode)) > 0 Then nyukaSet.Add NormCode(rawCode), True
                    End If
                    moved = moved + 1
                End If
            End If
        End If
    Next i

    '--- 4. F列・G列をクリア（判定入力をリセット） --------------
    wsM.Range(wsM.Cells(FIRST_ROW, 6), wsM.Cells(LAST_ROW, 7)).ClearContents

    Application.Calculation = savedCalc
    Application.CalculateFull
    Application.ScreenUpdating = True
    MsgBox moved & " 件を振り分けました。" & vbCrLf & _
           "（予約中＝水色 / 入荷待ち＝ピンク で一覧に残ります。バーコードは非表示になります）", _
           vbInformation, "予約・入荷待ちの振り分け"
    Exit Sub
EH:
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました：" & vbCrLf & Err.Description, vbExclamation, "予約・入荷待ちの振り分け"
End Sub

'--------------------------------------------------------------
' ② 予約分を復活（予定日が到来した予約を発注可能に戻す）
'    ・発注点割れデータに既に在る品目 → 予約発注リストから外すだけ
'      （＝バーコードが再表示され、色分けも解除される）
'    ・データに無い品目（過去の取込で除外済み）→ データ末尾へ追加
'--------------------------------------------------------------
Public Sub 予約分を復活()
    Dim wsY As Worksheet, wsD As Worksheet, wsMst As Worksheet
    Dim i As Long
    Dim savedCalc As XlCalculation

    On Error GoTo EH
    Set wsY = ThisWorkbook.Worksheets(SH_YOYAKU)
    Set wsD = ThisWorkbook.Worksheets(SH_DATA)
    Set wsMst = ThisWorkbook.Worksheets(SH_MASTER)

    savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' マスタ JAN(トリム) -> Array(YJコード, 入数)
    Dim mst As Object
    Set mst = CreateObject("Scripting.Dictionary")
    mst.CompareMode = vbTextCompare
    Dim mLast As Long, rr As Long, jc As String
    mLast = LastDataRow(wsMst, 3)
    For rr = 2 To mLast
        jc = Trim$(CStr(wsMst.Cells(rr, 3).Value))
        If Len(jc) > 0 And Not mst.Exists(jc) Then _
            mst.Add jc, Array(wsMst.Cells(rr, 2).Value, wsMst.Cells(rr, 4).Value)
    Next rr

    ' 発注点割れデータに現在ある YJ／薬品名 の集合
    Dim dYj As Object, dNm As Object
    Set dYj = CreateObject("Scripting.Dictionary"): dYj.CompareMode = vbTextCompare
    Set dNm = CreateObject("Scripting.Dictionary"): dNm.CompareMode = vbTextCompare
    Dim dLast As Long
    dLast = LastDataRow(wsD, 3)
    For rr = 2 To dLast
        Dim k1 As String, k2 As String
        k1 = Trim$(CStr(wsD.Cells(rr, 16).Value))
        k2 = Trim$(CStr(wsD.Cells(rr, 3).Value))
        If Len(k1) > 0 And Not dYj.Exists(k1) Then dYj.Add k1, True
        If Len(k2) > 0 And Not dNm.Exists(k2) Then dNm.Add k2, True
    Next rr

    Dim lastY As Long, restoredAdd As Long, restoredFree As Long
    lastY = LastDataRow(wsY, 1)

    Dim d As Variant, rawCode As String, nm As String, qty As Variant
    Dim dr As Long, yj As Variant, nyusu As Double

    For i = lastY To 2 Step -1
        d = wsY.Cells(i, 4).Value
        If IsDate(d) Then
            If CDate(d) <= Date Then
                rawCode = Trim$(CStr(wsY.Cells(i, 1).Value))
                nm = CStr(wsY.Cells(i, 2).Value)
                qty = wsY.Cells(i, 3).Value
                If Not IsNumeric(qty) Then qty = 0

                yj = ""
                nyusu = 1
                If mst.Exists(rawCode) Then
                    yj = mst(rawCode)(0)
                    If IsNumeric(mst(rawCode)(1)) Then
                        If CDbl(mst(rawCode)(1)) > 0 Then nyusu = CDbl(mst(rawCode)(1))
                    End If
                End If

                If (Len(CStr(yj)) > 0 And dYj.Exists(Trim$(CStr(yj)))) Or dNm.Exists(Trim$(nm)) Then
                    ' 既にデータにある → 予約解除のみ
                    restoredFree = restoredFree + 1
                Else
                    ' データに無い → 末尾へ追加
                    dr = LastDataRow(wsD, 3) + 1
                    If dr < 2 Then dr = 2
                    wsD.Cells(dr, 1).Value = dr - 1                  ' No.
                    wsD.Cells(dr, 3).Value = nm                      ' 薬品名
                    wsD.Cells(dr, 5).Value = CDbl(qty) * nyusu       ' 発注点（必要数量が qty になる調整）
                    wsD.Cells(dr, 6).Value = 0                       ' 在庫数量
                    wsD.Cells(dr, 7).Value = qty                     ' 発注量
                    If Len(CStr(yj)) > 0 Then wsD.Cells(dr, 16).Value = yj   ' YJコード
                    restoredAdd = restoredAdd + 1
                End If

                wsY.Rows(i).Delete
            End If
        End If
    Next i

    Application.Calculation = savedCalc
    Application.CalculateFull
    Application.ScreenUpdating = True
    MsgBox "予約期限が到来した " & (restoredAdd + restoredFree) & " 件を発注可能に戻しました。" & vbCrLf & _
           "（データへ追加 " & restoredAdd & " 件 / 予約解除のみ " & restoredFree & " 件）", _
           vbInformation, "予約分の復活"
    Exit Sub
EH:
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました：" & vbCrLf & Err.Description, vbExclamation, "予約分の復活"
End Sub

'==============================================================
' ③ 発注点割れデータを取込（レセコン CSV → 発注点割れデータ）
'
'    <ブックのドライブ>\発注割れシート\発注点割れ一覧表.csv を読み込み、
'    「発注点割れデータ」2行目以降を全入れ替えで書き込む。
'    予約発注リスト（状態＝予約中）に載っている薬品は取込から除外
'    （＝二重発注防止）。入荷待ちは除外しない。
'==============================================================
Public Sub 発注点割れデータを取込()
    Dim wsD As Worksheet, wsY As Worksheet, wsMst As Worksheet
    Dim savedCalc As XlCalculation
    Dim csvPath As String, ctlPath As String, taibiText As String

    On Error GoTo EH

    If ThisWorkbook.ReadOnly Then
        MsgBox "現在別のPCで編集中のため、更新・取り込みはできません。", vbExclamation, "発注点割れデータの取込"
        Exit Sub
    End If

    Set wsD = ThisWorkbook.Worksheets(SH_DATA)
    Set wsY = ThisWorkbook.Worksheets(SH_YOYAKU)
    Set wsMst = ThisWorkbook.Worksheets(SH_MASTER)

    csvPath = CsvSourcePath()
    If Len(csvPath) = 0 Then
        MsgBox "CSVファイル（発注点割れ一覧表.csv）が見つかりませんでした。" & vbCrLf & _
               "共有フォルダ内に発注点割れ一覧表.csvがあるか確認してください。", vbExclamation, "発注点割れデータの取込"
        Exit Sub
    End If

    ' .ctl から対象日を取得（表示用・任意）
    taibiText = ""
    ctlPath = Left$(csvPath, InStrRev(csvPath, ".")) & "ctl"
    If Len(Dir$(ctlPath)) > 0 Then
        Dim ctlLines() As String
        ctlLines = SplitLines(ReadTextSJIS(ctlPath))
        If UBound(ctlLines) >= 0 Then taibiText = Trim$(ctlLines(0))
    End If

    ' CSV 読み込み・パース
    Dim rawLines() As String
    rawLines = SplitLines(ReadTextSJIS(csvPath))

    Dim csvRows As Collection
    Set csvRows = New Collection
    Dim li As Long
    For li = LBound(rawLines) To UBound(rawLines)
        If Len(Trim$(rawLines(li))) > 0 Then csvRows.Add ParseCsvLine(rawLines(li))
    Next li
    If csvRows.Count = 0 Then
        MsgBox "CSVに行がありません。取込を中止します。" & vbCrLf & csvPath, vbExclamation, "発注点割れデータの取込"
        Exit Sub
    End If

    ' 1行目の見出し判定（先頭列が数値でなければ見出し）
    Dim startIdx As Long, h0 As String
    h0 = ""
    If UBound(csvRows(1)) >= 0 Then h0 = Trim$(csvRows(1)(0))
    If IsNumeric(h0) Then startIdx = 1 Else startIdx = 2
    If csvRows.Count < startIdx Then
        MsgBox "CSVにデータ行がありません。取込を中止します。", vbExclamation, "発注点割れデータの取込"
        Exit Sub
    End If

    ' 除外セット：予約発注リスト（状態＝予約中）の 薬品名 と JAN
    Dim exclName As Object, exclJan As Object
    Set exclName = CreateObject("Scripting.Dictionary"): exclName.CompareMode = vbTextCompare
    Set exclJan = CreateObject("Scripting.Dictionary"): exclJan.CompareMode = vbTextCompare
    Dim yLast As Long, r As Long, stt As String, nmv As String, janv As String
    yLast = LastDataRow(wsY, 1)
    For r = 2 To yLast
        stt = Trim$(CStr(wsY.Cells(r, 5).Value))
        If stt = "予約中" Then
            nmv = Trim$(CStr(wsY.Cells(r, 2).Value))
            janv = Trim$(CStr(wsY.Cells(r, 1).Value))
            If Len(nmv) > 0 And Not exclName.Exists(nmv) Then exclName.Add nmv, True
            If Len(janv) > 0 And Not exclJan.Exists(janv) Then exclJan.Add janv, True
        End If
    Next r

    ' マスタ YJコード(トリム) -> JANコード
    Dim yj2jan As Object
    Set yj2jan = CreateObject("Scripting.Dictionary"): yj2jan.CompareMode = vbTextCompare
    Dim mLast2 As Long, yjk As String
    mLast2 = LastDataRow(wsMst, 2)
    For r = 2 To mLast2
        yjk = Trim$(CStr(wsMst.Cells(r, 2).Value))
        If Len(yjk) > 0 And Not yj2jan.Exists(yjk) Then _
            yj2jan.Add yjk, Trim$(CStr(wsMst.Cells(r, 3).Value))
    Next r

    ' 取込対象を選別
    Dim outRows As Collection
    Set outRows = New Collection
    Dim total As Long, excluded As Long
    Dim idx As Long, f As Variant
    Dim cName As String, cYj As String, cJan As String
    For idx = startIdx To csvRows.Count
        f = csvRows(idx)
        total = total + 1
        cName = "": If UBound(f) >= 2 Then cName = Trim$(f(2))
        cYj = "":   If UBound(f) >= 15 Then cYj = Trim$(f(15))
        cJan = ""
        If Len(cYj) > 0 Then If yj2jan.Exists(cYj) Then cJan = yj2jan(cYj)

        If exclName.Exists(cName) Or (Len(cJan) > 0 And exclJan.Exists(cJan)) Then
            excluded = excluded + 1
        Else
            outRows.Add f
        End If
    Next idx

    Dim nOut As Long
    nOut = outRows.Count
    If nOut = 0 Then
        MsgBox "取り込む行がありません（全件が予約中で除外されました）。", vbExclamation, "発注点割れデータの取込"
        Exit Sub
    End If

    ' 確認ダイアログ
    Dim msg As String
    msg = IIf(Len(taibiText) > 0, taibiText & vbCrLf, "") & _
          "CSV : " & csvPath & vbCrLf & vbCrLf & _
          "CSV データ件数 ： " & total & vbCrLf & _
          "予約中で除外   ： " & excluded & vbCrLf & _
          "取込件数       ： " & nOut & vbCrLf & vbCrLf & _
          "「発注点割れデータ」の2行目以降をすべて置き換えます。よろしいですか？"
    If MsgBox(msg, vbQuestion + vbYesNo, "発注点割れデータの取込") <> vbYes Then Exit Sub

    savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 出力配列（16列）。数値列は数値に変換
    Dim numCols As String
    numCols = ",5,6,7,8,9,10,13,15,"
    Dim outArr() As Variant
    ReDim outArr(1 To nOut, 1 To 16)
    Dim c As Long, tok As String
    For idx = 1 To nOut
        f = outRows(idx)
        For c = 1 To 16
            tok = "": If UBound(f) >= (c - 1) Then tok = f(c - 1)
            If c = 1 Then
                outArr(idx, 1) = idx
            ElseIf InStr(numCols, "," & c & ",") > 0 And IsNumeric(tok) Then
                outArr(idx, c) = CDbl(tok)
            ElseIf Len(tok) = 0 Then
                outArr(idx, c) = Empty
            Else
                outArr(idx, c) = tok
            End If
        Next c
    Next idx

    ' 既存データをクリア（2行目以降 A:P）
    Dim oldLast As Long, clrLast As Long
    oldLast = LastDataRow(wsD, 3)
    If oldLast < 2 Then oldLast = 2
    clrLast = Application.Max(oldLast, nOut + 1)
    wsD.Range(wsD.Cells(2, 1), wsD.Cells(clrLast, 16)).ClearContents

    ' 書き込み
    wsD.Range(wsD.Cells(2, 1), wsD.Cells(nOut + 1, 16)).Value = outArr

    Application.Calculation = savedCalc
    Application.CalculateFull
    Application.ScreenUpdating = True

    Dim doneMsg As String
    doneMsg = "取込完了：" & nOut & " 件を「発注点割れデータ」に書き込みました。"
    If excluded > 0 Then doneMsg = doneMsg & vbCrLf & "（予約中 " & excluded & " 件を除外）"
    MsgBox doneMsg, vbInformation, "発注点割れデータの取込"
    Exit Sub
EH:
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました：" & vbCrLf & Err.Description, vbExclamation, "発注点割れデータの取込"
End Sub

'--------------------------------------------------------------
'  ヘルパー
'--------------------------------------------------------------
Private Function LastDataRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If r < 1 Then r = 1
    LastDataRow = r
End Function

Private Function NormCode(ByVal v As Variant) As String
    NormCode = Trim$(CStr(v))
End Function

' CSVの取得元パスを解決する。
'   1) ブックと同じドライブの \発注割れシート\発注点割れ一覧表.csv
'   2) ブックと同じフォルダの 発注点割れ一覧表.csv
'   3) 見つからなければファイル選択ダイアログ
Private Function CsvSourcePath() As String
    Dim base As String, p As String
    CsvSourcePath = ""
    base = ThisWorkbook.path
    If Len(base) >= 2 Then
        p = Left$(base, 2) & "\発注割れシート\発注点割れ一覧表.csv"
        If Len(Dir$(p)) > 0 Then CsvSourcePath = p: Exit Function
        p = base & "\発注点割れ一覧表.csv"
        If Len(Dir$(p)) > 0 Then CsvSourcePath = p: Exit Function
    End If
    Dim fsel As Variant
    fsel = Application.GetOpenFilename("CSVファイル (*.csv),*.csv", , "発注点割れ一覧表.csv を選択してください")
    If VarType(fsel) <> vbBoolean Then CsvSourcePath = CStr(fsel)
End Function

' Shift_JIS（システム ANSI = CP932）のテキストファイルを読み込んで文字列で返す
Private Function ReadTextSJIS(ByVal path As String) As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(path, 1, False)   ' 1 = ForReading
    If Not ts.AtEndOfStream Then ReadTextSJIS = ts.ReadAll
    ts.Close
End Function

' 改行（CRLF / CR / LF いずれも）で分割
Private Function SplitLines(ByVal s As String) As String()
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    SplitLines = Split(s, vbLf)
End Function

' CSV 1行を配列（0起点）へ分解。二重引用符のクォート・エスケープに対応
Private Function ParseCsvLine(ByVal line As String) As String()
    Dim res() As String
    Dim n As Long
    ReDim res(0 To 63)
    Dim i As Long, ch As String, cur As String
    Dim inQ As Boolean

    n = 0
    inQ = False
    cur = ""
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If inQ Then
            If ch = """" Then
                If i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inQ = True
            ElseIf ch = "," Then
                If n > UBound(res) Then ReDim Preserve res(0 To n + 32)
                res(n) = cur
                n = n + 1
                cur = ""
            Else
                cur = cur & ch
            End If
        End If
    Next i
    If n > UBound(res) Then ReDim Preserve res(0 To n)
    res(n) = cur
    ReDim Preserve res(0 To n)
    ParseCsvLine = res
End Function
