Attribute VB_Name = "MermaidToShapes"
Option Explicit

' Mermaid フローチャート -> Excel オートシェイプ変換
' アクティブシートの A 列に Mermaid の行を書いて MermaidToShapes を実行する。
' 配置は「記法」シートの配置方向。上から下は TD、左から右は LR。省略時は TD。
' 記法とオブジェクト名の対応は「記法」シート。お手本図形は「設定」シート（オブジェクト名で置く）。

Private Const SETTING_SHEET As String = "設定"
Private Const NOTATION_SHEET As String = "記法"

' 図形位置はポイント。ここを変えると全体の配置が変わる。
' 左から右 (LR)。段は Left 方向、同じ段の並びは Top 方向。
Private Const LR_ORIGIN_LEFT As Single = 250   ' 最初の段の Left
Private Const LR_RANK_GAP As Single = 50        ' 段と段のすき間
Private Const LR_SIBLING_GAP As Single = 30     ' 同じ段の図形のすき間
Private Const LR_MIN_TOP As Single = 20         ' Top の下限
' 上から下 (TD)。段は Top 方向、同じ段の並びは Left 方向。
Private Const TD_ORIGIN_TOP As Single = 20      ' 最初の段の Top
Private Const TD_RANK_GAP As Single = 40        ' 段と段のすき間
Private Const TD_SIBLING_GAP As Single = 30     ' 同じ段の図形のすき間
Private Const TD_MIN_LEFT As Single = 250       ' Left の下限

Private ShapeRules As Collection   ' 開始 & vbLf & 終了 & vbLf & オブジェクト名。開始が長い順
Private LineOpList As Collection   ' 矢印の記法。シートの順
Private LineMap As Object          ' 矢印の記法 -> オブジェクト名
Private DirectionMap As Object     ' 配置の記法 -> TD / LR
Private DefaultShape As String     ' 括弧のない ID に使うオブジェクト名

Sub MermaidToShapes()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim noteErr As String
    If Not SheetExists(NOTATION_SHEET) Then
        MsgBox "「記法」シートがありません。", vbExclamation
        Exit Sub
    End If
    If Not SheetExists(SETTING_SHEET) Then
        MsgBox "「設定」シートがありません。", vbExclamation
        Exit Sub
    End If
    If Not LoadNotation(noteErr) Then
        MsgBox noteErr, vbExclamation
        Exit Sub
    End If
    Dim cfg As Worksheet: Set cfg = Worksheets(SETTING_SHEET)
    Dim nodes As Object: Set nodes = CreateObject("Scripting.Dictionary")  ' id -> 表示文字
    Dim kinds As Object: Set kinds = CreateObject("Scripting.Dictionary")  ' id -> オブジェクト名
    Dim level As Object: Set level = CreateObject("Scripting.Dictionary")
    Dim edges As New Collection
    Dim r As Long, s As String, flow As String, d As String

    ' 1. 読み込み（A列）。記法が不正なセルはピンクにして、その行は図形にしない
    Dim re As Object, m As Object, pos As Long, prevId As String, curId As String, kind As String, label As String
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = ArrowPattern()
    For r = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        s = NormalizeLine(Trim(CStr(ws.Cells(r, 1).Value & "")))
        If s <> "" And Not LineOK(s, re) Then
            ws.Cells(r, 1).Interior.Color = RGB(255, 192, 203)
        Else
            If ws.Cells(r, 1).Interior.Color = RGB(255, 192, 203) Then ws.Cells(r, 1).Interior.ColorIndex = xlNone
            If flow = "" Then
                d = DirectionOf(s)
                If d = "TD" Or d = "LR" Then flow = d
            End If
            re.Pattern = ArrowPattern()
            If s <> "" And re.Test(s) Then
                re.Pattern = ArrowPattern()
                pos = 1: prevId = "": kind = "": label = ""
                For Each m In re.Execute(s)
                    curId = ParseNode(Trim(Mid(s, pos, m.FirstIndex + 1 - pos)), nodes, kinds)
                    If prevId <> "" Then edges.Add prevId & vbLf & curId & vbLf & kind & vbLf & label
                    prevId = curId
                    kind = LineKind(ArrowOp(m.Value))
                    label = ArrowLabel(m.Value)
                    pos = m.FirstIndex + m.Length + 1
                Next
                curId = ParseNode(Trim(Mid(s, pos)), nodes, kinds)
                If prevId <> "" And curId <> "" Then edges.Add prevId & vbLf & curId & vbLf & kind & vbLf & label
            End If
        End If
    Next

    If flow = "" Then flow = "TD"

    ' 2. 段を決める（矢印の先は元より1段進む）
    Dim k, e, changed As Boolean, n As Long, src As String, dst As String
    For Each k In nodes.Keys: level(k) = 0: Next
    Do
        changed = False: n = n + 1
        For Each e In edges
            src = EdgePart(CStr(e), 0)
            dst = EdgePart(CStr(e), 1)
            If level(dst) < level(src) + 1 Then level(dst) = level(src) + 1: changed = True
        Next
    Loop While changed And n < 100

    ' 3. 使うお手本が「設定」シートにあるか確認する
    Dim need As Object, miss As String
    Set need = CreateObject("Scripting.Dictionary")
    For Each k In nodes.Keys
        need(CStr(kinds(k))) = True
    Next
    For Each e In edges
        need(EdgePart(e, 2)) = True
    Next
    For Each k In need.Keys
        If FindTemplate(cfg, CStr(k)) Is Nothing Then miss = miss & CStr(k) & vbCrLf
    Next
    If miss <> "" Then
        MsgBox "設定シートに、次の名前のお手本図形がありません。" & vbCrLf & vbCrLf & miss, vbExclamation
        Exit Sub
    End If

    ' 4. お手本から箱を作る
    Dim shp As Object: Set shp = CreateObject("Scripting.Dictionary")
    Dim sh As Shape
    For Each k In nodes.Keys
        Set sh = CopyTemplate(FindTemplate(cfg, CStr(kinds(k))), ws)
        sh.Name = "node_" & k
        sh.TextFrame2.TextRange.Text = nodes(k)
        Set shp(k) = sh
    Next

    ' 5. 段ごとに並べ、合流は入ってくる図形の中央に置く
    PlaceNodes shp, level, edges, flow

    ' 6. 矢印の種類に応じたお手本で接続
    ' 書類などは接続点の番号が四角形と違う。最短経路への引き直しは横から付けてしまうのでしない。
    Dim c As Shape
    Dim failNo As Long, failSrc As String, failDesc As String
    On Error GoTo ConnectFail
    Application.ScreenUpdating = False
    For Each e In edges
        Set c = CopyTemplate(FindTemplate(cfg, EdgePart(e, 2)), ws)
        If flow = "LR" Then
            c.ConnectorFormat.BeginConnect shp(EdgePart(e, 0)), SideSite(shp(EdgePart(e, 0)), "right")
            c.ConnectorFormat.EndConnect shp(EdgePart(e, 1)), SideSite(shp(EdgePart(e, 1)), "left")
        Else
            c.ConnectorFormat.BeginConnect shp(EdgePart(e, 0)), SideSite(shp(EdgePart(e, 0)), "bottom")
            c.ConnectorFormat.EndConnect shp(EdgePart(e, 1)), SideSite(shp(EdgePart(e, 1)), "top")
        End If
        AddLineLabel ws, c, EdgePart(e, 3)
    Next
    Application.ScreenUpdating = True
    Exit Sub
ConnectFail:
    Application.ScreenUpdating = True
    failNo = Err.Number: failSrc = Err.Source: failDesc = Err.Description
    Err.Raise failNo, failSrc, failDesc
End Sub

' アクティブシートの図形・コネクタなどをすべて削除する（フォームボタンは残す）
Sub ClearAllObjects()
    Dim i As Long
    With ActiveSheet
        For i = .Shapes.Count To 1 Step -1
            If Not IsFormButton(.Shapes(i)) Then .Shapes(i).Delete
        Next
    End With
End Sub

' アクティブシートの図形・コネクタなどをすべて選択する（フォームボタンは除く）
Sub SelectAllObjects()
    Dim sh As Shape, first As Boolean
    first = True
    For Each sh In ActiveSheet.Shapes
        If Not IsFormButton(sh) Then
            sh.Select Replace:=first
            first = False
        End If
    Next
End Sub

' 段の位置を決める。横並びの基準は、その段で一番大きい図形
Private Sub PlaceNodes(shp As Object, level As Object, edges As Collection, ByVal flow As String)
    Dim rank As Object, bucket As Collection, k, id, lv As Long, maxLv As Long
    Set rank = CreateObject("Scripting.Dictionary")
    For Each k In shp.Keys
        lv = level(k)
        If lv > maxLv Then maxLv = lv
        If Not rank.Exists(lv) Then
            Set bucket = New Collection
            Set rank(lv) = bucket
        Else
            Set bucket = rank(lv)
        End If
        bucket.Add CStr(k)
    Next

    Dim mainAt As Single, mainGap As Single, crossGap As Single, minCross As Single
    If flow = "LR" Then
        mainAt = LR_ORIGIN_LEFT: mainGap = LR_RANK_GAP: crossGap = LR_SIBLING_GAP: minCross = LR_MIN_TOP
    Else
        mainAt = TD_ORIGIN_TOP: mainGap = TD_RANK_GAP: crossGap = TD_SIBLING_GAP: minCross = TD_MIN_LEFT
    End If

    Dim ids As Collection, sh As Shape
    Dim n As Long, i As Long, j As Long
    Dim arr() As String, desired() As Single
    Dim tmp As String, tmpD As Single
    Dim total As Single, mean As Single, startAt As Single, cursor As Single
    Dim cnt As Long, mainSize As Single

    For lv = 0 To maxLv
        If rank.Exists(lv) Then
            Set ids = rank(lv)
            n = ids.Count
            ReDim arr(1 To n)
            ReDim desired(1 To n)
            i = 0
            For Each id In ids
                i = i + 1
                arr(i) = CStr(id)
                desired(i) = ParentCenter(arr(i), shp, edges, flow)
            Next
            For i = 1 To n - 1
                For j = i + 1 To n
                    If desired(j) < desired(i) Then
                        tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                        tmpD = desired(i): desired(i) = desired(j): desired(j) = tmpD
                    End If
                Next
            Next

            total = 0: mean = 0: cnt = 0
            For i = 1 To n
                Set sh = shp(arr(i))
                If flow = "LR" Then total = total + sh.Height Else total = total + sh.Width
                If i < n Then total = total + crossGap
                If desired(i) >= 0 Then mean = mean + desired(i): cnt = cnt + 1
            Next
            If cnt > 0 Then mean = mean / cnt Else mean = minCross + total / 2
            startAt = mean - total / 2
            If startAt < minCross Then startAt = minCross

            mainSize = 0: cursor = startAt
            For i = 1 To n
                Set sh = shp(arr(i))
                If flow = "LR" Then
                    sh.Left = mainAt
                    sh.Top = cursor
                    If sh.Width > mainSize Then mainSize = sh.Width
                    cursor = cursor + sh.Height + crossGap
                Else
                    sh.Top = mainAt
                    sh.Left = cursor
                    If sh.Height > mainSize Then mainSize = sh.Height
                    cursor = cursor + sh.Width + crossGap
                End If
            Next
            mainAt = mainAt + mainSize + mainGap
        End If
    Next
End Sub

' 入ってくる矢印の元の中央。元がなければ -1
Private Function ParentCenter(ByVal id As String, shp As Object, edges As Collection, ByVal flow As String) As Single
    Dim e, sh As Shape, sum As Single, cnt As Long
    ParentCenter = -1
    For Each e In edges
        If EdgePart(CStr(e), 1) = id And shp.Exists(EdgePart(CStr(e), 0)) Then
            Set sh = shp(EdgePart(CStr(e), 0))
            If flow = "LR" Then sum = sum + sh.Top + sh.Height / 2 Else sum = sum + sh.Left + sh.Width / 2
            cnt = cnt + 1
        End If
    Next
    If cnt > 0 Then ParentCenter = sum / cnt
End Function

' フォームコントロールのボタンか
Private Function IsFormButton(sh As Shape) As Boolean
    If sh.Type <> msoFormControl Then Exit Function
    IsFormButton = (sh.FormControlType = xlButtonControl)
End Function

' 「設定」シートから名前でお手本を探す。無いときは Nothing
Private Function FindTemplate(cfg As Worksheet, ByVal name As String) As Shape
    Dim sh As Shape
    For Each sh In cfg.Shapes
        If StrComp(sh.Name, name, vbTextCompare) = 0 Then
            Set FindTemplate = sh
            Exit Function
        End If
    Next
End Function

' お手本と同じ種類・サイズの図形を作り、書式を写して返す
Private Function CopyTemplate(tpl As Shape, dest As Worksheet) As Shape
    Dim sh As Shape
    If tpl.Connector Then
        Set sh = dest.Shapes.AddConnector(tpl.ConnectorFormat.Type, 0, 0, 10, 10)
    Else
        Set sh = dest.Shapes.AddShape(tpl.AutoShapeType, 0, 0, tpl.Width, tpl.Height)
    End If
    tpl.PickUp
    sh.Apply
    If Not tpl.Connector Then
        With sh.TextFrame2
            .HorizontalAnchor = tpl.TextFrame2.HorizontalAnchor
            .VerticalAnchor = tpl.TextFrame2.VerticalAnchor
            .TextRange.Font.Size = tpl.TextFrame2.TextRange.Font.Size
            .TextRange.Font.Fill.ForeColor.RGB = tpl.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
            .TextRange.Font.NameFarEast = tpl.TextFrame2.TextRange.Font.NameFarEast
        End With
    End If
    Set CopyTemplate = sh
End Function

' 接続点。図形の種類で番号の並びが違うので、位置で上・右・下・左を選ぶ
Private Function SideSite(sh As Shape, ByVal side As String) As Long
    Dim i As Long, n As Long, best As Long
    Dim x As Single, y As Single, cx As Single, cy As Single
    Dim score As Single, bestScore As Single
    n = sh.ConnectionSiteCount
    If n < 1 Then SideSite = 1: Exit Function
    cx = sh.Left + sh.Width / 2
    cy = sh.Top + sh.Height / 2
    best = 1
    bestScore = -1E+30
    For i = 1 To n
        SiteXY sh, i, x, y
        Select Case side
            Case "top": score = (cy - y) * 10000 - Abs(x - cx)
            Case "bottom": score = (y - cy) * 10000 - Abs(x - cx)
            Case "left": score = (cx - x) * 10000 - Abs(y - cy)
            Case Else: score = (x - cx) * 10000 - Abs(y - cy)
        End Select
        If score > bestScore Then bestScore = score: best = i
    Next
    SideSite = best
End Function

' 接続点のシート上の座標。直線を一度つないで端を読む
Private Sub SiteXY(sh As Shape, ByVal site As Long, ByRef x As Single, ByRef y As Single)
    Dim c As Shape
    Dim cx As Single, cy As Single, d As Single, best As Single
    Dim pts(1 To 4, 1 To 2) As Single, i As Long
    cx = sh.Left + sh.Width / 2
    cy = sh.Top + sh.Height / 2
    Set c = sh.Parent.Shapes.AddConnector(msoConnectorStraight, cx, cy, cx + 5, cy + 5)
    c.ConnectorFormat.EndConnect sh, site
    pts(1, 1) = c.Left: pts(1, 2) = c.Top
    pts(2, 1) = c.Left + c.Width: pts(2, 2) = c.Top
    pts(3, 1) = c.Left: pts(3, 2) = c.Top + c.Height
    pts(4, 1) = c.Left + c.Width: pts(4, 2) = c.Top + c.Height
    c.Delete
    best = -1
    x = cx: y = cy
    For i = 1 To 4
        d = Abs(pts(i, 1) - cx) + Abs(pts(i, 2) - cy)
        If d > best Then best = d: x = pts(i, 1): y = pts(i, 2)
    Next
End Sub

' 「記法」シートの矢印を、長い順に並べた正規表現にする
Private Function ArrowPattern() As String
    Dim n As Long, i As Long, j As Long
    Dim ops() As String, tmp As String
    n = LineOpList.Count
    If n < 1 Then ArrowPattern = "$^": Exit Function
    ReDim ops(1 To n)
    For i = 1 To n: ops(i) = CStr(LineOpList(i)): Next
    For i = 1 To n - 1
        For j = i + 1 To n
            If Len(ops(j)) > Len(ops(i)) Then tmp = ops(i): ops(i) = ops(j): ops(j) = tmp
        Next
    Next
    tmp = ""
    For i = 1 To n
        If tmp <> "" Then tmp = tmp & "|"
        tmp = tmp & RegexEscape(ops(i))
    Next
    ArrowPattern = "(" & tmp & ")\s*(\|[^|]*\|)?"
End Function

Private Function ArrowOp(ByVal token As String) As String
    Dim p As Long
    p = InStr(token, "|")
    If p = 0 Then ArrowOp = Trim(token) Else ArrowOp = Trim(Left(token, p - 1))
End Function

Private Function ArrowLabel(ByVal token As String) As String
    Dim p As Long, q As Long
    p = InStr(token, "|")
    If p = 0 Then Exit Function
    q = InStr(p + 1, token, "|")
    If q = 0 Then Exit Function
    ArrowLabel = Trim(Mid(token, p + 1, q - p - 1))
End Function

' 条件文字を線の中央に置く。コネクタ自体には文字を載せられない
Private Sub AddLineLabel(ws As Worksheet, c As Shape, ByVal label As String)
    Dim tb As Shape
    If label = "" Then Exit Sub
    Set tb = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, c.Left, c.Top, 10, 10)
    tb.Line.Visible = msoFalse
    tb.Fill.Visible = msoTrue
    tb.Fill.ForeColor.RGB = RGB(255, 255, 255)
    With tb.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = label
        .TextRange.Font.Size = 11
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    tb.Left = c.Left + (c.Width - tb.Width) / 2
    tb.Top = c.Top + (c.Height - tb.Height) / 2
End Sub

' 矢印の記号 -> 記法シートのオブジェクト名
Private Function LineKind(ByVal op As String) As String
    If LineMap.Exists(op) Then LineKind = CStr(LineMap(op))
End Function

' 矢印を含む行の各ノードが、対応している記法かどうか
Private Function LineOK(ByVal s As String, re As Object) As Boolean
    Dim m As Object, pos As Long, d As String
    d = DirectionOf(s)
    If d = "?" Then Exit Function
    If d <> "" Then LineOK = True: Exit Function
    re.Pattern = ArrowPattern()
    If Not re.Test(s) Then LineOK = True: Exit Function
    re.Pattern = ArrowPattern()
    pos = 1
    For Each m In re.Execute(s)
        If Not ArrowEdgeOK(s, m.FirstIndex, ArrowOp(m.Value)) Then Exit Function
        If Not ChunkOK(Trim(Mid(s, pos, m.FirstIndex + 1 - pos))) Then Exit Function
        pos = m.FirstIndex + m.Length + 1
    Next
    LineOK = ChunkOK(Trim(Mid(s, pos)))
End Function

' ノード部分に残った | は、条件の書き方が閉じられていない
Private Function ChunkOK(ByVal s As String) As Boolean
    If InStr(s, "|") > 0 Then Exit Function
    ChunkOK = NodeOK(s)
End Function

' 定義より長い矢印（「--->」など）は、マッチの前後に矢印の文字が残る
Private Function ArrowEdgeOK(ByVal s As String, ByVal firstIndex As Long, ByVal op As String) As Boolean
    Dim ch As String
    If firstIndex >= 1 Then
        ch = Mid(s, firstIndex, 1)
        If ch = "-" Or ch = "." Or ch = "=" Or ch = ">" Then Exit Function
    End If
    If firstIndex + Len(op) + 1 <= Len(s) Then
        ch = Mid(s, firstIndex + Len(op) + 1, 1)
        If ch = "-" Or ch = "." Or ch = "=" Or ch = ">" Then Exit Function
    End If
    ArrowEdgeOK = True
End Function

' 記法シートの配置方向を読む。方向の行でなければ空、シートにない flowchart / graph は "?"
Private Function DirectionOf(ByVal s As String) As String
    Dim re As Object, key As String
    key = DirectionKey(s)
    If key <> "" And DirectionMap.Exists(key) Then
        DirectionOf = CStr(DirectionMap(key))
        Exit Function
    End If
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(flowchart|graph)(\s+\S+)?\s*$"
    If re.Test(s) Then DirectionOf = "?"
End Function

Private Function DirectionKey(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\s+"
    DirectionKey = UCase$(re.Replace(Trim$(NormalizeLine(s)), " "))
End Function

Private Function LayoutCode(ByVal meaning As String) As String
    Select Case meaning
        Case "上から下": LayoutCode = "TD"
        Case "左から右": LayoutCode = "LR"
    End Select
End Function

Private Sub AddDefaultDirections()
    DirectionMap(DirectionKey("flowchart TD")) = "TD"
    DirectionMap(DirectionKey("flowchart TB")) = "TD"
    DirectionMap(DirectionKey("flowchart")) = "TD"
    DirectionMap(DirectionKey("graph TD")) = "TD"
    DirectionMap(DirectionKey("graph TB")) = "TD"
    DirectionMap(DirectionKey("graph")) = "TD"
    DirectionMap(DirectionKey("flowchart LR")) = "LR"
    DirectionMap(DirectionKey("graph LR")) = "LR"
End Sub

' 閉じ記号が末尾にあるときだけ正しい記法とする
Private Function NodeOK(ByVal s As String) As Boolean
    Dim openAt As Long, rule As String, id As String
    If s = "" Then Exit Function
    If MatchShape(s, openAt, rule) Then
        NodeOK = IdOK(CleanId(Left(s, openAt - 1)))
        Exit Function
    End If
    If HasOpenMark(s) Then Exit Function
    NodeOK = IdOK(CleanId(s))
End Function

Private Function CloserAtEnd(ByVal s As String, ByVal startAt As Long, ByVal closer As String) As Boolean
    Dim q As Long
    q = InStr(startAt, s, closer)
    CloserAtEnd = (q > 0 And q + Len(closer) - 1 = Len(s))
End Function

' 記法シートの開始・終了に合うノードを解析して id を返す。kinds にはオブジェクト名を入れる
Private Function ParseNode(ByVal s As String, nodes As Object, kinds As Object) As String
    Dim openAt As Long, rule As String, id As String, op As String, cl As String
    s = NormalizeLine(s)
    If MatchShape(s, openAt, rule) Then
        id = CleanId(Left(s, openAt - 1))
        If Not IdOK(id) Then Exit Function
        op = EdgePart(rule, 0)
        cl = EdgePart(rule, 1)
        kinds(id) = EdgePart(rule, 2)
        nodes(id) = NodeText(s, openAt + Len(op), cl)
    Else
        id = CleanId(s)
        If IdOK(id) And Not HasOpenMark(s) Then
            If Not nodes.Exists(id) Then nodes(id) = id: kinds(id) = DefaultShape
        Else
            id = ""
        End If
    End If
    ParseNode = id
End Function

' 全角の括弧と矢印記号を半角にする
Private Function NormalizeLine(ByVal s As String) As String
    s = Replace(s, "［", "[")
    s = Replace(s, "］", "]")
    s = Replace(s, "（", "(")
    s = Replace(s, "）", ")")
    s = Replace(s, "｛", "{")
    s = Replace(s, "｝", "}")
    s = Replace(s, "－", "-")
    s = Replace(s, "ー", "-")
    s = Replace(s, "＞", ">")
    s = Replace(s, "｜", "|")
    s = Replace(s, "　", " ")
    NormalizeLine = s
End Function

' 見た目が同じ ID を同じ文字にそろえる
Private Function CleanId(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 9, 10, 13, 160, &H200B, &HFEFF
            Case &HFF10 To &HFF19: out = out & ChrW(code - &HFF10 + AscW("0"))
            Case &HFF21 To &HFF3A: out = out & ChrW(code - &HFF21 + AscW("A"))
            Case &HFF41 To &HFF5A: out = out & ChrW(code - &HFF41 + AscW("a"))
            Case Else: out = out & ch
        End Select
    Next
    CleanId = Trim(out)
End Function

Private Function EdgePart(ByVal item As String, ByVal index As Long) As String
    Dim parts
    parts = Split(item, vbLf)
    If index >= LBound(parts) And index <= UBound(parts) Then EdgePart = parts(index)
End Function

' 開始位置から閉じ記号の手前までを表示文字にする
Private Function NodeText(ByVal s As String, ByVal startAt As Long, ByVal closer As String) As String
    Dim q As Long
    q = InStr(startAt, s, closer)
    If q > 0 Then NodeText = Trim(Mid(s, startAt, q - startAt)) Else NodeText = Trim(Mid(s, startAt))
End Function

Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Worksheets
        If StrComp(ws.Name, name, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next
End Function

Private Function CellText(ws As Worksheet, ByVal row As Long, ByVal col As Long) As String
    CellText = Trim(CStr(ws.Cells(row, col).Value & ""))
End Function

' 「記法」シートから、図形の開始・終了と線の記法を読む
Private Function LoadNotation(ByRef errMsg As String) As Boolean
    Dim ws As Worksheet
    Dim headerRow As Long, c As Long, r As Long
    Dim openCol As Long, closeCol As Long, shapeNameCol As Long
    Dim lineOpCol As Long, lineNameCol As Long
    Dim dirMeanCol As Long, dirNoteCol As Long
    Dim foundOpen As Long, foundLineOp As Long, foundDir As Long
    Dim op As String, cl As String, nm As String, lay As String

    Set ws = Worksheets(NOTATION_SHEET)
    Set ShapeRules = New Collection
    Set LineOpList = New Collection
    Set LineMap = CreateObject("Scripting.Dictionary")
    Set DirectionMap = CreateObject("Scripting.Dictionary")
    DefaultShape = ""

    For headerRow = 1 To 15
        foundOpen = 0: foundLineOp = 0: foundDir = 0
        For c = 1 To 30
            If CellText(ws, headerRow, c) = "開始" And CellText(ws, headerRow, c + 1) = "終了" And CellText(ws, headerRow, c + 2) = "オブジェクト名" Then foundOpen = c
            If CellText(ws, headerRow, c) = "記法" And CellText(ws, headerRow, c + 1) = "オブジェクト名" Then foundLineOp = c
            If CellText(ws, headerRow, c) = "配置方向" And CellText(ws, headerRow, c + 1) = "記法" Then foundDir = c
        Next
        If foundOpen > 0 And foundLineOp > 0 Then
            openCol = foundOpen
            closeCol = foundOpen + 1
            shapeNameCol = foundOpen + 2
            lineOpCol = foundLineOp
            lineNameCol = foundLineOp + 1
            dirMeanCol = foundDir
            dirNoteCol = foundDir + 1
            Exit For
        End If
    Next
    If openCol = 0 Or lineOpCol = 0 Then
        errMsg = "「記法」シートに、開始・終了・オブジェクト名と、線の記法・オブジェクト名の見出しが必要です。"
        Exit Function
    End If

    For r = headerRow + 1 To headerRow + 300
        op = NormalizeLine(CellText(ws, r, openCol))
        cl = NormalizeLine(CellText(ws, r, closeCol))
        nm = CellText(ws, r, shapeNameCol)
        If op <> "" Or cl <> "" Or nm <> "" Then
            If op = "" Or cl = "" Or nm = "" Then
                errMsg = "「記法」シートの " & r & " 行目は、開始・終了・オブジェクト名をすべて入れてください。"
                Exit Function
            End If
            If DefaultShape = "" Or (op = "[" And cl = "]") Then DefaultShape = nm
            AddShapeRule op, cl, nm
        End If
    Next
    If ShapeRules.Count = 0 Then
        errMsg = "「記法」シートに図形の行がありません。"
        Exit Function
    End If

    For r = headerRow + 1 To headerRow + 300
        op = NormalizeLine(CellText(ws, r, lineOpCol))
        nm = CellText(ws, r, lineNameCol)
        If op <> "" Or nm <> "" Then
            If op = "" Or nm = "" Then
                errMsg = "「記法」シートの " & r & " 行目は、線の記法とオブジェクト名を入れてください。"
                Exit Function
            End If
            If Not LineMap.Exists(op) Then LineOpList.Add op
            LineMap(op) = nm
        End If
    Next
    If LineOpList.Count = 0 Then
        errMsg = "「記法」シートに線の行がありません。"
        Exit Function
    End If

    If dirMeanCol > 0 Then
        For r = headerRow + 1 To headerRow + 300
            nm = CellText(ws, r, dirMeanCol)
            op = DirectionKey(CellText(ws, r, dirNoteCol))
            If nm <> "" Or op <> "" Then
                If nm = "" Or op = "" Then
                    errMsg = "「記法」シートの " & r & " 行目は、配置方向と記法を入れてください。"
                    Exit Function
                End If
                lay = LayoutCode(nm)
                If lay = "" Then
                    errMsg = "「記法」シートの " & r & " 行目の配置方向は、「上から下」または「左から右」にしてください。"
                    Exit Function
                End If
                DirectionMap(op) = lay
            End If
        Next
    End If
    If DirectionMap.Count = 0 Then AddDefaultDirections
    LoadNotation = True
End Function

' 開始が長いルールを先に試す
Private Sub AddShapeRule(ByVal op As String, ByVal cl As String, ByVal nm As String)
    Dim item As String, i As Long
    item = op & vbLf & cl & vbLf & nm
    For i = 1 To ShapeRules.Count
        If Len(EdgePart(CStr(ShapeRules(i)), 0)) < Len(op) Then
            ShapeRules.Add item, Before:=i
            Exit Sub
        End If
    Next
    ShapeRules.Add item
End Sub

' 左から見て、開始が一致し終了が末尾にある最初のルール
Private Function MatchShape(ByVal s As String, ByRef openAt As Long, ByRef rule As String) As Boolean
    Dim i As Long, item, op As String, cl As String
    For i = 1 To Len(s)
        For Each item In ShapeRules
            op = EdgePart(CStr(item), 0)
            If Mid(s, i, Len(op)) = op Then
                cl = EdgePart(CStr(item), 1)
                If CloserAtEnd(s, i + Len(op), cl) Then
                    openAt = i
                    rule = CStr(item)
                    MatchShape = True
                    Exit Function
                End If
            End If
        Next
    Next
End Function

Private Function HasOpenMark(ByVal s As String) As Boolean
    Dim item, op As String
    For Each item In ShapeRules
        op = EdgePart(CStr(item), 0)
        If op <> "" And InStr(s, op) > 0 Then HasOpenMark = True: Exit Function
    Next
End Function

Private Function IdOK(ByVal id As String) As Boolean
    Dim i As Long, ch As String
    If id = "" Or InStr(id, "|") > 0 Then Exit Function
    For i = 1 To Len(id)
        ch = Mid(id, i, 1)
        If ch = "[" Or ch = "]" Or ch = "(" Or ch = ")" Or ch = "{" Or ch = "}" Then Exit Function
    Next
    IdOK = True
End Function

Private Function RegexEscape(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If InStr("\.+*?^${}()|[]-", ch) > 0 Then out = out & "\" & ch Else out = out & ch
    Next
    RegexEscape = out
End Function
