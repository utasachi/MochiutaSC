; MochiutaSC.ahk
; もちからuta-netスクロール歌詞付与 v0.5
; https://ahkwiki.net/

#Requires AutoHotkey v2.0
SetWorkingDir(A_ScriptDir)
;MsgBox A_AhkVersion

asshead  := "MochiutaSC_header.ass"
assshead := "MochiutaSC_syncheader.ass"
bgv := "image\bgv8min_s.mp4"
ytthumb := "image\ytthumb.jpg"
noimg := "image\noimgw.png"

FENRIR := "https://search.fenrir-inc.com/?hl=ja&channel=sleipnir_s&safe=off&lr=all&fr=ss&q="
FENRIR_U := FENRIR "歌ネット 歌詞ページ "
FENRIR_Y := FENRIR "youtube "

StdoutToVar(cmd) {              ;標準出力の値を変数に
    shell := ComObject("WScript.Shell")
    exec := shell.Exec(cmd)
    while (exec.Status = 0)
        Sleep(10)
    return exec.StdOut.ReadAll()
}
ar2txtb(arr){                   ;配列を改行区切りに
    text := ""
    for i, v in arr
        text .= i ": " v "`n"
    return text
}
GetF(path) {                    ;ファイル名部分だけ返す
    SplitPath path, &name
    return name
}
HtmlDecode(s) {                 ;htmlデコード（簡易）
    s := StrReplace(s, "＆", "&")
    s := StrReplace(s, "；", ";")
    s := StrReplace(s, "&nbsp;", " ")
    s := StrReplace(s, "&amp;", "&")
    s := StrReplace(s, "&lt;", "<")
    s := StrReplace(s, "&gt;", ">")
    s := StrReplace(s, "&quot;", '"')
    s := StrReplace(s, "&#039;", "'")
    s := StrReplace(s, "&#39;", "'")
    return s
}
UriEncode(str, encoding := "UTF-8"){        ;unicodeにエンコード
    out := ""
    buf := Buffer(StrPut(str, encoding))
    StrPut(str, buf, encoding)
    Loop buf.Size - 1 {
        b := NumGet(buf, A_Index - 1, "UChar")
        if (b >= 0x30 && b <= 0x39)      ; 0-9
         || (b >= 0x41 && b <= 0x5A)     ; A-Z
         || (b >= 0x61 && b <= 0x7A)     ; a-z
         || InStr("-_.~", Chr(b))
            out .= Chr(b)
        else
            out .= "%" Format("{:02X}", b)
    }
    return out
}
ResponseBodyToText(body, charset := "utf-8"){  ;レスポンスのbodyをテキストに設定
    stm := ComObject("ADODB.Stream")
    stm.Type := 1          ; binary
    stm.Open()
    stm.Write(body)
    stm.Position := 0
    stm.Type := 2          ; text
    stm.Charset := charset
    text := stm.ReadText()
    stm.Close()
    return text
}
rep(s) {                    ;文字列置き換え（簡易）
    if (s = "")
        return ""
    ; HTMLエンティティをデコード
    s := HtmlDecode(s)
    replacements := Map(
        "=", "＝",  ",", "，",  "'", "’",  '"', "“",
        "\", "＼", "/", "／", ":", "：",  ";", "；",
        "<", "＜",  ">", "＞", "|", "｜", "~", "～",
        "^", "＾",  "``", "｀", "*", "＊", "?", "？",
        "%", "％",  "$", "＄", "[", "［", "]", "］",
        "@", "＠",  "　", " ",  " ", " ", "+", "＋",
        "é", "e",   "&", "＆", "〜", "～"
    )
    for k, v in replacements
        s := StrReplace(s, k, v)
    return s
}
sec2mmss(duration) {        ;秒 → mm:ss に変換
    sec := duration + 0
    mm := Floor(sec / 60)
    ss := Mod(sec, 60)
    return Format("{:02}:{:02}", mm, ss)
}
TimeDiffMs(t1, t2){         ;t1とt2の差分の時間をミリ秒で返す
    RegExMatch(t1, "(\d{2}):(\d{2})\.(\d{2})", &m1)
    ms1 := (m1[1] * 60 + m1[2]) * 1000 + m1[3] * 10
    RegExMatch(t2, "(\d{2}):(\d{2})\.(\d{2})", &m2)
    ms2 := (m2[1] * 60 + m2[2]) * 1000 + m2[3] * 10
    return ms2 - ms1
}
SecToLrcTime(sec){          ;秒表記をLRCの[99:99.99]形式に
    min := Floor(sec / 60)
    sec2 := Floor(Mod(sec, 60))
    cs := Round((sec - Floor(sec)) * 100)
    if (cs = 100) {    ;丸めで100になった場合の補正
        cs := 0
        sec2++
        if (sec2 = 60) {
            sec2 := 0
            min++
        }
    }
    return Format("{:02}:{:02}.{:02}", min, sec2, cs)
}
HasLineTag(){       ;行タグかどうか判定
    for line in StrSplit(kashi.Value, "`n", "`r"){
        if RegExMatch(line, "^\[\d{2}:\d{2}\.\d{2,3}\]")
            return true
    }
    return false
}
ClearsInfo(){           ;曲情報消す
    title.Value := "" , artst.Value := "" , tieup.Value := "" , year.Value  := ""
    lyric.Value := "" , cmpst.Value := "" , arngm.Value := "" , kashi.Value := ""
    utaID.Value := "" , vidid.Value := "" , mtype.Value := "" , vname.Value := ""
    ystart.Value := "", yend.Value := "" ,  kstyle.Value := "", subdir.Value := ""
    loopvid.Value := ""
}
GetMPCPath() {          ;MPC-BEのパスを取得
    ini := "MochikaraSC.ini"
    key := "mpcpath"
    path := IniRead(ini, "path", key, "")
    if (path != "" && FileExist(path))
        return path
    MsgBox("MPC-BEのパスが設定されていません。")
    selected := FileSelect(1, , "mpc-be64.exe を選択してください", "mpc-be64.exe (*.exe)")
    if (selected = "")
        return ""   ; キャンセル
    IniWrite(selected, ini, "path", key)
    return selected
}
GetBaseDir() {          ;basedirを取得
    ini := "MochikaraSC.ini"
    key := "basedir"
    path := IniRead(ini, "path", key, "")
    if (path != "" && DirExist(path))
        return path
    MsgBox("basedirが設定されていません。")
    selected := DirSelect(,,"basedir を選択してください")
    if (selected = "")
        return ""   ; キャンセル
    IniWrite(selected, ini, "path", key)
    return selected
}

ReadAssf(assf) {        ;ass読込 スクロール歌詞/行同期歌詞 両対応
    if !FileExist(assf)
        return False
    ClearsInfo()
    fields := Map(
        "utaid", utaID,         "title", title,     "artist", artst,    "tieup", tieup,
        "year", year,           "subdir", subdir,   "lyrics", lyric,    "composition", cmpst,
        "arrangement", arngm,   "vidid", vidid,     "mtype", mtype,     "vidname", vname,
        "ystart", ystart,       "yend", yend,       "kstyle", kstyle,   "loopvid", loopvid
    )
    text := FileRead(assf, "UTF-8")
    isLineSync := InStr(text, ",LineSync")
    lastIdx := 0
    for line in StrSplit(text, "`n", "`r") {
        if RegExMatch(line, "^;([^=]+)=(.*?)\s*$", &m) {
            if (m[1] = "kstyle") {
                try fields[m[1]].Text := m[2]
            } else if fields.Has(m[1]) {
                fields[m[1]].Value := m[2]
            }
        }
        if isLineSync {
            ; 行同期歌詞
            if RegExMatch(line, "^Dialogue:\s*(\d+),([^,]+),([^,]+),.*?,LineSync.*,.*\}(.*)$", &m) {
                idx := Integer(m[1])
                if (idx > 0 && idx != lastIdx) {
                    lrcTime := RegExReplace(m[3], "^0:")
                    kashi.Value .= "[" lrcTime "] " m[4] "`n"
                    lastIdx := idx
                }
            }
        } else {
            ; 通常スクロール歌詞
            if RegExMatch(line, "^Dialogue:\s*1,") {
                if pos := InStr(line, "}", false, -1)
                    kashi.Value .= SubStr(line, pos + 1) "`n"
            }
        }
    }
    if (SubStr(kashi.Value, -1) = "`n")
        kashi.Value := SubStr(kashi.Value, 1, -1)
    return True
}

AppendSongInfo(lines){  ;ass書込共通関数　曲情報追加
    lines.Push(";[Song Info]")
    info := [
        ["title",       title],     ["artist",      artst],     ["tieup",       tieup],
        ["year",        year],      ["subdir",      subdir],    ["lyrics",      lyric],
        ["composition", cmpst],     ["arrangement", arngm],     ["utaid",       utaID],
        ["vidid",       vidid],     ["mtype",       mtype],     ["vidname",     vname],
        ["kstyle",      kstyle],    ["ystart",      ystart],    ["yend",        yend],
        ["loopvid",     loopvid] ]
        for item in info {
            if (item[1] = "kstyle")
                value := item[2].Text
            else
                value := item[2].Value
            if (value != "")
                lines.Push(";" item[1] "=" value)
        }
}
LoadAssHeader(headfile){    ;ass書込共通関数 ヘッダ読込
    lines := []
    f := Map()
    in_sinfo := false
    text := FileRead(headfile, "UTF-8")
    for line in StrSplit(text, "`n", "`r") {
        if RegExMatch(line, "^;f(\d+)=(.*)", &m) {
            num := Integer(m[1])
            val := m[2]
            if kstyle.Text != "" {
                val := RegExReplace(val, ",Kanji\d+,", ",Kanji" kstyle.Text ",")
                val := RegExReplace(val, ",sInfo\d+,", ",sInfo" kstyle.Text ",")
                val := RegExReplace(val, ",sRuby\d+,", ",sRuby" kstyle.Text ",")
            }
            f[num] := val
            continue
        }
        if (SubStr(line, 1, 13) = ";[Song Info]") {
            in_sinfo := true
            AppendSongInfo(lines)
            continue
        }
        if (in_sinfo && SubStr(line, 1, 2) = ";[") {
            in_sinfo := false
            lines.Push(line)
            continue
        }
        if (in_sinfo)
            continue
        lines.Push(line)
    }
    return {
        lines: lines,
        f: f
    }
}
AppendSongDialogue(lines, f){       ;ass書込共通関数 曲情報(Dialogue)を追加
    items := [
        [1, title], [2, artst], [3, tieup], [5, lyric], [6, cmpst], [7, arngm], [8, vname]
    ]
    for item in items
        if (item[2].Value != "")
            lines.Push(f[item[1]] item[2].Value)
    if (year.Value != "")
        lines.Push(f[tieup.Value != "" ? 4 : 3] year.Value)
}
AppendNormalLyrics(lines, f11) {    ;スクロール歌詞 歌詞付与
    f11 := StrReplace(f11, "ee:ee", sec2mmss(durat.Value))   ; duration埋め込み
    kashiText := RegExReplace(kashi.Value, "(\r?\n)+$")      ; 最後の空行のみ削除
    kashiLen := StrSplit(kashiText, "`n", "`r").Length       ; 歌詞行数
    t1 := RegExMatch(ystart.Value, "^-?\d+$") ? ystart.Value : 480
    t2 := RegExMatch(yend.Value, "^-?\d+$")
        ? yend.Value - ((kashiLen - 1) * 40)
        : 200 - ((kashiLen - 1) * 40)
    for line in StrSplit(kashiText, "`n", "`r") {
        line_f11 := StrReplace(f11, "t1", t1)
        line_f11 := StrReplace(line_f11, "t2", t2)
        lines.Push(line_f11 line)
        t1 += 40
        t2 += 40
    }
}
AppendSyncLyrics(lines, f11, f12){  ;同期歌詞 歌詞付与
    kashiText := RegExReplace(kashi.Value, "(\r?\n)+$")     ;最後の空行のみ削除
    lyrics := []
    for line in StrSplit(kashiText, "`n", "`r") {
        if RegExMatch(line, "^\[(\d{2}:\d{2}\.\d{2})(\d?)\]\s*(.*)$", &m) {
            lyrics.Push({
                time: m[1],
                text: m[3]
            })
        }
    }
    nextX := 480
    artmsg := ""
    Loop lyrics.Length {
        i := A_Index
        ss := (i = 1) ? "00:00.00" : lyrics[i - 1].time
        mm := lyrics[i].time
        ee := (i = lyrics.Length) ? SecToLrcTime(durat.Value) : lyrics[i + 1].time
        e2 := (i + 1 >= lyrics.Length) ? SecToLrcTime(durat.Value) : lyrics[i + 2].time
        if StrLen(lyrics[i].text) > 40{
            artmsg := "1行の文字数が40文字を超えています"
        }
        if (lyrics[i].text = "" || (i < lyrics.Length && lyrics[i + 1].text = "")) {
            xx := 640
            nextX := 480
        } else if (StrLen(lyrics[i].text) > 25) {
            xx := 640
            nextX := 480
        } else {
            xx := nextX
            nextX := (nextX = 480) ? 800 : 480
        }
        t1_0 := Max(TimeDiffMs(ss, mm), 0)
        t1_1 := Max(t1_0 - 500, 0)
        t2_0 := Max(TimeDiffMs(mm, ee), 0)
        t2_1 := Max(t2_0 - 500, 0)
        t3_0 := Min(Max(TimeDiffMs(mm, e2), 0), 30000)
        t3_1 := Max(t3_0 - 500, 0)
        vars := Map(
            "{ii}", i,  "{ss}", ss, "{mm}", mm, "{ee}", ee, "{e2}", e2, "{xx}", xx,
            "{t1_0}", t1_0, "{t1_1}", t1_1,
            "{t2_0}", t2_0, "{t2_1}", t2_1,
            "{t3_0}", t3_0, "{t3_1}", t3_1 )
        line_f11 := f11
        line_f12 := f12
        for k, v in vars {
            line_f11 := StrReplace(line_f11, k, v)
            line_f12 := StrReplace(line_f12, k, v)
        }
        lines.Push(line_f11 lyrics[i].text)
        lines.Push(line_f12 lyrics[i].text)
    }
    if (artmsg != ""){
        MsgBox(artmsg)
    }
}
SaveLines(path, lines) {            ;ass書込共通関数 行保存
    out := ""
    for i, line in lines
        out .= (i = 1 ? "" : "`r`n") line
    FileOpen(path, "w", "UTF-8").Write(out)
}
WriteAssf(assf){                    ;ass書き込み
    ass := LoadAssHeader(asshead)
    AppendSongDialogue(ass.lines, ass.f)
    AppendNormalLyrics(ass.lines, ass.f[11])
    SaveLines(assf, ass.lines)
}
WriteSyncAssf(assf){                ;ass同期書き込み
    ass := LoadAssHeader(assshead)
    AppendSongDialogue(ass.lines, ass.f)
    AppendSyncLyrics(ass.lines, ass.f[11], ass.f[12])
    SaveLines(assf, ass.lines)
}

getdur(mp4f){
    cmd := 'bin\ffprobe.exe -v error -show_entries format=duration -of default=nw=1:nk=1 "' mp4f '"'
    o := RegExReplace(StdoutToVar(cmd), "[^\d\.]") + 0
    if (o != "")
        durat.Value := Round(o, 0)
    else
        MsgBox("曲の長さ取得に失敗")
}

LoadMp4(mp4f){                     ;mp3/mp4読込
    stat.Value := ""
    if !RegExMatch(mp4f, "i)\.(mp3|mp4)$") || !FileExist(mp4f) {
        stat.Value := "mp3/mp4ファイルがありません" , MsgBox(stat.Value)
        return
    }
    assf := RegExReplace(mp4f, "\.[^\.]+$", ".ass")
    getdur(mp4f)
    ClearsInfo()
    oFile.Value := assf
    ReadAssf(assf)
    ytimg()
    stat.Value := "ロード : " GetF(mp4f)
}

HandleDrop(guiObj, guiCtrlObj, files, x, y) {       ;ファイルドロップのハンドル
    if (guiCtrlObj = vname) {
        vname.Value := files[1]
        mtype.Value := "youtube"
        return
    }
    LoadMp4(files[1])
}

GetNeighborFile(currentFile, offset := 1) {
    SplitPath(currentFile, , &dir, , &nameNoExt)
    if !DirExist(dir)
        return ""
    current := ""
    mp4 := dir "\" nameNoExt ".mp4"
    mp3 := dir "\" nameNoExt ".mp3"
    if FileExist(mp4)
        current := mp4
    else if FileExist(mp3)
        current := mp3
    files := []
    Loop Files dir "\*"
        if (A_LoopFileExt = "mp4" || A_LoopFileExt = "mp3")
            files.Push(A_LoopFileFullPath)
    if !files.Length
        return ""
    if !current
        return offset > 0 ? files[1] : files[files.Length]
    for i, file in files {
        if (file = current) {
            target := i + offset
            return (target >= 1 && target <= files.Length) ? files[target] : ""
        }
    }
    return offset > 0 ? files[1] : files[files.Length]
}

AjastAssf(ys,ye){       ;スクロール歌詞 ass位置修正
    stat.Value := ""
    if ! btnErrCk()
        return
    if HasLineTag(){
        stat.Value := "行タグ歌詞には使えません", MsgBox(stat.Value)
        return
    }
    if (ystart.Value = "")
        ystart.Value := 480
    if (yend.Value = "")
        yend.Value := 200
    ystart.Value := ystart.Value + ys
    yend.Value   := yend.Value   + ye
    WriteAssf(oFile.Value)
}

CreateAssf(fname){      ;一括実行用assファイル作成
    SplitPath(fname, &name)
    if GetKeyState("Esc", "P"){
        stat.Value := "[ESC]中断:" name
        Sleep(-1)
        return false
    }
    assf := RegExReplace(fname, "\.[^\.]+$", ".ass")
    if FileExist(assf){
        stat.Value := "assが存在する:" name
        Sleep(-1)
        return false
    }
    LoadMp4(fname)
    SplitPath(oFile.Value, , , , &name_no_ext)
    name_no_ext := RegExReplace(name_no_ext, "[-\[\]]", " ") ; 文字列除外になるようなので
    url := FENRIR_U rep(name_no_ext)
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", url, false)
    http.Send()
    texts := http.ResponseText
    if RegExMatch(texts, "https://www\.uta-net\.com/song/(\d+)/", &m){
        utaID.Value  := m[1]
    } else {
        stat.Value := "utaIDが見つからない:" name
        Sleep(-1)
        return false
    }
    btn01clk()
    WriteAssf(oFile.Value)
    stat.Value := "正常にassを作成:" name
    Sleep(-1)
    return true
}

btn01clk(*){            ;歌詞取得
    stat.Value := ""
    if utaID.Value = "" {
        r := MsgBox("utaIDが未入力です。assファイルを削除しますか？","確認","YesNo 32")
        if (r = "Yes" and FileExist(oFile.Value)){
            assf := oFile.Value
            FileDelete(oFile.Value)
            ClearsInfo()
            oFile.Value := assf
        }
        return false
    }
    url := "https://www.uta-net.com/song/" utaID.Value "/"
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", url, false)
    http.Send()
    texts := http.ResponseText
    if InStr(texts, "404 Not Found"){
        stat.Value := url "が見つかりません" , MsgBox(stat.Value)
        return false
    }
    utaID0 := utaID.Value , ClearsInfo() , utaID.Value := utaID0
    if RegExMatch(texts, '<h2 class="ms-2 ms-md-3 kashi-title">(.+?)</h2>', &m)
        title.Value := rep(m[1])
    if RegExMatch(texts, "(?s)<p class=`"ms-2 ms-md-3 mb-0`" style='font-size:12px;'>\s*(.*?)\s*</p>", &m)
        tieup.Value := Trim(rep(m[1]))
    if RegExMatch(texts, '<span itemprop="byArtist name">(.+?)</span></a></h3>', &m)
        artst.Value := rep(m[1])
    if RegExMatch(texts, '作詞：<a [^>]*itemprop="lyricist"[^>]*>(.+?)</a>', &m)
        lyric.Value := rep(m[1])
    if RegExMatch(texts, '作曲：<a [^>]*itemprop="composer"[^>]*>(.+?)</a>', &m)
        cmpst.Value := rep(m[1])
    if RegExMatch(texts, '編曲：<a [^>]*itemprop="arranger"[^>]*>(.+?)</a>', &m)
        arngm.Value := rep(m[1])
    if RegExMatch(texts, '発売日：(\d{4}/\d{2}/\d{2})', &m)
        year.Value := StrSplit(m[1], "/")[1]
    if RegExMatch(texts, '(?s)<div[^>]*id="kashi_area"[^>]*>(.*?)</div>', &m)
        kashi.Value := Trim(StrReplace(m[1], "<br />", "`n"))
    if RegExMatch(oFile.Value, "\[([^\[\]]+)\]\.ass$", &m)
        vidid.Value := m[1]
    stat.Value := "取得しました : " title.Value
}
btnErrCk(){             ;エラーチェック
    if title.Value = "" {
        MsgBox("titleなし")
        return false
    } else if artst.Value = "" {
        MsgBox("artistなし")
        return false
    } else if durat.Value = "" {
        MsgBox("曲の長さ なし")
        return false
    } else if oFile.Value = "" {
        MsgBox("出力ファイル なし")
        return false
    }
    if mtype.Value = "" {
        mtype.Value := "mv"
    }
    if InStr(vname.Value, "\") && InStr(vname.Value, ".") {
        SplitPath(vname.Value, , , , &nameNoExt)
        vname.Value := nameNoExt
    }
    return true
}
btn02clk(*){
    stat.Value := ""
    if ! btnErrCk()
        return
    ytimg()
    if HasLineTag()
        WriteSyncAssf(oFile.Value)
    else
        WriteAssf(oFile.Value)
    mp4f := RegExReplace(oFile.Value, "\.ass$", ".mp4")
    mp3f := RegExReplace(oFile.Value, "\.ass$", ".mp3")
    if FileExist(mp4f) {
        Run(mpcPath ' /play /sub "' oFile.Value '" "' mp4f '"')
        stat.Value := "再生 : " GetF(mp4f)
    } else if FileExist(mp3f) {
        Run(mpcPath ' /play /sub "' oFile.Value '" /dub "' mp3f '" "' bgv '"')
        stat.Value := "再生 : " GetF(mp3f)
    } else {
        stat.Value := "ファイルが存在しません" , MsgBox(stat.Value)
    }
}
btn03clk(*){
    stat.Value := ""
    SplitPath(oFile.Value, , , , &name_no_ext)
    name_no_ext := RegExReplace(name_no_ext, "[-\[\]]", " ") ; 文字列除外になるようなので
    Run(FENRIR_U rep(name_no_ext))
}
btn04clk(*){
    LoadMp4(GetNeighborFile(oFile.Value,-1))
}
btn05clk(*){
    LoadMp4(GetNeighborFile(oFile.Value,1))
}
btn06clk(*){
    stat.Value := ""
    SplitPath(oFile.Value, , &dir)
    if (dir = ""){
        MsgBox("出力ファイルを指定してください")
        return
    }
    msg := "フォルダ " dir " 配下のmp3/mp4ファイルに一括でassを付与します。`nよろしいですか？"
    r := MsgBox(msg,"一括","YesNo 32")
    if (r = "No"){
        return
    }
    Loop Files dir "\*.mp3","R" {
        CreateAssf(A_LoopFileFullPath)
    }
    Loop Files dir "\*.mp4","R" {
        CreateAssf(A_LoopFileFullPath)
        filecnt++
    }
    stat.Value := "一括処理が終了しました" , MsgBox(stat.Value)
}
btn07clk(*){
    stat.Value := ""
    if (title.Value = "" or artst.Value = ""){
        stat.Value := "titleとartistを入力してください" , MsgBox(stat.Value)
        return False
    }
    btn07.Enabled := false
    stat.Value := "LRCLIB の情報取得中..."
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    reqstr := title.Value " " artst.Value
    reqstr := RegExReplace(reqstr, "\([^)]*\)", " ")    ; (～) をスペースに置換
    reqstr := RegExReplace(reqstr, "[!$%&]", " ")       ; ! $ % & をスペースに置換
    reqstr := RegExReplace(reqstr, "\s+", " ")          ; スペースを1つにまとめる
    reqstr := Trim(reqstr)
    reqstr := UriEncode(reqstr)
    url := "https://lrclib.net/api/search?q=" reqstr
    try {
        http.Open("GET", url, false)
        http.Send()
    } catch Error as e {
        stat.Value := "通信エラー：" e.Message , MsgBox(stat.Value)
        btn07.Enabled := true
        return False
    }
    texts := ResponseBodyToText(http.ResponseBody)
    if RegExMatch(texts, '"syncedLyrics"\s*:\s*"((?:\\.|[^"])*)"', &m) {
        txt := m[1]
        txt := StrReplace(txt, "\n", "`n")
        txt := StrReplace(txt, "\r", "")
        txt := StrReplace(txt, "\\", "\")
        txt := StrReplace(txt, '\"', '"')
        kashi.Value := txt
    }
    btn07.Enabled := true
    stat.Value := "LRCLIB の情報取得:" http.Status " " http.StatusText
}
btn08clk(*){
    stat.Value := ""
    Run("https://lrclib.net/search/" rep(title.Value " " artst.Value))
}
btn09clk(*){
    stat.Value := ""
    WriteSyncAssf(oFile.Value)
    mp4f := RegExReplace(oFile.Value, "\.ass$", ".mp4")
    mp3f := RegExReplace(oFile.Value, "\.ass$", ".mp3")
    if FileExist(mp4f) {
        Run(mpcPath ' /play /sub "' oFile.Value '" "' mp4f '"')
    } else if FileExist(mp3f) {
        Run(mpcPath ' /play /sub "' oFile.Value '" /dub "' mp3f '" "' bgv '"')
    } else {
        MsgBox("ファイルが存在しない")
    }
}
btn10clk(*){
    stat.Value := ""
    if ! btnErrCk()
        return
    ystart.Value := ""
    yend.Value   := ""
    WriteAssf(oFile.Value)
}
btn11clk(*){
    AjastAssf(-120,0)
}
btn12clk(*){
    AjastAssf(120,0)
}
btn13clk(*){
    AjastAssf(0,-120)
}
btn14clk(*){
    AjastAssf(0,120)
}
btn15clk(*){        ; 削除
    if (oFile.Value) {
        msg := oFile.Value "`nを削除します。よろしいですか？"
        if (MsgBox(msg, "削除", "OKCancel 32") = "Cancel")
            return
        nextFile := GetNeighborFile(oFile.Value, 1)
        try FileDelete(oFile.Value)
        try FileDelete(RegExReplace(oFile.Value, "\.[^.]+$", ".mp4"))
        try FileDelete(RegExReplace(oFile.Value, "\.[^.]+$", ".mp3"))
        Sleep(1)
        LoadMp4(nextFile)
    }
}

btn16clk(*){        ; 配置
    subdir0 := subdir.Value
    if subdir0 = "" {
        try {
            y := Integer(year.Value)
            if (y <= 1989)
                subdir0 := "1989以前"
            else if (y <= 1999)
                subdir0 := "1990-1999"
            else if (y <= 2009)
                subdir0 := "2000-2009"
            else if (y <= 2019)
                subdir0 := "2010-2019"
            else
                subdir0 := year.Value
        } catch {
            subdir0 := year.Value
        }
    }
    mtype0 := mtype.Value
    if mtype0 != "" {
        mtype0 := " " mtype.Value
    }
    if tieup.Value != "" {
        fname := "[" tieup.Value "]" title.Value "／" artst.Value "" mtype0
    } else {
        fname := title.Value "／" artst.Value "" mtype0
    }
    fassf :=  oFile.Value
    fmp4f :=  RegExReplace(oFile.Value, "\.[^.]+$", ".mp4")
    fmp3f :=  RegExReplace(oFile.Value, "\.[^.]+$", ".mp3")
    tassf := baseDir "\" subdir0 "\" fname ".ass"
    tmp4f := baseDir "\" subdir0 "\" fname ".mp4"
    tmp3f := baseDir "\" subdir0 "\" fname ".mp3"
    msg := tassf "`nに配置します。よろしいですか？"
    if FileExist(tassf)
        msg := msg "`n(移動先にassファイルが存在します)"
    if FileExist(tmp4f)
        msg := msg "`n(移動先にmp4ファイルが存在します)"
    if FileExist(tmp3f)
        msg := msg "`n(移動先にmp3ファイルが存在します)"
    if (MsgBox(msg, "配置", "OKCancel 32") = "Cancel")
        return
    nextFile := GetNeighborFile(oFile.Value, 1)
    try FileMove(fassf, tassf, 1)
    try FileMove(fmp4f, tmp4f, 1)
    try FileMove(fmp3f, tmp3f, 1)
    Sleep(1)
    LoadMp4(nextFile)
}
ytdlp(loopf,loopvid){
    cmd := 'bin\yt-dlp.exe'
        . ' -f "bv[height<=1080]+ba"'
        . ' --merge-output-format mp4'
        . ' -N 1'
        . ' -o "' loopf '"'
        . ' -- "' loopvid '"'
    log := A_Temp "\yt-dlp.log"
    cmd2 := A_ComSpec ' /c "' cmd ' > "' log '" 2>&1"'
    exitCode := RunWait(cmd2, , "Hide")
    if (exitCode != 0) {
        MsgBox "ytdlpがエラーで終了しました。終了コード: " exitCode
        Run('notepad.exe "' log '"')
        return false
    } else if !FileExist(loopf) {
        MsgBox "ファイルが作成されませんでした`n" loopf
        return false
    }
    return true
}

loopedit(vidf,audf,fname){
    cmd := 'bin\ffmpeg.exe -y -stream_loop 5'
        . ' -i "' vidf '"'
        . ' -i "' audf '"'
        . ' -shortest -map 0:v:0 -map 1:a'
        . ' -c:v copy -c:a copy'
        . ' "' fname '"'
    exitCode := RunWait(cmd, , "Hide")
    if (exitCode != 0) {
        MsgBox "ffmpegがエラーで終了しました。終了コード: " exitCode
        return false
    } else if !FileExist(fname) {
        MsgBox "ファイルが作成されませんでした`n" fname
        return false
    }
    return true
}

ytimg() {
    if loopvid.Value != "" {
        videoid := loopvid.Value
    } else if vidid.Value != "" {
        videoid := vidid.Value
    } else {
        thumb.Value := noimg
        return false
    }
    url := "https://i.ytimg.com/vi/" videoid "/hqdefault.jpg"
    try {
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Open("GET", url, false)
        http.Send()
        if (http.Status != 200)
            return false
        stream := ComObject("ADODB.Stream")
        stream.Type := 1
        stream.Open()
        stream.Write(http.ResponseBody)
        stream.SaveToFile(ytthumb, 2)
        stream.Close()
        thumb.Value := ytthumb
        return true
    } catch {
        thumb.Value := noimg
        return false
    }
}

btn21clk(*){
    mp4f := RegExReplace(oFile.Value, "\.[^.]+$", ".mp4")
    if (vidid.Value && !loopvid.Value){
        stat.Value := "動画差し替え..." vname.Value
        loopf := RegExReplace(oFile.Value, "\.[^.]+$", "_loopvid.mp4")
        if (MsgBox("動画を差し替えます`n" vidid.Value, "確認", "OKCancel") != "OK")
            return
        if !ytdlp(loopf,vidid.Value)
            return
        try FileDelete(mp4f)
        FileMove(loopf,mp4f)
        getdur(mp4f)
    } else if (!vname.Value && loopvid.Value) {
        stat.Value := "loop動画ダウンロード..." vname.Value
        loopf := RegExReplace(oFile.Value, "\.[^.]+$", "_loopvid.mp4")
        fname := RegExReplace(mp4f, "\.mp4$", "_loopedit.mp4")
        if (MsgBox("loop動画をダウンロードします`n" loopvid.Value, "確認", "OKCancel") != "OK")
            return
        if !ytdlp(loopf,loopvid.Value)
            return
        if !loopedit(loopf,mp4f,fname)
            return
        try FileDelete(mp4f)
        try FileDelete(loopf)
        FileMove(fname,mp4f)
    } else if (vname.Value && FileExist(vname.Value)) {
        stat.Value := "DL済みファイルから作成..." vname.Value
        if !FileExist(mp4f) {
            MsgBox "動画が存在しません`n" mp4f
            return
        }
        vidf := vname.Value
        fname := RegExReplace(mp4f, "\.mp4$", "_loopedit.mp4")
        if (MsgBox("loop編集します`n" mp4f, "確認", "OKCancel") != "OK")
            return
        if !loopedit(vidf,mp4f,fname)
            return
        try FileDelete(mp4f)
        FileMove(fname,mp4f)
    } else {
        MsgBox "動画の指定がありません"
        return
    }
    stat.Value := "loop編集 完了"
    btn02clk()
}
btn22clk(*){
    stat.Value := ""
    Run(FENRIR_Y rep(title.Value " " artst.Value))    
}
; メイン
; ボタン=基準 Edit=+2 Text=+4 次行=+25
mpcPath := GetMPCPath()
baseDir := GetBaseDir()
myGui := Gui()
myGui.Title := "もちからuta-netスクロール歌詞付与 v0.5"
;y=0
myGui.AddText("x10 y4", "出力ファイル："),   oFile := myGui.AddEdit("x75 y2 w535")
;y=25
myGui.AddText("x10 y29", "曲の長さ："),      durat := myGui.AddEdit("x75 y27 w50")
myGui.AddText("x135 y29" , "mp3/mp4をドラッグ＆ドロップ")
btn04 := myGui.AddButton("x300 y25 w45", "▲前")
btn05 := myGui.AddButton("x350 y25 w45", "次▼")
btn06 := myGui.AddButton("x400 y25 w45", "一括")
btn06.Enabled := false
;y=50
myGui.AddLink("x10 y54" , '<a href="https://www.uta-net.com/">uta-net ID</a>：')
utaID := myGui.AddEdit("x75 y52 w55")
btn01 := myGui.AddButton("x135 y50 w60", "歌詞取得")
btn03 := myGui.AddButton("x200 y50 w75", "ファイル名検索")
btn15 := myGui.AddButton("x350 y50 w45", "削除")
btn16 := myGui.AddButton("x400 y50 w45", "配置")
;y=75
myGui.AddLink("x15 y79" , '<a href="https://lrclib.net/">LRCLIB</a>：')
btn07 := myGui.AddButton("x75 y75 w55", "LRC取得")
btn08 := myGui.AddButton("x135 y75 w60", "曲名検索")
;y=100 曲情報
myGui.AddText("x10  y102",  "[曲情報]")
;y=115
myGui.AddText("x10  y119", "title："),      title := myGui.AddEdit("x60  y117 w210")
;y=140
myGui.AddText("x10  y144", "artist："),     artst := myGui.AddEdit("x60  y142 w210")
;y=165
myGui.AddText("x10  y169", "tieup："),      tieup := myGui.AddEdit("x60  y167 w210")
;y=190
myGui.AddText("x10  y194", "year："),       year  := myGui.AddEdit("x60  y192 w50")
myGui.AddText("x120 y194", "subdir："),     subdir:= myGui.AddEdit("x160 y192 w110")
;y=215
myGui.AddText("x10 y219", "作詞："),        lyric := myGui.AddEdit("x60  y217 w210")
;y=240
myGui.AddText("x10 y244", "作曲："),        cmpst := myGui.AddEdit("x60  y242 w210")
;y=265
myGui.AddText("x10 y269", "編曲："),        arngm := myGui.AddEdit("x60  y267 w210")
;y=290 動画情報
myGui.AddText("x10 y292",  "[動画情報]")
;y=305
myGui.AddText("x10  y309", "videoID:"),    vidid  := myGui.AddEdit("x60  y307 w85")
myGui.AddText("x147 y309", "loopvid:"),    loopvid:= myGui.AddEdit("x185 y307 w85")
;y=330
myGui.AddText("x10  y334", "動画type:"),   mtype  := myGui.AddEdit("x60  y332 w75")
btn21 := myGui.AddButton("x140 y330 w60", "loop編集")
btn22 := myGui.AddButton("x200 y330 w60", "動画検索")
;y=355
myGui.AddText("x10  y359", "動画名:"),     vname  := myGui.AddEdit("x60  y357 w210")
;y=380 動画情報
myGui.AddText("x10  y382", "[歌詞位置修正]")
;y=395
btn02 := myGui.AddButton("x10  y400 w45", "ass作成＆再生")
btn11 := myGui.AddButton("x60  y395 w35", "始㊤")
btn12 := myGui.AddButton("x100 y395 w35", "始㊦")
btn13 := myGui.AddButton("x140 y395 w35", "終㊤")
btn14 := myGui.AddButton("x180 y395 w35", "終㊦")
btn10 := myGui.AddButton("x220 y395 w40", "reset")
myGui.AddText("x65  y429", "開始座標："),   ystart  := myGui.AddEdit("x120 y425 w40")
myGui.AddText("x165 y429", "終了座標："),   yend    := myGui.AddEdit("x220 y425 w40")
;y=50
thumb := myGui.AddPicture("x450 y25 w160 h90", noimg)
;y=100 歌詞
myGui.AddText("x280 y102", "[歌詞]")
myGui.AddText("x320 y102", "歌詞style："),  
kstyle:= myGui.AddDropDownList("x380 y100 w30", ["", "1", "2", "3", "4", "5", "6", "7", "8"])
;y=120
kashi := myGui.AddEdit("x280 y120 w330 h300 +Multi +VScroll +HScroll")
stat  := myGui.AddText("x280 y427 w330 h40 +Wrap", "[ステータス]")

btn01.OnEvent("Click", btn01clk)
btn02.OnEvent("Click", btn02clk)
btn03.OnEvent("Click", btn03clk)
btn04.OnEvent("Click", btn04clk)
btn05.OnEvent("Click", btn05clk)
btn06.OnEvent("Click", btn06clk)
btn07.OnEvent("Click", btn07clk)
btn08.OnEvent("Click", btn08clk)

btn10.OnEvent("Click", btn10clk)
btn11.OnEvent("Click", btn11clk)
btn12.OnEvent("Click", btn12clk)
btn13.OnEvent("Click", btn13clk)
btn14.OnEvent("Click", btn14clk)
btn15.OnEvent("Click", btn15clk)
btn16.OnEvent("Click", btn16clk)
btn21.OnEvent("Click", btn21clk)
btn22.OnEvent("Click", btn22clk)
myGui.OnEvent("DropFiles", HandleDrop)

myGui.Show("w620 h460")

if (A_Args.Length >= 1){
    if (A_Args.Length >= 2 && A_Args[2] != "")
        asshead := A_Args[2]
    LoadMp4(A_Args[1])
}                                                                                                                                                                                                                                                                                                                                                                                                                      