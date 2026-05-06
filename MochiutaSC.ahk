; MochiutaSC.ahk
; もちからuta-netスクロール歌詞付与 v0.1
; https://ahkwiki.net/

#Requires AutoHotkey v2.0
SetWorkingDir(A_ScriptDir)

ASSHEAD := "MochiutaSC_header.ass"

; 標準出力の値を変数に
StdoutToVar(cmd) {
    shell := ComObject("WScript.Shell")
    exec := shell.Exec(cmd)
    return exec.StdOut.ReadAll()
}
;htmlデコード（簡易）
HtmlDecode(s) {
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
;文字列置き換え（簡易）
rep(s) {
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
; 秒 → mm:ss に変換
sec2mmss(duration) {
    sec := duration + 0
    mm := Floor(sec / 60)
    ss := Mod(sec, 60)
    return Format("{:02}:{:02}", mm, ss)
}
;曲情報消す
ClearsInfo(){
    title.Value := "" , artst.Value := "" , tieup.Value := "" , year.Value  := ""
    lyric.Value := "" , cmpst.Value := "" , arngm.Value := "" , kashi.Value := ""
    utaID.Value := "" , vidid.Value := "" , mtype.Value := "" , vname.Value := ""
    ystart.Value := "" , yend.Value := "" , kstyle.Value := ""
}
;MPC-BEのパスを取得
GetMPCPath() {
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
; ass読込
ReadAssf(assf){
    if !FileExist(assf) { 
        Return False
    }
    ClearsInfo()
    for line in StrSplit(FileRead(assf, "UTF-8"), "`n", "`r") {
        if RegExMatch(line, "^;utaid=(\d+)", &m) {
            utaID.Value := m[1]
        } else if RegExMatch(line, "^;title=(.+?)\s*$", &m) {
            title.Value := m[1]
        } else if RegExMatch(line, "^;artist=(.+?)\s*$", &m) {
            artst.Value := m[1]
        } else if RegExMatch(line, "^;tieup=(.+?)\s*$", &m) {
            tieup.Value := m[1]
        } else if RegExMatch(line, "^;year=(.+?)\s*$", &m) {
            year.Value := m[1]
        } else if RegExMatch(line, "^;lyrics=(.+?)\s*$", &m) {
            lyric.Value := m[1]
        } else if RegExMatch(line, "^;composition=(.+?)\s*$", &m) {
            cmpst.Value := m[1]
        } else if RegExMatch(line, "^;arrangement=(.+?)\s*$", &m) {
            arngm.Value := m[1]
        } else if RegExMatch(line, "^;vidid=(.+?)\s*$", &m) {
            vidid.Value := m[1]
        } else if RegExMatch(line, "^;mtype=(.+?)\s*$", &m) {
            mtype.Value := m[1]
        } else if RegExMatch(line, "^;vidname=(.+?)\s*$", &m) {
            vname.Value := m[1]
        } else if RegExMatch(line, "^;ystart=(.+?)\s*$", &m) {
            ystart.Value := m[1]
        } else if RegExMatch(line, "^;yend=(.+?)\s*$", &m) {
            yend.Value := m[1]
        } else if RegExMatch(line, "^;kstyle=(.+?)\s*$", &m) {
            kstyle.Value := m[1]
        } else if RegExMatch(line, "^Dialogue:\s*1,") {
            pos := InStr(line, "}", false, -1)  ; 後ろから検索
            if pos {
                text := SubStr(line, pos + 1)
                kashi.Value .= text . "`n"
            }
        }
    }
    if SubStr(kashi.Value, -1) = "`n" {
        kashi.Value := SubStr(kashi.Value, 1, -1)
    }
}
; ass書き込み
WriteAssf(assf) {
    lines := []
    in_sinfo := false
    f01 := "", f02 := "", f03 := "", f04 := "", f05 := ""
    f06 := "", f07 := "", f08 := "", f11 := ""
    text := FileRead(ASSHEAD, "UTF-8")
    kstyle0 := (kstyle.Value is number) ? kstyle.Value + 0 : 0
    for line in StrSplit(text, "`n", "`r") {
        if RegExMatch(line, "^;f(\d+?)=(.*)", &m) {     ;ASSヘッダからDialogue定義行取得
            num := Integer(m[1])
            val := m[2]
            if ( kstyle0 >= 1 && kstyle0 <= 8) {   ;スタイル書き換え
                val := RegExReplace(val, ",Kanji\d+,", ",Kanji" kstyle0 ",")
                val := RegExReplace(val, ",sInfo\d+,", ",sInfo" kstyle0 ",")
                val := RegExReplace(val, ",sRuby\d+,", ",Kanji" kstyle0 ",")
            }
            switch num {
                case 1:  f01 := val
                case 2:  f02 := val
                case 3:  f03 := val
                case 4:  f04 := val
                case 5:  f05 := val
                case 6:  f06 := val
                case 7:  f07 := val
                case 8:  f08 := val
                case 11: f11 := val
            }
        }
        if SubStr(line, 1, 13) = ";[Song Info]" {       ;Songinfo挿入
            in_sinfo := true
            lines.Push(";[Song Info]")
            if (title.Value != "") 
                lines.Push(";title=" title.Value)
            if (artst.Value != "") 
                lines.Push(";artist=" artst.Value)
            if (tieup.Value != "") 
                lines.Push(";tieup=" tieup.Value)
            if (year.Value != "")  
                lines.Push(";year="  year.Value)
            if (lyric.Value != "") 
                lines.Push(";lyrics=" lyric.Value)
            if (cmpst.Value != "") 
                lines.Push(";composition=" cmpst.Value)
            if (arngm.Value != "") 
                lines.Push(";arrangement=" arngm.Value)
            if (utaID.Value != "") 
                lines.Push(";utaid=" utaID.Value)
            if (vidid.Value != "") 
                lines.Push(";vidid=" vidid.Value)
            if (mtype.Value != "") 
                lines.Push(";mtype=" mtype.Value)
            if (vname.Value != "") 
                lines.Push(";vidname=" vname.Value)
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
    ; Dialogue部分記入
    if (title.Value != "") 
        lines.Push(f01 title.Value)
    if (artst.Value != "") 
        lines.Push(f02 artst.Value)
    if (tieup.Value != "") 
        lines.Push(f03 tieup.Value)
    if (year.Value != "") {
        if (tieup.Value != "") 
            lines.Push(f04 year.Value)
        else
            lines.Push(f03 year.Value)
    }  
    if (lyric.Value != "") 
        lines.Push(f05 lyric.Value)
    if (cmpst.Value != "") 
        lines.Push(f06 cmpst.Value)
    if (arngm.Value != "") 
        lines.Push(f07 arngm.Value)
    if (vname.Value != "") 
        lines.Push(f08 vname.Value)
    f11 := StrReplace(f11, "ee:ee", sec2mmss(durat.Value))   ;duration埋め込み
    kashiText := RegExReplace(kashi.Value, "(\r?\n)+$")     ;最後の空行のみ削除
    kashiLen := StrSplit(kashiText, "`n", "`r").Length      ;歌詞行数
    t1 := 480
    if RegExMatch(ystart.Value, "^-?\d+$")
        t1 := ystart.Value
    t2 := 200 - ((kashiLen - 1) * 40)
    if RegExMatch(yend.Value, "^-?\d+$")
        t2 := yend.Value  - ((kashiLen - 1) * 40)
    for line in StrSplit(kashiText, "`n", "`r") {
        f11a := StrReplace(f11 , "t1", t1)
        f11a := StrReplace(f11a, "t2", t2)
        lines.Push(f11a line)
        t1 := t1 + 40
        t2 := t2 + 40
    }

    out := ""
    for i, l in lines {
        out .= (i > 1 ? "`r`n" : "") l
    }
    file := FileOpen(assf, "w", "UTF-8")
    file.Write(out)
    file.Close()
}

; ドロップイベント
HandleDrop(guiObj, guiCtrlObj, files, x, y) {
    if !RegExMatch(files[1], "\.mp4$"){
        MsgBox("mp4ファイルをドロップしてください")
        return
    }
    mp4f := files[1]
    assf := RegExReplace(files[1], "\.[^\.]+$", ".ass")
    cmd := 'bin\ffprobe -v error -show_entries format=duration -of default=nw=1:nk=1 "' mp4f '"'
    o := RegExReplace(StdoutToVar(cmd), "[^\d\.]") + 0
    if (o != "")
        durat.Value := Round(o, 0)
    else
        MsgBox("曲の長さ取得に失敗")
    oFile.Value := assf
    ReadAssf(assf)
}
;歌詞取得
btn01clk(*){
    if utaID.Value = "" {
        MsgBox("utaIDが未入力です")
        return false
    }
    url := "https://www.uta-net.com/song/" utaID.Value "/"
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", url, false)
    http.Send()
    texts := http.ResponseText
    if InStr(texts, "404 Not Found"){
        MsgBox( url "が見つかりません")
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
}
;ass作成
btnErrCk(){
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
    if mtype.Value = ""
        mtype.Value := "mv"
    return true
}
btn02clk(*){
    if ! btnErrCk()
        return
    WriteAssf(oFile.Value)
    mp4f := RegExReplace(oFile.Value, "\.ass$", ".mp4")
    Run(mpcPath ' /play "' mp4f '"')
}
;ass位置修正
AjastAssf(ys,ye){
    if ! btnErrCk()
        return
    if (ystart.Value = "")
        ystart.Value := 480
    if (yend.Value = "")
        yend.Value := 200
    ystart.Value := ystart.Value + ys
    yend.Value   := yend.Value   + ye
    WriteAssf(oFile.Value)
}
btn10clk(*){
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
; メイン
mpcPath := GetMPCPath()
myGui := Gui()
myGui.Title := "もちからuta-netスクロール歌詞付与 v0.1"
myGui.AddText("x5 y7" , "uta-net ID："),    utaID := myGui.AddEdit("x80 y5 w50")
btn01 := myGui.AddButton("x140 y3 w60", "歌詞取得")
myGui.AddText("x220 y7" , "uta-net URL：")
myGui.AddLink("x300 y7", '<a href="https://www.uta-net.com/">https://www.uta-net.com/</a>')
myGui.AddText("x5 y32", "曲の長さ："),      durat := myGui.AddEdit("x80 y30 w50")
myGui.AddText("x140 y32" , "mp4ファイルをここにドロップすると曲の長さ、出力ファイルを取得します")
myGui.AddText("x5 y57", "出力ファイル："),  oFile := myGui.AddEdit("x80 y55 w500")
;SongInfo表示
myGui.AddText("x10 y90", "[曲情報]")
myGui.AddText("x10 y112", "title："),       title := myGui.AddEdit("x60 y110 w200")
myGui.AddText("x10 y137", "artist："),      artst := myGui.AddEdit("x60 y135 w200")
myGui.AddText("x10 y162", "tieup："),       tieup := myGui.AddEdit("x60 y160 w200")
myGui.AddText("x10 y187", "year："),        year  := myGui.AddEdit("x60 y185 w50")

myGui.AddText("x10 y217", "作詞："),        lyric := myGui.AddEdit("x60 y215 w200")
myGui.AddText("x10 y242", "作曲："),        cmpst := myGui.AddEdit("x60 y240 w200")
myGui.AddText("x10 y267", "編曲："),        arngm := myGui.AddEdit("x60 y265 w200")

myGui.AddText("x10 y297", "videoID："),     vidid := myGui.AddEdit("x70 y295 w60")
myGui.AddText("x170 y297", "歌詞style："),  
kstyle:= myGui.AddDropDownList("x230 y295 w30", ["", "1", "2", "3", "4", "5", "6", "7", "8"])
myGui.AddText("x10 y322", "動画type："),    mtype := myGui.AddEdit("x70 y320 w100")
myGui.AddText("x10 y347", "動画名："),      vname := myGui.AddEdit("x70 y345 w190")
myGui.AddText("x10  y372", "開始座標："),   ystart:= myGui.AddEdit("x70  y370 w60")
myGui.AddText("x140 y372", "終了座標："),   yend  := myGui.AddEdit("x200 y370 w60")

myGui.AddText("x280 y92", "[歌詞]")
kashi := myGui.AddEdit("x280 y110 w320 h330 +Multi +VScroll +HScroll")

btn02 := myGui.AddButton("x15  y405 w50", "ass作成 ＆再生")
myGui.AddText("x85 y400", "[スクロール歌詞 位置修正]")
btn10 := myGui.AddButton("x220 y395 w40", "reset")
btn11 := myGui.AddButton("x85  y420 w40", "始㊤")
btn12 := myGui.AddButton("x130 y420 w40", "始㊦")
btn13 := myGui.AddButton("x175 y420 w40", "終㊤")
btn14 := myGui.AddButton("x220 y420 w40", "終㊦")

btn01.OnEvent("Click", btn01clk)
btn02.OnEvent("Click", btn02clk)
btn10.OnEvent("Click", btn10clk)
btn11.OnEvent("Click", btn11clk)
btn12.OnEvent("Click", btn12clk)
btn13.OnEvent("Click", btn13clk)
btn14.OnEvent("Click", btn14clk)
myGui.OnEvent("DropFiles", HandleDrop)

myGui.Show("w620 h460")
