' Haruki1234
' atd_jyanken.vbs version 1.0
' 2021.10.15
' https://github.com/haruki1234/neknaj

function CSCRIPT()
    CSCRIPT_EXE="cscript.exe"
    if not LCase(Right(WScript.FullName, Len(CSCRIPT_EXE))) = CSCRIPT_EXE then
        CreateObject("Wscript.Shell").run "cscript //nologo "+WScript.ScriptFullName
        WScript.Quit(1)
    end if
end function
CSCRIPT()
Randomize
function Input()
    Input = WScript.StdIn.ReadLine()
end function
function Print(text)
    WScript.StdOut.Write(text)
end function
function PrintLine(text)
    WScript.StdOut.WriteLine(text)
end function
function Random(min , max)
    Random = Int((max-min+1)*Rnd()+min)    
end function
function RandomArray(array)
    RandomArray = array(Random(Lbound(array),UBound(array)))    
end function

janh = Array("グー","チョキ","パー")
janha = Array("g","t","p")
res = Array("勝って","負けて")
function  chand(ahand,bhand,mode)
    if StrComp(ahand, janh(0)) = 0 then
        if StrComp(bhand, janh(0)) = 0 then
            if StrComp(mode, res(0)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            elseif StrComp(mode, res(1)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            else
                chand="エラー"
            end if
        elseif StrComp(bhand, janh(1)) = 0 then
            if StrComp(mode, res(0)) = 0 Then
                chand="正解"
                true_ans = true_ans+1
            elseif StrComp(mode, res(1)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            else
                chand="エラー"
            end if
        elseif StrComp(bhand, janh(2)) = 0 then
            if StrComp(mode, res(0)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            elseif StrComp(mode, res(1)) = 0 Then
                chand="正解"
                true_ans = true_ans+1
            else
                chand="エラー"
            end if
        else
            chand="エラー"
        end if
    elseif StrComp(ahand, janh(1)) = 0 then
        if StrComp(bhand, janh(0)) = 0 then
            if StrComp(mode, res(0)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            elseif StrComp(mode, res(1)) = 0 Then
                chand="正解"
                true_ans = true_ans+1
            else
                chand="エラー"
            end if
        elseif StrComp(bhand, janh(1)) = 0 then
            if StrComp(mode, res(0)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            elseif StrComp(mode, res(1)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            else
                chand="エラー"
            end if
        elseif StrComp(bhand, janh(2)) = 0 then
            if StrComp(mode, res(0)) = 0 Then
                chand="正解"
                true_ans = true_ans+1
            elseif StrComp(mode, res(1)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            else
                chand="エラー"
            end if
        else
            chand="エラー"
        end if
    elseif StrComp(ahand, janh(2)) = 0 then
        if StrComp(bhand, janh(0)) = 0 then
            if StrComp(mode, res(0)) = 0 Then
                chand="正解"
                true_ans = true_ans+1
            elseif StrComp(mode, res(1)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            else
                chand="エラー"
            end if
        elseif StrComp(bhand, janh(1)) = 0 then
            if StrComp(mode, res(0)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            elseif StrComp(mode, res(1)) = 0 Then
                chand="正解"
                true_ans = true_ans+1
            else
                chand="エラー"
            end if
        elseif StrComp(bhand, janh(2)) = 0 then
            if StrComp(mode, res(0)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            elseif StrComp(mode, res(1)) = 0 Then
                chand="間違い"
                false_ans = false_ans+1
            else
                chand="エラー"
            end if
        else
            chand="エラー"
        end if
    else
        chand="エラー"
    end if
end function
function do_atodasi(qnum)
    PrintLine("["+CStr(qnum)+"/"+CStr(ques_all)+"]")
    WScript.Sleep(500)
    Print("最初はグー")
    WScript.Sleep(500)
    PrintLine(" じゃんけん")
    WScript.Sleep(200)
    bhand = RandomArray(janh)
    PrintLine(bhand)
    startt = Timer
    mode = RandomArray(res)
    Print(mode+": ")
    ahand = replace_h(input())
    alltime = alltime + Timer-startt
    PrintLine(chand(ahand,bhand,mode))
    PrintLine("")
end function
function replace_h(hand)
    hand = Replace(hand,janha(0),janh(0))
    hand = Replace(hand,janha(1),janh(1))
    hand = Replace(hand,janha(2),janh(2))
    replace_h = hand
end function

PrintLine("Janken vbs")
PrintLine("後だしじゃんけん")
PrintLine("")

do
    alltime = 0
    allcunt = 0
    true_ans = 0
    false_ans = 0

    ques_all = 10

    for qnum = 1 to ques_all
        do_atodasi(qnum)
    next
    all_time = CStr(Int(alltime*10)/10)
    allqc = CStr(true_ans+false_ans)
    tfs = CStr(Int(true_ans/(true_ans+false_ans)*1000)/10)
    score = CStr(Int((true_ans*10/(true_ans+false_ans+1)/((alltime+100)/1000))*10)/10)
    PrintLine("お疲れ様です")
    PrintLine("")
    PrintLine("合計時間: "+all_time+"秒")
    PrintLine("問題数　: "+allqc+"問")
    PrintLine("正答率　: "+tfs+"%")
    PrintLine("スコア: "+score)
    PrintLine("")
    WScript.Sleep(1000)
    Print("もう一度しますか[y/N]: ")
    next_ = Input()
    if not ((StrComp(next_, "y")=0) or (StrComp(next_, "yes")=0) or (StrComp(next_, "Y")=0)) then
        exit do
    end if
loop