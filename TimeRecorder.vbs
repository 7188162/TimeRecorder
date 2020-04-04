Option Explicit

Dim tab, _
    q, _
    theDate, _
    startTime, _
    endTime, _
    work, _
    theContents, _
    strTitle, _
    strLogFile, _
    theUrl

strTitle = "TimeRecorder"
Const DEFAULT_ANSWER = ""
Const OVERWORK_START = "17:15"

'---------------------------------------------------------------------------
theUrl = ""
strLogFile = "c:\Users\user\Documents\log.txt"
'---------------------------------------------------------------------------

'========================================================================
tab = chr(9)
q = chr(34)

theDate = InputBox("���t����͂��Ă��������F", strTitle, Date())

startTime = InputBox("�J�n��������͂��Ă��������F" & vbCrLf & _
"�i�ʏ�Ɩ����F" & OVERWORK_START & "�A���̑��F�ύX���Ԃɂ��j", strTitle, OVERWORK_START)

endTime = InputBox("�I����������͂��Ă��������F", strTitle, FormatDateTime(Now(), vbShortTime))
work = InputBox("�Ɩ����e����͂��Ă��������F", strTitle, DEFAULT_ANSWER)

Dim res
res = MsgBox("�Ζ���: " & theDate & vbCrLf & _
"�J�n����: " & starttime & vbCrLf &_
"�I������: " & endTime & vbCrLf &_
"�Ɩ����e: " & work & vbCrLf & _
"�œo�^���܂��B��낵���ł����H", vbYesNo + vbQuestion, strTitle)

if res = vbNo then
    res = MsgBox("�����𒆎~���܂��B", vbInformation, strTitle)
    Wscript.quit
end if

theContents = theDate & tab & startTime & tab & endTime & tab & tab & work

Dim fso,ts
Set fso = CreateObject("Scripting.FileSystemObject")

Set ts = fso.OpenTextFile(strLogFile, 8, True)
ts.WriteLine(theContents)
ts.Close

Set ts = Nothing
Set fso = Nothing

' Google �X�v���b�h�V�[�g�ɓo�^
' added on Nov 11, 2019
Dim ws
Set ws = CreateObject("WScript.Shell")
'ws.Run "powershell.exe -command " & q & "Invoke-RestMethod -Method Post -Uri " & q & theUrl & q & " -Body @{startTime='17:15'; endTime='" & endTime & "';task='" & work & "'} ", 0, True

res = MsgBox("�o�^���܂����B�����l�ł����B", vbInformation, strTitle)

res = MsgBox("�V�X�e�����I�����܂����H", vbQuestion + vbYesNoCancel, strTitle)
If res = vbYes Then
    ws.run "%WINDIR%\system32\shutdown.exe -s -t 0", 0
Else
    Wscript.quit
End If