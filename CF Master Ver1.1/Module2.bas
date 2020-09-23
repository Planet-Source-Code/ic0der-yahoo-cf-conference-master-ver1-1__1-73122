Attribute VB_Name = "Module2"

Public Function RandomS11()
Dim strInputString As String
Dim intLength As Long
Dim intNameLength As Integer
Dim strName As String
Dim intStep As Integer
Dim intRnd As Long
   strInputString = "¾"
   
   intLength = Len(strInputString)
   
   intNameLength = 1000
   
   Randomize
   
   strName = ""
   
   For intStep = 1 To intNameLength
       intRnd = Int((intLength * Rnd) + 1)
   
       strName = strName & Mid(strInputString, intRnd, 1)
   Next
   
   RandomS11 = strName
End Function

Public Function RandomS22()
Dim strInputString As String
Dim intLength As Long
Dim intNameLength As Integer
Dim strName As String
Dim intStep As Integer
Dim intRnd As Long
   strInputString = "¶§¥"
   
   intLength = Len(strInputString)
   
   intNameLength = 1000
   
   Randomize
   
   strName = ""
   
   For intStep = 1 To intNameLength
       intRnd = Int((intLength * Rnd) + 1)
   
       strName = strName & Mid(strInputString, intRnd, 1)
   Next
   
   RandomS22 = strName
End Function

Public Function RandomS33()
Dim strInputString As String
Dim intLength As Long
Dim intNameLength As Integer
Dim strName As String
Dim intStep As Integer
Dim intRnd As Long
   strInputString = "Øß"
   
   intLength = Len(strInputString)
   
   intNameLength = 1000
   
   Randomize
   
   strName = ""
   
   For intStep = 1 To intNameLength
       intRnd = Int((intLength * Rnd) + 1)
   
       strName = strName & Mid(strInputString, intRnd, 1)
   Next
   
   RandomS33 = strName
End Function
Public Function RandomS44()
Dim strInputString As String
Dim intLength As Long
Dim intNameLength As Integer
Dim strName As String
Dim intStep As Integer
Dim intRnd As Long
   strInputString = "Ã€$§"
   
   intLength = Len(strInputString)
   
   intNameLength = 1000
   
   Randomize
   
   strName = ""
   
   For intStep = 1 To intNameLength
       intRnd = Int((intLength * Rnd) + 1)
   
       strName = strName & Mid(strInputString, intRnd, 1)
   Next
   
   RandomS44 = strName
End Function

Public Function RandomS55()
Dim strInputString As String
Dim intLength As Long
Dim intNameLength As Integer
Dim strName As String
Dim intStep As Integer
Dim intRnd As Long
   strInputString = "!@#$%^&*()_+yfYF"

   
   intLength = Len(strInputString)
   
   intNameLength = 1000
   
   Randomize
   
   strName = ""
   
   For intStep = 1 To intNameLength
       intRnd = Int((intLength * Rnd) + 1)
   
       strName = strName & Mid(strInputString, intRnd, 1)
   Next
   
   RandomS55 = strName
End Function

Function RandomSS()
Dim TheNum As Integer, TheString As String
TheNum% = 5 * Rnd
Select Case TheNum%
Case 1: TheString$ = RandomS11
Case 2: TheString$ = RandomS22
Case 3: TheString$ = RandomS33
Case 4: TheString$ = RandomS44
Case 5: TheString$ = RandomS55
End Select
RandomSS = TheString$
End Function

Function RandomSize()
Dim TheNum As Integer, TheString As String
TheNum% = 5 * Rnd
Select Case TheNum%
Case 1: TheString$ = "32"
Case 2: TheString$ = "28"
Case 3: TheString$ = "24"
Case 4: TheString$ = "14"
Case 5: TheString$ = "18"
End Select
RandomSize = TheString$
End Function
Function RandomAlt()
Dim TheNum As Integer, TheString As String
TheNum% = 5 * Rnd
Select Case TheNum%
Case 1: TheString$ = "<ALT #f3e212,#e2239f,#1e23e8,#1ddee9,#7dea1c,#fb0b0b>"
Case 2: TheString$ = "<ALT #e2239f,#1e23e8,#1ddee9,#101010>"
Case 3: TheString$ = "<ALT #1e23e8,#7dea1c,#858184>"
Case 4: TheString$ = "<ALT #f3e212,#e2239f,#1ddee9,#7dea1c>"
Case 5: TheString$ = "<ALT #e2239f,#1e23e8,#fb0b0b,#858184,#101010>"
End Select
RandomAlt = TheString$
End Function
Function RandomFont()
Dim TheNum As Integer, TheString As String
TheNum% = 4 * Rnd
Select Case TheNum%
Case 1: TheString$ = "<font face=""Wingdings"" size=""" & RandomSize & """>"
Case 2: TheString$ = "<font face=""Webdings"" size=""" & RandomSize & """>"
Case 3: TheString$ = "<font face=""Symbol"" size=""" & RandomSize & """>"
Case 4: TheString$ = "<font face=""Comic Sans MS"" size=""" & RandomSize & """>"
End Select
RandomFont = TheString$
End Function
Function Randomlaggcode()
Dim TheNum As Integer, TheString As String
TheNum% = 9 * Rnd
Select Case TheNum%
Case 1: TheString$ = "Ã€â‚¬Å¸clients suk Å¸Å¸Å¸Å¸ÅThe Only One®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸®Â®Â®Â®Â®Â®Â®Ã¶Â¶Â¶â€°â€°â€°â€"
Case 2: TheString$ = "Ã€â‚¬Å¸helloÅ¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®ÂThe Only OneÂ®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Âbyebye¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸®Â®Â®Â®Â®Â®Â®Ã¶Â¶Â¶â€°â€°â€°â€"
Case 3: TheString$ = "Ã€â‚¬Å¸Å¸YFÅYou Got owned!!®Â®Â®Â®Â®YFÂ®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸®Â®Â®Â®Â®Â®Â®Ã¶Â¶Â¶â€°â€°â€°â€"
Case 4: TheString$ = "Ã€â‚¬Å¸Å¸Å¸Å¸ÅThe Only One¸Â®Â®Â®Â®Â®Â®Â®Â®I OWNED YOUÂ®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶ÂYF-Inc¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸®Â®Â®Â®Â®Â®Â®Ã¶Â¶Â¶â€°â€°â€°â€"
Case 5: TheString$ = "Ã€â‚¬Å¸Å¸I OWNED YOUÅ¸Å¸Å¸!!hahahahaha!!Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®YFÂ®Â®Â®The Only OneÂ®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸®Â®Â®Â®Â®Â®Â®Ã¶Â¶Â¶â€°â€°â€°â€"
Case 6: TheString$ = "Ã€â‚¬Å¸Å¸YahmartÅ¸ÅThe Only OneÂ®Â®Â®We OWNED YOUÂ®Â®Â®lolÂ®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸®Â®Â®Â®Â®Â®Â®Ã¶Â¶Â¶â€°â€°â€°â€"
Case 7: TheString$ = "Ã€â‚¬Å¸OwnedÅ¸The Only OneÅ¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Âlol®Â®Â®Â®Â®Â®Â®Â®ÃYMC¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Å¸Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Â®Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Ã¦Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶Â¶â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°â€°%%%%%%%%%%%%%%%%%Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**Ã¥**ÂµÂ¡Ã¤Ã¤Ã¤Ã¤Ã¥Ã¥Ã¥Ã¥Ã¥Ã›Ã›Ã›Ã›Ã›ÃŸÃ Ã OUÂ©Â©Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Ã¶Å¸Å¸Å¸Å¸Å¸®Â®Â®Â®Â®Â®Â®Ã¶Â¶Â¶â€°â€°â€°â€"
Case 8: TheString$ = "QWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxz"
Case 9: TheString$ = "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@#################################################################################$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!****************************************************************&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++???????????????????????????????????????::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
End Select
Randomlaggcode = TheString$
End Function

