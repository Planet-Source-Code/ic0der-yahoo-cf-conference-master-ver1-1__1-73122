Attribute VB_Name = "Module2"

Public Function RandomS11()
Dim strInputString As String
Dim intLength As Long
Dim intNameLength As Integer
Dim strName As String
Dim intStep As Integer
Dim intRnd As Long
   strInputString = "�"
   
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
   strInputString = "���"
   
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
   strInputString = "��"
   
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
   strInputString = "À$�"
   
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
Case 1: TheString$ = "À€Ÿclients suk ŸŸŸŸ�The Only One�®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰ŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸ�®®®®®®ö¶¶‰‰‰�"
Case 2: TheString$ = "À€ŸhelloŸŸŸŸŸŸŸŸŸŸŸ®®®®®�The Only One®æææææææææææææææææ¶¶¶¶¶�byebye�¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰ŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸ�®®®®®®ö¶¶‰‰‰�"
Case 3: TheString$ = "À€ŸŸYF�You Got owned!!�®®®®YF®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰ŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸ�®®®®®®ö¶¶‰‰‰�"
Case 4: TheString$ = "À€ŸŸŸŸ�The Only One�®®®®®®®®I OWNED YOU®®®®®®®®æææææææææææææææææ¶¶¶¶¶�YF-Inc�¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰ŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸ�®®®®®®ö¶¶‰‰‰�"
Case 5: TheString$ = "À€ŸŸI OWNED YOUŸŸŸ!!hahahahaha!!ŸŸŸŸŸŸŸ®®®YF®®®The Only One®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰ŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸ�®®®®®®ö¶¶‰‰‰�"
Case 6: TheString$ = "À€ŸŸYahmartŸ�The Only One®®®We OWNED YOU®®®lol®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰ŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸ�®®®®®®ö¶¶‰‰‰�"
Case 7: TheString$ = "À€ŸOwnedŸThe Only OneŸŸŸŸ®®®®®®®®�lol�®®®®®®®�YMC�ææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰ŸŸŸŸŸŸŸŸŸŸŸŸ®®®®®®®®®®®®®®®®æææææææææææææææææ¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰‰%%%%%%%%%%%%%%%%%å**å**å**å**å**å**µ¡ääääåååååÛÛÛÛÛßààOU©©ööööööööööööööööŸŸŸŸŸ�®®®®®®ö¶¶‰‰‰�"
Case 8: TheString$ = "QWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxzQWERTYUIOPLKJHGFDSAZXCVBNM<>?:{}+_)(*&^%$#@!1234567890-=][poiuytrewqasdfghjkl;'/.,mnbvcxz"
Case 9: TheString$ = "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@#################################################################################$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!****************************************************************&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++???????????????????????????????????????::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
End Select
Randomlaggcode = TheString$
End Function

