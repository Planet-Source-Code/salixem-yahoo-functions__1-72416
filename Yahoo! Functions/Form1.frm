VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yahoo! Functions by Salar Zeynali - salixem@gmail.com"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   3795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Expand 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   22
      Top             =   6420
      Width           =   3675
   End
   Begin VB.TextBox txtInPMs 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6660
      Left            =   3840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   60
      Width           =   5595
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      TabIndex        =   18
      Top             =   2160
      Width           =   675
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   17
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3060
      TabIndex        =   15
      Top             =   1440
      Width           =   675
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   7
      Top             =   1740
      Width           =   1695
   End
   Begin VB.TextBox txtTo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CheckBox chkBusy 
      Caption         =   "Busy"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1290
      TabIndex        =   4
      Top             =   1050
      Width           =   1665
   End
   Begin VB.ListBox lstOns 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   60
      TabIndex        =   12
      Top             =   4800
      Width           =   3675
   End
   Begin VB.CommandButton cmdOKStatus 
      Caption         =   "ok"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3060
      TabIndex        =   5
      Top             =   750
      Width           =   705
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1290
      TabIndex        =   3
      Top             =   750
      Width           =   1695
   End
   Begin VB.ListBox lstContacts 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   60
      TabIndex        =   10
      Top             =   2820
      Width           =   3675
   End
   Begin MSWinsockLib.Winsock Wsck 
      Left            =   8940
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3060
      TabIndex        =   2
      Top             =   60
      Width           =   705
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1290
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1290
      TabIndex        =   0
      Top             =   60
      Width           =   1695
   End
   Begin VB.Line Line5 
      X1              =   60
      X2              =   3720
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Line Line4 
      X1              =   60
      X2              =   3720
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line3 
      X1              =   60
      X2              =   3720
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   3720
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   3750
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Online contacts"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   90
      TabIndex        =   21
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contact list"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   60
      TabIndex        =   20
      Top             =   2580
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Add :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   16
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   14
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "PM To :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   11
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   9
      Top             =   390
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Y! ID:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Yahoo Functions
'By SaLar Zeynali
'Salixem@Gmail.Com
'  _________      .____    .______  ___        _____
' /   _____/____  |    |   |__\   \/  /____   /     \
' \_____  \\__  \ |    |   |  |\     // __ \ /  \ /  \
' /        \/ __ \|    |___|  |/     \  ___//    Y    \
'/_______  (____  /_______ \__/___/\  \___  >____|__  /
'        \/     \/        \/        \_/   \/        \/

Public ID As String
Public Pass As String
Public Key As String
Public Server As String

Private Sub cmdLogin_Click()
'Connecting proccess
ID = txtID.Text
Pass = txtPass.Text
Server = "scs.msg.yahoo.com" 'You can change server name to other Yahoo servers
Wsck.Close
Wsck.Connect Server, 5050 '5050 is the default port for connectin, in some cases
                          'you can use 23 or 8080 or etc
Me.Caption = "Connecting ..."
End Sub

Private Sub cmdOKStatus_Click()
'Changing status proccess
SendPack StatusMsg(txtStatus.Text, chkBusy.Value)
End Sub
Private Sub cmdSend_Click()
'Sending PM
PM txtTo.Text, txtMsg.Text
txtMsg.Text = Empty
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Expand_Click()
If Expand.Caption = ">>" Then
Me.Width = 9585
Expand.Caption = "<<"
Else
Me.Width = 3885
Expand.Caption = ">>"
End If
End Sub

Private Sub Wsck_Connect()
'After connecting login proccess will start
SendPack PreLogin(ID)
'After sending login request, we have to wait for response in Wsck_DataArrival
End Sub

Private Sub Wsck_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Wsck.GetData Data
Debug.Print Asc(Mid(Data, 12, 1)) & "- " & Data

Dim sptData() As String
sptData = Split(Data, "À€")
Select Case Asc(Mid(Data, 12, 1))
 Case 87 ' Logging in
    'In this case our login request has been accepted and we have to
    'send username and password
    SendPack PostLogin(ID, Pass, sptData(3))
    Key = Mid(Data, 17, 4)
    Me.Caption = "Logging In"
 Case 85 ' Logged in
    'In this case Yahoo ID has been logged in successfully and
    'loads contact list
    Incomming_ContactList sptData(1)
    Me.Caption = "Logged In"
 Case 1 'Online contacts
    'Online contacts will load, also if someone gets online after login proccess
    'this will update
    Incomming_ContactListOnline Data
 Case 2 'Offline contacts
    'Offline contacts will load, also if someone gets offline after login proccess
    'this will update
    Contact_Off sptData(1)
 Case 6 'Recieving PM
    txtInPMs.Text = txtInPMs.Text & sptData(1) & ": " & sptData(5) & vbNewLine
End Select
End Sub
Sub Incomming_ContactList(Contacts As String)
'Recieving all contacts
On Error Resume Next
Dim Bud() As String
Bud = Split(Replace(Contacts, Chr(&HA), ","), ",")
For i = 0 To UBound(Bud)
    If InStr(Bud(i), ":") Then Bud(i) = Mid(Bud(i), InStr(Bud(i), ":") + 1)
    If Not Bud(i) = "" Then lstContacts.AddItem Trim(Bud(i))
Next
End Sub
Sub Incomming_ContactListOnline(Data As String)
'Recieving online contacts
On Error Resume Next
Dim Bud() As String, n As Integer
Bud = Split(Data, "À€7À€")
For i = 1 To UBound(Bud)
    n = InStr(Bud(i), "À€")
    If n > 1 Then Bud(i) = Left(Bud(i), n - 1)
    Contact_On Bud(i)
Next
End Sub
Sub Contact_Off(OffBud As String)
'Recieving offline contacts
On Error Resume Next
For j = 0 To lstOns.ListCount - 1
    If lstOns.List(j) = OffBud Then
        lstOns.RemoveItem j
    End If
Next
With lstContacts
 For i = 0 To .ListCount - 1
  If LCase(.List(i)) = LCase(OffBud & " (online)") Then
    .List(i) = OffBud
    Exit Sub
  End If
 Next
End With
End Sub


Sub Contact_On(OnBud As String)
On Error Resume Next
Dim IsInList As Boolean
IsInList = False
For j = 0 To lstContacts.ListCount - 1
    If lstContacts.List(j) = OnBud & " (Online)" Then
        Exit Sub
    End If
Next
With lstContacts
 For i = 0 To .ListCount - 1
  If LCase(.List(i)) = LCase(OnBud) Then
    .List(i) = OnBud & " (Online)"
    IsInList = True
    Exit For
  End If
 Next
If IsInList = False Then
    .AddItem OnBud & " (Online)"
End If
End With
lstOns.AddItem OnBud
End Sub
Function SendPack(Packet As String) As Boolean
On Error GoTo Error
 Wsck.SendData Packet
 Debug.Print "   " & Packet
 SendPack = True
 Exit Function
Error:
 SendPack = False
End Function
Sub Sleep(ByVal Sec As Long)
Sec = Timer & Sec
Do Until Timer > Sec
    DoEvents
Loop
End Sub
Private Sub cmdAdd_Click()
'Adding a contact
SendPack AddContact(ID, txtAdd.Text)
End Sub
Sub PM(ToWho As String, Msg As String)
'Sending message
SendPack SendPM(ID, ToWho, Msg)
End Sub

