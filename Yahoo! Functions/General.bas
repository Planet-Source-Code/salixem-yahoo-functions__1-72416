Attribute VB_Name = "General"
'Yahoo Functions
'By SaLar Zeynali
'Salixem@Gmail.Com
Declare Function YMSG12_ScriptedMind_Encrypt Lib "YMSG12ENCRYPT.dll" (ByVal username As String, ByVal password As String, ByVal Seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean
Function StatusMsg(Msg As String, Optional Busy As Integer = 0)
If Msg = "Idle" Then Busy = 2
If LCase(Msg) = "invisible" Then
    StatusMsg = Packet("c5", "13��2��")
Else
    StatusMsg = Packet("c6", "10��99��19��" & Msg & "��47��" & Busy & "��187��0��")
End If
End Function
Function SendPM(From As String, Who As String, Msg As String)
SendPM = Packet(17, "1��" & From & "��5��" & Who & "��14��" & Msg & "��97��1��")
End Function
Function AddContact(ID As String, Who As String, Optional Grp As String = "Contacts", Optional Msg As String)
AddContact = Packet(83, "1��" & ID & "��7��" & Who & "��14��" & Msg & "��65��" & Grp & "��")
End Function
Function PostLogin(ID As String, PW As String, SD As String)
Dim Enc(1) As String
On Error GoTo Error
Enc(0) = String(80, 0)
Enc(1) = String(80, 0)
If YMSG12_ScriptedMind_Encrypt(ID, PW, SD, Enc(0), Enc(1), 1) = False Then
    MsgBox "Error on: YMSG12ENCRYPT.DLL", vbCritical, "YMSG12ENCRYPT.DLL"
    GoTo Error
End If
 For i = 0 To 1
    Enc(i) = Left$(Enc(i), InStr(1, Enc(i), Chr(0)) - 1)
 Next
PostLogin = Packet(54, "6��" & Enc(0) & "��96��" & Enc(1) & "��0��" & ID & "��2��" & ID & "��192��-1��1��" & ID & "��135��6,0,0,0000��148��360��")
Exit Function
Error:
PostLogin = Err
End Function
Function PreLogin(ID As String)
PreLogin = Packet(57, "1��" & ID & "��")
End Function

Function Packet(PackType As String, Pack As String, Optional ByVal Key As String)
If Key = "" Then Key = String(4, 0)
Packet = "YMSG" & Chr(0) & Chr(12) & String(2, 0) & _
Chr(Fix(Len(Pack) / 256)) & Chr(Len(Pack) Mod 256) & _
Chr(0) & Chr("&h" & PackType) & String(4, 0) & Key & _
Pack
End Function

