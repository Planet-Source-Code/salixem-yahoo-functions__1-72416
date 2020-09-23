Attribute VB_Name = "General"
'Yahoo Functions
'By SaLar Zeynali
'Salixem@Gmail.Com
Declare Function YMSG12_ScriptedMind_Encrypt Lib "YMSG12ENCRYPT.dll" (ByVal username As String, ByVal password As String, ByVal Seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean
Function StatusMsg(Msg As String, Optional Busy As Integer = 0)
If Msg = "Idle" Then Busy = 2
If LCase(Msg) = "invisible" Then
    StatusMsg = Packet("c5", "13À€2À€")
Else
    StatusMsg = Packet("c6", "10À€99À€19À€" & Msg & "À€47À€" & Busy & "À€187À€0À€")
End If
End Function
Function SendPM(From As String, Who As String, Msg As String)
SendPM = Packet(17, "1À€" & From & "À€5À€" & Who & "À€14À€" & Msg & "À€97À€1À€")
End Function
Function AddContact(ID As String, Who As String, Optional Grp As String = "Contacts", Optional Msg As String)
AddContact = Packet(83, "1À€" & ID & "À€7À€" & Who & "À€14À€" & Msg & "À€65À€" & Grp & "À€")
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
PostLogin = Packet(54, "6À€" & Enc(0) & "À€96À€" & Enc(1) & "À€0À€" & ID & "À€2À€" & ID & "À€192À€-1À€1À€" & ID & "À€135À€6,0,0,0000À€148À€360À€")
Exit Function
Error:
PostLogin = Err
End Function
Function PreLogin(ID As String)
PreLogin = Packet(57, "1À€" & ID & "À€")
End Function

Function Packet(PackType As String, Pack As String, Optional ByVal Key As String)
If Key = "" Then Key = String(4, 0)
Packet = "YMSG" & Chr(0) & Chr(12) & String(2, 0) & _
Chr(Fix(Len(Pack) / 256)) & Chr(Len(Pack) Mod 256) & _
Chr(0) & Chr("&h" & PackType) & String(4, 0) & Key & _
Pack
End Function

