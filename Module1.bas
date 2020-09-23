Attribute VB_Name = "Module1"
Option Explicit

Public Sub PreScan()
Dim hFile As Long, retVal As Long
Dim sRegMonClass As String, sFileMonClass As String
    
'this is used to detect softice(tm) or regmon(tm)/filemon(tm)
sRegMonClass = Chr(82) & Chr(101) & Chr(103) & Chr(77) & Chr(111) & Chr(110) & Chr(67) & Chr(108) & Chr(97) & Chr(115) & Chr(115)
sFileMonClass = Chr(70) & Chr(105) & Chr(108) & Chr(101) & Chr(77) & Chr(111) & Chr(110) & Chr(67) & Chr(108) & Chr(97) & Chr(115) & Chr(115)
 
 Select Case True
   Case FindWindow(sRegMonClass, vbNullString) <> 0
    End
   Case FindWindow(sFileMonClass, vbNullString) <> 0
    End
 End Select

hFile = CreateFile("\\.\SICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
If hFile <> -1 Then
   End
Else
 hFile = CreateFile("\\.\NTICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
 If hFile <> -1 Then
    End
 End If
End If

End Sub

Public Sub ResetForWrite()
 
Call CloseCOMM            'make sure port is closed if its open

'open the port with windows API
hPort = CreateFile(port, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)

If hPort = INVALID_HANDLE_VALUE Then         'check if port opened ok
  MsgBox port & " Error: invalid port or already in use."
  Exit Sub
End If

icond = SetupComm(hPort, 8192, 1024)         'set send/recv buffers

timeouts.ReadTotalTimeoutConstant = 10
timeouts.ReadTotalTimeoutMultiplier = 0
timeouts.ReadIntervalTimeout = 0
timeouts.WriteTotalTimeoutConstant = 100
timeouts.WriteTotalTimeoutMultiplier = 0

icond = SetCommTimeouts(hPort, timeouts)
icond = GetCommState(hPort, DCB)

SetDCBits fBinary, 1
SetDCBits fParity, 0
SetDCBits fOutxCtsFlow, 1
SetDCBits fOutxDsrFlow, 1
SetDCBits fDtrControl, DTR_CONTROL_ENABLE
SetDCBits fDsrSensitivity, 1
SetDCBits fTXContinueOnXoff, 0
SetDCBits fOutX, 0
SetDCBits fInX, 0
SetDCBits fErrorChar, 0
SetDCBits fNull, 0
SetDCBits fDtrControl, DTR_CONTROL_HANDSHAKE
SetDCBits fRtsControl, RTS_CONTROL_ENABLE
SetDCBits fAbortOnError, 0

DCB.ByteSize = 8
DCB.Parity = NOPARITY
DCB.StopBits = ONESTOPBIT
DCB.BaudRate = 9600                   'set baud to atr speed 9600

icond = SetCommState(hPort, DCB)
icond = SetCommTimeouts(hPort, timeouts)

icond = EscapeCommFunction(hPort, CLRRTS) 'init atr

If icond = False Then
  MsgBox "Error Setting COM State."
  CloseHandle hPort
  Exit Sub
End If

'read input BufferIn to get ATR data
Call ReadATR
'show the ATR data
Call ShowATR
DelaySecs 0.5

icond = SetupComm(hPort, 8192, 1024)

timeouts.ReadTotalTimeoutConstant = 10
timeouts.ReadTotalTimeoutMultiplier = 0
timeouts.ReadIntervalTimeout = 0
timeouts.WriteTotalTimeoutConstant = 100
timeouts.WriteTotalTimeoutMultiplier = 0

icond = SetCommTimeouts(hPort, timeouts)
icond = GetCommState(hPort, DCB)

SetDCBits fBinary, 1
SetDCBits fParity, 0
SetDCBits fOutxCtsFlow, 1
SetDCBits fOutxDsrFlow, 1
SetDCBits fDtrControl, DTR_CONTROL_ENABLE
SetDCBits fDsrSensitivity, 1
SetDCBits fTXContinueOnXoff, 0
SetDCBits fOutX, 0
SetDCBits fInX, 0
SetDCBits fErrorChar, 0
SetDCBits fNull, 0
SetDCBits fDtrControl, DTR_CONTROL_HANDSHAKE
SetDCBits fRtsControl, RTS_CONTROL_ENABLE
SetDCBits fAbortOnError, 0

DCB.BaudRate = 38400
DCB.ByteSize = 8
DCB.Parity = ODDPARITY
DCB.StopBits = TWOSTOPBITS
DCB.DCBlength = Len(DCB)

icond = SetCommState(hPort, DCB)
icond = SetCommTimeouts(hPort, timeouts)
icond = EscapeCommFunction(hPort, CLRRTS)

If icond = False Then
  MsgBox "Error Setting Com State."
  CloseHandle hPort
  Exit Sub
End If


End Sub

Public Sub CheckCOM(xPort)

Form1.PORTLITE.Picture = Form1.PortOFF.Picture
Call CloseCOMM  'make sure we close prev. comport first

'open COM port for generic read/write
hPort = CreateFile(port, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)

If hPort = INVALID_HANDLE_VALUE Then
   
   Form1.StatusLabel.Caption = " ACTION: " + port + " already in use or not available."
   Form1.PORTLITE.Picture = Form1.PortOFF.Picture
  Exit Sub
  
Else
  
  Form1.StatusLabel.Caption = " ACTION: " + port + " opened OK"
  Form1.PORTLITE.Picture = Form1.PortON.Picture
     
End If

End Sub

Public Sub ResetForATR()
Dim f

Call CloseCOMM

'open COM port for generic read/write
hPort = CreateFile(port, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)

If hPort = INVALID_HANDLE_VALUE Then
  MsgBox port & " Error: invalid port or already in use."
  Exit Sub
End If

icond = SetupComm(hPort, 8192, 1024)
timeouts.ReadTotalTimeoutConstant = 10
timeouts.ReadTotalTimeoutMultiplier = 0
timeouts.ReadIntervalTimeout = 0
timeouts.WriteTotalTimeoutConstant = 100
timeouts.WriteTotalTimeoutMultiplier = 0

icond = SetCommTimeouts(hPort, timeouts)
icond = GetCommState(hPort, DCB)

SetDCBits fBinary, 1
SetDCBits fParity, 0
SetDCBits fOutxCtsFlow, 1
SetDCBits fOutxDsrFlow, 1
SetDCBits fDtrControl, DTR_CONTROL_ENABLE
SetDCBits fDsrSensitivity, 1
SetDCBits fTXContinueOnXoff, 0
SetDCBits fOutX, 0
SetDCBits fInX, 0
SetDCBits fErrorChar, 0
SetDCBits fNull, 0
SetDCBits fDtrControl, DTR_CONTROL_HANDSHAKE
SetDCBits fRtsControl, RTS_CONTROL_ENABLE
SetDCBits fAbortOnError, 0

DCB.BaudRate = 9600
DCB.ByteSize = 8
DCB.Parity = EVENPARITY
DCB.StopBits = ONESTOPBIT

icond = SetCommState(hPort, DCB)
icond = SetCommTimeouts(hPort, timeouts)

'init the ATR by dropping RTS
icond = EscapeCommFunction(hPort, CLRRTS)

If icond = False Then
  MsgBox "Error Setting COM State."
  CloseHandle hPort
  Exit Sub
End If

'read input BufferIn to get ATR data
Call ReadATR

'show the ATR data
Call ShowATR
DelaySecs 0.25

'flush the BufferIn
icond = PurgeComm(hPort, PURGE_RXCLEAR Or PURGE_TXCLEAR)


End Sub

Public Sub WriteCOMM(DATA As String)

'you can send a formatted data byte here
icond = WriteFile(hPort, DATA, 1, written, 0)
InverseBuffer = ""

End Sub

Public Sub ReadATR()

BufferIncount = 0
BufferIn = ""
DelaySecs 0.25

Do
icond = ReadFile(hPort, InBuff, 1, numRead, 0)
If numRead = 0 Then Exit Sub
 BufferIn = BufferIn & InBuff
 BufferIncount = Len(BufferIn)
Loop

End Sub

Public Sub ReadDATA()

Form1.RXLITE.Picture = Form1.RXOFF.Picture

If Len(BufferIn) >= 1 Then
   If Len(BufferIn) > 3 Then CardInserted = True: Exit Sub
   BufferIn = "" 'clear out possible trash
End If

Do
Form1.RXLITE.Picture = Form1.RXON.Picture: DoEvents
icond = ReadFile(hPort, InBuff, 1, numRead, 0)
If numRead = 0 Then Form1.RXLITE.Picture = Form1.RXOFF.Picture: Exit Sub
 BufferIn = BufferIn & InBuff
 BufferIncount = Len(BufferIn)
 Form1.BuffCntText = BufferIncount
Loop

End Sub

Public Sub CloseCOMM()

If hPort = 0 Or hPort = INVALID_HANDLE_VALUE Then Exit Sub
 icond = EscapeCommFunction(hPort, CLRDTR)
 CloseHandle (hPort)

End Sub

Public Sub SetDCBits(pos As Long, val As Integer)
Dim ip As Integer
Dim imul As Long
Dim poz As Long

imul = 1
poz = pos
For ip = 1 To 32
If (poz And 1) Then Exit For
poz = poz / 2
imul = imul * 2
Next

 DCB.Bits1 = DCB.Bits1 And (Not pos) Or (imul * val)

End Sub

Public Sub DelaySecs(ByVal seconds As Single)
Static start As Single

start = Timer
Do While Timer < start + seconds
  DoEvents
Loop

End Sub

Public Sub Inverse(databyte As String)
Dim Cpos As Integer
Dim xxx As Integer
Dim RealByte
Dim TmpBuffer As String
 
 Nibble = 1
 RealByte = 0
         
 For Npos = 1 To Len(databyte)
     Temp3 = Mid$(databyte, Npos, 1)
            
If Nibble >= 1 Then
  
Select Case Temp3
   Case "0" To "9"
      RealByte = RealByte + (val(Temp3) * 16)
      Nibble = 0
   Case "A" To "F"
      RealByte = RealByte + ((Asc(Temp3) - 55) * 16)
      Nibble = 0
  End Select

Else

Select Case Temp3
   Case "0" To "9"
      RealByte = RealByte + (val(Temp3))
   Case "A" To "F"
      RealByte = RealByte + (Asc(Temp3) - 55)
  End Select
      
End If
    
    Next Npos
       
   Temp1 = (RealByte Xor 255)
   RealByte = Temp1
   Temp3 = 0
        
For Cpos = 7 To 0 Step -1
   Select Case Cpos
     Case 7:  Temp2 = RealByte And 1
     Case 6:  Temp2 = RealByte And 2
     Case 5:  Temp2 = RealByte And 4
     Case 4:  Temp2 = RealByte And 8
     Case 3:  Temp2 = RealByte And 16
     Case 2:  Temp2 = RealByte And 32
     Case 1:  Temp2 = RealByte And 64
     Case 0:  Temp2 = RealByte And 128
   End Select
     
If Temp2 > 0 Then
  If (Cpos = 0) Then
      Temp3 = Temp3 + 1
 Else
      Temp3 = Temp3 + (2 ^ Cpos)
  End If
Else
      Temp3 = Temp3
     End If
  
Next Cpos
 
   InverseBuffer = Temp3
 
End Sub

Public Sub CardInfo2A(theString As String)
Dim ret, zzz, ttt, ooo As Integer
Dim TempBuf As String
Dim CardIDlong

'clear all the variables first
TempBuf = ""
CardIDlong = 0
ret = 0
zzz = 0
ttt = 0
ooo = 0
tmpCARDID = ""
tmpIRD = ""
tmpUSW = ""
tmpGUIDE = ""
tmpRATING = ""
tmpSPENDING = ""


ooo = Len(theString)

For ret = 1 To ooo / 2
             
    On Error Resume Next
     TempBuf = Left(theString, 2)
     If Trim(TempBuf) = "" Then Exit For
     zzz = InStr(1, theString, Len(TempBuf))
     theString = Mid(theString, 3, Len(theString))
     
     If ret = 1 Then
        tmpFUSE = TempBuf
        GoTo NEGST
     End If
     
     If ret = 11 Then
        tmpRATING = TempBuf
        GoTo NEGST
     End If
     
     If ret = 12 Or ret = 13 Then
       If ret = 12 Then tmpSPENDING = TempBuf: GoTo NEGST
        If ret = 13 Then
          tmpSPENDING = tmpSPENDING + TempBuf
          tmpSPENDING = val("&H" + tmpSPENDING)
          tmpTRASH = Left(tmpSPENDING, Len(tmpSPENDING) - 2)
          tmpSPENDING = "$" + tmpTRASH + "." + Right(tmpSPENDING, 2)
         GoTo NEGST
        End If
     End If
     
     If ret = 21 Or ret = 22 Or ret = 23 Or ret = 24 Then
      If ret = 21 Then tmpCARDID = TempBuf: GoTo NEGST
       If ret = 22 Then tmpCARDID = tmpCARDID + TempBuf: GoTo NEGST
        If ret = 23 Then tmpCARDID = tmpCARDID + TempBuf: GoTo NEGST
         If ret = 24 Then
          tmpCARDID = tmpCARDID + TempBuf
          CardIDlong = tmpCARDID
          CardIDlong = val("&H" + CardIDlong)
          tmpCARDID = CardIDlong
          tmpCARDID = "000" & tmpCARDID & "_"
        GoTo NEGST
       End If
     End If
     
     If ret = 25 Or ret = 26 Or ret = 27 Or ret = 28 Then
      If ret = 25 Then tmpIRD = TempBuf: GoTo NEGST
       If ret = 26 Then tmpIRD = tmpIRD + TempBuf: GoTo NEGST
        If ret = 27 Then tmpIRD = tmpIRD + TempBuf: GoTo NEGST
         If ret = 28 Then
          tmpIRD = tmpIRD + TempBuf
          tmpIRD = val("&H" + tmpIRD)
          tmpIRD = CardIDlong Xor tmpIRD
          tmpIRD = Hex(val(tmpIRD))
          If tmpIRD = "1" Then tmpIRD = "00000001"
        GoTo NEGST
       End If
     End If
     
    If ret = 30 Then
     tmpUSW = TempBuf
    End If

NEGST:
 Next ret
 

CardInfoBuffer = ""                 'clear this buffer for 58 cmd

Exit Sub

ERRORED:
  MsgBox "We Hit an error at ret = " & ret
      
End Sub

Public Sub CardInfo58(theString As String)
Dim ret, zzz, ttt, ooo As Integer
Dim TempBuf As String

ooo = Len(theString)

For ret = 1 To ooo / 2
             
    On Error Resume Next
     TempBuf = Left(theString, 2)
     If Trim(TempBuf) = "" Then Exit For
     zzz = InStr(1, theString, Len(TempBuf))
     theString = Mid(theString, 3, Len(theString))
     
    If ret = 11 Then
        tmpTIMEZONE = TempBuf
       GoTo NEGST
     End If
     
     If ret = 13 Then
        tmpGUIDE = TempBuf
       GoTo NEGST
     End If

NEGST:
     
 Next ret
 
CardInfoBuffer = ""                 'clear this buffer

'display the info`s here
Form1.CardIDtext = tmpCARDID
Form1.IRDText = tmpIRD
Form1.USWtext = tmpUSW
Form1.FUSEtext = tmpFUSE
Form1.GUIDEtext = tmpGUIDE
Form1.TIMEZONEtext = tmpTIMEZONE
Form1.RATINGtext = tmpRATING
Form1.SPENDINGLIMITtext = tmpSPENDING

Exit Sub

ERRORED:
  MsgBox "We Hit an error at ret = " & ret

End Sub

Public Sub CardInfoPPV(theString As String)
Dim ret, zzz, xxx, ooo As Integer
Dim TempBuf As String
Dim PPVStr

For xxx = 1 To 25
   PPV(xxx) = ""
Next xxx

ooo = Len(theString)

xxx = 0

For ret = 1 To ooo / 2
             
    On Error Resume Next
     xxx = ret
     TempBuf = Left(theString, 2)
     If Trim(TempBuf) = "" Then Exit For
     zzz = InStr(1, theString, Len(TempBuf))
     theString = Mid(theString, 3, Len(theString))
     
Select Case ret
   Case Is = 1, 2, 3: PPV(&H1) = PPV(&H1) + TempBuf
   Case Is = 4, 5, 6: PPV(&H2) = PPV(&H2) + TempBuf
   Case Is = 7, 8, 9: PPV(&H3) = PPV(&H3) + TempBuf
   Case Is = 10, 11, 12: PPV(4) = PPV(4) + TempBuf
   Case Is = 13, 14, 15: PPV(5) = PPV(5) + TempBuf
   Case Is = 16, 17, 18: PPV(6) = PPV(6) + TempBuf
   Case Is = 19, 20, 21: PPV(7) = PPV(7) + TempBuf
   Case Is = 22, 23, 24: PPV(8) = PPV(8) + TempBuf
   Case Is = 25, 26, 27: PPV(9) = PPV(9) + TempBuf
   Case Is = 28, 29, 30: PPV(10) = PPV(10) + TempBuf
   Case Is = 31, 32, 33: PPV(11) = PPV(11) + TempBuf
   Case Is = 34, 35, 36: PPV(12) = PPV(12) + TempBuf
   Case Is = 37, 38, 39: PPV(13) = PPV(13) + TempBuf
   Case Is = 40, 41, 42: PPV(14) = PPV(14) + TempBuf
   Case Is = 43, 44, 45: PPV(15) = PPV(15) + TempBuf
   Case Is = 46, 47, 48: PPV(16) = PPV(16) + TempBuf
   Case Is = 49, 50, 51: PPV(17) = PPV(17) + TempBuf
   Case Is = 52, 53, 54: PPV(18) = PPV(18) + TempBuf
   Case Is = 55, 56, 57: PPV(19) = PPV(19) + TempBuf
   Case Is = 58, 59, 60: PPV(20) = PPV(20) + TempBuf
   Case Is = 61, 62, 63: PPV(21) = PPV(21) + TempBuf
   Case Is = 64, 65, 66: PPV(22) = PPV(22) + TempBuf
   Case Is = 67, 68, 69: PPV(23) = PPV(23) + TempBuf
   Case Is = 70, 71, 72: PPV(24) = PPV(24) + TempBuf
   Case Is = 73, 74, 75: PPV(25) = PPV(25) + TempBuf
    Case Else
       GoTo NEGST
    End Select
    
NEGST:
     
 Next ret
 
CardInfoBuffer = ""                 'clear this buffer

xxx = 0

For xxx = 0 To 24
 Form1.Text1(xxx).Text = PPV(xxx + 1)
Next xxx

Exit Sub

ERRORED:
  MsgBox "We Hit an error at ret = " & ret


End Sub

Public Sub FlipBuffer()
Dim ret, zzz, ttt, ooo As Integer
Dim FlipTemp As String * 3
Dim TempBuf As String
Dim BufLen As Integer

R0byte = ""

BufLen = Len(BufferIn)

For ret = 1 To BufLen  '(BufferIn)
             
    On Error Resume Next
     TempBuf = Left(BufferIn, 1)
     If Trim(TempBuf) = "" Then Exit For
     zzz = InStr(1, BufferIn, TempBuf)
     BufferIn = Mid(BufferIn, zzz + 1, Len(BufferIn))
     TempBuf = Hex(Asc(BufferIn))
      
     If Len(TempBuf) = 1 Then
       TempBuf = "0" + TempBuf
     End If
       
    Call Inverse(TempBuf)
       
    TempBuf = Hex(InverseBuffer)
     
     If Len(TempBuf) = 1 Then
        TempBuf = "0" + TempBuf
     End If
     
     ByteToFlip(ret) = TempBuf
     
    If ret = BufLen - 2 Then
      R0byte = ByteToFlip(ret)
    End If
    
    If ret = BufLen - 1 Then
      R0byte = R0byte + " " + ByteToFlip(ret)
      BufferIn = ""
      GoTo FIN                                   'we know this is last byte sent from
     End If                                      'card so clear any trash and exit
     
 Next ret
 
FIN:
     BytesToRead = ret

Exit Sub

ERRORED:
  MsgBox "We Hit an error at ret = " & ret
  BufferIn = ""
      

End Sub

Public Sub SendData(StrName As String)
Dim yyy As Integer
Dim zzz As Integer
Dim num As Integer

PurgeComm hPort, PURGE_RXCLEAR Or PURGE_TXCLEAR

Form1.TXLITE.Picture = Form1.TXOFF.Picture: DoEvents

If AtrLen < 59 Then Exit Sub


num = Len(StrName) / 2                                 '# of bytes for this data string
xxx = 0                                                'set xxx to 0
                    
For xxx = 1 To num
   
   If xxx = 1 Then
     If Trim(StrName) = "" Then Exit For               'if data string empty then stop
     ByteStr$ = Trim$(Left(StrName, 2))                'grab 2 bytes from data string
     zzz = InStr(1, StrName, ByteStr$)                 'set len of data to minus 1st 2 bytes
     StrName = Mid(StrName, zzz + 2, Len(StrName))     'remove the 2 bytes from orig string
   Else
     If Trim(StrName) = "" Then Exit For               'if data string empty then stop
     ByteStr$ = Trim$(Left(StrName, 2))                'grab 2 bytes from data string
     zzz = InStr(1, StrName, ByteStr$)                 'set len of data to -2 1st places
     StrName = Mid(StrName, zzz + 2, Len(StrName))     'remove the 2 bytes from orig string
   End If
  
   Call CheckHexLen(ByteStr$, 1)
   
   SendStr(xxx) = ByteStr$                             'set the array for the data/header bytes
   
Next xxx

 
 xxx = 0                                               'clear xxx again


For xxx = 1 To num
    
    Form1.TXLITE.Picture = Form1.TXOFF.Picture: DoEvents
    
    Call Inverse(SendStr(xxx))                        'flip the bits/bytes
   
    If Trim(StrName) = "" Then
     TmpStr$ = Chr(InverseBuffer)                     'format the data
    Else
     MsgBox "Data in header or packet did not parse correctly!", 0, "ERROR"
     Exit Sub
    End If
    
    Form1.TXLITE.Picture = Form1.TXON.Picture: DoEvents
    WriteFile hPort, TmpStr$, 1, written, 0         'write the byte to card
    DelaySecs (0.019)
    ReadFile hPort, InBuff, 1, numRead, 0           'read the echo bytes
    BytesTotalSent = BytesTotalSent + 1
    Form1.BYTESsentText = Str(BytesTotalSent)
    
    If CheckINS Then
'      Stop
      DoEvents
      If InBuff = TmpStr Then
       DoEvents
      Else
       MsgBox "Byte did not echo correctly!"
       Stop
       Call CloseCOMM
       End
       Exit Sub
      End If
    End If
    
    'the following 2 lines are a MUST, it clears the way for RX/TX data!
    Call ReadDATA 're-read the echo bytes that may remain in buffer
    PurgeComm hPort, PURGE_RXCLEAR Or PURGE_TXCLEAR  'flush the buffer!
     
Next
 
    Form1.TXLITE.Picture = Form1.TXOFF.Picture

End Sub

Public Sub CheckHexLen(dxStr As String, xNum As Integer)

Select Case xNum
Case Is = 1: If Len(dxStr) = 1 Then ByteStr = "0" + ByteStr
Case Is = 2: If Len(dxStr) = 1 Then WorkByte = "0" + WorkByte
Case Is = 3: If Len(dxStr) = 1 Then preDATA = "0" + preDATA
Case Is = 4: If Len(PreATR) = 1 Then PreATR = "0" + PreATR
End Select

End Sub
Public Sub GetReturn()
  
  'this sub gets the 90 00 / 90 80 type response after sending a packet
  postDATA = "": preDATA = ""
  WorkByte = "":  R0byte = ""
  
  For xxx = 1 To Len(BufferIn)
      
      WorkByte = Hex(Asc(Mid(BufferIn, xxx, 1)))
   
      Call CheckHexLen(WorkByte, 2)
      
      Call Inverse(WorkByte)
      
      preDATA = Hex(InverseBuffer$)
      
      Call CheckHexLen(preDATA, 3)
      
      postDATA = postDATA + " " + preDATA

  Next
  
      R0byte = LTrim(RTrim(postDATA))
      Form1.R02Label.Text = R0byte
      
End Sub

Public Sub ShowATR()
  
  ATR$ = ""
  PreATR = ""
  PostATR = ""
  
  For xxx = 1 To Len(BufferIn)
     WorkByte = Hex(Asc(Mid(BufferIn, xxx, 1))) 'grab 1 byte from BufferIn
      
      Call CheckHexLen(WorkByte, 2)
          
      Call Inverse(WorkByte)                    'reverse and invert it
      
      PreATR = Hex(InverseBuffer)               'convert to Hex
          
      Call CheckHexLen(PreATR, 4)
      
      PostATR = PostATR + " " + PreATR          'add each hex byte to holder string
      ATR$ = Trim$(PostATR)                     'trim off any end spaces if any
      
   Next xxx
   
   
 Select Case Mid(ATR$, 1, 5)
  Case Is = "3F 7F"
    Form1.ATRlabel.Caption = ""
    Form1.ATRlabel.Caption = " ATR:    " + ATR$
    Form1.StatusLabel.Caption = " ACTION: HU series (P3) ATR detected"
    AtrLen = Len(ATR): BufferIn = "": ATR$ = ""
    
  Case Is = "3F 78"
    Form1.ATRlabel.Caption = ""
    Form1.ATRlabel.Caption = " ATR:    " + ATR$
    Form1.StatusLabel.Caption = " ACTION:  H series (P2) ATR detected"
    AtrLen = Len(ATR): BufferIn = "": ATR$ = ""
    
   Case Is = "3F 76"
    Form1.ATRlabel.Caption = ""
    Form1.ATRlabel.Caption = " ATR:    " + ATR$
    Form1.StatusLabel.Caption = " ACTION:  F series (P1) ATR detected"
    AtrLen = Len(ATR): BufferIn = "": ATR$ = ""
    
   Case Else
    MsgBox "Unknown or corrupt ATR"
    Form1.ATRlabel.Caption = ""
    Form1.ATRlabel.Caption = " ATR:    " + ATR$
    Form1.StatusLabel.Caption = " ACTION:  ? series (P?) ATR detected"
    AtrLen = Len(ATR): BufferIn = "": ATR$ = ""
 End Select
    
    
End Sub

Public Sub ShowDATA()

'convert it to hex so we can parse it below
 Call FlipBuffer

'clear out our variables
  WorkByte = ""
  preDATA = ""
  postDATA = ""
  
'parse the coverted Hex into an array
  For xxx = 1 To BytesToRead
     WorkByte = ByteToFlip(xxx)
     CardInfoBuffer = CardInfoBuffer + WorkByte
     postDATA = WorkByte
     
    If xxx = BytesToRead - 1 Then
      R0byte = postDATA
      GoTo skip
    End If
     
    If xxx = BytesToRead Then
      R0byte = R0byte + " " + postDATA
      GoTo skip
    End If
     
    DATA$ = DATA$ + " " + Trim$(postDATA)

skip:
  
  Next xxx
   
     Form1.R02Label.Text = R0byte
     Form1.txtOut.Text = Trim(DATA$)
     CardInfoBuffer = Trim(CardInfoBuffer)
     DATA = ""

End Sub

Public Sub ClearVariables()
 
 Form1.txtOut.Text = "":
 Form1.TextInReadBuffer.Text = ""
 Form1.CardIDtext.Text = ""
 Form1.IRDText.Text = ""
 Form1.USWtext.Text = ""
 Form1.FUSEtext.Text = ""
 Form1.GUIDEtext.Text = ""
 Form1.TIMEZONEtext.Text = ""
 Form1.RATINGtext.Text = ""
 Form1.SPENDINGLIMITtext.Text = ""
 Form1.BuffCntText.Text = ""
 Form1.R02Label.Text = ""
 Form1.BYTESsentText.Text = ""
 ATR = "":
 PreATR = "":
 PostATR = "":
 DATA = "":
 preDATA = "":
 postDATA = "":
 BufferIn = "":
 BytesTotalSent = 0
 
End Sub

Public Sub ToggleButtons()

If Form1.CARDinfoBtn.Enabled = True Then
  Form1.ATRlabel.Caption = ""
  Form1.COMMlist.Enabled = False
  Form1.CARDinfoBtn.Enabled = False
  Form1.Command1.Enabled = False
Else
 Form1.CARDinfoBtn.Enabled = True
 Form1.Command1.Enabled = True
 Form1.COMMlist.Enabled = True
 End If

End Sub

Public Sub SaveState()

'save port the user chose
SaveSetting "HKEY_CLASSES_ROOT\Interface\{F9043C87-F6F2-101A-A3C9-08002B2F49FF}\TypeLib", "Properties", "Port", port  'save COM# to registry
    
End Sub

Public Sub GetState()

'restore port the user chose
port = GetSetting("HKEY_CLASSES_ROOT\Interface\{F9043C87-F6F2-101A-A3C9-08002B2F49FF}\TypeLib", "Properties", "Port", "")

End Sub

Public Sub ShowStatus(xMsg)

xMsg = " ACTION: " + xMsg
Form1.StatusLabel.Caption = xMsg

End Sub

