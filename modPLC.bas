Attribute VB_Name = "modPLC"

Public Settings As TypeSettings
Public Type typeFolder
    DB As String
    Confiq As String
    Datalog As String
End Type

Public Type typeSerial
    Com As String
    Baudrate As Integer
    Parity As String
    bits As Integer
End Type

Public Type typeIP
    Address As String
    Port As Integer
End Type

Type TypeSettings
     Folder As typeFolder
     Serial As typeSerial
     IP As typeIP
End Type



'---- WINSOCK
Public holdread As Boolean
Public MbusResponse As String
Public MbusByteArray(255) As Byte

Public ModbusTimeOut As Integer
Public Modbuswait As Boolean
'-
Dim MbusQuery(11) As Byte
Global PLCmw(255) As Long
Global MbusRead As Boolean
Global Mbuswrite As Boolean
'Public Const StartAddress = 0
Public Const ModbusLength = 120 '100



Public Function ReadModbus(ByVal StartAddress As Integer)
    Dim StartLow As Byte
    Dim StartHigh As Byte
    Dim LengthLow As Byte
    Dim LengthHigh As Byte
    With frmMain
        If (.Winsock1.State = 7) Then
            StartLow = StartAddress Mod 256
            StartHigh = StartAddress \ 256
            LengthLow = ModbusLength Mod 256
            LengthHigh = ModbusLength \ 256
            
            MbusQuery(0) = 0
            MbusQuery(1) = 0
            MbusQuery(2) = 0
            MbusQuery(3) = 0
            MbusQuery(4) = 0
            MbusQuery(5) = 6
            MbusQuery(6) = 1
            MbusQuery(7) = 3 'change 3
            MbusQuery(8) = StartHigh
            MbusQuery(9) = StartLow
            MbusQuery(10) = LengthHigh
            MbusQuery(11) = LengthLow
            MbusRead = True
            Mbuswrite = False
            .Winsock1.SendData MbusQuery
            Modbuswait = True
            ModbusTimeOut = 0
            'tmrTimeOut.Enabled = True
        End If
    End With
End Function

Public Function WriteModbus(ByVal MwAddress As Integer, ByVal intValue As Integer)
    Dim StartLow As Byte
    Dim StartHigh As Byte
    Dim ByteLow As Byte
    Dim ByteHigh As Byte
    Dim i As Integer
    
    With frmMain
    Dim LengthLow As Byte
    Dim LengthHigh As Byte
    
        If (.Winsock1.State = 7) Then
            StartLow = MwAddress Mod 256
            StartHigh = MwAddress \ 256
            LengthLow = 1 Mod 256
            LengthHigh = 1 \ 256
            
            MbusWriteQuery = Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(7 + 2 * 1) + Chr(1) + Chr(16) + Chr(StartHigh) + Chr(StartLow) + Chr(0) + Chr(1) + Chr(2 * 1)
            ByteLow = intValue Mod 256
            ByteHigh = intValue \ 256
            MbusWriteQuery = MbusWriteQuery + Chr(ByteHigh) + Chr(ByteLow)
            
            MbusRead = False
            Mbuswrite = True
            .Winsock1.SendData MbusWriteQuery
            Modbuswait = True
            ModbusTimeOut = 0
            'Sleep 0.01
        End If
    End With
End Function

Public Sub ClosedWinsock()
With frmMain
    If (.Winsock1.State <> sckClosed) Then
    .Winsock1.Close
    End If
    Do While (.Winsock1.State <> sckClosed)
    DoEvents
    Loop
End With
End Sub
Public Function Winsock_Connect() As Boolean
    Dim StartTime
    With frmMain
    If .Winsock1.State <> 7 Then
        If (.Winsock1.State <> sckClosed) Then
            .Winsock1.Close
        End If
        .Winsock1.RemoteHost = Settings.IP.Address
        .Winsock1.RemotePort = Settings.IP.Port
        .Winsock1.Connect
        
        StartTime = Timer
        Do While ((Timer < StartTime + 2) And (.Winsock1.State <> 7))
        DoEvents
        Loop
    End If
        If (.Winsock1.State = 7) Then
            Winsock_Connect = True
        End If
    End With
End Function
Public Sub DataArrival(ByVal datalength As Long)
On Error Resume Next

    With frmMain
            Dim b As Byte
            Dim j As Byte
            Dim i As Integer
            
            For i = 1 To datalength
                If i > 255 Then Exit Sub
                .Winsock1.GetData b
                MbusByteArray(i) = b
            Next
            j = 1 'Start register/ index
            If MbusByteArray(8) = 3 Then
                For i = 10 To MbusByteArray(9) + 9 Step 2
                    If MbusByteArray(i) > 127 Then
                        PLCmw(j) = "-" & Str((Val(Not MbusByteArray(i)) * 256) + (Not MbusByteArray(i + 1)) + 1)
                    Else
                        PLCmw(j) = Str(Val(MbusByteArray(i) * 256) + MbusByteArray(i + 1))
                        If holdread = True Then Form1.Text4(j).Text = Str(Val(MbusByteArray(i) * 256) + MbusByteArray(i + 1))
                        'Form1.Text4(j).Text = Str(Val(MbusByteArray(i) * 256) + MbusByteArray(i + 1))
                    End If
                    j = j + 1
                Next i
            ElseIf MbusByteArray(8) = 1 Then
                For i = 10 To MbusByteArray(9) + 9
                    Dim k As Integer
                    For k = 0 To 7
                        PLCmw(j + k) = ValBit(MbusByteArray(i), k)
                    Next k
                    j = j + 8
                Next i
            End If
'////////


            If Mbuswrite Then
                If (MbusByteArray(8) = 16) And (MbusByteArray(12) = 1) Then
                    Modbuswait = False
                    ModbusTimeOut = 0
                    Else
                End If
            End If
    End With
End Sub


Function SetBit(aByte As Byte, bitId As Integer, outBit As Integer) As Byte
If outBit = 1 Then
    SetBit = aByte Or (2 ^ (bitId And 7))
ElseIf outBit = 0 Then
    SetBit = aByte And (Not (2 ^ (bitId And 7)))
End If
End Function

Function ValBit(aByte As Byte, bitId As Integer) As Integer
If (aByte And (2 ^ bitId)) = 0 Then
    ValBit = 0
Else
    ValBit = 1
End If
End Function

Function SetByte(strBit As String) As Byte
For b = 1 To Len(strBit)
    aByte = Val(aByte) * 2
    bitB = Mid(strBit, b, 1)
    aByte = Val(aByte) + bitB
Next b
SetByte = aByte
End Function

