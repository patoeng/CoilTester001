Attribute VB_Name = "Module1"
'''''variabel global
Public statusECG As Integer
Public statusLaser As Boolean
Public tunggu As Integer
Public tunggu2 As Integer
Public goNextStep As Boolean
Public firstload As Boolean
Public DebugMode As Boolean
Public ECGtriger As Boolean
Public stsDelay As Boolean
Public Conn As New ADODB.Connection
Public DatatxtR As String
Public DatatxtL As String


Public Sub koneksi()
    On Error GoTo konekErr
    If Conn.State = 1 Then Conn.Close
        Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\db_tes.mdb;Persist Security Info=False"
    Exit Sub
konekErr:
        MsgBox "Gagal menghubungkan ke Database ! Kesalahan pada : " & Err.Description, vbCritical, "Peringatan"
End Sub
