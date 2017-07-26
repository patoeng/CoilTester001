Attribute VB_Name = "ModLog"

'Region "File Operation"
    Sub Update_Datalog_Laser()
        'Exit Sub 'sementara lewati
        
        On Error Resume Next
        Dim FolderName, FileName As String
        'Week = Format(Date, "WW", vbSunday, vbFirstFullWeek)
        'FolderName = App.Path & "\DataLaser\"
        FolderName = "\\192.168.0.10\DATA\"
        FileName = "Data.txt"
        
        If Dir(FolderName) = "" Then
            MkDir (FolderName)
        End If
        If Dir(FolderName & FileName) = "" Then
            Open FolderName & FileName For Output As #1
            'Open FolderName & FileName For Append As #1
            'Print #1, "SN,Description,Date"
        Else
            Open FolderName & FileName For Output As #1
            'Open FolderName & FileName For Append As #1
        End If
'            Print #1, SN16 _
'            & ", " & ProductDesc _
'            & ", " & Now

            Print #1, frmMain.txtLaser1(0).Text
            Print #1, frmMain.txtLaser1(1).Text
            Print #1, frmMain.txtLaser1(2).Text
            Print #1, frmMain.txtLaser1(3).Text
            Print #1, frmMain.txtLaser1(4).Text

            Print #1, frmMain.txtLaser2(0).Text
            Print #1, frmMain.txtLaser2(1).Text
            Print #1, frmMain.txtLaser2(2).Text
            Print #1, frmMain.txtLaser2(3).Text
            Print #1, frmMain.txtLaser2(4).Text
            Close #1
    End Sub
 'End Region

