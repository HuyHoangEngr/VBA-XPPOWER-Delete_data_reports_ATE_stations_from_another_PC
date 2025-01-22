Dim strip As String
Dim danhapdataip As Boolean
Dim truycapduocATE As Boolean
Dim strSN As String
Dim strSNATE As String
Dim dungIPcuatramATE As Boolean
Dim daFindpart As String
Dim chuatimxong As String
Dim dung As String

Sub btncheckip_Click()
    strip = ""
    strip = txtip.Text
    danhapdataip = False
    truycapduocATE = False
    strSN = ""
    'DOI SN O DAY DE SU DUNG CHO TRAM KHAC
    strSNATE = "S12018"
    dungIPcuatramATE = False
    
    Dim objFolder As Object
    Dim objFile As Object
    
    If strip = "" Then
        MsgBox "IP khong duoc rong", vbInformation
    Else
        danhapdataip = True
    End If
    
    If danhapdataip Then
        On Error Resume Next
        If Dir("\\" & strip & "\Reports\TEST GO-NOGO", vbDirectory) = "" Then
            MsgBox "IP khong phai cua tram ATE !", vbCritical
        Else
            truycapduocATE = True
        End If
    End If
    
    If truycapduocATE Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFolder = objFSO.GetFolder("\\" & strip & "\Reports\TEST GO-NOGO\HTML")
        
        For Each objFile In objFolder.files
            With objFile
                strSN = Mid(objFile.Name, 2, 6)
            End With
            Exit For
        Next objFile
        
        Debug.Print strSN
    
        If strSN = strSNATE Then
            MsgBox "Da truy cap thanh cong !", vbInformation
            dungIPcuatramATE = True
            txtip.Locked = True
            btncheckip.Locked = True
        Else
            MsgBox "Khong dung IP cua tram ATE !", vbCritical
        End If
    End If
End Sub
'CONTROL CHO WORKBOOK BACK DE COI SU THAY DOI CUA TEXTBOX
'MO CHON TASK MANAGER TRONG KHI CHAY
Sub btnfindpath_Click()
    Dim folder As Object
    Dim folderhtml As Object
    Dim file As Object
    Dim fileNgoai As Object
    
'    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim demsoluongfiletrongmotfoler As Long
'    i = 0
    j = CLng(ThisWorkbook.Sheets("Quet Part").Range("H3").Text)

    Dim dungquet As String
    Dim daquetxongfile As String
    
    daquetxongfile = "False"
    daquetfilexong = ThisWorkbook.Sheets("Delete").Range("B1").Text
    
    Dim dakiemtramocpart As Boolean
    dakiemtramocpart = False
    Dim dakiemtramocfile As Boolean
    dakiemtramocfile = False
    'KIEM TRA DA FIND PART CHUA
    daFindpart = ThisWorkbook.Sheets(2).Range("B1").Text
    
    If dungIPcuatramATE <> True Then MsgBox "Thuc hien kiem tra IP truoc !", vbCritical
    If dungIPcuatramATE = True Then
        If daquetfilexong <> "True" Then
            btnfindpath.Caption = "Finding ..."
            btnfindpath.BackColor = &H80FF&
            txtfilefound.Text = ThisWorkbook.Sheets("Quet Part").Range("H3").Text
            Application.Wait Now + TimeValue("00:00:03")
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFSO.GetFolder("\\" & strip & "\Reports\FAIL")
                            k = 0
                            demsoluongfiletrongmotfoler = 0
                            'DEM SO LUONG FILE TRONG 1 FOLDER
                            'For Each fileNgoai In objFolder.files
                                Set fileNgoai = objFolder.files
                                demsoluongfiletrongmotfoler = fileNgoai.Count
                                Application.Wait Now + TimeValue("00:00:01")
                            'Next fileNgoai
                            
                            If objFolder.Name = "fail" Or objFolder.Name = "Fail" Or objFolder.Name = "FAIL" Then
                                For Each fileNgoai In objFolder.files
                                    With fileNgoai
                                        'DINH VI CAC FILE DA QUET
                                        k = k + 1
                                        Debug.Print k
                                        If k Mod 500 = 0 Then
                                            Application.Wait Now + TimeValue("00:00:01")
                                        End If
                                        If k = CLng(ThisWorkbook.Sheets("Quet Part").Range("H2").Text) Or dakiemtramocfile Then
                                            dakiemtramocfile = True
                                            'SUA NAM O DAY DE DOI NAM KHAC
                                            If Year(CDate(fileNgoai.DateLastModified)) = 2018 Or Year(CDate(fileNgoai.DateLastModified)) = 2019 Or Year(CDate(fileNgoai.DateLastModified)) = 2020 Or Year(CDate(fileNgoai.DateLastModified)) = 2021 Then
                                                j = j + 1
                                                Debug.Print j & " " & fileNgoai.Name & " -> " & fileNgoai.DateLastModified
    '                                            ThisWorkbook.Sheets("Quet Files").Range("A" & j) = i
                                                ThisWorkbook.Sheets("Quet Files").Range("B" & j) = j
                                                ThisWorkbook.Sheets("Quet Files").Range("C" & j) = objFolder.Name
                                                ThisWorkbook.Sheets("Quet Files").Range("D" & j) = fileNgoai.DateLastModified
                                                ThisWorkbook.Sheets("Quet Files").Range("E" & j) = fileNgoai.Path
                                                
                                                'LUU SO LUONG PART DA QUET VAO SHEET QUET PART
    '                                            ThisWorkbook.Sheets("Quet Part").Range("H1") = i
                                                'LUU SO LUONG FILE DA LAY PATH VAO SHEET QUET PART
                                                ThisWorkbook.Sheets("Quet Part").Range("H3") = j
                                                'LUU SO LUONG FILE DA QUET (CHI TRONG THU MUC PART TUONG UNG) VAO SHEET QUET PART - MATRIX 2 CHIEU i,k
                                                ThisWorkbook.Sheets("Quet Part").Range("H2") = k + 1
                                                
                                                If j Mod 100 = 0 Then
                                                    txtfilefound.Text = j
                                                    Application.Wait Now + TimeValue("00:00:01")
                                                    ThisWorkbook.Save
                                                End If
                                                
                                                If j Mod 1000000 = 0 Then
                                                    txtfilefound.Text = j
    '                                                    Application.Wait Now + TimeValue("00:00:01")
                                                    ThisWorkbook.Save
                                                    dungquet = InputBox("Type y/Y to stop finding:", "Stop box")
                                                    'THOAT KHOI VIEC QUET
                                                    If dungquet = "Y" Or dungquet = "y" Then
                                                        ThisWorkbook.Save
                                                        MsgBox "Da tim duoc " & j & " paths"
                                                        Exit Sub
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End With
                                Next fileNgoai
                            End If
                            'CAP NHAT THONG TIN SAU KHI
                            txtfilefound.Text = j
                            ThisWorkbook.Save
            
            'DA QUET FILE XONG
            'CHUA XAC DINH KHI NAO XOA XONG
            If k >= demsoluongfiletrongmotfoler Then
                MsgBox "Da quet xong !"
                btnfindpath.Caption = "Find path completed"
                btnfindpath.Font.Size = 8
                btnfindpath.BackColor = &H80FF80
                btnfindpath.Locked = True
                daquetxongfile = "True"
                ThisWorkbook.Sheets("Delete").Range("B1") = daquetxongfile
            End If
    
            ThisWorkbook.Save
        Else
            MsgBox "Da quet xong !"
            btnfindpath.Caption = "Find path completed"
            btnfindpath.Font.Size = 8
            btnfindpath.BackColor = &H80FF80
            btnfindpath.Locked = True
            daquetxongfile = "True"
            ThisWorkbook.Sheets("Delete").Range("B1") = daquetxongfile
            txtfilefound.Text = ThisWorkbook.Sheets("Quet Part").Range("H3").Text
        End If
    End If
End Sub

Sub btndelete_Click()
    Dim i As Long
    Dim dungxoa As String
    Dim dakiemtramocdelete As Boolean
    
    dakiemtramocdelete = False
    
    If ThisWorkbook.Sheets("Delete").Range("B1").Text = "False" Or ThisWorkbook.Sheets("Delete").Range("B1").Text = "" Then MsgBox "Hoan thanh viec quet Files truoc khi Delete !", vbCritical
    If ThisWorkbook.Sheets("Delete").Range("B1").Text = "True" Then
        btndelete.Locked = True
        btndelete.Caption = "Deleting..."
        btndelete.BackColor = &H80FF&
        Application.Wait Now + TimeValue("00:00:03")
        For i = 1 To CLng(ThisWorkbook.Sheets("Quet Part").Range("H3").Text) Step 1
            If i = CLng(ThisWorkbook.Sheets("Delete").Range("B2").Text) Or dakiemtramocdelete Then
                dakiemtramocdelete = True
                On Error Resume Next
                If Dir(ThisWorkbook.Sheets("Quet Files").Range("E" & i).Text) <> "" Then
                    Kill (ThisWorkbook.Sheets("Quet Files").Range("E" & i).Text)
                    ThisWorkbook.Sheets("Delete").Range("E" & i) = ThisWorkbook.Sheets("Quet Files").Range("E" & i)
                    Debug.Print i & " Deleted -> " & ThisWorkbook.Sheets("Quet Files").Range("E" & i).Text
                    
                    ThisWorkbook.Sheets("Delete").Range("B2") = i + 1
                    
                    'LUU SO LUONG PART DA QUET VAO SHEET QUET PART
                    ThisWorkbook.Sheets("Delete").Range("B2") = i + 1
                    
                    If i Mod 100 = 0 Then
                        txtdelete.Text = i
                        Application.Wait Now + TimeValue("00:00:01")
                        ThisWorkbook.Save
                    End If
                    
                    If i Mod 1000000 = 0 Then
                        dungxoa = InputBox("Type y/Y to stop deleting:", "Stop box")
                        'THOAT KHOI VIEC XOA
                        If dungxoa = "Y" Or dungquet = "y" Then
                            ThisWorkbook.Save
                            MsgBox "Da xoa duoc " & i & " files"
                            Exit For
                        End If
                    End If
                End If
                
                If i = CLng(ThisWorkbook.Sheets("Quet Part").Range("H3").Text) Then
                    MsgBox "Da xoa xong !", vbInformation
                    btndelete.Caption = "Delete completed"
                    btndelete.Font.Size = 8
                    btndelete.Locked = True
                    btndelete.BackColor = &H80FF80
                    txtdelete.Text = i
                End If
            End If
        Next i
    End If
End Sub
