Option Explicit
Dim objShell, intShutdown, fso, file, msgFile,strShutdown, strAbort

' Goal : If the user meets the 'click no' limit on the message box, then forcefully shutdown the computer
Function MsgShutdown
    set fso = CreateObject("Scripting.FileSystemObject")
    set objShell = CreateObject("WScript.Shell")
    strShutdown = "shutdown.exe -s -t 120 -f" ' shutdown the computer within 2 minutes
    strAbort = "shutdown.exe -a" ' abort the shutdown

    objShell.Run strShutdown, 0, false ' start first shutdown script
    ' Message box appear
    intShutdown = (MsgBox("WAKTU SUDAH HABIS, APAKAH INGIN MEMATIKAN KOMPUTER? JIKA DIDIAMKAN AKAN MATI DALAM 2 MENIT. Pilih YES UNTUK MEMATIKAN KOMPUTER, PILIH NO UNTUK LOGIN VOUCHER",vbYesNo+vbExclamation+vbSystemModal,"WAKTU INTERNET SUDAH HABIS !!!"))

    If intShutdown = vbYes Then ' if the user clicks yes button      
        If fso.FileExists("1.Slg") or fso.FileExists("2.Slg") Then
            fso.DeleteFile("*.Slg")
        End If
	objShell.Run strAbort, 0, false
        strShutdown = "shutdown.exe -s -t 0 -f" 
        objShell.Run strShutdown, 0, false ' shutdown the computer immediately
    ElseIf intShutdown = vbNo Then ' If the user clicks no button  
        If fso.FileExists("2.Slg") Then
	    objShell.Run strAbort, 0, false
            strShutdown = "shutdown -s -f -t 7" ' shutdown the computer within 7 seconds and force close all apps
            objShell.Run strShutdown, 0, false
            fso.DeleteFile("*.Slg") ' delete the all .Slg File
            msgFile= (MsgBox("KOMPUTER AKAN DI MATIKAN. SILAHKAN MEMBELI VOUCHER DAN LOGIN SUPAYA PESAN INI TIDAK MUNCUL",vbExclamation,"PERINGATAN MENCAPAI BATAS MAKSIMAL !!!"))
        Else
            createDecisionLog ' call the sub
            '  Abort Shutdown
            objShell.Run strAbort, 0, false
            ' Display the default login page that has been set in MikroTik
            objShell.Run "http://warnetghz.ghz", 9 'Adjust DNS Name according to the settings in MikroTik.
        End if
    End if
    set objShell = Nothing
    set fso = Nothing
    Wscript.Quit
End Function

Sub createDecisionLog
If fso.FileExists("1.Slg") Then '
    fso.MoveFile "1.Slg", "2.Slg" ' changing file 1 to 2
Else
   fso.CreateTextFile "1.Slg" ' create a file named `1.Slg` to count the 'click no' limit on the message box
End If
End Sub

MsgShutdown