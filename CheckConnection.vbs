Option Explicit
Dim objShell, intShutdown, fso, msgFile,strShutdown, strAbort,folderPath,file
folderPath = "."

Do
    If Not CheckInternetConnection() Then
        set objShell = CreateObject("WScript.Shell")
        set fso = CreateObject("Scripting.FileSystemObject")

		If Not fso.FileExists("1.Slg") or fso.FileExists("2.Slg") Then
			Wscript.Sleep 120000 ' Wait 2 Minute for user to prepare and enter the voucher or other 
			objShell.Run("shutdown.vbs") ' If the waiting time is up, then execute the shutdown script
		Else
			objShell.Run("shutdown.vbs") ' If file log exists, then execute shutdown script without delay 
		End If
    End If
    set fso = Nothing
    set objShell = Nothing
    Wscript.Sleep 132000 
Loop

Function CheckInternetConnection()
    On Error Resume Next
    Dim objHTTP

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

    objHTTP.Open "GET", "https://www.google.com", False
    objHTTP.Send

    If Err.Number = 0 Then
        If objHTTP.Status = 200 Then
            set fso = CreateObject("Scripting.FileSystemObject")
            If fso.FileExists("1.Slg") or fso.FileExists("2.Slg") Then
                fso.DeleteFile("*.Slg")
            End If
            CheckInternetConnection = True
        Else
            CheckInternetConnection = False
        End If
    Else
        CheckInternetConnection = False
    End If

    Set objHTTP = Nothing
    On Error GoTo 0
End Function
