' Generates LLD string for SMART requests
' Uses Smamontools
' Smamontools folder suggested to be added to %PATH% before script starts
' Outputs JSON
' v 1.2

' CONSTANTS
Const LogMaxSize   = 16777216 ' bytes

Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8

Const LogPath      = "C:\Program Files\Zabbix Agent\Scripts\ScriptData\Logs\smlld.log"
Const LogPrevPath  = "C:\Program Files\Zabbix Agent\Scripts\ScriptData\Logs\smlld_prev.log"

Const OutPath      = "C:\Program Files\Zabbix Agent\Scripts\ScriptData\smlld_out.txt"

' VARIABLES
Set objShell       = WScript.CreateObject("WScript.Shell")
Set objExecObject  = objShell.Exec("cmd /c smartctl --scan-open")
Set objFSO         = CreateObject("Scripting.FileSystemObject")

' FUNCTIONS
Function FormatNow
	dnow = Now()
	logday = Day(dnow)
	If logday < 10 Then logday = "0" & logday
	logmonth = Month(dnow)
	If logmonth < 10 Then logmonth = "0" & logmonth
	loghour = Hour(dnow)
	If loghour < 10 Then loghour = "0" & loghour
	logminute = Minute(dnow)
	If logminute < 10 Then logminute = "0" & logminute
	logsec = Second(dnow)
	If logsec < 10 Then logsec = "0" & logsec
	FormatNow = logday & "/" & logmonth & "/" & Year(dnow) & " " & _
				loghour & ":" &logminute & ":" & logsec
End Function

Sub LogAddLine(line)
	If objFSO.FileExists(LogPath) Then
		Set objFile = objFSO.GetFile(LogPath)
		If ObjFile.Size < LogMaxSize Then
			Set objFile = Nothing
			Set outputFile = objFSO.OpenTextFile(LogPath, ForAppending, True, -1)
			outputFile.WriteLine(FormatNow & " - " & line)
			outputFile.Close
			Set outputFile = Nothing
		Else
			Set objFile = Nothing
			objFSO.CopyFile LogPath, LogPrevPath, True
			Set outputFile = objFSO.CreateTextFile(LogPath, ForWriting, True)
			outputFile.WriteLine(FormatNow & " - " & line)
			outputFile.Close
			Set outputFile = Nothing
		End If
	Else
		Set outputFile = objFSO.CreateTextFile(LogPath, True, -1)
		outputFile.WriteLine(FormatNow & " - " & line)
		outputFile.Close
		Set outputFile = Nothing
	End If
End Sub

' SCRIPT
LogAddLine "Script started"
strOut = ""
strOutput = objExecObject.StdOut.ReadAll
If strOutput = "" Then
	WScript.Echo retSMTUnavail
	WScript.Quit
End If
strSearch = "/dev"
arrSpl = Split(strOutput, vbCrLf)
arrSN = Array()
strOut = "["
strDiskType = "unknown"
strDiskName = ""
strDevType = ""
strDevName = ""
strSN = ""
fSNEx = True
fNoDiskType = True
For I = 0 To UBound(arrSpl)
	If InStr(arrSpl(I), strSearch) <> 0 Then
		fSNEx = True
		fNoDiskType = True
		lineSpl = Split(arrSpl(I))
		strDevName = lineSpl(0)
		strDevType = LCase(lineSpl(UBound(lineSpl) - 1))
		Set objExecObject = Nothing
		Set objExecObject = objShell.Exec("cmd /c smartctl -a """ + strDevName + "")
		strOutput = objExecObject.StdOut.ReadAll
		If InStr(strOutput, "Unavailable") = 0 And InStr(strOutput, "ERROR") = 0 And InStr(strOutput, "Identity failed") = 0 Then
			infoSpl = Split(strOutput, vbCrLf)
			For J = 0 To UBound(infoSpl)
				If InStr(infoSpl(J), "Serial Number:") <> 0 Then
					snSpl = Split(infoSpl(J), " ")
					strSN = snSpl(UBound(snSpl))
					For K = 0 To UBound(arrSN)
						If arrSN(K) = strSN Then
							fSNEx = False
							Exit For
						End If
					Next
					If fSNEx = True Then
						Redim Preserve arrSN(UBound(arrSN) + 1)
						arrSN(UBound(arrSN)) = strSN
					End If
				End If
				If fSNEx = True Then
					If strDevType = "ata" Then
						If fNoDiskType = True Then
							If InStr(infoSpl(J), "Device Model:") <> 0 Or InStr(infoSpl(J), "Model Number:") <> 0 Then
								fNoDiskType = False
								nameSpl = Split(infoSpl(J), " ")
								strDiskName = LCase(nameSpl(UBound(nameSpl)))
								If InStr(strDiskName, "ssd") <> 0 Or (InStr(strOutput, "Program_Fail") <> 0 Or InStr(strOutput, "Erase_Fail") <> 0) Then
									strDiskType = "ssd"
								Else
									If InStr(strOutput, "Rotation Rate:") <> 0 Then
										strDiskType = "hdd"
									Else
										If InStr(strOutput, "Spin") <> 0 Then
											strDiskType = "hddmin"
										End If
									End If
								End If
							End If
						End If
					ElseIf strDevType = "nvme" Then
						strDiskType = "nvme"
					End If
				End If
			Next
		Else
			fSNEx = False
		End If
		If fSNEx = True Then
			If Len(strOut) > 1 Then strOut = strOut + ","
			dName = Mid(strDevName, 6, Len(strDevName) - 4)
			strOut = strOut + "{""{#NAME}"":""" + dName + """,""{#SN}"":""" + strSN + """,""{#TYPE}"":""" + strDevType + """,""{#DISKTYPE}"":""" + strDiskType + """}"
		End If
	End If
Next
If Len(strOut) = 0 Then
	LogAddLine "No supported devices"
Else
	LogAddLine "LLD string generated"
End If
strOut = strOut & "]"
Set outFile = objFSO.CreateTextFile(OutPath, True, False)
outFile.Write strOut
outFile.Close
Set outFile = Nothing
LogAddLine "Script finished"
Set objExecObject = Nothing
Set objShell = Nothing
Set objFSO = Nothing
WScript.Echo 0
