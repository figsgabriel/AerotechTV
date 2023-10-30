'Programa de configuracao dos computadores da Aerotech TV v1.0
'Ultima atualizacao 30/10/2023

Dim UpdateUrl, ShutdownTime, objNetwork, computerName, DefaultPassword, version, RefreshInterval
Set objNetwork = CreateObject("WScript.Network")

computerName = objNetwork.ComputerName
ShutdownTime = ""

'================= INÍCIO DA ÁREA DE TRABALHO =================

RefreshInterval = 60000 'milissegundos

DefaultPassword = ""

UpdateUrl = "https://raw.githubusercontent.com/figsgabriel/AerotechTV/main/password.txt"

version = "1.0"

Select Case computerName

    Case "1", "2"
        ShutdownTime = "00:10" '2º Turno

    Case "3"
        ShutdownTime = "nunca" '3º Turno

    Case Else
        ShutdownTime = "10:00" '1º Turno

End Select

'=================== FIM DA ÁREA DE TRABALHO =================
Set objNetwork = Nothing
Set objShell = CreateObject("WScript.Shell")
WScript.Echo "       ____         _________    ___"
WScript.Echo "    .d888888b.  .d88888888b. \   888"
WScript.Echo "   d88P****Yb  d88P******Y88b  \ 888"
WScript.Echo "  888*        888   /     888.  |888"
WScript.Echo "  888    88888888  |       888  |888"
WScript.Echo "  888.   ***88888   \     888*  |888"
WScript.Echo "   Y88baaaad88PY8b.  *--.d8P   / 888aaaaaaa"
WScript.Echo "     *Y888888P   *Y8b._______/   8888888888"
WScript.Echo vbCrLf
WScript.Echo "Programa de configuracao Aerotech TV v"& version &" maquina "& computerName
WScript.Echo vbCrLf
WScript.Echo "Este computador sera desligado automaticamente "& ShutdownTime
WScript.Echo "Em execucao, nao feche esta janela..."
WScript.Sleep 5000

'===================== INICIALIZA CHROME ====================

'objShell.run "chrome.exe"

'=========================== LOOP ===========================

If ShutdownTime = "nunca" Then
    objShell.Exec("cmd /c REG ADD ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"" /v ""DefaultPassword"" /t REG_SZ /d "& DefaultPassword &" /f /reg:64")
    Do While True
	WScript.Sleep RefreshInterval
	objShell.SendKeys "{F5}"
    Loop
Else
    endtime = CDate(ShutdownTime)
    Do Until Time > endtime
	WScript.Sleep RefreshInterval
	objShell.SendKeys "{F5}"
    Loop
End If

'======================= LÊ TEXTO ONLINE ====================

WScript.Echo "Atualizando software..."

Dim objHTTP, responseText
Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")

objHTTP.Open "GET", UpdateUrl, False
objHTTP.setRequestHeader "Content-Type", "text/xml"
objHTTP.send ""
responseText = objHTTP.responseText

WScript.Echo "Chave online lida: " & responseText

'======================= SALVA ARQUIVO ====================

'Dim objFSO, objFile
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFile = objFSO.CreateTextFile("C:\Users\gafpereira\Desktop\tv_setup_test.vbs", True)

'objFile.Write responseText
'objFile.Close

WScript.Echo "Atualizacao concluida."

'================ MUDA CHAVES DO REGISTRO ==============

'objShell.Exec("cmd /c REG ADD ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"" /v ""DefaultPassword"" /t REG_SZ /d "& DefaultPassword &" /f /reg:64")
'objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\legalnoticecaption","","REG_SZ"
'objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\legalnoticetext","","REG_SZ"

WScript.Echo "Registro alterado com sucesso."

'======================= DESLIGA =====================

'WScript.Sleep 1000
'objShell.Run "shutdown /s /f /t 0",0
