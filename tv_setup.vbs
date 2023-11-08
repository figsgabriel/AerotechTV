'Programa de configuracao dos computadores da Aerotech TV v1.0
'Ultima atualizacao 06/11/2023

'=================== VERIFICA HOSTNAME =====================

Dim UpdateUrl, ShutdownTime, objNetwork, computerName, DefaultPassword, version, RefreshInterval, filePath, currentCode
Set objNetwork = CreateObject("WScript.Network")

computerName = objNetwork.ComputerName
ShutdownTime = ""

'================= INÍCIO DA ÁREA DE TRABALHO =================

version = "1.0"

RefreshInterval = 60000 '<- Inserir intervalo de atualização do Chrome em milissegundos

filePath = "tv_setup_test_1.vbs"

DefaultPassword = "senha" '<- Inserir senha entre parenteses. Não alterar forma devido ao regex

UpdateUrl = "https://raw.githubusercontent.com/figsgabriel/AerotechTV/main/tv_setup.vbs" '<- Inserir link do da atualização do codigo

Select Case computerName

    Case "CMM230DT084522", "CMM230DT122146", "CMM230DT084543", "CMM230DT083979", "CMM230DT79943", "CMM230DT84815", "CMM230DT85203"
        ShutdownTime = "00:10" '<- Inserir hostnames de computadores do 2º Turno

    Case "CMM230DT084818", "CMM230DT084981", "CMM230DT121801", "CMM230DTCR1T8S1"
        ShutdownTime = "nunca" '<- Inserir hostnames de computadores do 3º Turno

    Case Else
        ShutdownTime = "23:00" '<- 1º Turno

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
WScript.Echo "Programa de configuracao Aerotech TV v"& version &" Hostname "& computerName
WScript.Echo vbCrLf
WScript.Echo "AVISO: Este computador sera desligado automaticamente "& ShutdownTime

'========= DEFINIÇÃO DA FUNÇÃO DE LER TEXTO ONLINE ==========

WScript.Echo "Definindo funcoes.."
Dim objHTTP, regEx, matches
Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
Set regEx = New RegExp

Function Update(url, ByRef responseText, ByRef password)
	objHTTP.Open "GET", url, False
	objHTTP.setRequestHeader "Content-Type", "text/xml"
	objHTTP.send ""
	responseText = objHTTP.responseText

	'== EXTRAI NOVA SENHA DO CODIGO VIA REGEX ==

	regEx.Global = True
	regEx.IgnoreCase = True
	regEx.Pattern = "DefaultPassword = ""(.+)"""
	Set matches = regEx.Execute(responseText)

	If matches.Count > 0 Then
		password = matches(0).Submatches(0)
	Else
    		password = "0"
	End If
End Function

'============= DEFINIÇÃO DO CODE UPDATER ===============

Function Updater(ByRef currentCode)
        
	Update UpdateUrl, updatedCode, Password
	If updatedCode <> currentCode Then
		WScript.Echo "Atualizacao encontrada. Atualizando..."
		WScript.Echo "Senha detectada: " & Password

		'======= REESCREVE ARQUIVO =======
		Dim objFSO, objFile
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.CreateTextFile("filePath", True)
		objFile.Write responseText
		objFile.Close

		'===== ATUALIZA SENHA CHROME ====== [em desenvolvimento]
		'objShell.run "chrome.exe https://www.office.com/?auth=2"

		'======= REESCREVE REGISTRO =======
		'objShell.Exec("cmd /c REG ADD ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"" /v ""DefaultPassword"" /t REG_SZ /d "& Password &" /f /reg:64")
		'objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\legalnoticecaption","","REG_SZ"
		'objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\legalnoticetext","","REG_SZ"
		WScript.Sleep 1000
		'objShell.Run "shutdown /r /t 0",0

		WScript.Echo "Atualizacao concluida."
	'======= FIM DO UPDATER =======
	Else
		WScript.Echo "Nenhuma atualizacao encontrada."
	End If
End Function

'=========== LÊ CHAVE DEFAULTPASSWORD DO REGISTRO ==========

WScript.Echo "Verificando registro..."

PasswordRegRead = objShell.Exec("cmd /c REG QUERY ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"" /v ""DefaultPassword"" /f /reg:64").StdOut.ReadAll

lines = Split(PasswordRegRead, vbCrLf) 'Separa as linhas do resultado em um array

Dim extractedValue 'Extrai o valor da linha que contem "REG_SZ"
For Each line In lines
    If InStr(line, "REG_SZ") > 0 Then
        parts = Split(line, "REG_SZ")
        If UBound(parts) > 0 Then
            extractedValue = Trim(parts(1))
            Exit For
        End If
    End If
Next

WScript.Echo "Registro lido. " & extractedValue

'==================== LÊ CÓDIGO ATUAL ===================

Dim file, content, fso
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile(filePath, 1)
currentCode = file.ReadAll

WScript.Echo "Codigo atual lido."

'===================== INICIALIZA CHROME ====================

WScript.Echo "Em execucao, nao feche esta janela..."
Dim updatedCode, Password
'WScript.Sleep 5000
'objShell.run "chrome.exe"

'=========================== LOOP ===========================

If ShutdownTime = "nunca" Then
    Do While True
	WScript.Sleep RefreshInterval
	'======= CODE UPDATER =======
	Update UpdateUrl, updatedCode, Password
	Updater currentCode
	'======= CODE UPDATER =======
	objShell.SendKeys "{F5}"
    Loop
Else
    endtime = CDate(ShutdownTime)
    Do Until Time > endtime
	WScript.Sleep RefreshInterval
	WScript.Echo "Procurando atualizacoes..."
        '======= CODE UPDATER =======
	Update UpdateUrl, updatedCode, Password
	Updater currentCode
	'======= CODE UPDATER =======
	objShell.SendKeys "{F5}"
    Loop
End If

'========= MUDA CHAVES DO REGISTRO AO DESLIGAR ========

'objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\legalnoticecaption","","REG_SZ"
'objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\legalnoticetext","","REG_SZ"

WScript.Echo "Registro alterado com sucesso. Desligando..."

'======================= DESLIGA =====================

'WScript.Sleep 1000
'objShell.Run "shutdown /s /f /t 0",0
