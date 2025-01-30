#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#singleinstance, force  		; force, ignore, off
; definisco la finestra principale
Gui, Main:New, +Caption +ToolWindow
Gui, Main:Add, DateTime, vDataInizio x90 y9 w90 h20 Choose20180101,
Gui, Main:Add, Text, x12 y9 w70 h20 , Data inizio:
Gui, Main:Add, Text, x12 y39 w70 h20 , Data fine:
Gui, Main:Add, DateTime, vDataFine x90 y39 w90 h20 Choose%A_YYYY%%A_MM%%A_DD%,
Gui, Main:Add, GroupBox, x4 y69 w185 h50 , Impianti COAL
Gui, Main:Add, CheckBox, vBS x12 y89 w35 h20 , BS
Gui, Main:Add, CheckBox, vFS x57 y89 w35 h20 , FS
Gui, Main:Add, CheckBox, vSU x102 y89 w35 h20 , SU
Gui, Main:Add, CheckBox, vTN x148 y89 w35 h20 , TN
Gui, Main:Add, GroupBox, x4 y129 w185 h70 , Impianti GAS
Gui, Main:Add, CheckBox, vLC x12 y149 w35 h20 , LC
Gui, Main:Add, CheckBox, vMC x57 y149 w35 h20 , MC
Gui, Main:Add, CheckBox, vPC x102 y149 w35 h20 , PC
Gui, Main:Add, CheckBox, vPE x148 y149 w35 h20 , PE
Gui, Main:Add, CheckBox, vPF x12 y169 w35 h20 , PF
Gui, Main:Add, CheckBox, vPG x57 y169 w35 h20 , PG
Gui, Main:Add, CheckBox, vSB x102 y169 w35 h20 , SB
Gui, Main:Add, CheckBox, vTI x148 y169 w35 h20 , TI
Gui, Main:Add, GroupBox, x4 y209 w185 h50 , Implementazioni Future
Gui, Main:Add, CheckBox, vRO x12 y229 w35 h20 , RO
Gui, Main:Add, CheckBox, vFSG x57 y229 w35 h20 , FS
Gui, Main:Add, CheckBox, vOE x102 y229 w35 h20 , OE
Gui, Main:Add, CheckBox, vOF x148 y229 w35 h20 , OF
Gui, Main:Add, GroupBox, x4 y269 w185 h50 , Estrazioni
Gui, Main:Add, CheckBox, vAdM x12 y289 w40 h20 , AdM
Gui, Main:Add, CheckBox, vOdM x65 y289 w40 h20 , OdM
Gui, Main:Add, CheckBox, vAdMSicu x115 y289 w70 h20 , AdM Sicu
Gui, Main:Add, Progress, x4 y317 w185 h5 cAqua BackgroundNavy Range0-3 vMyProgress
Gui, Main:Add, Button, x60 y330 w70 h20 gStart, Elabora
ConfigIcon := A_ScriptDir . "\Icons\Config.png"
Gui, Main:Add, Picture, x155 y327 w30 h-1 gConfig, %ConfigIcon%
gui, Main:Font, cBlack s9 italic
gui, Main:Font, italic
gui, Main:Add, StatusBar, +Left,  by Luca Abrusci
; definisco la finestra di setting
Gui, Settings:New, +Caption +ToolWindow
Gui, Settings:Add, Text, x12 y29 w310 h30 vT_AdMPath +Border,
Gui, Settings:Add, Text, x12 y89 w310 h30 vT_OdMPath +Border,
Gui, Settings:Add, Text, x12 y149 w310 h30 vT_AdMSPath +Border,
Gui, Settings:Add, Button, x332 y29 w60 h30 gBtn_SettingBrowseAdM, Browse
Gui, Settings:Add, Button, x332 y89 w60 h30 gBtn_SettingBrowseOdM, Browse
Gui, Settings:Add, Button, x332 y149 w60 h30 gBtn_SettingBrowseAdMS, Browse
gui, Settings:Font, % "cBlue"
Gui, Settings:Add, Text, x12 y9 w80 h20 , AdM
Gui, Settings:Add, Text, x12 y69 w80 h20 , OdM
Gui, Settings:Add, Text, x12 y129 w80 h20 , AdM Sicurezza
gui, Settings:Font, % "cBlack"
Gui, Settings:Add, Button, x132 y199 w60 h30 gBtn_SettingOK, OK
Gui, Settings:Add, Button, x212 y199 w60 h30 gBtn_SettingCancel, Cancel
; Generated using SmartGUI Creator for SciTE
	;>> leggo il file di configurazione per determinare il percorso in cui salvare i file <<
	; Verifico l'esistenza del file di cfg e che contenga informazioni
	FileConfig := A_ScriptDir . "\config.ini"
	FileGetSize, FileConfigSize, %FileConfig%
	;msgbox % FileConfigSize . " - " . FileConfig
	if (!FileExist(FileConfig) or (FileConfigSize = 0))
	{
	; se non esiste o non contiene valori allora visualizzo la finestra per inserirli
		Gui, Main:Hide
		GuiControl,,T_AdMPath, "Selezionare destinazione estrazioni"
		GuiControl,,T_OdMPath, "Selezionare destinazione estrazioni"
		GuiControl,,T_AdMSPath, "Selezionare destinazione estrazioni"
		Gui, Settings:Show, w408 h300, Seleziona cartella
	}
	; se esite e contiene info recupero i dati contenuti
	else if (FileExist(FileConfig) and (FileConfigSize > 0))
	{
		IniRead, AdMFolder, %FileConfig%, WorkingDirectory, AdMFolder
		IniRead, OdMFolder, %FileConfig%, WorkingDirectory, OdMFolder
		IniRead, AdMSFolder, %FileConfig%, WorkingDirectory, AdMSFolder
		; Aggiorno i valori nella GUI
		GuiControl,,T_AdMPath, %AdMFolder%
		GuiControl,,T_OdMPath, %OdMFolder%
		GuiControl,,T_AdMSPath, %AdMSFolder%
		; verifico se ci sono errori
		if (AdMFolder != "ERROR" and OdMFolder != "ERROR" and AdMSFolder != "ERROR")
		{
			Gui, Settings:Hide
			Gui, Main:Show, w193 h385, Estrai Dati
		}
		else
		{
			MsgBox, 48, Selezionare destinazione estrazioni SAP, Errore nella determinazione della destinazione delle estrazioni SAP.`nSpecificare una destinzione valida.
			Gui, Main:Hide
			Gui, Settings:Show, w408 h250, Seleziona cartella
		}
	}
	; se non esiste mostro la finestra per impostare i percorsi in cui salvare le estrazioni
return

GuiClose:
ExitApp

Config:
Gui, Main:Hide
Gui, Settings:Show, w408 h250, Seleziona cartella
return

Btn_SettingBrowseAdM:
	FileSelectFolder, AdMFolder, *%A_ScriptDir%
	GuiControl,,T_AdMPath, %AdMFolder%
return
Btn_SettingBrowseOdM:
	FileSelectFolder, OdMFolder, *%A_ScriptDir%
	GuiControl,,T_OdMPath, %OdMFolder%
return
Btn_SettingBrowseAdMS:
	FileSelectFolder, AdMSFolder, *%A_ScriptDir%
	GuiControl,,T_AdMSPath, %AdMSFolder%
return
Btn_SettingOK:
; salvo i dati in un file di configurazione
		IniWrite, %AdMFolder%, %FileConfig%, WorkingDirectory, AdMFolder
		IniWrite, %OdMFolder%, %FileConfig%, WorkingDirectory, OdMFolder
		IniWrite, %AdMSFolder%, %FileConfig%, WorkingDirectory, AdMSFolder
		Gui,  Settings:Hide
		Gui, Main:Show, w193 h385, Estrai Dati
return
Btn_SettingCancel:
	Gui,  Settings:Hide
	Gui, Main:Show, w193 h385, Estrai Dati
return

Start:
	Gui, Main:Submit, NoHide
	; Create the array, initially empty:
	Divisioni := [] ; or Array := Array()
	; Write to the array:
	; leggo i valori della variabile Divisione e la modifico per utilizzarla negli scritp SAP
	if (BS = 0) and (FS = 0) and (SU = 0) and (TN = 0) and (AdMSicu != 1)
	{
		MsgBox, 48, Selezionare impianto, Selezionare almeno un impianto
		return
	}
	else
	{
		if (BS = 1)
			Divisioni.Push("ITBS") ; Append this line to the array.
		if (FS = 1)
			Divisioni.Push("ITFS") ; Append this line to the array.
		if (SU = 1)
			Divisioni.Push("ITSU") ; Append this line to the array.
		if (TN = 1)
			Divisioni.Push("ITTN") ; Append this line to the array.
	}
	FormatTime, DataInizio, %DataInizio%, dd.MM.yy
	FormatTime, DataFine, %DataFine%, dd.MM.yy
	; calcolo l'intervallo di anni
	StringRight, Anno_inizio, DataInizio, 2
	StringRight, Anno_fine, DataFine, 2
	Anni := []
	Anni_intervallo := Anno_fine - Anno_inizio + 1
	Loop, %Anni_intervallo%
	{
		if (A_index = 1) {
			Anni[A_Index, 1] := DataInizio
		}
		else {
			Anni[A_Index, 1] := "01.01." . (Anno_inizio + A_index - 1)
		}
		if (A_index = Anni_intervallo) {
			Anni[A_Index, 2] := DataFine
		}
		else {
			Anni[A_Index, 2] := "31.12." . (Anno_inizio + A_index - 1)
		}
		;MsgBox, % Divisioni[%A_Index%, 1] . " - " . Divisioni[%A_Index%, 2]
	}
	gosub Elabora
	msgbox Elaborazione terminata!
return


ESC::ExitApp

;msgbox % StarDate_N . " - " . EndDate_N . "`n" . StarDate_D . " - " . EndDate_D

Elabora:
{
	Progress_Limit := AdM + OdM + AdMSicu
		GuiControl, +Range0-%Progress_Limit%, MyProgress, 0
	if WinExist("ahk_class SAP_FRONTEND_SESSION")
	{
		;try
		{
			WinActivate ; Use the window found by WinExist.
			WinGet, SAP_PID, PID, A ; memorizzo il PID del processo SAP che sto iniziando
			;; Stabilisco una connessione con SAP
			;>>>>>>>>>>>>>>>>>>>>>>>>>>
			_oSAP := ComObjGet("SAPGUI").GetScriptingEngine  ; Get the Already Running Instance
			session := _oSAP.Activesession
			If (session.Busy())
			{
				MsgBox, The Session is Busy
				return
			}
			If (session.Info.IsLowSpeedConnection())
			{
				MsgBox, The Conction is LOW Speed
				return
			}
			;>>>>>>>>>>>>>>>>>>>>>>>>>>
			; --------- Estrazione AdM ---------------
			MyProgress := 0
			GuiControl,, MyProgress, 0
			if (AdM = 1)
			{
				; Leggo gli impianti selezionati dall'array
				for index, element in Divisioni ; Enumeration is the recommended approach in most cases.
				{
					Divisione := element
					Loop, %Anni_intervallo% ; Enumeration is the recommended approach in most cases.
					{
						; Avvio gli Script
						ExportPath := AdMFolder . "\"
						NomeFile := "AdM_" . Divisione . "_" . (Anno_inizio + A_index - 1) . ".xlsx"
						;DB("TimeStamp file: " . T_File)
						;session.findById("wnd[0]").maximize
						session.findById("wnd[0]/tbar[0]/okcd").text := "/nIW29"
						session.findById("wnd[0]/tbar[0]/btn[0]").press
						sleep, 500
						session.findById("wnd[0]/usr/chkDY_OFN").selected := true
						session.findById("wnd[0]/usr/chkDY_RST").selected := true
						session.findById("wnd[0]/usr/chkDY_IAR").selected := true
						session.findById("wnd[0]/usr/chkDY_MAB").selected := true
						session.findById("wnd[0]/usr/ctxtDATUV").text := ""
						session.findById("wnd[0]/usr/ctxtDATUB").text := ""
						session.findById("wnd[0]/usr/ctxtERDAT-LOW").text := Anni[A_Index, 1]
						session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text := Anni[A_Index, 2]
						session.findById("wnd[0]/usr/ctxtSWERK-LOW").text := Divisione
						; session.findById("wnd[0]/usr/ctxtVARIANT").text := "/KPIADMPOOL"
						session.findById("wnd[0]/usr/ctxtVARIANT").text := "KPIADMPOOL"  ; uso un layout specifico utente per evitare manomissioni
						session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
						session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition := 11
						session.findById("wnd[0]/tbar[1]/btn[8]").press
						while session.Busy()
						{
							sleep, 1000
							DB("Busy")
						}
						sleep, 1000
						session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
						sleep, 1000
						win_Title := "Visualizzare avvisi: lista avvisi"
						WinWaitActive, %win_Title%, , 3
						session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
						session.findById("wnd[1]/tbar[0]/btn[0]").press
						session.findById("wnd[1]/usr/ctxtDY_PATH").text := ExportPath
						session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := NomeFile
						session.findById("wnd[1]/tbar[0]/btn[11]").press

						; attendo che si apra il workbook Excel
						; attendo che si apra il workbook Excel
						WinWait, % NomeFile . " - Excel",, 90 ; max 90 secondi
						if ErrorLevel
						{
							MsgBox, WinWait timed out.
							;return false
						}
						else
						{
							sleep, 5000
							xl := ComObjActive("Excel.Application")
							wb := xl.application.Workbooks.Item(NomeFile)
							wb.close(1) ; save and close
							xl.quit()
							xl:=wb:=""
						}
					}
				}
				DB("Estrazione AdM terminata")
				GuiControl,, MyProgress, +1
			}
			; --------- Estrazione OdM ---------------
			if (OdM = 1)
			{
				; Leggo gli impianti selezionati dall'array
				for index, element in Divisioni ; Enumeration is the recommended approach in most cases.
				{
					Divisione := element
					Loop, %Anni_intervallo% ; Enumeration is the recommended approach in most cases.
					{
						; Avvio gli Script
						ExportPath := OdMFolder . "\"
						NomeFile := "OdM_" . Divisione . "_" . (Anno_inizio + A_index - 1) . ".xlsx"
						;DB("TimeStamp file: " . T_File)
						session.findById("wnd[0]/tbar[0]/okcd").text := "/nIW39"
						session.findById("wnd[0]").sendVKey(0)
						session.findById("wnd[0]/usr/chkDY_OFN").selected := true
						session.findById("wnd[0]/usr/chkDY_IAR").selected := true
						session.findById("wnd[0]/usr/chkDY_HIS").selected := true
						session.findById("wnd[0]/usr/chkDY_MAB").selected := true
						session.findById("wnd[0]/usr/ctxtDATUV").text := ""
						session.findById("wnd[0]/usr/ctxtDATUB").text := ""
						session.findById("wnd[0]/usr/ctxtERDAT-LOW").text := Anni[A_Index, 1]
						session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text := Anni[A_Index, 2]
						session.findById("wnd[0]/usr/ctxtIWERK-LOW").text := "" ; DivPianifManut
						session.findById("wnd[0]/usr/ctxtSWERK-LOW").text := Divisione ; Divis. ubic.
						;session.findById("wnd[0]/usr/ctxtVARIANT").text := "/KPIODMPOOL"
						session.findById("wnd[0]/usr/ctxtVARIANT").text := "KPIODMPOOL"				; uso un layout specifico utente per evitare manomissioni
						session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
						session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition := 11
						session.findById("wnd[0]/tbar[1]/btn[8]").press
						while session.Busy()
						{
							sleep, 1000
							DB("Busy")
						}
						sleep, 1000
						session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
						sleep, 1000
						win_Title := "Visualizzare ordini PM: lista ordini"
						WinWaitActive, %win_Title%, , 3
						session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
						session.findById("wnd[1]/tbar[0]/btn[0]").press
						session.findById("wnd[1]/usr/ctxtDY_PATH").text := ExportPath
						session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := NomeFile
						session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition := 13
						session.findById("wnd[1]/tbar[0]/btn[11]").press
						session.findById("wnd[0]/tbar[0]/btn[15]").press
						session.findById("wnd[0]/tbar[0]/btn[15]").press
						; attendo che si apra il workbook Excel
						WinWait, % NomeFile . " - Excel",, 90 ; max 30 secondi
						if ErrorLevel
						{
							MsgBox, WinWait timed out.
							;return false
						}
						else
						{
							sleep, 5000
							xl := ComObjActive("Excel.Application")
							wb := xl.application.Workbooks.Item(NomeFile)
							wb.close(1) ; save and close
							xl.quit()
							xl:=wb:=""
						}
					}
				}
				DB("Estrazione OdM terminata")
				GuiControl,, MyProgress, +1
			}
			; --------- Estrazione AdM SICU ---------------
			if (AdMSicu = 1)
			{
				; Avvio gli Script
				ExportPath := AdMSFolder . "\"
				NomeFile := "export_SICUAMBI.XLSX"
				;DB("TimeStamp file: " . T_File)
				;session.findById("wnd[0]").maximize
				session.findById("wnd[0]/tbar[0]/okcd").text := "/nIW29"
				session.findById("wnd[0]").sendVKey(0)
				session.findById("wnd[0]/usr/ctxtDATUV").text := ""
				session.findById("wnd[0]/usr/ctxtDATUB").text := ""
				session.findById("wnd[0]/usr/chkDY_OFN").selected := true
				session.findById("wnd[0]/usr/chkDY_RST").selected := true
				session.findById("wnd[0]/usr/chkDY_IAR").selected := true
				session.findById("wnd[0]/usr/chkDY_MAB").selected := true
				session.findById("wnd[0]/usr/ctxtDATUB").setFocus
				session.findById("wnd[0]/usr/ctxtDATUB").caretPosition := 1
				session.findById("wnd[0]/usr/btn%P013040_1000").press
				session.findById("wnd[0]/usr/subCLASS_AND_ECM:SAPLCLSD:0210/ctxtCLSELINPUT-CLASS").text := "IT_QUALIFICAZ_AVV"
				session.findById("wnd[0]/usr/subCLASS_AND_ECM:SAPLCLSD:0210/ctxtCLSELINPUT-KLART").text := "015"
				;session.findById("wnd[0]/usr/subCLASS_AND_ECM:SAPLCLSD:0210/ctxtCLSELINPUT-CLASS").caretPosition := 17
				session.findById("wnd[0]").sendVKey(0)
				session.findById("wnd[0]").sendVKey(4)
				session.findById("wnd[1]/usr/tblSAPLCTMSVALUE_S/chkRCTMS-SEL01[0,5]").selected := true
				session.findById("wnd[1]/usr/tblSAPLCTMSVALUE_S/chkRCTMS-SEL01[0,7]").selected := true
				session.findById("wnd[1]/usr/tblSAPLCTMSVALUE_S/chkRCTMS-SEL01[0,7]").setFocus
				session.findById("wnd[1]/tbar[0]/btn[8]").press
				session.findById("wnd[0]/tbar[1]/btn[9]").press
				while session.Busy()
					{
						sleep, 1000
						DB("Busy")
					}
				session.findById("wnd[0]/shellcont[1]/shell").contextMenu
				win_Title := "Ricercare oggetti in classi"
				WinWaitActive, %win_Title%, , 3
				sleep, 1000
				session.findById("wnd[0]/shellcont[1]/shell").selectContextMenuItem("&XXL")
				session.findById("wnd[1]/tbar[0]/btn[0]").press
				session.findById("wnd[1]/usr/ctxtDY_PATH").text := ExportPath
				session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := NomeFile
				session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
				session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition := 20
				session.findById("wnd[1]/tbar[0]/btn[11]").press
				; attendo che si apra il workbook Excel
				WinWait, % NomeFile . " - Excel",, 30 ; max 30 secondi
				if ErrorLevel
				{
					MsgBox, WinWait timed out.
					;return false
				}
				else
				{
					sleep, 1000
					xl := ComObjActive("Excel.Application")
					wb := xl.application.Workbooks.Item(NomeFile)
					wb.close(1) ; save and close
					xl.quit()
					xl:=wb:=""
				}
				DB("Estrazione AdM Sicurezza terminata")
				GuiControl,, MyProgress, +1
			}
			session.findById("wnd[0]/tbar[0]/okcd").text := "/n"
			session.findById("wnd[0]").sendVKey(0)
			_oSAP := session := ""
		}
/*		catch
		{
				_oSAP := session := ""
				msgbox Errore SAP
				ExitApp
		}
*/
	}
	else
	{
		msgbox sessione SAP NON attiva
	}
}

return

; >>>>> FUNZIONI <<<<<<<
;------------------------------------------------------------
CheckWB(xl, WBname)
;------------------------------------------------------------
; verifica l'avvenuta scrittura del file su disco.
{
	TIMEOUT_FileWrite := 180000
	StartTime := A_TickCount
		while ((A_TickCount - StartTime) < TIMEOUT_FileWrite) ;in questo intervallo di tempo
		{
			for wb in xl.application.Workbooks
			{
				msgbox % wb.Name
				if wb.Name = WBname
					return true
			}
			return false
		}
}
;------------------------------------------------------------
CheckFileWrite(TimeStampPre, ExportFile)
;------------------------------------------------------------
; verifica l'avvenuta scrittura del file su disco.
{
	TIMEOUT_FileWrite := 180000

	StartTime := A_TickCount
	if (TimeStampPre > 0)
	{
		FileGetTime, timeStamp, %ExportFile%
		while (timeStamp <= TimeStampPre)
		{
			if  ((A_TickCount - StartTime) < TIMEOUT_FileWrite) ; aspetto
			{
				FileGetTime, timeStamp, %ExportFile%
				sleep, 250
			}
			else
			{
				;DB("ERRORE!")
				return 0
			}
		}
		;DB("1. Tempo trascorso: " . timeStamp - TimeStampPre . " s.")
	}
	else
	{
		ExistBoolean := !FileExist(ExportFile)
		while ExistBoolean
		{
			if  ((A_TickCount - StartTime) < TIMEOUT_FileWrite) ; aspetto
			{
				ExistBoolean := !FileExist(ExportFile)
				sleep, 250
			}
			else
			{
				;DB("ERRORE!")
				return 0
			}
		}
		;DB("2. Tempo trascorso: " . A_TickCount - StartTime . " ms.")
	}
	return 1
}

;------------------------------------------------------------
CheckTimeStampFile(ExportFile)
;------------------------------------------------------------
; verifica l' avvenuta scrittura del file su disco.
{
	if FileExist(ExportFile)
	{
		FileGetTime, timeStamp, %ExportFile%
		;DB("TimeStamp: " . timeStamp)
		return timeStamp
	}
	else
	{
		;DB("Il file non esiste: " . ExportFile)
		return 0
	}
}

;------------------------------------------------------------
NumberOfRowsFile(ExportFile)
;------------------------------------------------------------
{
	FileRead, text, %ExportFile%
	StrReplace(text, "`n", "`n", count)
	count += -6
	;MsgBox % count
	return count
}

;------------------------------------------------------------
NumberOfRows(ExportString)
;------------------------------------------------------------
{
	if (ExportString <> "")
	{
		RegExReplace(ExportString, "(\R)",,count)
		return count - 5
	}
	else
		return 0

}

DB(Text, Clear=0, LineBreak=1,Exit=0) {
oScite := ComObjActive("SciTe4AHK.Application") ; get pointer to activer SciTe window

IfEqual, Clear, 1, SendMessage, oSciTe.Message(0x111, 420)	; if clear = 1 Clear output windows

IfEqual, LineBreak,1, SetEnv, Text, `r`n%text% ; if lineBreak=1 prepend text with 'r'n

oSciTe.Output(Text)	; send text to SciTe output pane

IfEqual, Exit,1,MsgBox, 36, Exit App?, Exit Application?	; if Exit = 1 ask if want to exit application

ifMsgBox, Yes, ExitApp	;If Msgbox = yes then Exit the application
}
