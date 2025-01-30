#Requires AutoHotkey v2.0

#Include SAP_Connection.ahk

class SAP_Estrai_Manager {

    static __New() {
        SAP_Estrai_Manager.InitializeVariables()
    }

    Static InitializeVariables() {

    }

    Static WriteToTextFile(content, filePath) {
        try {
            ; Verifica che il contenuto non sia vuoto
            if (content = "") {
                throw Error("Il contenuto da scrivere è vuoto")
            }
            
            ; Verifica che il percorso non sia vuoto
            if (filePath = "") {
                throw Error("Il percorso del file non è specificato")
            }
            
            ; Verifica che la directory esista
            SplitPath(filePath,, &dir)
            if !DirExist(dir) {
                throw Error("La directory di destinazione non esiste: " . dir)
            }
            
            ; Se il file esiste lo elimina
            if FileExist(filePath)
                FileDelete(filePath)
            
            ; Scrive il file
            FileAppend(content, filePath, "UTF-8")
            return true
            
        } catch Error as err {
            throw Error("Errore nella scrittura del file: " . err.Message)
        }
    }    

    Static EstraiDati(info) {
        ; info := {estrazioni: arr_SAP_estrazioni, divisioni: arr_divisioni, dataInizio:"" ,dataFine: ""}
        ; calcolo l'intervallo di anni
        global isRunning := true

        Anno_inizio := SubStr(info.dataInizio, -2)
        Anno_fine := SubStr(info.dataFine, -2)
        Anni := []
        Anni_intervallo := Anno_fine - Anno_inizio + 1
        Loop Anni_intervallo
        {
            date:={dataInizio:"", dataFine:""}
            if (A_index = 1) {
                date.dataInizio := info.dataInizio
            }
            else {
                date.dataInizio := "01.01." . (Anno_inizio + A_index - 1)
            }
            if (A_index = Anni_intervallo) {
                date.dataFine := info.dataFine
            }
            else {
                date.dataFine := "31.12." . (Anno_inizio + A_index - 1)
            }
            Anni.Push(date)
        }
        ; avvio una sessione SAP
        session := SAPConnection.GetSession()
        errori := ""
        if (session) {
            try {
                for estrazione in info.estrazioni {
                    if (!isRunning) {  ; Controlla se l'esecuzione è stata interrotta
                        break
                    }
                    for divisione in info.divisioni {
                        if (!isRunning) {  ; Controlla se l'esecuzione è stata interrotta
                            break
                        }
                        Loop Anni_intervallo {
                            if (!isRunning) {  ; Controlla se l'esecuzione è stata interrotta
                                break
                            }
                            date := anni[a_index]
                            ; Rendi la finestra SAP sempre in primo piano
                            WinSetAlwaysOnTop(true, "ahk_class SAP_FRONTEND_SESSION")
                            if (estrazione = "AdM")
                                MainGUI.SetStatusBarText("AdM - " . divisione . " - " . date.dataInizio . " - " . date.dataFine)
                                errori .= SAP_Estrai_Manager.Estrai_AdM(session, divisione, date)
                            if (!isRunning) {  ; Controlla se l'esecuzione è stata interrotta
                                break
                            }
                            if (estrazione = "OdM")
                                MainGUI.SetStatusBarText("OdM - " . divisione . " - " . date.dataInizio . " - " . date.dataFine)
                                errori .= SAP_Estrai_Manager.Estrai_OdM(session, divisione, date)                               
                        } ; estrazioni da fare per ogni divisione
                        if (estrazione = "Ptw") 
                            MainGUI.SetStatusBarText("PtW - " . divisione . " - " . info.dataFine)
                            errori .= SAP_Estrai_Manager.Estrai_Ptw(session, divisione, info.dataFine) ; è una estrazione riferita al mese corrente. Dalla data fine calcolo l'intervallo del mese.
                    }
                    ; queste estrazioni devono essere eseguite senza divisione ne date
                    if (estrazione = "SicuAmbi")
                        MainGUI.SetStatusBarText("SicuAmbi - " . divisione)
                        errori .= SAP_Estrai_Manager.Estrai_SicuAmbi(session) ; è una singola estrazione senza specificare divisione e data
                }
                ; rimuovo finestra in primo piano
                if (isRunning) ; se ho interrotto non esiste più la sessione SAP
                    WinSetAlwaysOnTop(false, "ahk_class SAP_FRONTEND_SESSION")
                if(StrLen(errori)>0) {
                    SAP_Estrai_Manager.WriteToTextFile(errori, A_ScriptDir . "\debug.txt")                    
                }
                if (isRunning) {  ; Se l'esecuzione non è stata interrotta
                    MainGUI.SetStatusBarText("Elaborazione termianta")
                    msgbox("Elaborazione terminata:`n" . errori, "Info Estrazioni", 4160)
                }                
                
 
            } catch as err {
                ; rimuovo finestra in primo piano
                WinSetAlwaysOnTop(false, "ahk_class SAP_FRONTEND_SESSION")                
                MsgBox("Errore nell'esecuzione dell'azione SAP: " err.Message, "Errore", 4112)
                return false
            }
        }
        else {
            MsgBox("Impossibile ottenere una sessione SAP valida.", "Errore", 4112)
            return false
        }
    } 
    
    Static Estrai_AdM(session, divisione, date) {
        exportPath := DirectoryManager.GetCategoryPath("AdM") . "\"
        nomeFile := "AdM_" . divisione . "_" . (SubStr(date.dataInizio, -2)) . ".xlsx"
        save := false
        errorMsg := ""
        if (session) {
            try {
                ; eseguo transazione
                ; memorizzo hwnd della finestra per renderal attiva
                hwnd := session.findById("wnd[0]").Handle
                WinActivate("ahk_id " . hwnd)
                session.findById("wnd[0]/tbar[0]/okcd").text := "/nIW29"
                session.findById("wnd[0]/tbar[0]/btn[0]").press
                session.findById("wnd[0]/usr/chkDY_OFN").selected := true
                session.findById("wnd[0]/usr/chkDY_RST").selected := true
                session.findById("wnd[0]/usr/chkDY_IAR").selected := true
                session.findById("wnd[0]/usr/chkDY_MAB").selected := true
                session.findById("wnd[0]/usr/ctxtDATUV").text := ""
                session.findById("wnd[0]/usr/ctxtDATUB").text := ""
                session.findById("wnd[0]/usr/ctxtERDAT-LOW").text := date.dataInizio
                session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text := date.dataFine
                session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text := divisione . "*"
                session.findById("wnd[0]/usr/ctxtSWERK-LOW").text := "" ; divisione
                session.findById("wnd[0]/usr/ctxtBUKRS-LOW").text := "IT1B" ; società
                ; session.findById("wnd[0]/usr/ctxtVARIANT").text := "/KPIADMPOOL"
                session.findById("wnd[0]/usr/ctxtVARIANT").text := "KPIADMPOOL"  ; uso un layout specifico utente per evitare manomissioni
                session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
                session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition := 11
                ; rilevo il nome e la classe della finestra attiva
                WinActivate("ahk_id " . hwnd)
                activeTitle := WinGetTitle("A")
                activeClass := WinGetClass("A")
                OutputDebug("1) ActiveTitle:" activeTitle . "`n")
                OutputDebug("2) activeClass:" activeClass . "`n")
                session.findById("wnd[0]/tbar[1]/btn[8]").press
                while session.Busy()
                    {
                        sleep 1000
                    }
                sleep 1000
                ; Ci sono tre possibilità:                
                ; 1) --> nessun AdM -> nessuna nuova finsetra
                StatusBarMessage := session.FindById("wnd[0]/sbar").Text
                if(InStr(StatusBarMessage, "Non sono stati selezionati oggetti")) {
                    OutputDebug("Nessun AdM per la divisione: " . divisione . " nell'intervallo: " . date.dataInizio . " - " . date.dataFine . "`n")
                    ;errorMsg := "Nessun AdM per la divisione: " . divisione . " nell'intervallo: " . date.dataInizio . " - " . date.dataFine . "`n"
                    return ""
                }
                ; Verifico il tipo di nuova finestra che viene attivata
                WinActivate("ahk_id " . hwnd)
                if WinWaitActive("A",, 5) {  ; attende fino a 5 secondi
                    activeTitle := WinGetTitle("A")
                    activeClass := WinGetClass("A")
                    OutputDebug("3) ActiveTitle:" activeTitle . "`n")
                    OutputDebug("4) activeClass:" activeClass . "`n")

                    ; 2) --> lista di AdM
                    if (InStr(activeTitle, "Visualizzare avvisi: lista avvisi") && activeClass = "SAP_FRONTEND_SESSION") {
                        OutputDebug("Trovata finestra Lista Avvisi`n")
                        save := true
                    }                   
                    ; 3) --> un solo AdM
                    else if (InStr(activeTitle, "Visualizzare avviso PM:") && activeClass = "SAP_FRONTEND_SESSION") {
                        OutputDebug("Trovata finestra Singolo Avviso`n")
                        errorMsg := "Rilevato un unico AdM per la divisione: " . divisione . " nell'intervallo: " . date.dataInizio . " - " . date.dataFine . "`n"
                        return errorMsg
                    }
                }               
                ; verifica ulteriore contanto le righe della tabella lista avvisi  
                try {
                    countRows := session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
                    OutputDebug("IW29 " divisione . " - " date.dataInizio " - " date.dataFine " - Numero di elementi in tabella: " . countRows . "`n")
                    save := true
                }
                catch {
                    OutputDebug("IW29 " divisione . " - " date.dataInizio " - Non sono presenti elementi in tabella" . "`n")
                    save := false
                }
                if(save = true) { ; salvo solo se la tabella contiene dati                
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
                    sleep 1000
                    win_Title := "Visualizzare avvisi: lista avvisi"
                    WinWaitActive(win_Title, , 3)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
                    session.findById("wnd[1]/tbar[0]/btn[0]").press
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text := exportPath
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := nomeFile
                    session.findById("wnd[1]/tbar[0]/btn[11]").press
                    while session.Busy()
                        {
                            sleep 1000
                        }                
                    ; attendo che si apra il workbook Excel
                    if WinWait(nomeFile . " - Excel",, 90) {
                        sleep 5000
                        xl := ComObjActive("Excel.Application")
                        wb := xl.application.Workbooks.Item(nomeFile)
                        wb.close(1) ; save and close
                        ;xl.quit()
                        xl:= wb:= ""
                    }
                    else
                        MsgBox("WinWait Excel timed out.", "Errore", 4112)
                }
            } catch as err {
                MsgBox("Errore nell'esecuzione dell'azione SAP: " err.Message, "Errore", 4112)
                return ""
            } finally {
                SAPConnection.Disconnect() 
            }
        }
        else {
            MsgBox("Impossibile ottenere una sessione SAP valida.", "Errore", 4112)
            return ""
        }
    }
    
    Static Estrai_OdM(session, divisione, date) {
        exportPath := DirectoryManager.GetCategoryPath("OdM") . "\"
        nomeFile := "OdM_" . divisione . "_" . (SubStr(date.dataInizio, -2)) . ".xlsx"
        save := false
        if (session) {
            try {
                ; memorizzo hwnd della finestra per renderal attiva
                hwnd := session.findById("wnd[0]").Handle
                WinActivate("ahk_id " . hwnd)                
                session.findById("wnd[0]/tbar[0]/okcd").text := "/nIW39"
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[0]/usr/chkDY_OFN").selected := true
                session.findById("wnd[0]/usr/chkDY_IAR").selected := true
                session.findById("wnd[0]/usr/chkDY_HIS").selected := true
                session.findById("wnd[0]/usr/chkDY_MAB").selected := true
                session.findById("wnd[0]/usr/ctxtDATUV").text := ""
                session.findById("wnd[0]/usr/ctxtDATUB").text := ""
                session.findById("wnd[0]/usr/ctxtERDAT-LOW").text := date.dataInizio
                session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text := date.dataFine
                session.findById("wnd[0]/usr/ctxtIWERK-LOW").text := "" ; DivPianifManut
                session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text := divisione . "*"
                session.findById("wnd[0]/usr/ctxtBUKRS-LOW").text := "IT1B" ; società
                ;session.findById("wnd[0]/usr/ctxtVARIANT").text := "/KPIODMPOOL"
                session.findById("wnd[0]/usr/ctxtVARIANT").text := "KPIODMPOOL"				; uso un layout specifico utente per evitare manomissioni
                session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
                session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition := 11
                ; rilevo il nome e la classe della finestra attiva
                WinActivate("ahk_id " . hwnd)
                activeTitle := WinGetTitle("A")
                activeClass := WinGetClass("A")
                OutputDebug("1) ActiveTitle:" activeTitle . "`n")
                OutputDebug("2) activeClass:" activeClass . "`n")                
                session.findById("wnd[0]/tbar[1]/btn[8]").press
                while session.Busy()
                {
                    sleep 1000
                }
                sleep 1000
                ; Ci sono tre possibilità:                
                ; 1) --> nessun OdM -> nessuna nuova finsetra
                StatusBarMessage := session.FindById("wnd[0]/sbar").Text
                if(InStr(StatusBarMessage, "Non sono stati selezionati oggetti")) {
                    OutputDebug("Nessun OdM per la divisione: " . divisione . " nell'intervallo: " . date.dataInizio . " - " . date.dataFine . "`n")
                    ;errorMsg := "Nessun OdM per la divisione: " . divisione . " nell'intervallo: " . date.dataInizio . " - " . date.dataFine . "`n"
                    return ""
                }
                ; Verifico il tipo di nuova finestra che viene attivata
                WinActivate("ahk_id " . hwnd)
                if WinWaitActive("A",, 5) {  ; attende fino a 5 secondi
                    activeTitle := WinGetTitle("A")
                    activeClass := WinGetClass("A")
                    OutputDebug("3) ActiveTitle:" activeTitle . "`n")
                    OutputDebug("4) activeClass:" activeClass . "`n")

                    ; 2) --> lista di OdM
                    if (InStr(activeTitle, "Visualizzare ordini PM: lista ordini") && activeClass = "SAP_FRONTEND_SESSION") {
                        OutputDebug("Trovata finestra Lista Ordini`n")
                        save := true
                    }                   
                    ; 3) --> un solo OdM
                    else if (InStr(activeTitle, "Visualizzare Manutenzione Correttiva Termo") && activeClass = "SAP_FRONTEND_SESSION") {
                        OutputDebug("Trovata finestra Singolo Ordine`n")
                        errorMsg := "Rilevato un unico AdM per la divisione: " . divisione . " nell'intervallo: " . date.dataInizio . " - " . date.dataFine . "`n"
                        return errorMsg
                    }
                }               
                ; verifica ulteriore contanto le righe della tabella lista avvisi  
                try {
                    countRows := session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
                    OutputDebug("IW39 " divisione . " - " date.dataInizio " - " date.dataFine " - Numero di elementi in tabella: " . countRows . "`n")
                    save := true
                }
                catch {
                    OutputDebug("IW39 " divisione . " - " date.dataInizio " - Non sono presenti elementi in tabella" . "`n")
                    save := false
                }
                if(save = true) { ; salvo solo se la tabella contiene dati
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
                    sleep 1000
                    win_Title := "Visualizzare ordini PM: lista ordini"
                    WinWaitActive(win_Title, , 3)                
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
                    session.findById("wnd[1]/tbar[0]/btn[0]").press
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text := exportPath
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := nomeFile
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition := 13
                    session.findById("wnd[1]/tbar[0]/btn[11]").press
                    session.findById("wnd[0]/tbar[0]/btn[15]").press
                    session.findById("wnd[0]/tbar[0]/btn[15]").press
                    while session.Busy()
                        {
                            sleep 1000
                        }                 
                    ; attendo che si apra il workbook Excel
                    if WinWait(nomeFile . " - Excel",, 90) {
                        sleep 5000
                        xl := ComObjActive("Excel.Application")
                        wb := xl.application.Workbooks.Item(nomeFile)
                        wb.close(1) ; save and close
                        ;xl.quit()
                        xl:= wb:= ""
                    }
                    else
                        MsgBox("WinWait Excel timed out.", "Errore", 4112)
                }
            } 
            catch as err {
                MsgBox("Errore nell'esecuzione dell'azione SAP: " err.Message, "Errore", 4112)
                return ""
            } finally {
                SAPConnection.Disconnect() 
            }
        }
        else {
            MsgBox("Impossibile ottenere una sessione SAP valida.", "Errore", 4112)
            return ""
        }
    }

    Static Estrai_SicuAmbi(session) {
        if (session) {
            exportPath := DirectoryManager.GetCategoryPath("SicuAmbi") . "\"
            nomeFile := "export_SICUAMBI.XLSX"            
            try {
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
                        sleep 500
                        OutputDebug("SAP is busy" . "`n")
                    }
                session.findById("wnd[0]/shellcont[1]/shell").contextMenu
                win_Title := "Ricercare oggetti in classi"
                WinWaitActive(win_Title, , 3)
                sleep 1000
				session.findById("wnd[0]/shellcont[1]/shell").selectContextMenuItem("&XXL")
				session.findById("wnd[1]/tbar[0]/btn[0]").press
				session.findById("wnd[1]/usr/ctxtDY_PATH").text := ExportPath
				session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := NomeFile
				session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
				session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition := 20
				session.findById("wnd[1]/tbar[0]/btn[11]").press
                while session.Busy()
                    {
                        sleep 1000
                    }                 
                ; attendo che si apra il workbook Excel
                if WinWait(nomeFile . " - Excel",, 90) {
                    sleep 5000
                    xl := ComObjActive("Excel.Application")
                    wb := xl.application.Workbooks.Item(nomeFile)
                    wb.close(1) ; save and close
                    ;xl.quit()
                    xl:= wb:= ""
                }
                else
                    MsgBox("WinWait Excel timed out.", "Errore", 4112) 
            } catch as err {
                MsgBox("Errore nell'esecuzione dell'azione SAP: " err.Message, "Errore", 4112)
                return ""
            } finally {
                SAPConnection.Disconnect() 
            }                
            }
        else {
            MsgBox("Impossibile ottenere una sessione SAP valida.", "Errore", 4112)
            return ""
        }
    }    

    Static Estrai_Ptw(session, divisione, dataFine) {
        save := false
        dataInizio := SAP_Estrai_Manager.GetFirstDayOfMonth(dataFine)
        exportPath := DirectoryManager.GetCategoryPath("Ptw") . "\"
        nomeFile := "IW49N_export_" . divisione . "_" . dataInizio . "-" . dataFine . ".XLSX"
        if (session) {
            try {
                session.findById("wnd[0]/tbar[0]/okcd").text := "/nIW49N"
                session.findById("wnd[0]/tbar[0]/btn[0]").press
                session.findById("wnd[0]/usr/chkSP_OFN").selected := True
                session.findById("wnd[0]/usr/chkSP_IAR").selected := True
                session.findById("wnd[0]/usr/chkSP_MAB").selected := True
                session.findById("wnd[0]/usr/chkSP_HIS").selected := True
                session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_STRNO-LOW").text := divisione . "*"
                session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text := ""
                session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").select
                sleep 500
                session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GSTRP-LOW").text := dataInizio
                session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GSTRP-HIGH").text := dataFine
                session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").select
                sleep 500
                session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/ctxtS_BUKRS-LOW").text := "IT1B" ; società
                session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").select
                sleep 500
                session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text := "/KPI_PTW"
                ; avvio la transazione
                session.findById("wnd[0]/tbar[1]/btn[8]").press
                sleep 500
                while session.Busy()
                    {
                        sleep 500
                        OutputDebug("SAP is busy" . "`n")
                    }
                ; verifico se sono stati estratti dei dati
                StatusBarMessage := session.FindById("wnd[0]/sbar").Text
                if(InStr(StatusBarMessage, "Non sono stati selezionati oggetti")) {
                    OutputDebug("Nessun OdM per la divisione: " . divisione . " nell'intervallo: " . dataInizio . " - " . dataFine . "`n")
                    ;errorMsg := "Nessun OdM per la divisione: " . divisione . " nell'intervallo: " . date.dataInizio . " - " . date.dataFine . "`n"
                    return ""
                }                
                try {
                    countRows := session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
                    OutputDebug("IW49N " divisione . " - " dataInizio " - " dataFine " - Numero di elementi in tabella: " . countRows . "`n")
                    save := true
                }
                catch {
                    OutputDebug("IW49N " divisione . " - " dataInizio " - " dataFine " - Non sono presenti elementi in tabella" . "`n")
                    return ""
                }                    
                ; esporto i valori se presenti
                if(save = true) { ; salvo solo se la tabella contiene dati            
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
                    sleep 1000
                    win_Title := "Selezionare calcolo costi tabella"
                    WinWaitActive(win_Title, , 3)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
                    session.findById("wnd[1]/tbar[0]/btn[0]").press
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text := exportPath
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := nomeFile
                    session.findById("wnd[1]/tbar[0]/btn[11]").press
                    while session.Busy()
                        {
                            sleep 500
                            OutputDebug("SAP is busy" . "`n")
                        }
                    ; attendo che si apra il workbook Excel
                    if WinWait(nomeFile . " - Excel",, 90) {
                        sleep 5000
                        xl := ComObjActive("Excel.Application")
                        wb := xl.application.Workbooks.Item(nomeFile)
                        wb.close(1) ; save and close
                        ;xl.quit()
                        xl:= wb:= ""
                    }
                    else
                        MsgBox("WinWait Excel timed out.", "Errore", 4112)
                }                 
            } catch as err {
                MsgBox("Errore nell'esecuzione dell'azione SAP: " err.Message, "Errore", 4112)
                return ""
            } finally {
                SAPConnection.Disconnect() 
            }
        }
        else {
            MsgBox("Impossibile ottenere una sessione SAP valida.", "Errore", 4112)
            return ""
        }
    }

    static GetFirstDayOfMonth(data) {
        ; Divide la data nelle sue componenti
        dateParts := StrSplit(data, ".")
        
        ; Ricostruisce la data usando "01" come giorno
        return Format("01.{}.{}", dateParts[2], dateParts[3])
    }
}