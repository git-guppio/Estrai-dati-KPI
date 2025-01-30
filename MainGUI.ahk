#Requires AutoHotkey v2.0

#Include GlobalConstants.ahk
#Include ButtonIconManager.ahk
#Include GestioneDirectory.ahk
#Include GestioneImpianti.ahk
#Include SAP\SAP_Connection.ahk
#Include MainGUI.ahk


class GUI_Manager {
    __New() {
        ; Crea la GUI principale
        this.mainGui := MainGUI(this)
    }
    
    ; Metodo per mostrare la GUI secondaria
    GestioneImpiantiGui() {
            this.mainGui.gui.Opt("+Disabled")  ; Disabilita la finestra principale
            this.ImpiantiGui := ImpiantiGUI(this)
    }

    ; Metodo per mostrare la GUI secondaria
    GestioneDirectoryGui() {
        this.mainGui.gui.Opt("+Disabled")  ; Abilita la finestra principale
        this.DirectoryManagerGui := DirectoryManager(this)
    }
    
    EnableMainGui() {
        this.mainGui.gui.Opt("-Disabled")  ; Disabilita la finestra principale
    }

}

class ImpiantiConfig {
    static configFile := G_CONSTANTS.IMPIANTI_CONFIG_FILE
    static defaultImpianti := Map(
        "COAL", ["BS", "FS", "SU", "TN"],
        "GAS", ["LC", "MC", "PC", "PE", "PF", "PG", "SB", "TI"],
        "ImpFuture", ["RO", "FSP", "0E", "0F"]
    )
    
    static GetImpianti() {
        if !FileExist(this.configFile) {
            this.InitializeDefaultValues()
        }
        return this.LoadFromIni()
    }
    
    static InitializeDefaultValues() {
        for categoria, impianti in this.defaultImpianti {
            impiantiStr := ArrayTools.Join(impianti)
            IniWrite(impiantiStr, this.configFile, categoria, "impianti")
        }
    }
    
    static LoadFromIni() {
        impiantiMap := Map()
        for categoria, _ in this.defaultImpianti {
            try {
                impiantiStr := IniRead(this.configFile, categoria, "impianti")
                impiantiMap[categoria] := StrSplit(impiantiStr, ";", " ")
            }
        }
        return impiantiMap
    }
}

class CheckboxManager {
    static GetCheckedControls(gui) {
        checkedControls := []
        
        ; Itera attraverso tutti i controlli della GUI
        for ctrl in gui {
            ; Verifica se è un checkbox e se è selezionato
            if (ctrl.Type = "Checkbox" && ctrl.Value) {
/*                 checkedControls.Push({
                    name: ctrl.Name,        ; Nome del controllo
                    text: ctrl.Text,        ; Testo del checkbox
                    value: ctrl.Value       ; Stato (1 = checked, 0 = unchecked)
                }) */
                checkedControls.Push(
                    ctrl.Text        ; Testo del checkbox
                )               
            }
        }
        
        return checkedControls ; ritrorna un array con il testo dei checkbox selezionati
    }
}

class MainGUI {
    ; Definizione delle costanti come proprietà statiche
    static MARGIN := 10
    static CHECKBOX_WIDTH := 45
    static ROW_HEIGHT := 20
    static BUTTON_DIM := 40
    static SB := "" ; status bar
    
    __New(mainApp) {
        ; Salva il riferimento all'app principale
        this.mainApp := mainApp        
        this.isChecked := false
        this.CreateMainGui()
    }
    
    UpdateMainGUI() {
        ; Ricrea completamente la GUI principale
        this.mainApp.mainGui.gui.Destroy()  ; Distrugge la GUI corrente
        this.mainApp.mainGui := MainGUI(this.mainApp)  ; Crea una nuova istanza
    }

    CreateMainGui() {
        ; Creazione della GUI principale
        this.gui := Gui("-MinimizeBox -MaximizeBox")
        this.gui.Title := "Estrai dati x KPI"
        this.gui.MarginX := 10
        this.gui.MarginY := 5
        
        ; Aggiunta dei controlli data
        this.AddDateControls()
        
        ; Aggiunta delle checkbox per gli impianti
        this.impiantiMap := ImpiantiConfig.GetImpianti()
        this.AddImpiantiCheckboxes()
        
        ; Aggiunta del gruppo "Estrai"
        this.AddExtractionGroup()
        
        ; Aggiunta dei pulsanti di controllo
        this.AddControlButtons()

        ; aggiungi la StatusBar
        this.AddStatusBar()
        
        ; Mostra la GUI
        this.gui.Show()
    }
  
    AddDateControls() {
        this.gui.Add("Text", "xm ym", "Data Inizio:")
        this.DataInizio := this.gui.Add("DateTime", "x80 yp vData1 w90 Choose20180101 ", "dd.MM.yy")
        lastDay := This.GetLastDayPreviousMonth()
        this.gui.Add("Text", "xm yp+25", "Data Fine:")
        this.dataFine := this.gui.Add("DateTime", "x80 yp vData2 w90 Choose" lastDay, "dd.MM.yy")
    }
    
    AddStatusBar() {
        ;this.SB := this.gui.Add("StatusBar",,"Ready!")
        MainGUI.SB := this.gui.Add("StatusBar")
        MainGUI.SetStatusBarText("Ready!")
    }
    
    static SetStatusBarText(text) {
        MainGUI.SB.SetText(text)
    }
    
    AddImpiantiCheckboxes() {
        currentY := 60
        for categoria, impianti in this.impiantiMap {
            currentY := this.CreateImpiantiGroup(categoria, impianti, currentY)
        }
    }
    
    CreateImpiantiGroup(categoria, impianti, startY) {
        ; Crea il gruppo
        gb := this.gui.AddGroupBox("xm y" startY " h100 w190", categoria)

        currentY := startY + MainGUI.ROW_HEIGHT
        currentX := this.gui.MarginX + MainGUI.MARGIN
        
        ; Aggiungi le checkbox
        boxHeight := MainGUI.ROW_HEIGHT
        for index, impianto in impianti {
            if (Mod(index-1, 4) = 0) and (index > 1) {
                currentY += MainGUI.ROW_HEIGHT
                currentX := this.gui.MarginX + MainGUI.MARGIN
                boxHeight += MainGUI.ROW_HEIGHT
            }
            
            cb := this.gui.Add("CheckBox", "x" currentX " y" currentY " v" impianto, impianto)
            cb.OnEvent("DoubleClick", this.CheckboxDoubleClick.Bind(this))
            currentX += MainGUI.CHECKBOX_WIDTH
        }
        
        ; Ridimensiona il gruppo
        gb.Move(,,, Floor(boxHeight + MainGUI.MARGIN*2.5))
        gb.GetPos(&x, &y, &w, &h)
        return Floor(y + h + MainGUI.MARGIN)
    }

    CheckboxDoubleClick(clickedCheckbox, *) {
        ; Trova il GroupBox contenente la checkbox cliccata
        clickedCheckbox.GetPos(&cbX, &cbY)
        
        ; Cerca il GroupBox che contiene queste coordinate
        for ctl in this.gui {
            if ctl is Gui.GroupBox {
                ctl.GetPos(&gbX, &gbY, &gbW, &gbH)
                
                ; Se la checkbox è in questo GroupBox
                if (cbX >= gbX && cbX <= gbX + gbW && 
                    cbY >= gbY && cbY <= gbY + gbH) {
                    
                    ; Determina il nuovo stato (opposto della checkbox cliccata)
                    newState := clickedCheckbox.Value
                    
                    ; Applica lo stato a tutte le checkbox in questo GroupBox
                    for checkbox in this.gui {
                        if checkbox is Gui.Checkbox {
                            checkbox.GetPos(&ctlX, &ctlY)
                            if (ctlX >= gbX && ctlX <= gbX + gbW && 
                                ctlY >= gbY && ctlY <= gbY + gbH) {
                                checkbox.Value := newState
                            }
                        }
                    }
                    break
                }
            }
        }
    }
    
    AddExtractionGroup() {
        LastPosition := this.GetLowestGroup()
        gb := this.gui.AddGroupBox("xm y" . LastPosition.y + MainGUI.MARGIN . " h100 w190", "Estrai")
        
        ; Aggiungi le checkbox di estrazione
        this.gui.Add("CheckBox", "xm+10 yp+20 vAdM", "AdM")
        this.gui.Add("CheckBox", "xp+45 yp vOdM", "OdM")
        this.gui.Add("CheckBox", "xp+45 yp vSicuAmbi", "SicuAmbi")
        cb := this.gui.Add("CheckBox", "xm+10 yp+20 vPtw", "Ptw")
        
        ; Ridimensiona il gruppo
        gb.GetPos(&x_start, &y_start, &w_start, &h_start)
        cb.GetPos(&x, &y, &w, &h)
        gb.Move(,,,y-y_start+30)
    
    }

    GetLowestGroup() {
        maxY := 0
        
        ; Itera attraverso tutti i controlli della GUI
        for hwnd, ctl in this.gui {
            ; Verifica se il controllo è un GroupBox
            if ctl is Gui.GroupBox {
                ; Ottiene la posizione e le dimensioni del gruppo
                ctl.GetPos(&x, &y, &w, &h)
                
                ; Calcola la posizione Y più bassa del gruppo (y + altezza)
                bottomY := y + h
                
                ; Aggiorna maxY se necessario
                if (bottomY > maxY)
                    maxY := bottomY
            }
        }
        
        return {x:x, y:maxY, w:w, h:h}
    }    
    
    AddControlButtons() {
        LastPosition := this.GetLowestGroup()
        buttonY := LastPosition.y + MainGUI.MARGIN
        
        ; Crea i pulsanti
        this.checkBtn := this.gui.Add("Button", "xm y" buttonY " w" MainGUI.BUTTON_DIM . " h" . MainGUI.BUTTON_DIM, "")
        ButtonIconManager.SetIcon(this.checkBtn, 'shell32.dll',294, 's' . 32)
        this.checkBtn.OnEvent("Click", this.ToggleCheckboxes.Bind(this))
        
        this.extractBtn := this.gui.Add("Button", "x+10 yp w" MainGUI.BUTTON_DIM . " h" . MainGUI.BUTTON_DIM, "")
        ButtonIconManager.SetIcon(this.extractBtn, 'shell32.dll',265, 's' . 32)
        this.extractBtn.OnEvent("Click", this.EstraiDati.Bind(this))      

        this.configDirBtn := this.gui.Add("Button", "x+10 yp w" MainGUI.BUTTON_DIM . " h" . MainGUI.BUTTON_DIM, "")
        ButtonIconManager.SetIcon(this.configDirBtn, 'shell32.dll',267, 's' . 32)
        this.configDirBtn.OnEvent("Click", this.ConfiguraDirectory.Bind(this))        
        
        this.configBtn := this.gui.Add("Button", "x+10 yp w" MainGUI.BUTTON_DIM . " h" . MainGUI.BUTTON_DIM, "")
        ButtonIconManager.SetIcon(this.configBtn, 'shell32.dll',315, 's' . 32)
        this.configBtn.OnEvent("Click", this.ConfiguraImpianti.Bind(this))       
        
        ; Calcola lo spazio per centrare i pulsanti
        totalButtonWidth := 4 * MainGUI.BUTTON_DIM
        space := (LastPosition.w - totalButtonWidth) / 3
        this.extractBtn.Move(this.gui.MarginX + MainGUI.BUTTON_DIM + space,,,)
        this.configDirBtn.Move(this.gui.MarginX + (2 * MainGUI.BUTTON_DIM) + 2*space,,,)
        this.configBtn.Move(this.gui.MarginX + (3 * MainGUI.BUTTON_DIM) + 3*space,,,)
    }
    
    ToggleCheckboxes(*) {
        this.isChecked := !this.isChecked
        
        for ctl in this.gui {
            if ctl is Gui.Checkbox {
                ctl.Value := this.isChecked
            }
        }
        
       ; this.checkBtn.Text := this.isChecked ? "Uncheck" : "Check"
    }
    
    ConfiguraDirectory(*){
        this.mainApp.GestioneDirectoryGui()
    }

    EstraiDati(*) {
        ; Implementare la logica di estrazione dati
        checkedBoxes := CheckboxManager.GetCheckedControls(this.gui) ; contiene i nomi dei checkbox selezionati

        ; Prepara il messaggio
        if (checkedBoxes.Length > 0) {
/*             message := "Checkbox selezionati:`n`n"
            for checkbox in checkedBoxes {
                message .= "- " . checkbox . "`n"
            } */
            ; verifico che sia selezionato almeno un checkbox di estrazione
            SAP_estrazioni := ["AdM", "OdM", "SicuAmbi", "Ptw"]
            arr_SAP_estrazioni:=[]
            arr_divisioni := []
            for estrazione in SAP_estrazioni {
                if ArrayTools.HasElement(checkedBoxes, estrazione) {
                    arr_SAP_estrazioni.Push(estrazione)
                }
            }
            ; verifico se l'array contiene elementi, ovvero se sono state selezionate estrazioni da fare
            if (arr_SAP_estrazioni.Length = 0) {
                message .= "Selezionare almeno un'estrazione da compiere"    
            }
            ; se ci sono estrazioni allora verifico se ci sono impianti selezionati
            else if (arr_SAP_estrazioni.Length > 0) and  (checkedBoxes.Length > arr_SAP_estrazioni.Length) {
                arr_divisioni := ArrayTools.ArrayDifference(checkedBoxes, arr_SAP_estrazioni)
                if (arr_divisioni.Length > 0) {
                    result := MsgBox("Desideri eliminare tutte le estrazioni precedenti?",
                                   "Elimina dati precedenti",
                                   "YesNo Icon?")
                    if (result = "Yes") {
                        ; elimino il contenuto delle estrazioni precedenti
                        ; -- directory AdM
                        Path_to_delete := DirectoryManager.GetCategoryPath("AdM")
                        this.ClearDirectoryByPrefix(Path_to_delete, "AdM_")
                        ; -- directory OdM
                        Path_to_delete := DirectoryManager.GetCategoryPath("OdM")
                        this.ClearDirectoryByPrefix(Path_to_delete, "OdM_")
                        ; -- directory PtW
                        Path_to_delete := DirectoryManager.GetCategoryPath("PtW")
                        this.ClearDirectoryByPrefix(Path_to_delete, "IW49N_export_")
                        ; -- directory SicuAmbi
                        Path_to_delete := DirectoryManager.GetCategoryPath("SicuAmbi")
                        this.ClearDirectoryByPrefix(Path_to_delete, "export_SICUAMBI") 
                    } 
                    ; avvia la routine per l'estrazione in SAP
                    info := {estrazioni: arr_SAP_estrazioni, divisioni: arr_divisioni, dataInizio: FormatTime(this.DataInizio.value, "dd.MM.yy") ,dataFine: FormatTime(this.dataFine.value, "dd.MM.yy")}
                    SAP_Estrai_Manager.EstraiDati(info)
/*                     message .= "`nEstrazioni:`n`n"
                    for estrazioni in info.estrazioni {
                        message .= "- " . estrazioni . "`n"
                    }
                    message .= "`nDivisioni:`n`n"
                    for divisioni in info.divisioni {
                        message .= "- " . divisioni . "`n"
                    } */
                }
            }
            else {
                MsgBox("Selezionare almeno un impianto", "Errore", 4112)
            }
        }
        else {
            MsgBox("Nessun checkbox selezionato", "Errore", 4112)
        } 
    }
    
    ConfiguraImpianti(*) {
        ; Crea una nuova istanza della GUI di configurazione
        this.mainApp.GestioneImpiantiGui()
    }

    ClearDirectoryByPrefix(dirPath, filePrefix) {
        try {
            ; Verifica che la directory esista
            if !DirExist(dirPath) {
                throw Error("La directory non esiste: " . dirPath)
            }
    
            ; Verifica che il prefisso non sia vuoto
            if (filePrefix = "") {
                throw Error("Il prefisso non può essere vuoto")
            }
    
            ; Conta i file eliminati
            deletedCount := 0
    
            ; Cerca ed elimina i file che iniziano con il prefisso specificato
            Loop Files dirPath . "\" . filePrefix . "*.xlsx" {
                FileDelete(A_LoopFilePath)
                deletedCount++
            }
    
            return deletedCount
            
        } catch Error as err {
            throw Error("Errore nella pulizia della directory: " . err.Message)
        }
    }

    GetLastDayPreviousMonth() {
        ; Ottieni data corrente nel formato dd.mm.yy
        currentDate := FormatTime(A_Now, "dd.MM.yy")
        
        ; Estrai componenti della data
        dateParts := StrSplit(currentDate, ".")
        giorno := Integer(dateParts[1])
        mese := Integer(dateParts[2])
        anno := Integer("20" . dateParts[3])  ; Converte yy in yyyy
        
        ; Determina l'ultimo giorno del mese corrente
        giornoFinale := 31  ; Default per i mesi con 31 giorni
        if (mese = 4 || mese = 6 || mese = 9 || mese = 11)
            giornoFinale := 30
        else if (mese = 2) {  ; Febbraio
            if (Mod(anno, 4) = 0 && (Mod(anno, 100) != 0 || Mod(anno, 400) = 0))
                giornoFinale := 29  ; Anno bisestile
            else
                giornoFinale := 28  ; Anno non bisestile
        }
        
        ; Se è l'ultimo giorno del mese, restituisci la data corrente
        if (giorno = giornoFinale)
            return Format("{:04d}{:02d}{:02d}", anno, mese, giorno)
            
        ; Altrimenti calcola l'ultimo giorno del mese precedente
        mese := mese - 1
        if (mese = 0) {
            mese := 12
            anno := anno - 1
        }
        
        ; Calcola l'ultimo giorno del mese precedente
        giornoFinale := 31  ; Default per i mesi con 31 giorni
        if (mese = 4 || mese = 6 || mese = 9 || mese = 11)
            giornoFinale := 30
        else if (mese = 2) {  ; Febbraio
            if (Mod(anno, 4) = 0 && (Mod(anno, 100) != 0 || Mod(anno, 400) = 0))
                giornoFinale := 29  ; Anno bisestile
            else
                giornoFinale := 28  ; Anno non bisestile
        }
        
        ; Restituisci la data nel formato YYYYmmdd
        return Format("{:04d}{:02d}{:02d}", anno, mese, giornoFinale)
    }    
}