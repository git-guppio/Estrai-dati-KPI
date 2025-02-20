#Requires AutoHotkey v2.0

; Classe: MakeMainGUI
; Descrizione: Classe per la creazione dell'interfaccia grafica principale
; Esempio:

class MainApp {
    __New() {
        this.mainGui := MakeMainGUI(this)
        this.invTecGui := MakeInvTecGUI(this)
        this.parameters := {inverterTechnology: 0}
        ; Questo metodo viene chiamato automaticamente quando la classe viene caricata
        ;this.interval := 200  ; Intervallo di aggiornamento in millisecondi
        ;this.isActive := false
        ;this.baseText := ""
        ;this.indicatorType := 1
        ;this.maxIndicatorTypes := 9  ; Numero totale di tipi di indicatori
        ;this.timer := ObjBindMethod(this, "Tick")
        this.SetupEventListeners()
        this.mainGui.CreateGUI()
        ;this.invTecGui.CreateGUI()              
        }
    
        SetupEventListeners() {
            EventManager.Subscribe("ShowInv", (*) => this.ShowInv())
            EventManager.Subscribe("AddLV", (data) => this.mainGui.AddLineLV(data.icon, data.element, data.text))
            EventManager.Subscribe("PI_Start", (data) => this.mainGui.PI_Start(data.inputValue))
            EventManager.Subscribe("PI_Stop", (data) => this.mainGui.PI_Stop(data.inputValue))
            EventManager.Subscribe("ProcessStatusUpdated",(data) => this.ProcessStatusUpdated(data.processId, data.status, data.details, data.result))            
        }
       
        ProcessStatusUpdated(processId, status, details, result:={}) {
            if (processID = "CheckFL") and ((status = "Started") OR (status = "In Progress")) {
                this.mainGui.ClearBtn.Enabled := false
                this.mainGui.CheckBtn.Enabled := false
                this.mainGui.ActionBtn.Enabled := false
            }
            if (processID = "CheckFL") and ((status = "Error") OR (status = "Completed")) {
                this.mainGui.ClearBtn.Enabled := true
                this.mainGui.CheckBtn.Enabled := true
                this.mainGui.ActionBtn.Enabled := true
            }

        }

        ShowMain() {
            this.mainGui.GuiShow()
        }

        ShowInv() {
            this.invTecGui.InvShow()
        }        
    }

class MakeMainGUI {
    __New(mainApp) {
    this.mainApp := mainApp
    this.baseText := ""
    this.isActive := false
    this.interval := 200
    this.maxIndicatorTypes := 9  ; Numero totale di tipi di indicatori
    this.indicatorType := 1
    this.timer := ObjBindMethod(this, "Tick")
    this.CreateGUI()
    }

/*     ProcessStatusUpdated(processId, status, details, result:={}) {
        if (processID = "CheckFL") and ((status = "Started") OR (status = "In Progress"))
            this.ClearBtn.Enabled := false
        else if (processID = "CheckOdMDettail") and ((status = "Started") OR (status = "In Progress"))
            this.ClearBtn.Enabled := false
        else
            this.ClearBtn.Enabled := true
    } */

    AddLineLV(icon, element, text) {
        this.gui.LV.Add(icon , element, text)
        this.gui.LV.ModifyCol(1, "autoHdr")    
    }

    PI_Start(inputValue){
        this.Start(inputValue)
    }

    PI_Stop(inputValue){
        this.Stop(inputValue)
    }

    Start(baseText := "") {
        this.baseText := baseText
        this.isActive := true
        this.tickCount := 0
        SetTimer(this.timer, this.interval)
    }

    Stop(baseText := "Pronto per il controllo FL") {
        this.isActive := false
        SetTimer(this.timer, 0)
        ;~ this.gui.SB.SetText(this.baseText)  ; Ripristina il testo base
        this.SetSBText("  " . baseText)  ; Ripristina il testo base
    }

    Tick() {
        if (this.isActive) {
            indicator := this.AsciiProgressIndicator(this.indicatorType)
            ;OutputDebug("this.baseText = " . this.baseText . "`n")
            this.gui.SB.SetText("  " . this.baseText . " " . indicator)
        }
    }

    AsciiProgressIndicator(type := 1) {
        static index := 0
        indicators := [
            ["-", "=", "-", "="],
            ["_", "_", "=", "=", "≡", "≡", "=", "="],
            ["|", "/", "─", "\"],
            ["▁", "▂", "▃", "▄", "▅", "▆", "▇", "█", "▇", "▆", "▅", "▄", "▃", "▂"],
            ["◐", "◓", "◑", "◒"],
            ["◰", "◳", "◲", "◱"],
            ["←", "↖", "↑", "↗", "→", "↘", "↓", "↙"],
            ["┌", "┐", "┘", "└"],
            [".   ", "..  ", "... ", "...."]
        ]
        index := Mod(index + 1, indicators[type].Length)
        return indicators[type][index + 1]
    }

    ChangeIndicatorType() {
        this.indicatorType := Mod(this.indicatorType, this.maxIndicatorTypes) + 1
        if (this.isActive) {
            this.Tick()  ; Aggiorna immediatamente l'indicatore
        }
    }

    ClickToChangeIndicator() {
        TempText := this.baseText
        this.ChangeIndicatorType()
        if (!this.isActive) {
            ; Se l'indicatore non è attivo, mostriamo brevemente il nuovo tipo
            this.Start("Tipo indicatore cambiato ")
            SetTimer(() => this.Stop(TempText), -2000)  ; Ferma dopo 2 secondi
        }
    }


    CreateGUI() {
        this.gui := Gui("+Resize -MinimizeBox -MaximizeBox")
        mX := 10
        mY := 10
        this.gui.MarginX := mX
        this.gui.MarginY := mY

        ; Calcola le dimensioni iniziali
        screenWidth := A_ScreenWidth
        screenHeight := A_ScreenHeight

        ; Imposta dimensioni come percentuale dello schermo
        guiWidth := Floor(screenWidth * 0.5)
        guiHeight := Floor(screenHeight * 0.5)

        ; imposto la grandezza del primo Groupbox pari a 2/3 dello dimensione della GUI
        ; imposto la grandezza del secondo Groupbox pari a 1/3 dello dimensione della GUI
        
        ; Inserisco un gruppo per la lista degli OdM
        this.GroupListaOdM := this.gui.Add("GroupBox",
        "xm ym " .
        " w" . (guiWidth * 2/3) .
        " h" . (guiHeight - 30))

        this.GroupListaOdM.GetPos(&x, &y, &w, &h)
        ; Elemento interno al Groupbox
        this.gui.LV := this.gui.Add("ListView",
            "xp+" . mX .
            " yp+" . (2*mY) . 
            " w" . (w - 2*mX) .
            " h" . (0.8*h) .
        " Grid", ["Ptw", "OdM", "Testo breve", "Stato S.", "Stato U."])
        
        ; Bottoni per le attività su OdM
        this.ActionBtn := this.gui.Add("Button", "xp y+10 section w100", " -- Undefined --")
        this.ActionBtn.Enabled := false
        this.ClearBtn := this.gui.Add("Button", "xm+10 ys w90", "Clear all")
        this.ClearBtn.Enabled := false

        ; Crea il Gruppo per i dettagli dell'OdM selezionato
        this.gui.AddGroupBox("xp+420 ym w300 h200 section", "Dettagli OdM")
        ; Crea box in cui inserire il testo
        this.gui.OdMDettail := this.gui.Add("Edit", "-VScroll xs+10 ys+20 w280 h170 vOdMDettail")        

        ; Crea il Gruppo per i dettagli del KPI
        this.gui.AddGroupBox("xp-10 yp+190 w300 h200 section", "Dettagli KPI")
        ; Crea box in cui inserire il testo
        this.gui.OdMDettail := this.gui.Add("Edit", "-VScroll xs+10 ys+20 w280 h170 vKPIDettail")


        this.ClearBtn := this.gui.Add("Button", "xm+10 ys w90", "Clear all")
        this.ClearBtn.Enabled := false
        ;~ MakeMainGUI.ExpandBtnToggle := MakeMainGUI.gui.Add("Button", "x+10 yp w90", "Expand")
        this.CheckBtn := this.gui.Add("Button", "x+10 yp w90", "Check")
        this.CheckBtn.Enabled := false

        ; Tasto per esportare
        ;~ MakeMainGUI.ExportBtn := MakeMainGUI.gui.Add("Button", "w100 xm", "Export Selected")

        ; Crea una imagelist
        this.imageListID := IL_Create(3)
        this.gui.LV.SetImageList(this.imageListID)

        icons := [295, 236, 132, 4] ; [OK, Allert, NOK, File]
        for icon in icons {
            IL_Add(this.imageListID, "shell32.dll", icon)
        }

        ; Aggiunge StatusBar
        this.gui.SB := this.gui.Add("StatusBar",, "Bar's starting text (omit to start off empty).")
        this.SetSBText("  Pronto per il controllo FL")

        ; Impostazione eventi
        this.CheckBtn.OnEvent("Click", (*) => this.OnCheckButtonClick())
        this.ClearBtn.OnEvent("Click", (*) => this.ClearAll())
        this.ActionBtn.OnEvent("Click", (*) =>  EventManager.Publish("UpLoadFiles", {}))

        ; Aggiungi l'evento di doppio clic alla StatusBar
        this.gui.SB.OnEvent("DoubleClick", (*) => this.ClickToChangeIndicator())
        ; Aggiungi l'evento Change al controllo Edit
        this.gui.OdMDettail.OnEvent("Change", (*) => this.CheckFLListContent())
        ; Aggiungi il menu contestuale
        this.gui.LV.OnEvent("ContextMenu", (*) => this.ShowContextMenu())        

        this.gui.OnEvent("Close", (*) => this.Exit())
        this.gui.Title := "Check FL"
        this.gui.Opt("+AlwaysOnTop")
    }

    Exit() {
        ; Completamento con successo
        EventManager.Publish("ProcessCompleted", {processId: "MainGUI", details: "Esecuzione completata con successo", result: {}})
        ExitApp()
    }

    ; Definisco il menu contestuale da mostrare sulla LV
    ShowContextMenu(*) {
        contextMenu := Menu()
        contextMenu.Add("Copia", (*) => this.CopySelectedText())
        contextMenu.Show()
    }

    CopySelectedText(*) {
        copiedText := ""
        rowNumber := 0

        ; Itera attraverso le righe selezionate
        while (rowNumber := this.gui.LV.GetNext(rowNumber)) {
            loop this.gui.LV.GetCount("Col") {
                cellText := this.gui.LV.GetText(rowNumber, A_Index)
                copiedText .= cellText . "`t"  ; Usa tab come separatore tra colonne
            }
            copiedText := RTrim(copiedText, "`t") . "`n"  ; Rimuovi l'ultimo tab e aggiungi un newline
        }

        if (copiedText != "") {
            A_Clipboard := RTrim(copiedText, "`n")  ; Rimuovi l'ultimo newline
            MsgBox("Testo copiato negli appunti!")
        }
    }

    ; Funzione per la lettura del contenuto del controllo edit <OdMDettail>
    ; Restituisce un array composto dalle linee di testo presenti    
    OnCheckButtonClick(*) {
        ;this.mainApp.ShowInv

        ; EventManager.Publish("ProcessStarted", {processId: "OnCheckButtonClick", details: "Check button pressed"})     
        if (this.ClearListCB.Value = true)
            this.gui.LV.Delete()
                
        ; Otteniamo il contenuto del controllo OdMDettail dalla GUI principale
        flListContent := this.gui.OdMDettail.Value

        ; Verifichiamo che il contenuto non sia vuoto
        if (flListContent = "") {
            ; EventManager.Publish("ProcessError", {processId: "OnCheckButtonClick", details: "la Lista FL è vuota. LN: " . A_LineNumber})
            MsgBox("Errore: Il contenuto della lista FL è vuoto.", "Errore", 4112)
            return false
        }
        else {
            EventManager.Publish("CheckDataRequest", {flArray: flListContent})
        }
        
        ;EventManager.Publish("ProcessCompleted", {processId: "OnCheckButtonClick", details: "Funzione OnCheckButtonClick completata", result: {}})
    }

    SetSBText(Text) {
        this.baseText := Text
        this.gui.SB.SetText(Text)
    }
    
    ; verifico se è  presente del testo nel controllo OdMDettail
    CheckFLListContent(*) {
        if (this.gui.OdMDettail.Value != "") {
            ; Il controllo contiene del testo
            this.CheckBtn.Enabled := true
            this.ClearBtn.Enabled := true
            this.SetSBText("  Pronto per il controllo FL")
        } else {
            ; Il controllo è vuoto
            this.CheckBtn.Enabled := false
            this.ClearBtn.Enabled := false
            this.ActionBtn.Enabled := false
            this.SetSBText("  Inserisci le FL da controllare")
        }
    }

    ; elimino il contenuto della OdMDettail e della LV
    ClearAll(*) {
            this.gui.OdMDettail.value := ""
            this.gui.LV.Delete()
            this.gui.LV.ModifyCol(1, "autoHdr")
            this.ActionBtn.Enabled := false
            this.CheckBtn.Enabled := false
            this.ClearBtn.Enabled := false
    }

    GuiShow() {
        this.gui.Show()
    }
}

; Classe: MakeInvTecGUI
; Descrizione: Classe per creazione del menu di selezione della tipologia di inverter negli impianti Solar
; Esempio:

class MakeInvTecGUI {
    __New(mainApp) {
        this.mainApp := mainApp
        this.result := { success: false, value: false, error: "", class: "MakeGUI.ahk", function: "MakeInvTecGUI" }
        ;this.CreateGUI()
    }

    CreateGUI() {
        this.gui := Gui("-MinimizeBox -MaximizeBox +Owner" . this.mainApp.mainGui.gui.Hwnd)
        this.gui.MarginX := 10
        this.gui.MarginY := 10

        this.gui.Add("Text", "xm ym", "Select inverter type")

        ; Aggiungi i tre radio button
        this.radioGroup := this.gui.Add("Radio", "xm+10 yp+20 h23 vRadioGroup", " Central Inverter")
        this.gui.Add("Radio", "xm+10 yp+20 h23", " String Inverter")
        this.gui.Add("Radio", "xm+10 yp+20 h23", " Inverter Module")

        this.OKBtn := this.gui.Add("Button", "xm+10 yp+30 w80", "OK")
        this.CancelBtn := this.gui.Add("Button", "xp+90 yp w80", "Cancel")

        ; Impostazione eventi
        this.OKBtn.OnEvent("Click", (*) => this.HandleButtonOKClick())
        this.CancelBtn.OnEvent("Click", (*) => this.HandleButtonCancelClick())
        this.gui.OnEvent("Close", (*) => this.HandleCloseButtonClick())

        this.gui.Title := "Solar Inverter Type"
    }

    InvShow() {
        this.mainApp.mainGui.gui.Opt("+Disabled")  ; Disabilita la finestra principale
        EventManager.Publish("ProcessStarted", {processId: "SelectInverter", status: "Started", details: "Avvio funzione", result: {}})
        this.CreateGUI()
        EventManager.Publish("ProcessProgress", {processId: "SelectInverter", status: "In Progress", details: "Esecuzione in corso", result: {}})      
        this.gui.Show()
    }

    HandleButtonCancelClick(*) {
        this.result.value := 0
        EventManager.Publish("ProcessError", {processId: "SelectInverter", details: "Nessuna tipologia di inverter selezionato.", result: this.result}) 
        this.mainApp.mainGui.gui.Opt("-Disabled")  ; Abilita la finestra principale
        this.Close()
    }

    HandleCloseButtonClick(*) {
        this.result.value := 0
        EventManager.Publish("ProcessError", {processId: "SelectInverter", details: "Nessuna tipologia di inverter selezionato.", result: this.result})      
        this.mainApp.mainGui.gui.Opt("-Disabled")  ; Disabilita la finestra principale
        this.Close()
    }

    HandleButtonOKClick(*) {
        itemState := this.gui.Submit(false)  ; false significa "non nascondere la GUI"
        this.result.value := itemState.RadioGroup
        EventManager.Publish("ProcessCompleted", {processId: "SelectInverter", status: "Completed", details: "", result: this.result})
        this.Close()
    }

    Close() {
        this.gui.Destroy()
        this.mainApp.mainGui.gui.Opt("-Disabled")  ; Abilita la finestra principale
    }
}

