#Requires AutoHotkey v2.0

class DirectoryManager {
    configFile := G_CONSTANTS.DIRECTORY_CONFIG_FILE
    categories := G_CONSTANTS.ESTRAZIONI_SAP
    gui := ""
    editControls := Map()
    
    __New(mainApp) {
        ; Salva il riferimento all'app principale
        this.mainApp := mainApp        ; Crea la finestra principale
        this.gui := Gui("+ToolWindow +Owner" . this.mainApp.mainGui.gui.Hwnd, "Configurazione Directory")
        this.gui.MarginX := 10
        this.gui.MarginY := 10
        
        ; Carica le configurazioni esistenti o usa valori di default
        this.LoadConfig()
        
        ; Crea i controlli per ogni categoria
        this.CreateControls()
        
        ; Aggiunge i pulsanti di salvataggio e annulla
        this.CreateButtons()
        
        ; Mostra la GUI
        this.gui.Show()
    }
    
    LoadConfig() {
        ; Se il file di configurazione non esiste, lo crea con valori di default
        if !FileExist(this.configFile) {
            for category in this.categories {
                IniWrite(A_ScriptDir "\" category, this.configFile, "Directories", category)
            }
        }
    }
    
    CreateControls() {
        ; Aggiunge una breve descrizione
        this.gui.Add("Text","xm ym", "Seleziona le directory per il salvataggio dei dati:")
        ;this.gui.Add("Text", "y+5", "")  ; Spaziatura
        
        ; Per ogni categoria, crea una riga di controlli
        for category in this.categories {
            ; Leggi il valore salvato o usa il default
            savedPath := IniRead(this.configFile, "Directories", category, A_ScriptDir "\" category)
            
            ; Crea i controlli per questa categoria
            this.gui.Add("Text", "xm y+20", category ":")
            
            ; Edit box per il percorso
            edit := this.gui.Add("Edit", "xm yp+15 w405 v" category "Path", savedPath)
            this.editControls[category] := edit
            
            ; Pulsante Browse
            btn := this.gui.Add("Button", "x+5 yp-1 w" MainGUI.BUTTON_DIM . " h" . MainGUI.BUTTON_DIM, "")
            ButtonIconManager.SetIcon(btn, 'shell32.dll',267, 's' . 32)
            btn.OnEvent("Click", this.BrowseFolder.Bind(this, category))
        }
    }
    
    CreateButtons() {
        ; Aggiunge una linea di separazione
        this.gui.Add("Text", "xm y+20 w450 h2 +0x10")
        
        ; Pulsanti di controllo
        saveBtn := this.gui.Add("Button", "xm y+10 w100", "Salva")
        saveBtn.OnEvent("Click", this.SaveConfig.Bind(this))
        
        cancelBtn := this.gui.Add("Button", "x+10 yp w100", "Annulla")
        cancelBtn.OnEvent("Click", this.Annulla.Bind(this))
        
        ; Pulsante per aprire il file di configurazione
        openConfigBtn := this.gui.Add("Button", "x+10 yp w140", "Apri File Config")
        openConfigBtn.OnEvent("Click", (*) => Run(this.configFile))
        
        ; Pulsante per resettare ai valori di default
        resetBtn := this.gui.Add("Button", "x+10 yp w100", "Reset Default")
        resetBtn.OnEvent("Click", this.ResetToDefault.Bind(this))

        ; evento per la chiusura della finestra
        this.gui.OnEvent("Close", (*) => this.HandleCloseButtonClick())
        this.gui.OnEvent("Escape", (*) => this.HandleCloseButtonClick())
    }
    
    HandleCloseButtonClick(*) {
        this.mainApp.EnableMainGui()
        this.gui.Destroy()
        this.mainApp.DirectoryManagerGui := ""  ; Rimuovi il riferimento
    }


    Annulla(*) {
        this.mainApp.EnableMainGui()
        this.gui.Destroy()        
    }

    BrowseFolder(category, *) {
        ; Ottieni il percorso corrente
        currentPath := this.editControls[category].Value
        
        ; Apri il selettore di directory
        if (selectedDir := DirSelect("*" . currentPath, 3, "Seleziona directory per " category)) {
            ; Aggiorna il controllo Edit con il nuovo percorso
            this.editControls[category].Value := selectedDir
        }
    }
    
    SaveConfig(*) {
        try {
            ; Salva ogni percorso nel file INI
            for category, edit in this.editControls {
                path := edit.Value
                
                ; Verifica che la directory esista o creala
                if !DirExist(path) {
                    result := MsgBox("La directory per " category " non esiste.`n`nVuoi crearla?",
                                   "Directory non trovata",
                                   "YesNo Icon?")
                    if (result = "Yes") {
                        DirCreate(path)
                    } else {
                        continue
                    }
                }
                
                ; Salva il percorso
                IniWrite(path, this.configFile, "Directories", category)
            }
            
            MsgBox("Configurazione salvata con successo!", "Successo", "Iconi")
            this.mainApp.EnableMainGui()
            this.gui.Destroy()
        } catch as err {
            MsgBox("Errore durante il salvataggio: " err.Message, "Errore", "Icon!")
        }
    }
    
    ResetToDefault(*) {
        if (MsgBox("Vuoi davvero ripristinare i valori predefiniti?",
                   "Conferma Reset",
                   "YesNo Icon?") = "Yes") {
            ; Ripristina ogni percorso al valore predefinito
            for category in this.categories {
                defaultPath := A_ScriptDir "\" category
                this.editControls[category].Value := defaultPath
            }
        }
    }
    
    ; Metodo pubblico per ottenere il percorso di una categoria
    GetPath(category) {
        return IniRead(this.configFile, "Directories", category, A_ScriptDir "\" category)
    }

    ; Metodo pubblico per ottenere il percorso di una categoria
    static GetCategoryPath(category) {
        return IniRead(G_CONSTANTS.DIRECTORY_CONFIG_FILE, "Directories", category, A_ScriptDir "\" category)
    }
    
    ; Metodo pubblico per verificare se tutte le directory esistono
    VerifyDirectories() {
        missingDirs := []
        for category in this.categories {
            path := this.GetPath(category)
            if !DirExist(path) {
                missingDirs.Push(category)
            }
        }
        return missingDirs
    }
}


/* DirectoryManager.__New() */

; Per ottenere un percorso specifico da altre parti del codice:
; admPath := DirectoryManager.GetPath("AdM")

; Per verificare che tutte le directory esistano:
; missingDirs := DirectoryManager.VerifyDirectories()
; if missingDirs.Length {
;     MsgBox("Directory mancanti: " ArrayTools.Join(missingDirs))
; }
