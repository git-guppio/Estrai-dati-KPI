#Requires AutoHotkey v2.0

; ======================================================================================================================
; AHK Settings
; ======================================================================================================================
#SingleInstance Force
SetWorkingDir(A_ScriptDir)
SetTitleMatchMode(2)

class ImpiantiManager {
    static configFile := A_ScriptDir "\impianti.ini"
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
                IniWrite(ArrayUtils.Join(impianti), this.configFile, categoria, "impianti")
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

    ; Legge tutti gli impianti di una categoria
    static LeggiImpianti(categoria) {
        try {
            impiantiStr := IniRead(this.configFile, categoria, "impianti")
            return StrSplit(impiantiStr, ";", " ")  ; Converte la stringa in array
        } catch Error as err {
            MsgBox("Errore nella lettura degli impianti: " err.Message)
            return []
        }
    }

    ; Aggiunge un nuovo impianto a una categoria
    static AggiungiImpianto(categoria, nuovoImpianto) {
        try {
            impianti := this.LeggiImpianti(categoria)
            
            ; Verifica se l'impianto esiste già
            if (ArrayUtils.HasElement(impianti, nuovoImpianto))
                return "Impianto già esistente"
                
            impianti.Push(nuovoImpianto)
            IniWrite(ArrayUtils.Join(impianti), this.configFile, categoria, "impianti")
            return "Impianto aggiunto con successo"
        } catch Error as err {
            return "Errore nell'aggiunta dell'impianto: " err.Message
        }
    }    
    
    ; Rimuove un impianto da una categoria
    static RimuoviImpianto(categoria, impiantoDaRimuovere) {
        try {
            impianti := this.LeggiImpianti(categoria)
            
            ; Cerca e rimuove l'impianto
            for index, impianto in impianti {
                if (impianto = impiantoDaRimuovere) {
                    impianti.RemoveAt(index)
                    IniWrite(ArrayUtils.Join(impianti), this.configFile, categoria, "impianti")
                    return "Impianto rimosso con successo"
                }
            }
            return "Impianto non trovato"
        } catch Error as err {
            return "Errore nella rimozione dell'impianto: " err.Message
        }
    }
    
    ; Legge tutte le categorie disponibili
    static LeggiCategorie() {
        try {
            return StrSplit(IniRead(this.configFile), '`n') 
        } catch Error as err {
            return ["COAL", "GAS", "ImpFuture"]
        }
    }
}

class ArrayUtils {
    /*
     * Verifica se un elemento esiste in un array
     * @param array Array in cui cercare
     * @param element Elemento da cercare
     * @param caseSensitive (opzionale) Se true, la ricerca è case-sensitive. Default: true
     * @returns {Boolean} True se l'elemento è presente, False altrimenti
     */
    static HasElement(array, element, caseSensitive := true) {
        if !IsObject(array)
            return false
            
        if caseSensitive {
            for item in array {
                if (item = element)
                    return true
            }
        } else {
            for item in array {
                if (StrLower(item) = StrLower(element))
                    return true
            }
        }
        return false
    }
    
    /*
     * Restituisce l'indice di un elemento nell'array
     * @param array Array in cui cercare
     * @param element Elemento da cercare
     * @param caseSensitive (opzionale) Se true, la ricerca è case-sensitive. Default: true
     * @returns {Integer} Indice dell'elemento se trovato, 0 se non trovato
     */
    static IndexOf(array, element, caseSensitive := true) {
        if !IsObject(array)
            return 0
            
        if caseSensitive {
            for index, item in array {
                if (item = element)
                    return index
            }
        } else {
            for index, item in array {
                if (StrLower(item) = StrLower(element))
                    return index
            }
        }
        return 0
    }
    
    /*
     * Conta quante volte un elemento appare nell'array
     * @param array Array in cui cercare
     * @param element Elemento da contare
     * @param caseSensitive (opzionale) Se true, la ricerca è case-sensitive. Default: true
     * @returns {Integer} Numero di occorrenze dell'elemento
     */
    static Count(array, element, caseSensitive := true) {
        if !IsObject(array)
            return 0
            
        count := 0
        if caseSensitive {
            for item in array {
                if (item = element)
                    count++
            }
        } else {
            for item in array {
                if (StrLower(item) = StrLower(element))
                    count++
            }
        }
        return count
    }
    
    /*
     * Converte un array in una stringa con elementi separati da punto e virgola
     * @param array Array da convertire in stringa
     * @param delimiter (opzionale) Separatore da usare. Default: ";"
     * @returns {String} Stringa con elementi separati dal delimiter
     */    

    static Join(array, delimiter := ";") {
        if !IsObject(array)
            return ""
            
        result := ""
        for index, element in array {
            if (index > 1)
                result .= delimiter
            result .= element
        }
        return result
    }
    
    /**
     * Versione alternativa usando un approccio diverso
     * Utile per confronto prestazioni con array grandi
     */
    static JoinAlt(array, delimiter := ";") {
        if !IsObject(array)
            return ""
            
        tempArray := []
        for element in array
            tempArray.Push(String(element))
            
        return tempArray.Length ? tempArray.Join(delimiter) : ""
    }
}

; Crea una GUI per gestire gli impianti
class ImpiantiGUI {
    static __New() {
        ; Inizializza il file con i valori di default
        ImpiantiManager.GetImpianti()
        
        ; Crea la finestra principale
        ImpiantiGUI.MainGui := Gui("+ToolWindow", "Gestione Impianti")
        
        ; Dropdown per selezionare la categoria
        ImpiantiGUI.MainGui.Add("Text",, "Categoria:")
        categorie := ImpiantiManager.LeggiCategorie()
        ImpiantiGUI.categoriaDropdown := ImpiantiGUI.MainGui.Add("DropDownList", "vCategoria w200", categorie)
        ImpiantiGUI.categoriaDropdown.Choose(1)
        
        ; ListBox per mostrare gli impianti
        ImpiantiGUI.MainGui.Add("Text", "y+10", "Impianti:")
        ImpiantiGUI.implantList := ImpiantiGUI.MainGui.Add("ListBox", "vImpianti w200 h200")
        
        ; Controlli per aggiungere/rimuovere impianti
        ImpiantiGUI.MainGui.Add("Text", "y+10", "Nuovo Impianto:")
        ImpiantiGUI.MainGui.Add("Edit", "vNuovoImpianto w200")
        ImpiantiGUI.MainGui.Add("Button", "y+5 w95", "Aggiungi").OnEvent("Click", this.AggiungiClick)
        ImpiantiGUI.MainGui.Add("Button", "x+10 w95", "Rimuovi").OnEvent("Click", this.RimuoviClick)
        
        ; Aggiorna la lista quando si cambia categoria
        ImpiantiGUI.categoriaDropdown.OnEvent("Change", this.AggiornaLista)
/*         ; Restiuisce il valore selezionato nella ListBox
        ImpiantiGUI.implantList.OnEvent("Change", this.AggiornaLista) */
        ; Mostra la GUI e aggiorna la lista iniziale
        ImpiantiGUI.MainGui.Show()
        ImpiantiGUI.AggiornaLista()
        
        MyMainGui := ImpiantiGUI.MainGui
    }
    
    static AggiornaLista(*) {
        categoria := ImpiantiGUI.categoriaDropdown.Text
        impianti := ImpiantiManager.LeggiImpianti(categoria)
        ImpiantiGUI.MainGui["Impianti"].Delete()
        ImpiantiGUI.MainGui["Impianti"].Add(impianti)
    }
    
    static AggiungiClick(ctrl, *) {
        categoria := ImpiantiGUI.MainGui["Categoria"].Text
        nuovoImpianto := ImpiantiGUI.MainGui["NuovoImpianto"].Text
        
        if (nuovoImpianto = "") {
            MsgBox("Inserire un nome per il nuovo impianto")
            return
        }
        
        risultato := ImpiantiManager.AggiungiImpianto(categoria, nuovoImpianto)
        MsgBox(risultato)
        ImpiantiGUI.AggiornaLista(ctrl)
        ImpiantiGUI.MainGui["NuovoImpianto"].Text := ""
    }
    
    static RimuoviClick(ctrl, *) {
        categoria := ImpiantiGUI.MainGui["Categoria"].Text
        impiantoSelezionato := ImpiantiGUI.MainGui["Impianti"].Text
        
        if (impiantoSelezionato = "") {
            MsgBox("Selezionare un impianto da rimuovere")
            return
        }
        
        risultato := ImpiantiManager.RimuoviImpianto(categoria, impiantoSelezionato)
        MsgBox(risultato)
        ImpiantiGUI.AggiornaLista(ctrl)
    }
}

; Avvia l'applicazione
app := ImpiantiGUI()