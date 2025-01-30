#Requires AutoHotkey v2.0

class TooltipManager {
    static tooltips := Map()
    
    /**
     * Aggiunge un tooltip a un controllo
     * @param {Gui.Button} control - Controllo GUI a cui aggiungere il tooltip
     * @param {String} text - Testo del tooltip
     * @param {String} title - (Opzionale) Titolo del tooltip
     * @param {Integer} icon - (Opzionale) Icona: 0=None, 1=Info, 2=Warning, 3=Error
     */
    static AddToolTip(control, text, title := "", icon := 0) {
        ; Crea il tooltip
        hTT := DllCall("CreateWindowEx", "UInt", 0x8, "Str", "tooltips_class32", "Ptr", 0
            , "UInt", 0x02 | 0x40, "Int", 0x80000000, "Int", 0x80000000
            , "Int", 0x80000000, "Int", 0x80000000, "Ptr", control.Gui.Hwnd
            , "Ptr", 0, "Ptr", 0, "Ptr", 0)
            
        ; Configura il tooltip
        TOOLINFO := Buffer(A_PtrSize = 8 ? 72 : 48, 0)
        NumPut("UInt", A_PtrSize = 8 ? 72 : 48, TOOLINFO, 0)  ; cbSize
        NumPut("UInt", 0x11, TOOLINFO, 4)                     ; uFlags
        NumPut("Ptr", control.Gui.Hwnd, TOOLINFO, 8)          ; hwnd
        NumPut("Ptr", control.Hwnd, TOOLINFO, 8 + A_PtrSize)  ; uId
        NumPut("Ptr", StrPtr(text), TOOLINFO, 36 + A_PtrSize) ; lpszText
        
        ; Aggiunge il tool
        DllCall("SendMessage", "Ptr", hTT, "UInt", 0x404, "Ptr", 0, "Ptr", TOOLINFO.Ptr) ; TTM_ADDTOOL
        
        ; Imposta il titolo se specificato
        if title {
            DllCall("SendMessage", "Ptr", hTT, "UInt", 0x420, "Ptr", icon, "Str", title) ; TTM_SETTITLE
        }
        
        ; Memorizza il tooltip per riferimento futuro
        this.tooltips[control.Hwnd] := { hwnd: hTT, text: text, title: title }
    }
    
    /**
     * Aggiorna il testo di un tooltip esistente
     * @param {Gui.Button} control - Controllo con il tooltip
     * @param {String} newText - Nuovo testo del tooltip
     */
    static UpdateText(control, newText) {
        if this.tooltips.Has(control.Hwnd) {
            tooltip := this.tooltips[control.Hwnd]
            TOOLINFO := Buffer(A_PtrSize = 8 ? 72 : 48, 0)
            NumPut("UInt", A_PtrSize = 8 ? 72 : 48, TOOLINFO, 0)
            NumPut("UInt", 0x11, TOOLINFO, 4)
            NumPut("Ptr", control.Gui.Hwnd, TOOLINFO, 8)
            NumPut("Ptr", control.Hwnd, TOOLINFO, 8 + A_PtrSize)
            NumPut("Ptr", StrPtr(newText), TOOLINFO, 36 + A_PtrSize)
            
            DllCall("SendMessage", "Ptr", tooltip.hwnd, "UInt", 0x40C, "Ptr", 0, "Ptr", TOOLINFO.Ptr) ; TTM_UPDATETIPTEXT
            this.tooltips[control.Hwnd].text := newText
        }
    }
    
    /**
     * Rimuove il tooltip da un controllo
     * @param {Gui.Button} control - Controllo da cui rimuovere il tooltip
     */
    static RemoveToolTip(control) {
        if this.tooltips.Has(control.Hwnd) {
            tooltip := this.tooltips[control.Hwnd]
            DllCall("DestroyWindow", "Ptr", tooltip.hwnd)
            this.tooltips.Delete(control.Hwnd)
        }
    }
}

; Esempio di implementazione:
class ExampleGUI {
    static __New() {
        ; Crea la GUI
        this.gui := Gui()
        this.gui.MarginX := 10
        this.gui.MarginY := 10
        
        ; Aggiungi alcuni pulsanti di esempio
        
        ; Esempio 1: Tooltip semplice
        btn1 := this.gui.Add("Button", "w120 h30", "Hover Me")
        TooltipManager.AddToolTip(btn1, "Questo è un tooltip semplice")
        
        ; Esempio 2: Tooltip con titolo e icona
        btn2 := this.gui.Add("Button", "x+10 yp w120 h30", "Info Tooltip")
        TooltipManager.AddToolTip(btn2, "Tooltip con informazioni aggiuntive", "Informazione", 1)
        
        ; Esempio 3: Tooltip con warning
        btn3 := this.gui.Add("Button", "xm y+10 w120 h30", "Warning")
        TooltipManager.AddToolTip(btn3, "Attenzione! Questo è un avviso.", "Attenzione", 2)
        
        ; Esempio 4: Tooltip con errore
        btn4 := this.gui.Add("Button", "x+10 yp w120 h30", "Error")
        TooltipManager.AddToolTip(btn4, "Questo è un messaggio di errore", "Errore", 3)
        
        ; Esempio 5: Tooltip dinamico
        this.dynBtn := this.gui.Add("Button", "xm y+10 w120 h30", "Dynamic")
        TooltipManager.AddToolTip(this.dynBtn, "Tooltip iniziale")
        this.dynBtn.OnEvent("Click", this.UpdateTooltip.Bind(this))
        
        this.gui.Show()
    }
    
    static UpdateTooltip(*) {
        static count := 0
        count++
        TooltipManager.UpdateText(this.dynBtn, "Cliccato " count " volte")
    }
}

; Avvia l'esempio
ExampleGUI.__New()