#Requires AutoHotkey v2.0

/*
This icon selection tool uses the ImageList in the ListView command to create an icon selection box.
I added the FileSelectFile command to make exploring various files easier.
Double-click on any image to save the file path and name plus the icon number to the Windows clipboard.
Both C:\Windows\System32\shell32.dll and C:\Windows\System32\imageres.dll contain hundreds of icons.
*/

class IconViewer {
    static gui := ""
    static listView := ""
    static currentFile := ""
    static imageListID := ""
    
    static __New() {
        ; Seleziona il file iniziale
        if !this.SelectFile()
            return
            
        ; Crea e configura la GUI
        this.CreateGui()
        
        ; Carica le icone
        this.LoadIcons()
        
        ; Mostra la GUI
        this.gui.Show()
    }
    
    static CreateGui() {
        ; Crea la finestra principale
        this.gui := Gui()
        this.gui.OnEvent("Close", (*) => ExitApp())
        
        ; Imposta il font
        this.gui.SetFont("s20")
        
        ; Aggiunge il ListView
        this.listView := this.gui.Add("ListView", "h415 w150", ["Icon"])
        this.listView.OnEvent("DoubleClick", this.IconSelected.Bind(this))
        
        ; Crea il menu
        fileMenu := Menu()
        fileMenu.Add("&Open File`tCtrl+O", this.SelectFile.Bind(this))
        
        MyMenuBar := MenuBar()
        MyMenuBar.Add("&File", fileMenu)
        
        this.gui.MenuBar := MyMenuBar
    }
    
    static SelectFile() {
        ; Apre il dialogo di selezione file
        ; selectedFile := FileSelect(32,, "Pick a file to check icons", "*.*")
        selectedFile := "C:\Windows\System32\shell32.dll"
        if !selectedFile
            return false
            
        this.currentFile := selectedFile
        
        ; Se la GUI esiste già, ricarica le icone
        if this.gui {
            this.listView.Delete()
            this.LoadIcons()
        }
        
        return true
    }
    
    static LoadIcons() {
        ; Crea una nuova ImageList
        this.imageListID := IL_Create(10, 1, false)
        this.listView.SetImageList(this.imageListID)
        
        ; Carica le icone
        count := 0
        loop {
            count++
            image := IL_Add(this.imageListID, this.currentFile, A_Index)
            
            if (image = 0)  ; Quando non ci sono più icone
                break
        }
        
        ; Aggiunge le righe al ListView
        loop count {
            this.listView.Add("Icon" . A_Index, A_Index, "n/a")
        }
        
        ; Auto-aggiusta la larghezza della colonna
        this.listView.ModifyCol(1, "AutoHdr")
    }
    
    static IconSelected(*) {
        ; Ottiene l'indice dell'icona selezionata
        if (row := this.listView.GetNext()) {
            ; Copia nel clipboard il percorso del file e il numero dell'icona
            A_Clipboard := this.currentFile ", " row
            
            MsgBox(A_Clipboard "`r     added to Clipboard! `r" A_ScriptDir)
        }
    }
}

; Avvia l'applicazione
IconViewer.__New()