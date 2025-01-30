; Programma principale
; Descrizione: Utilizza l'event manager per coordinare le attività tra i diversi componenti del programma, inizializza e configura tutti i componenti dell'applicazione.
; Esempio:

#Requires AutoHotkey v2.0
#SingleInstance Force

#Include GlobalConstants.ahk
#Include ButtonIconManager.ahk
#Include GestioneDirectory.ahk
#Include GestioneImpianti.ahk
#Include Utils\ArrayTools.ahk
#Include SAP\SAP_Connection.ahk
#Include SAP\SAP_Estrai_Manager.ahk
#Include MainGUI.ahk

; Aggiungi hotkey per ESC
Hotkey "Escape", StopExecution

; Flag globale per il controllo dell'esecuzione
global isRunning := false

StopExecution(*) {
    global isRunning
       if isRunning {
           result := MsgBox("Vuoi interrompere l'esecuzione?", "Conferma", 4)
           if (result = "Yes") {
               isRunning := false
                While ProcessExist("saplogon.exe") {
                    ProcessClose("saplogon.exe")
                }
               ;SAPConnection.TerminateSession()
               MainGUI.SetStatusBarText("Esecuzione interrotta dall'utente")
           }
       }
       else {
            result := MsgBox("Vuoi chiudere l'applicazione?", "Conferma", 4)
            if (result = "Yes") {
                ExitApp
            }
        }
    }

Main()

Main() {
    ; Avvia l'applicazione
    app := GUI_Manager()
}