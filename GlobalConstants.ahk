#Requires AutoHotkey v2.0
 
    global G_CONSTANTS := {
        ;file_Country: A_ScriptDir . "\Config\country.csv",
        DIRECTORY_CONFIG_FILE: A_ScriptDir "\directories.ini",
        IMPIANTI_CONFIG_FILE: A_ScriptDir "\impianti.ini",
        ESTRAZIONI_SAP: ["AdM", "OdM", "SicuAmbi", "Ptw"],
        }

        /* 
        Esempio di utilizzo:
        G_CONSTANTS.test
        */