#Requires AutoHotkey v2.0

class ArrayTools {
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

    ; Metodo: ArrayDifference
    ; Descrizione: Effettua la differenza tra il contenuto di due array
    ; Parametri:
    ;   - param1: array contenente un insieme di elementi
    ;   - param2: array contenente un insieme di elementi (presumibilmente sottoinsieme del primo)
    ; Restituisce:
    ; - un array contenente i valori presenti nel primo array meno gli elementi presenti nel secondo (che è un sottoinsieme del primo)
    ; - un array privo di elementi
    ; Esempio: ArrayDifference(fl_array, array)

    Static ArrayDifference(array1, array2) {
        result := []
        for _, item in array1 {
            if !ArrayTools.HasElement(array2, item) {
                result.Push(item)
            }
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