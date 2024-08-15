on run argv
    if (count of argv) is 0 then
        return "Erreur : Aucun nom de case à cocher fourni"
    end if
    
    set checkboxName to item 1 of argv
    
    tell application "Microsoft Excel"
        set workbookPath to "/Users/remillieux/Documents/Proneo/logiciel/test/appel_script/Plan et Rapport d'audit certification V32.xlsm"
        
        try
            open workbookPath
            
            tell workbook 1
                if exists sheet "OPAC" then
                    tell sheet "OPAC"
                        try
                            set checkboxState to value of checkbox checkboxName
                            return checkboxState
                        on error
                            return "checkbox off"
                        end try
                    end tell
                else
                    return "Erreur : La feuille OPAC n'existe pas dans ce classeur"
                end if
            end tell
            
        on error errMsg
            return "Erreur : " & errMsg
        end try
    end tell
end run