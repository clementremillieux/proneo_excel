on run argv
    if (count of argv) is less than 2 then
        return "Erreur : Nom de case à cocher et chemin du fichier Excel requis"
    end if
    
    set checkboxName to item 1 of argv
    set workbookPath to item 2 of argv
    
    tell application "Microsoft Excel"
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