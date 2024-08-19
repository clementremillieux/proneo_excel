on run argv
    if (count of argv) is less than 2 then
        return "Erreur : Nom de case à cocher et chemin du fichier Excel requis"
    end if
    
    set checkboxName to item 1 of argv
    set workbookName to item 2 of argv
    
    tell application "Microsoft Excel"
        try

            tell workbook workbookName
                if exists sheet "OPAC" then
                    tell sheet "OPAC"
                        try
                            return value of checkbox checkboxName
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