; CONTROL-F5 Convert numeric dates to spelled out dates (e.g., 01/10/2022 to January 10, 2022).  Works only in an open Microsoft Word document.
^F5:: {
    try {
        oWord := ComObjActive("Word.Application")
        FullText := oWord.ActiveDocument.Content.Text
    } catch {
        MsgBox("Please ensure a Word document is open.")
        return
    }

    ; Month Map for expansion
    Months := ["January", "February", "March", "April", "May", "June", 
               "July", "August", "September", "October", "November", "December"]

    ; Pattern matches 1-2 digit month, slash, 1-2 digit day, slash, 2-4 digit year
    DatePattern := "\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b"
    Pos := 1
    
    while (FoundPos := RegExMatch(FullText, DatePattern, &Match, Pos)) {
        MonthNum := Integer(Match.1)
        DayNum := Integer(Match.2)
        YearNum := Match.3
        
        ; Safety: Only process if month is 1-12
        if (MonthNum >= 1 && MonthNum <= 12) {
            SpelledDate := Months[MonthNum] . " " . DayNum . ", " . YearNum
            FullText := StrReplace(FullText, Match.0, SpelledDate, , , 1)
            Pos := FoundPos + StrLen(SpelledDate)
        } else {
            Pos := FoundPos + StrLen(Match.0)
        }
    }

    oWord.ActiveDocument.Content.Text := FullText
    MsgBox("Numeric dates converted to spelled-out format!")
}
