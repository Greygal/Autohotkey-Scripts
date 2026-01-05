; Set to ALT-F5. This converts spelled out dates to numeric dates (e.g., change April 6, 2019 to 4/6/2019) in an open Microsoft Word document.

!F5:: {
    try {
        oWord := ComObjActive("Word.Application")
        FullText := oWord.ActiveDocument.Content.Text
    } catch {
        MsgBox("Please ensure a Word document is open.")
        return
    }

    ; Reverse Month Map for conversion
    MonthMap := Map("January", 1, "February", 2, "March", 3, "April", 4, "May", 5, "June", 6, 
                    "July", 7, "August", 8, "September", 9, "October", 10, "November", 11, "December", 12)

    ; Pattern matches Month Name, Space, Day, Comma, Space, Year
    ; i) = Case-insensitive
    SpelledPattern := "i)\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{2,4})\b"
    Pos := 1
    
    while (FoundPos := RegExMatch(FullText, SpelledPattern, &Match, Pos)) {
        MonthName := Match.1
        ; Get numeric month from map, ensuring we handle the case-insensitivity
        MonthNum := MonthMap[Format("{:T}", MonthName)]
        DayNum := Integer(Match.2) ; Removes leading zeros automatically
        YearNum := Match.3
        
        NumericDate := MonthNum . "/" . DayNum . "/" . YearNum
        FullText := StrReplace(FullText, Match.0, NumericDate, , , 1)
        Pos := FoundPos + StrLen(NumericDate)
    }

    oWord.ActiveDocument.Content.Text := FullText
    MsgBox("Spelled-out dates converted to numeric format!")
}
