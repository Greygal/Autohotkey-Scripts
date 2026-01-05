; ###### CLEAN UP CONTRACTIONS #####
; Expands any contractions (e.g., changes I'm to I am, I'll to I will).  Works only on an open Microsoft Word document.  
; It's best to run this script BEFORE doing any serious formatting because  some formatting/paragraph breaks may get stripped out due to quirks in COM Object and Microsoft Word formatting. 

F9:: {
    try {
        oWord := ComObjActive("Word.Application")
        FullText := oWord.ActiveDocument.Content.Text
    } catch {
        MsgBox("Please make sure a Word document is open.")
        return
    }

    ; Defines the most common English contractions
    ; Note: We use lowercase keys for the search to keep the logic simple
    Contractions := Map(
        "i'll", "I will",
        "it's", "it is",
        "doesn't", "does not",
        "don't", "do not",
        "can't", "cannot",
        "won't", "will not",
        "shouldn't", "should not",
        "couldn't", "could not",
        "wouldn't", "would not",
        "isn't", "is not",
        "aren't", "are not",
        "wasn't", "was not",
        "weren't", "were not",
        "hasn't", "has not",
        "haven't", "have not",
        "i'm", "I am",
        "you're", "you are",
        "he's", "he is",
        "she's", "she is",
        "we're", "we are",
        "they're", "they are",
        "i've", "I have",
        "you've", "you have",
        "we've", "we have",
        "they've", "they have",
        "that's", "that is",
        "there's", "there is",
        "what's", "what is",
        "who's", "who is"
    )

    ; Loop through the map and replace each contraction
    for Contraction, Expansion in Contractions {
        ; Use RegexReplace to ensure we only catch whole words
        ; i) = Case-insensitive
        ; \b = Word boundary
        FullText := RegExReplace(FullText, "i)\b" . Contraction . "\b", Expansion)
    }

    ; Put the text back into Word
    oWord.ActiveDocument.Content.Text := FullText
    MsgBox("Contractions expanded!")
}
