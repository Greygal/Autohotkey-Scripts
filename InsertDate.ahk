;####### INSERT TODAY'S DATE #######

F8::  ; inserts today's date
{
Now := FormatTime(A_Now, "MMMM d, yyyy")
Send(Now)
return
}
