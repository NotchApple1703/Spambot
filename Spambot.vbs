set shell = createobject ("wscript.shell")

strtext = inputbox ("What to spam")
strtimes = inputbox ("How many times?")
strspeed = inputbox ("How fast? (1000 = 1/s, 100 = 10/sec, 10 = 100/s, 1 = 1000/s)")
strtimeneed = inputbox ("Countdown to spam")

If not isnumeric (strtimes & strspeed & strtimeneed) then
msgbox "Please type a number"
wscript.quit
End If
strtimeneed2 = strtimeneed * 1000
do
msgbox "You have " & strtimeneed & " seconds left"
wscript.sleep strtimeneed2
for i=0 to strtimes
shell.sendkeys (strtext & "{enter}")
wscript.sleep strspeed
Next
wscript.sleep strspeed * strtimes / 10
returnvalue=MsgBox ("Want to spam again?",36)
If returnvalue=6 Then
Msgbox "Ok"
End If
If returnvalue=7 Then
msgbox "Bye"
wscript.quit
End IF
loop