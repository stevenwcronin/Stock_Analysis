Dim browobj

P = InputBox("What are you considering buying?", _
    "Search for what Products?")


If P = "" Then
    Wscript.Quit
Else
    
End If
' NEW CODE HERE
' NEW CODE HERE
PP     = Replace(P,"//","/",1,-1)
PPAMZN = Replace(P,"//","/",1,-1)

PRODUCTS     = SPLIT(PP, "/")
PRODUCTSAMZN = SPLIT(PPAMZN, "/")

END_OF_STRING_COUNT =   (UBound(PRODUCTS))

' Wscript.Echo "You are interested in " & END_OF_STRING_COUNT+1 & " PRODUCTS.  " & vbCrLf & vbCrLf & (END_OF_STRING_COUNT+1)*16  & " web pages will be opened in total." & vbCrLf  & vbCrLf & "I hope your computer can handle it!! ;)"



'intAnswer = _
'    Msgbox("You are interested in " & END_OF_STRING_COUNT+1 & " PRODUCTS.  "  & vbCrLf & vbCrLf & (END_OF_STRING_COUNT+1)*17  & " web pages will be opened in total."  & vbCrLf &  vbCrLf  & vbCrLf & vbCrLf  &  "Are you sure your computer can handle it??" , _
'    vbYesNo , "Can your computer handle it??")

intAnswer = vbYes

If intAnswer = vbYes Then

Wscript.Echo END_OF_STRING_COUNT 
Wscript.Echo PRODUCTS(END_OF_STRING_COUNT)

WScript.Sleep 450
Set browobj = CreateObject("WScript.Shell") 
WScript.Sleep 450
browobj.Run "chrome -url -new-window https://www.amazon.com/dps/&field-keywords=" & PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 450
browobj.Run "chrome -url -new-window https://www.amazon.com/dp/" & PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 450
browobj.Run "chrome -url https://www.fakespot.com" ' & "/product/" & PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 150
browobj.Run "chrome -url https://reviewmeta.com/amazon/" & PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 150
browobj.Run "chrome -url https://camelcamelcamel.com/" ' & PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 150


Set browobj= Nothing




Else
    Msgbox "Your realism is appreciated.  Maybe you should get a better computer!!! ;)"
End If