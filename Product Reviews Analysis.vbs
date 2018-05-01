Dim browobj

P = InputBox("What are you considering buying?", _
    "Search for what Products?")


If P = "" Then
    Wscript.Quit
Else
    
End If
' NEW CODE HERE
' NEW CODE HERE


PPAMZN = Replace(P,"//","/",1,-1)
PRODUCTSAMZN = SPLIT(PPAMZN, "/")

END_OF_STRING_COUNT =   (UBound(PRODUCTSAMZN))

' Wscript.Echo "You are interested in " & END_OF_STRING_COUNT+1 & " PRODUCTS.  " & vbCrLf & vbCrLf & (END_OF_STRING_COUNT+1)*16  & " web pages will be opened in total." & vbCrLf  & vbCrLf & "I hope your computer can handle it!! ;)"

'intAnswer = _
'    Msgbox("You are interested in " & END_OF_STRING_COUNT+1 & " PRODUCTS.  "  & vbCrLf & vbCrLf & (END_OF_STRING_COUNT+1)*17  & " web pages will be opened in total."  & vbCrLf &  vbCrLf  & vbCrLf & vbCrLf  &  "Are you sure your computer can handle it??" , _
'    vbYesNo , "Can your computer handle it??")

intAnswer = vbYes

If intAnswer = vbYes Then

 Wscript.Echo END_OF_STRING_COUNT 
' Wscript.Echo PRODUCTS(END_OF_STRING_COUNT)
Wscript.Echo PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 450

Set browobj = CreateObject("WScript.Shell") 
' NEED TO JUST BE ABLE TO OPEN UP THE CUT AND PASTED PRODUCT URL FROM AMAZON BELOW... HOW?
' ADD AN IF STATEMENT IF TO MOVE TO PREVIOUS STRING IF THERE'S A / IN THE END OF THE COPY CUT PASTED TEXT....

WScript.Sleep 450
browobj.Run "chrome -url -new-window https://www.amazon.com/s?field-keywords=" & PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 450
browobj.Run "chrome -url https://www.amazon.com/dp/" & PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 450
browobj.Run "chrome -url https://www.fakespot.com" ' & "/product/" & PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 150
browobj.Run "chrome -url https://reviewmeta.com/amazon/" & PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 150
browobj.Run "chrome -url https://camelcamelcamel.com/" ' & PRODUCTSAMZN(END_OF_STRING_COUNT)
WScript.Sleep 150
' browobj.Run "chrome -url -new-window & P(1)

Set browobj= Nothing


Else
    Msgbox "Your realism is appreciated.  Maybe you should get a better computer!!! ;)"
End If