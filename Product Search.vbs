Dim browobj

SS = InputBox("What are you considering buying?", _
    "Analyze Which Stocks?")


If SS = "" Then
    Wscript.Quit
Else
    
End If
' NEW CODE HERE
' NEW CODE HERE
SS =Replace(SS," ","%20%",1,-1)


STOCKS = SPLIT(SS, ",")

NUMBER_OF_STOCKS =   (UBound(STOCKS))

' Wscript.Echo "You are interested in " & NUMBER_OF_STOCKS+1 & " stocks.  " & vbCrLf & vbCrLf & (NUMBER_OF_STOCKS+1)*16  & " web pages will be opened in total." & vbCrLf  & vbCrLf & "I hope your computer can handle it!! ;)"



'intAnswer = _
'    Msgbox("You are interested in " & NUMBER_OF_STOCKS+1 & " stocks.  "  & vbCrLf & vbCrLf & (NUMBER_OF_STOCKS+1)*17  & " web pages will be opened in total."  & vbCrLf &  vbCrLf  & vbCrLf & vbCrLf  &  "Are you sure your computer can handle it??" , _
'    vbYesNo , "Can your computer handle it??")

intAnswer = vbYes

If intAnswer = vbYes Then
    


For i = 0 To NUMBER_OF_STOCKS
 ' Wscript.Echo STOCKS(i)
WScript.Sleep 5
CURRENT_STRING = STOCKS(i)
WScript.Sleep 5
FIRST =  Left(CURRENT_STRING, 1)
WScript.Sleep 450
Set browobj = CreateObject("WScript.Shell") 
WScript.Sleep 450
browobj.Run "chrome -url -new-window https://www.amazon.com/s/&field-keywords=" & STOCKS(i)
WScript.Sleep 450
browobj.Run "chrome -url https://express.google.com/u/0/s?q=" & STOCKS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.google.com/search?&tbm=shop&q=" & STOCKS(i)
Set browobj= Nothing



Next



Else
    Msgbox "Your realism is appreciated.  Maybe you should get a better computer!!! ;)"
End If