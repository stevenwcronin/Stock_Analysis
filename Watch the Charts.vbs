Dim browobj

SS = InputBox("Enter the stock symbols you are interested in separated by commas:", _
    "Analyze Which Stocks?")


If SS = "" Then
    Wscript.Quit
Else
    
End If


STOCKS = SPLIT(SS, ",")

NUMBER_OF_STOCKS =   (UBound(STOCKS))

' Wscript.Echo "You are interested in " & NUMBER_OF_STOCKS+1 & " stocks.  " & vbCrLf & vbCrLf & (NUMBER_OF_STOCKS+1)*16  & " web pages will be opened in total." & vbCrLf  & vbCrLf & "I hope your computer can handle it!! ;)"



'intAnswer = _
'    Msgbox("You are interested in " & NUMBER_OF_STOCKS+1 & " stocks.  "  & vbCrLf & vbCrLf & (NUMBER_OF_STOCKS+1)*17  & " web pages will be opened in total."  & vbCrLf &  vbCrLf  & vbCrLf & vbCrLf  &  "Are you sure your computer can handle it??" , _
'    vbYesNo , "Can your computer handle it??")

intAnswer = vbYes

If intAnswer = vbYes Then
    
WScript.Sleep 450
Set browobj = CreateObject("WScript.Shell") 

browobj.Run "chrome -url https://finance.yahoo.com/quotes/" & SS


For i = 0 To NUMBER_OF_STOCKS
 ' Wscript.Echo STOCKS(i)
WScript.Sleep 5
CURRENT_STRING = STOCKS(i)
WScript.Sleep 5
FIRST =  Left(CURRENT_STRING, 1)


WScript.Sleep 450
browobj.Run "chrome -url http://finance.yahoo.com/echarts?s=" & STOCKS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://stockrow.com/" & STOCKS(i) & "/snapshots"
WScript.Sleep 450
browobj.Run "chrome -url http://www.stockta.com/cgi-bin/analysis.pl?symb=" & STOCKS(i) & "&cobrand=&mode=stock" 
WScript.Sleep 450
browobj.Run "chrome -url http://finviz.com/quote.ashx?t=" & STOCKS(i)
WScript.Sleep 450
browobj.Run "chrome -url https://stockflare.com/stocks/" & STOCKS(i) & ".o"




Next



Else
    Msgbox "Your realism is appreciated.  Maybe you should get a better computer!!! ;)"
End If



Set browobj= Nothing