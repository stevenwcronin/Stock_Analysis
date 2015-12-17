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
Set browobj = CreateObject("WScript.Shell") 
browobj.Run "chrome -url -new-window http://www.google.com" 



For i = 0 To NUMBER_OF_STOCKS
 ' Wscript.Echo STOCKS(i)
WScript.Sleep 5
CURRENT_STRING = STOCKS(i)
WScript.Sleep 5
FIRST =  Left(CURRENT_STRING, 1)
WScript.Sleep 45
Set browobj = CreateObject("WScript.Shell") 
WScript.Sleep 45
' browobj.Run "chrome -url -new-window http://www.gurufocus.com/stock/" & STOCKS(i)
browobj.Run "chrome -url http://www.gurufocus.com/stock/" & STOCKS(i)
Set browobj= Nothing
Next



For i = 0 To NUMBER_OF_STOCKS
 ' Wscript.Echo STOCKS(i)
WScript.Sleep 5
CURRENT_STRING = STOCKS(i)
WScript.Sleep 5
FIRST =  Left(CURRENT_STRING, 1)
WScript.Sleep 45
Set browobj = CreateObject("WScript.Shell") 
WScript.Sleep 45
browobj.Run "chrome -url http://analysisreport.morningstar.com/stock/research?t=" & STOCKS(i)
Set browobj= Nothing
Next


For i = 0 To NUMBER_OF_STOCKS
 ' Wscript.Echo STOCKS(i)
WScript.Sleep 5
CURRENT_STRING = STOCKS(i)
WScript.Sleep 5
FIRST =  Left(CURRENT_STRING, 1)
WScript.Sleep 45
Set browobj = CreateObject("WScript.Shell") 
WScript.Sleep 45
browobj.Run "chrome -url https://www.etrade.wallst.com/v1/stocks/research/research.asp?symbol=" & STOCKS(i)
Set browobj= Nothing
Next



WScript.Quit







For i = 0 To NUMBER_OF_STOCKS
 ' Wscript.Echo STOCKS(i)
WScript.Sleep 5
CURRENT_STRING = STOCKS(i)
WScript.Sleep 5
FIRST =  Left(CURRENT_STRING, 1)
WScript.Sleep 45
Set browobj = CreateObject("WScript.Shell") 
WScript.Sleep 45
browobj.Run "chrome -url http://caps.fool.com/Ticker/" & STOCKS(i) & ".aspx"
Set browobj= Nothing
Next


For i = 0 To NUMBER_OF_STOCKS
 ' Wscript.Echo STOCKS(i)
WScript.Sleep 5
CURRENT_STRING = STOCKS(i)
WScript.Sleep 5
FIRST =  Left(CURRENT_STRING, 1)
WScript.Sleep 45
Set browobj = CreateObject("WScript.Shell") 
WScript.Sleep 45
browobj.Run "chrome -url http://finance.yahoo.com/echarts?s=" & STOCKS(i)
Set browobj= Nothing
Next






Else
    Msgbox "Your realism is appreciated.  Maybe you should get a better computer!!! ;)"
End If

WScript.Quit





















'  CODE BELOW IS HERE IN CASE I'D LIKE TO ADD ADDITIONAL FOR LOOPS ABOVE


WScript.Sleep 450
browobj.Run "chrome -url http://finance.yahoo.com/echarts?s=" & STOCKS(i)
WScript.Sleep 150
browobj.Run "chrome -url http://stockcharts.com/h-sc/ui?s=" & STOCKS(i)
WScript.Sleep 150
browobj.Run "chrome -url http://analysisreport.morningstar.com/stock/research?t=" & STOCKS(i)
WScript.Sleep 250
browobj.Run "chrome -url http://financials.morningstar.com/valuation/price-ratio.html?t=" & STOCKS(i)
WScript.Sleep 150
browobj.Run "chrome -url http://financials.morningstar.com/cash-flow/cf.html?t=" & STOCKS(i)
WScript.Sleep 150
browobj.Run "chrome -url http://financials.morningstar.com/ratios/r.html?t=" & STOCKS(i)
WScript.Sleep 150
browobj.Run "chrome -url http://financials.morningstar.com/balance-sheet/bs.html?t=" & STOCKS(i)
WScript.Sleep 150
browobj.Run "chrome -url http://financials.morningstar.com/income-statement/is.html?t=" & STOCKS(i)
WScript.Sleep 150
browobj.Run "chrome -url http://caps.fool.com/Ticker/" & STOCKS(i) & ".aspx"
WScript.Sleep 150
browobj.Run "chrome -url http://caps.fool.com/Ticker/" & STOCKS(i) & "/EarningsGrowthRates.aspx"
WScript.Sleep 150
browobj.Run "chrome -url https://www.etrade.wallst.com/v1/stocks/research/research.asp?symbol=" & STOCKS(i)
WScript.Sleep 150
browobj.Run "chrome -url http://biz.yahoo.com/research/earncal/" & FIRST   & "/" & STOCKS(i)  & ".html"
WScript.Sleep 150
browobj.Run "chrome -url http://www.thestreet.com/quote/" & STOCKS(i) & ".html"
WScript.Sleep 150
browobj.Run "chrome -url http://seekingalpha.com/symbol/" & STOCKS(i) & "/news"
WScript.Sleep 150
browobj.Run "chrome -url https://www.google.com/finance?q=" & STOCKS(i)
WScript.Sleep 450
browobj.Run "chrome -url https://secure.sigfig.com/l/stock-quote/" & STOCKS(i)
WScript.Sleep 450
