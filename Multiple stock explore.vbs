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
    


For i = 0 To NUMBER_OF_STOCKS
 ' Wscript.Echo STOCKS(i)
WScript.Sleep 5
CURRENT_STRING = STOCKS(i)
WScript.Sleep 5
FIRST =  Left(CURRENT_STRING, 1)
WScript.Sleep 450

Set browobj = CreateObject("WScript.Shell") 
WScript.Sleep 450
browobj.Run "chrome -url -new-window http://www.gurufocus.com/stock/" & STOCKS(i)
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
browobj.Run "chrome -url http://www.marketwatch.com/investing/stock/" & STOCKS(i)
WScript.Sleep 450
browobj.Run "chrome -url http://finviz.com/quote.ashx?t=" & STOCKS(i)
WScript.Sleep 450
browobj.Run "chrome -url https://www.zacks.com/stock/quote/" & STOCKS(i) & "?q=" & STOCKS(i)
WScript.Sleep 450
browobj.Run "chrome -url http://www.bloomberg.com/quote/" & STOCKS(i) & ":US" 
WScript.Sleep 450
browobj.Run "chrome -url http://www.barchart.com/plmodules/?module=cashFlow&symbol=" & STOCKS(i) & "&1_type=2" 
WScript.Sleep 450
browobj.Run "chrome -url http://investorplace.com/stock-quotes/" & STOCKS(i) & "-stock-quote/" 
WScript.Sleep 450
browobj.Run "chrome -url http://www.dataroma.com/m/stock.php?sym=" & STOCKS(i)
WScript.Sleep 450
browobj.Run "chrome -url https://stockrow.com/" & STOCKS(i) & "/snapshots"
WScript.Sleep 450
browobj.Run "chrome -url http://www.insidermonkey.com/search/all?q=" & STOCKS(i)
WScript.Sleep 450
browobj.Run "chrome -url http://www.stockta.com/cgi-bin/analysis.pl?symb=" & STOCKS(i) & "&cobrand=&mode=stock" 
WScript.Sleep 450
browobj.Run "chrome -url https://www.stockrover.com/markets/cash_flow/" & STOCKS(i)
WScript.Sleep 450
browobj.Run "chrome -url https://www.quandl.com/search?query=" & STOCKS(i)
WScript.Sleep 450
browobj.Run "chrome -url http://stocktwits.com/symbol/" & STOCKS(i) & "?q="
WScript.Sleep 450
browobj.Run "chrome -url https://stockflare.com/stocks/" & STOCKS(i) & "?key=ticker"




Set browobj= Nothing



Next



Else
    Msgbox "Your realism is appreciated.  Maybe you should get a better computer!!! ;)"
End If