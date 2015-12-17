Dim browobj

SS = InputBox("Enter a symbol:", _
    "Analyze What?")


If SS = "" Then
    Wscript.Quit
Else
    
End If

FIRST =  Left(SS, 1)


Set browobj = CreateObject("WScript.Shell") 
browobj.Run "chrome -url -new-window http://www.gurufocus.com/stock/" & SS
WScript.Sleep 750
browobj.Run "chrome -url http://finance.yahoo.com/echarts?s=" & SS
WScript.Sleep 50
browobj.Run "chrome -url http://stockcharts.com/h-sc/ui?s=" & SS
WScript.Sleep 50
browobj.Run "chrome -url http://analysisreport.morningstar.com/stock/research?t=" & SS
WScript.Sleep 250
browobj.Run "chrome -url http://financials.morningstar.com/valuation/price-ratio.html?t=" & SS
WScript.Sleep 150
browobj.Run "chrome -url http://financials.morningstar.com/cash-flow/cf.html?t=" & SS
WScript.Sleep 150
browobj.Run "chrome -url http://financials.morningstar.com/ratios/r.html?t=" & SS
WScript.Sleep 150
browobj.Run "chrome -url http://financials.morningstar.com/balance-sheet/bs.html?t=" & SS
WScript.Sleep 150
browobj.Run "chrome -url http://financials.morningstar.com/income-statement/is.html?t=" & SS
WScript.Sleep 150
browobj.Run "chrome -url http://caps.fool.com/Ticker/" & SS & ".aspx"
WScript.Sleep 50
browobj.Run "chrome -url http://caps.fool.com/Ticker/" & SS & "/EarningsGrowthRates.aspx"
WScript.Sleep 150
browobj.Run "chrome -url https://www.etrade.wallst.com/v1/stocks/research/research.asp?symbol=" & SS
WScript.Sleep 150
browobj.Run "chrome -url http://biz.yahoo.com/research/earncal/" & FIRST   & "/" & SS  & ".html"
WScript.Sleep 50
browobj.Run "chrome -url http://www.thestreet.com/quote/" & SS & ".html"
WScript.Sleep 50
browobj.Run "chrome -url http://seekingalpha.com/symbol/" & SS & "/news"
WScript.Sleep 50
browobj.Run "chrome -url https://www.google.com/finance?q=" & SS
WScript.Sleep 50
browobj.Run "chrome -url https://secure.sigfig.com/l/stock-quote/" & SS


Set browobj= Nothing
