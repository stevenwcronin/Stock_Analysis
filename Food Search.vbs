Dim browobj

P = InputBox("What are you hungry for?", _
    "Search for what kind of Food?")


If P = "" Then
    Wscript.Quit
Else
    
End If
' NEW CODE HERE
' NEW CODE HERE
PP =Replace(P," ","+",1,-1)
PPAMZN =Replace(P," ","%20",1,-1)

PRODUCTS = SPLIT(PP, ",")
PRODUCTSAMZN = SPLIT(PPAMZN, ",")

NUMBER_OF_PRODUCTS =   (UBound(PRODUCTS))

' Wscript.Echo "You are interested in " & NUMBER_OF_PRODUCTS+1 & " PRODUCTS.  " & vbCrLf & vbCrLf & (NUMBER_OF_PRODUCTS+1)*16  & " web pages will be opened in total." & vbCrLf  & vbCrLf & "I hope your computer can handle it!! ;)"



'intAnswer = _
'    Msgbox("You are interested in " & NUMBER_OF_PRODUCTS+1 & " PRODUCTS.  "  & vbCrLf & vbCrLf & (NUMBER_OF_PRODUCTS+1)*17  & " web pages will be opened in total."  & vbCrLf &  vbCrLf  & vbCrLf & vbCrLf  &  "Are you sure your computer can handle it??" , _
'    vbYesNo , "Can your computer handle it??")

intAnswer = vbYes

If intAnswer = vbYes Then
    


For i = 0 To NUMBER_OF_PRODUCTS
' Wscript.Echo PRODUCTS(i)
' Wscript.Echo PRODUCTSAMZN(i)
WScript.Sleep 5
CURRENT_STRING = PRODUCTS(i)
WScript.Sleep 5
FIRST =  Left(CURRENT_STRING, 1)
WScript.Sleep 450
Set browobj = CreateObject("WScript.Shell") 
WScript.Sleep 450
browobj.Run "chrome -url -new-window https://www.amazon.com/s/&field-keywords=" & PRODUCTSAMZN(i)
WScript.Sleep 450
browobj.Run "chrome -url https://express.google.com/u/0/s?q=" & PRODUCTS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.google.com/search?&tbm=shop&q=" & PRODUCTS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.google.com/search?&tbm=shop&q=" & PRODUCTS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.ebay.com/sch/" & PRODUCTS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.samsclub.com/sams/search/searchResults.jsp?searchTerm=" & PRODUCTS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.overstock.com/search?keywords=" & PRODUCTS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.aliexpress.com/wholesale?&SearchText=" & PRODUCTS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.walmart.com/search/?query=" & PRODUCTSAMZN(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.etsy.com/search?q=" & PRODUCTSAMZN(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.wish.com/search/" & PRODUCTSAMZN(i)
WScript.Sleep 150
browobj.Run "chrome -url http://www.alibaba.com/trade/search?&SearchText=" & PRODUCTS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://www.lightinthebox.com/index.php?&" & PRODUCTS(i)

WScript.Sleep 150
browobj.Run "chrome -url https://www.producthunt.com/search?q=" & PRODUCTS(i)






browobj.Run "chrome -url https://orderacme.com/products?Search=" & PRODUCTS(i)
WScript.Sleep 150
browobj.Run "chrome -url https://curbsideexpress.gianteagle.com/store/27391278/#/search/" & PRODUCTS(i)








Set browobj= Nothing
Next



Else
    Msgbox "Your realism is appreciated.  Maybe you should get a better computer!!! ;)"
End If