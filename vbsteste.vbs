dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP").

dim bStrm: Set bStrm = createobject("Adodb.Stream").

xHttp.Open "GET", "https://ftp.mozilla.org/pub/firefox/releases/62.0b9/win64/br/Firefox%20Setup%2062.0b9.exe", False.

xHttp.Send



with bStrm

    .type = 1 '//binary
    
    .open
    
    .write xHttp.responseBody
    
    .savetofile "installfirefox.exe", 2 '//overwrite
    
end with

