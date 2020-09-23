<div align="center">

## Read Query String on Client Side Using VBScript


</div>

### Description

This function reads all query string dynamically.

Read for copying, pasting and testing
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marcio Coelho](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marcio-coelho.md)
**Level**          |Intermediate
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VbScript \(browser/client side\)

**Category**       |[Algorithims](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithims__4-29.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marcio-coelho-read-query-string-on-client-side-using-vbscript__4-6987/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<Form name="stringsample" id="stringsample"  action="clientquerystring.asp" method=GET id=form1 name=form1>
<P>&nbsp;</P>
<INPUT type="text" id=text1 name=text1><BR>
<INPUT type="text" id=text2 name=text2><BR>
<INPUT type="text" id=text3 name=text3><BR>
<INPUT type="submit" value="Submit" id=submit1 name=submit1>
</Form>
</BODY>
</HTML>
<SCRIPT LANGUAGE=vbscript>
<!--
private queryId()
private queryvalue()
Private maxBound
Sub Window_OnLoad()
		call ClientQueryString()
		REDIM Preserve queryid(maxBound)
		REDIM Preserve queryvalue(maxBound)
		for i = 0 TO maxBound
					msgbox queryid(i) & " value is " & queryvalue(i)
					Select Case i
							Case 0
									document.stringsample.text1.value = queryvalue(i)
							Case 1
									document.stringsample.text2.value = queryvalue(i)
							Case 2
									document.stringsample.text3.value = queryvalue(i)
					End Select
					If len(queryvalue(i)) = 0 Then
						msgbox "there is no value for " & queryid(i)
				end If
		next
end sub
Function ClientQueryString()
	Dim urlString, bPos, ePos, firstPartofPair, secondPartofPair, exitDo, i, h, take
	Dim countquery
	exitDo = False
	countquery = 0
	urlString = document.url 'retrieve complete url
	bPos = Instr(1, urlString, "?", 1) + 1 'question mark (?) will determine if there is any query
	If bPos > 1 Then
							' We have at least one valid value pair in the QueryString
							Do Until exitDo
												ePos = Instr(bPos, urlString, "=", 1) 'get the position that separate variable of query name and the value of the query string
												firstPartofPair = Mid(urlString, bPos, ePos - bPos) 'retrieve the variable name
												bPos = ePos + 1 'move for the next position after the =
												If Instr(bPos, urlString, "&", 1) > 0 Then
													ePos = Instr(bPos, urlString, "&", 1)
													secondPartofPair = Mid(urlString, bPos, ePos - bPos) ' retrieve the variable value
												Else
													' End of QueryString has been reached
													ePos = Len(urlString)
													take = ePos - BPos
													If take = 0 Then take = 1
													if take <> -1 Then 	secondPartofPair = Mid(urlString, bPos, take)
													exitDo = True
												End If
											  REDIM Preserve queryid(countquery)
												queryid(countquery) = firstPartofPair
												REDIM  Preserve queryvalue(countquery)
												queryvalue(countquery) 	= secondPartofPair
												countquery = countquery + 1 'increment the number of query string passed
												bPos = ePos + 1
							Loop
	End If
	maxBound = countquery - 1
End Function
//-->
</SCRIPT>
```

