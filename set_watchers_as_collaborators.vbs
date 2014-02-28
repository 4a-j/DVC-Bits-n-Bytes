on error resume next

Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function
Function Base64Encode(inData)
'ripped from: 
'http://www.pstruh.cz/tips/detpg_Base64Encode.htm
  'rfc1521
  '2001 Antonin Foller, PSTRUH Software, http://pstruh.cz
  Const Base64 = _
"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim sOut, I
  
  'For each group of 3 bytes
  For I = 1 To Len(inData) Step 3
    Dim nGroup, pOut
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _
      &H100 * MyASC(Mid(inData, I + 1, 1)) + _
      MyASC(Mid(inData, I + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
  Next
  Select Case Len(inData) Mod 3
    Case 1: '8 bit final
      sOut = Left(sOut, Len(sOut) - 2) + "=="
    Case 2: '16 bit final
      sOut = Left(sOut, Len(sOut) - 1) + "="
  End Select
  Base64Encode = sOut
End Function

user = inputbox("Enter GitHub username")
pass = inputbox("Enter GitHub password")
repo = inputbox("Enter repository in the form USER/REPOSITORY")

set req = createobject("MSXML2.XMLHTTP.3.0")
req.open "GET", "https://api.github.com/repos/" & repo & "/subscribers", false, user, pass
req.setRequestHeader "Authorization", "Basic " & Base64Encode(user & ":" & pass)
req.setrequestheader "Accept", "application/vnd.github.v3+json"
req.send
response = req.responsetext

set list = createobject("Scripting.Dictionary")

n = 0
do
	n = instr(n + 1, response, """login"":""")
	if n = 0 then exit do
	n = n + 9
	index = index + 1
	list.add index, mid(response, n, instr(n, response, """") - n)
loop

for each n in list
	collab = list.item(n)
	result = result & chr(13) & chr(10) & collab & chr(13) & chr(10)
	set req = createobject("MSXML2.XMLHTTP.3.0")
	req.open "PUT", "https://api.github.com/repos/" & repo & "/collaborators/" & collab, false, user, pass
	req.setRequestHeader "Authorization", "Basic " & Base64Encode(user & ":" & pass)
	req.setrequestheader "Accept", "application/vnd.github.v3+json"
	req.send
	result = result & req.responsetext
next

msgbox result

