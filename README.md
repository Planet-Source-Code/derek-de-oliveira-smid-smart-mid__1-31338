<div align="center">

## SMid \- "Smart" Mid


</div>

### Description

This function is similar to the Mid function. Except, you can specify the starting, and ending strings (capture the data in between).
 
### More Info
 
orig_string As String : Source string

start As Long : Location in source to start from

str_start As String : beginning string (capture from)

str_end As String : ending string (capture to)

This code will return the data that is found between the "str_start" and "str_end" strings.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Derek de Oliveira](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/derek-de-oliveira.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/derek-de-oliveira-smid-smart-mid__1-31338/archive/master.zip)





### Source Code

```
'Example:
'newStr = SMid("Hello 1Between2 world!", 1, "1", "2")
'will return: "Between"
Function SMid(orig_string As String, start As Long, str_start As String, str_end As String)
On Error GoTo handler
'SMid (Smart MID)
'By: Derek de Oliveira
'Use this function in any program. No need to thank me :)
'o_string = Origional String
's_start = Start From string
's_end = Ending string
step1 = InStr(start, orig_string, str_start, vbTextCompare)
result = Mid(orig_string, step1 + Len(str_start), InStr(step1 + Len(str_start), orig_string, str_end, vbTextCompare) - step1 - Len(str_start))
SMid = result
Exit Function
handler:
SMid = ""
End Function
```

