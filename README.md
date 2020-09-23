<div align="center">

## Fast Code 2 Create Excel Files


</div>

### Description

I have many troubles when try to export big amount of records from database into

excel. These troubles because I use Excel.Application, Excel.Workbook and

Excel.Worksheet. Unfortunately all of them working too slow, spend many resources,

and not compatable between OSs/Versions.

There is other way to create all excel compatable file.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tair Abdurman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tair-abdurman.md)
**Level**          |Intermediate
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tair-abdurman-fast-code-2-create-excel-files__1-6491/archive/master.zip)





### Source Code

```
'by Tair Abdurman
'visit http://www.tair.freeservers.com
'   for other examples
'e-mail: excelz@tair.freeservers.com
Function CreateExcelFile() As Long
 On Error GoTo CatchErr
   Const LF_SYMBOL As Byte = &HA
   Const TAB_SYMBOL As Byte = &H9
   Dim szFilePath As String
   Dim szFileName As String
   Dim szDefaultBuffer As String
   Dim lFieldCount As Long
   Dim lRowCount As Long
   Dim ltempCount As Long
   Dim ltempCount2 As Long
   szFilePath = App.Path
   If Right(szFilePath, 1) <> "\" Then szFilePath = szFilePath & "\"
   szFileName = "TestExcel"
   lFieldCount = 10
   lRowCount = 10
   Open szFilePath & szFileName & ".xls" For Append As #1
     szDefaultBuffer = ""
    'save field names
     ltempCount = 1
     Do While ltempCount <= lFieldCount
       szDefaultBuffer = szDefaultBuffer & Chr(TAB_SYMBOL) & "Field" & ltempCount
       ltempCount = ltempCount + 1
     Loop
     'can be skipped because Print put that symbol
     'szDefaultBuffer=szDefaultBuffer & chr(LF_SYMBOL)
     Print #1, szDefaultBuffer
    'save field values
     ltempCount = 1
     Do While ltempCount <= lRowCount
       szDefaultBuffer = ""
       ltempCount2 = 1
       Do While ltempCount2 <= lFieldCount
        szDefaultBuffer = szDefaultBuffer & Chr(TAB_SYMBOL) & "Value" & ltempCount & ":" & ltempCount2
        ltempCount2 = ltempCount2 + 1
       Loop
       'can be skipped because Print put that symbol
       'szDefaultBuffer=szDefaultBuffer & chr(LF_SYMBOL)
       Print #1, szDefaultBuffer
       ltempCount = ltempCount + 1
     Loop
   Close 1
   CreateExcelFile = 0
   Exit Function
CatchErr:
   CreateExcelFile = Err.Number
End Function
```

