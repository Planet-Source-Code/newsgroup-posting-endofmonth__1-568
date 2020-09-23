<div align="center">

## EndOfMonth


</div>

### Description

Last day of the month

Calculates the last day of the month.

"Jim Doherty" <jdoherty@proweb.co.uk>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-endofmonth__1-568/archive/master.zip)





### Source Code

```
Function EndOfMonth (D As Variant) As Variant
 EndOfMonth = DateSerial(Year(D), Month(D) + 1, 0)
End Function
```

