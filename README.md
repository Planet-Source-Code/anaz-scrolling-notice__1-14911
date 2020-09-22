<div align="center">

## Scrolling Notice


</div>

### Description

The basic part is to have a message poping up from the system tray area with a hyperlink in it, then after a while it returns back.If you click the hyperlink the site will be opened.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-02-02 01:03:38
**By**             |[Anaz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/anaz.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD14403212001\.zip](https://github.com/Planet-Source-Code/anaz-scrolling-notice__1-14911/archive/master.zip)

### API Declarations

```
Option Explicit 'For helping to declare all variables
'To open the site.This is not needed for the animation
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
 ByVal lpFile As String, ByVal lpParameters As String, _
 ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
```





