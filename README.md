<div align="center">

## Animate your program cursor


</div>

### Description

This little snippet will make the cursor animated. I.E a globe that is spinning.
 
### More Info
 
Be sure to know that the cursor you are gona use is a Animated cursor, if you don't got one you can probably find one at a windows desktop theme site.

Not as far as I know.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/animate-your-program-cursor__1-24150/archive/master.zip)

### API Declarations

```
'This part will go into a module
'Its only 2 functions and 1 constant long :)
Declare Function LoadCursorFromFile Lib "user32" _
Alias "LoadCursorFromFileA" _
(ByVal lpFileName As String) As Long
Declare Function SetClassLong Lib "user32" _
Alias "SetClassLongA" _
(ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
Public Const GCL_HCURSOR = (-12)
```


### Source Code

```
'put this in your Form_Load event or in whatever you would like to trigger the code.
'Note be sure that the cursor you want to use is a Animated cursor.
Dim sCursorFile As String
Dim hCursor As Long
Dim hOldCursor As Long
Dim lReturn As Long
'Pointing to the place where to cursor is.
sCursorFile = App.Path & "\animantedcursor.ani"
hCursor = LoadCursorFromFile(sCursorFile)
'Change the Form1.hwnd to yourformname.hwnd
hOldCursor = SetClassLong(Form1.hwnd, GCL_HCURSOR, hCursor)
'Use this to get back to normal cursor again.
'you can trigger it on form unload event or whatever you want to use to end it.
lReturn = SetClassLong(Form1.hWnd, GCL_HCURSOR, hOldCursor)
```

