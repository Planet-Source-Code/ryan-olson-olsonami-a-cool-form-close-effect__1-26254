<div align="center">

## A \!\! Cool Form Close Effect \!\!


</div>

### Description

This is a cool closing effect for your app. Tell me watcha think!!
 
### More Info
 
you need to add a Form_Unload() into your code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ryan Olson \- OLSONAMI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ryan-olson-olsonami.md)
**Level**          |Beginner
**User Rating**    |4.1 (41 globes from 10 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ryan-olson-olsonami-a-cool-form-close-effect__1-26254/archive/master.zip)





### Source Code

```
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Me.WindowState = 0
Do
Me.Top = Me.Top + 10
Me.Left = Me.Left + 10
Me.Width = Me.Width - 20
Me.Height = Me.Height - 20
Loop Until Me.Top => Screen.Height
'you can change those numbers to make it faster
'or slower. right now it is pretty slow.
'if the height and width #'s are twice as much
'as the top and left #'s, it will make a
' "zooming out" effect and then will fall to the
'bottom of the screen.
End Sub
```

