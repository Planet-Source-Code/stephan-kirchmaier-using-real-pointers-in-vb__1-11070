<div align="center">

## Using real pointers in VB\!


</div>

### Description

This code uses undocumented functions of VB that gives you a pointer to a string or a number!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stephan Kirchmaier](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephan-kirchmaier.md)
**Level**          |Advanced
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephan-kirchmaier-using-real-pointers-in-vb__1-11070/archive/master.zip)

### API Declarations

```
Private Declare Sub CopyMemByPtr Lib "kernel32" Alias _
  "RtlMoveMemory" (ByVal lpTo As Long, ByVal lpFrom As Long, _
  ByVal lLen As Long)
```


### Source Code

```
Private Declare Sub CopyMemByPtr Lib "kernel32" Alias _
  "RtlMoveMemory" (ByVal lpTo As Long, ByVal lpFrom As Long, _
  ByVal lLen As Long)
Private Sub Form_Click()
  Dim a As Long, b As String, c As Long, d As String
  Dim i As Integer, j As Long, k As Integer, l As Long
  Dim u(2) As Byte, o As Long
  b = "HELLO!"
  d = Space(Len(b))
  i = 20
  u(0) = 23
  u(1) = 243
  u(2) = 124
  o = VarPtr(u(0))
  j = VarPtr(i)
  l = VarPtr(k)
  a = StrPtr(b)
  c = StrPtr(d)
  CopyMemByPtr o + 1, j, Len(u(0)) * 2
  CopyMemByPtr l, j, Len(i) * 2
  CopyMemByPtr c, a, Len(b) * 2
  MsgBox d & vbCr & k & vbCr & u(1)
End Sub
```

