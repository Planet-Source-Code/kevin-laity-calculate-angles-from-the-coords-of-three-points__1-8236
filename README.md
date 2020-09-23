<div align="center">

## Calculate Angles from the coords of three points


</div>

### Description

This code is useful in games where you need to calculate an angle quickly. Using three points it gives you the angle, in degrees, of the vertex.
 
### More Info
 
Three sets of coordinates: A center point, and the ends of two lines coming from the center.

The angle of the vertex in degrees


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kevin Laity](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kevin-laity.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kevin-laity-calculate-angles-from-the-coords-of-three-points__1-8236/archive/master.zip)





### Source Code

```
Function LineLen(x1, y1, x2, y2)
	'This function will simply give you the length
	'of a line using the coordinates of its two
	'endpoints.
	Dim A, B As Single
	A = Abs(x2 - x1)
	B = Abs(y2 - y1)
	LineLen = Sqr(A ^ 2 + B ^ 2)
End Function
Function Arccos(X As Single)
	If X = 1 Then Arccos = 0: Exit Function
	Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function
Public Function CalcAnAngle(CenterX, CenterY, x2, y2, x3, y3)
	'This function will take three coordinates and
	'automagically turn them into an angle.
	'The angle is the one located at CenterX, CenterY
	'For example:
	'  / X2,Y2
	'  /
	' /
	' < CenterX,CenterY
	' \
	'  \
	'  \ X3,Y3
	'CalcAnAngle will return the angle, in degrees,
	'of the center vertex.
	On Error Resume Next
	Dim SideA, SideB, SideC As Single
	SideC = lineLen(CenterX, CenterY, x2, y2)
	SideB = lineLen(CenterX, CenterY, x3, y3)
	SideA = lineLen(x3, y3, x2, y2)
	a = Arccos((SideA ^ 2 - SideB ^ 2 - SideC ^ 2) / (SideB * SideC * -2))
	CalcAnAngle = a * (180 / 3.141)
	'VB seems to like to work in confusing units
	'called Radians instead of good ol' degrees.
	'Multiplying by (180 / 3.141) converts radians
	'to degrees.
End Function
```

