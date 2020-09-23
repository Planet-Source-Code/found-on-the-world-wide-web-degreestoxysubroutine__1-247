<div align="center">

## DegreesToXYsubroutine


</div>

### Description

The DegreesToXYsubroutine, calculates the X (horizontal) and Y (vertical) coordinates of any point, measured in degrees, on the circumference of a circle or ellipse.
 
### More Info
 
Pass the subroutine the center X, Y of your ellipse, the degree position,

and the horizontal and vertical radii (if they are equal, you're specifying a

circle, if not, it is an elongated ellipse).

Returns the coordinates in the X and Y parameters.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-degreestoxysubroutine__1-247/archive/master.zip)





### Source Code

```
Public Sub DegreesToXY(CenterX As _
    Long, CenterY As Long, degree _
    As Double, radiusX As Long, _
    radiusY As Long, X As Long, Y _
    As Long)
Dim convert As Double
    convert = 3.141593 / 180
    'pi divided by 180
    X = CenterX - (Sin(-degree * _
        convert) * radiusX)
    Y = CenterY - (Sin((90 + _
        (degree)) * convert) * radiusY)
End Sub
```

