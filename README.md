<div align="center">

## Convert Common Dialog Control Color to WEB Hex


</div>

### Description

This code takes the common dialog color, extracts the R, G, B values, and

converts each value to the correct HEX equivilant supported by HTML.
 
### More Info
 
Function requires a common dialog control color being selected and passed to the

function

You must add the common dialog control to your application. After the ShowOpen event

occurs, pass the color selected by the user to the function:

Dim sColHex as String

sColHex = HexRGB(cdlCont.color)

Function returns properly formatted HEX color value for HTML


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Charlie Wilson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/charlie-wilson.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/charlie-wilson-convert-common-dialog-control-color-to-web-hex__1-1015/archive/master.zip)

### API Declarations

None


### Source Code

```
Private Function HexRGB(lCdlColor As Long)
  Dim lCol As Long
  Dim iRed, iGreen, iBlue As Integer
  Dim vHexR, vHexG, vHexB As Variant
  'Break out the R, G, B values from the common dialog color
  lCol = lCdlColor
  iRed = lCol Mod &H100
    lCol = lCol \ &H100
  iGreen = lCol Mod &H100
    lCol = lCol \ &H100
  iBlue = lCol Mod &H100
  'Determine Red Hex
  vHexR = Hex(iRed)
      If Len(vHexR) < 2 Then
         vHexR = "0" & vHexR
      End If
  'Determine Green Hex
  vHexG = Hex(iGreen)
      If Len(vHexG) < 2 Then
         vHexG = "0" & iGreen
      End If
  'Determine Blue Hex
  vHexB = Hex(iBlue)
      If Len(vHexB) < 2 Then
         vHexB = "0" & vHexB
      End If
  'Add it up, return the function value
  HexRGB = "#" & vHexR & vHexG & vHexB
End Function
```

