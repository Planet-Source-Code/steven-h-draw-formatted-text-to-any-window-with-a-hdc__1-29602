<div align="center">

## Draw FORMATTED Text to any window with a hDc


</div>

### Description

Allows you to draw Formatted (diffrent fonts, sizes, colours) text to anything with a valad hDc - only tested against Screen Objects - Converts Font Object to LOGFONT struct

IMPUTS: StdFontEx, Rect - Bottom Top Left Right, hDc, Text, Text Allignment Flags.

Either provided as Paramaters of the .Draw procedure or as Properties of the cTextEx Object

RETURNS: no Returns

SIDE EFFECTS: no Knowen Side Effects

Misc: you MUST pass a StdFontEx object as either a Paremter or a Property of the cTextEx object. cStdFontEx mirrors functonalaty provided by the StdFont object. i only added a Colour Property (it makes sence to do this) future revisions of this class will encorperate diffrent brush styles thus the reasoning for mirroring the vb font object
 
### More Info
 
StdFontEx, Rect - Bottom Top Left Right, hDc, Text, Text Allignment Flags.

Either provided as Paramaters of the .Draw procedure or as Properties of the cTextEx Object

you MUST pass a StdFontEx object as either a Paremter or a Property of the cTextEx object. cStdFontEx mirrors functonalaty provided by the StdFont object. i only added a Colour Property (it makes sence to do this) future revisions of this class will encorperate diffrent brush styles thus the reasoning for mirroring the vb font object

no Returns

no Knowen Side Effects


<span>             |<span>
---                |---
**Submitted On**   |2001-12-09 18:14:12
**By**             |[Steven H](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steven-h.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Draw\_FORMA404301292001\.zip](https://github.com/Planet-Source-Code/steven-h-draw-formatted-text-to-any-window-with-a-hdc__1-29602/archive/master.zip)








