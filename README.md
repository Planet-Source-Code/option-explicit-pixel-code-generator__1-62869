<div align="center">

## Pixel Code Generator


</div>

### Description

This small, almost totally useless project scans all the pixels in a bitmap stored in a picturebox and generates SetPixelV code to replicate that bitmap. Automatically generates For loop code if enough consecutive pixels are identical to exceed a FORLOOP_THRESHOLD value.

Generated code is written to the clipboard for pasting into other projects. Why did I do something so useless, you ask? Well, as many

of you know, my hobby at PSC is inflicting usercontrols on you. I prefer things like checkmarks for checkboxes drawn by code as opposed to being stored in imagelists or the like. I don't like dependencies of any kind in controls. So, I draw them using LineTo or SetPixelV. The problem is, I can't even draw a stick figure. I see these nice custom checkmarks in web sites and I couldn't replicate them, so I wrote this. As far as I know checkmarks aren't copyrighted, so hopefully this wouldn't be considered stealing! To use, just use your favorite screen grabber to grab the checkmark or whatever, then MS Paint or whatever to save just the area you want to disk. (You may want to edit colors, or resize, or some such.) Place this bitmap in Picture1 and run the program. Stop the program, and paste the newly generated code into the "Redraw" sub. Run again and click the button under Picture2. Voila! You can tweak the code to take out parts you don't want fairly easily. Only try this with SMALL (32x32 size or less) bitmaps. I tried with larger bitmaps and while the code generates properly, the code for even a smaller photgraph can easily exceed 10,000 lines. A VB procedure can't be larger than 64K. I included a small calculator icon example with the code already loaded into the Redraw routine to get you started. I wouldn't even use this for something the size of the calculator icon but did it just to give you the idea. I'm sure this could be optimized (especially the string concatenation) but it runs great on small bitmaps and I'm not going to endlessly tweak this. Let the flames begin :-)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2005-10-13 01:02:08
**By**             |[Option Explicit](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/option-explicit.md)
**Level**          |Beginner
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Pixel\_Code19398910132005\.zip](https://github.com/Planet-Source-Code/option-explicit-pixel-code-generator__1-62869/archive/master.zip)








