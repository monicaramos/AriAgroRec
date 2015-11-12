==============================
Enumerate Icon Resources
Visual Basic source code
by Giorgio Brausi (aka GIBRA)
http://www.vbcorner.net/eng
email: vbcorner@vbcorner.net
==============================

April 2004

This file is part of the "Enum Icon Resources" VB project.

===============
WARNING
===============
If you have downloaded this project from Planet Source-Code
four DLL is missing (because PSC don't permit to add them).
This 4 DLL is very useful to know better how to work the project:
- is XP_ICONS_16.DLL
- is XP_ICONS_24.DLL
- is XP_ICONS_48.DLL
- is XP_ICONS_ALL.DLL

If you want, you can download the entire project with this and
other file used by the project from my web site (see above)
from section "Projects"


Topic:
 1) How to use 'really' in VB the 32bpp icons (Alpha channel)
 2) How to create icons library (DLL) with 32bpp icons  (Alpha channel)

======================================================================
1) How to use 'really' in VB the 32bpp icons (Alpha channel)
======================================================================

When we try to load a 32bpp icon - with Alpha Channel - 
to any graphic control (like Form, PictureBox, Image, ...) 
Visual Basic return the error: "Invalid image."

Many icon files embedded more than one image format, like
below:
16x16 4bpp      16x16 16bpp     16x16 32bpp
32x32 4bpp      32x32 16bpp     32x32 32bpp
48x48 4bpp      48x48 16bpp     48x48 32bpp

By loading this 'misc' icon file (with 9 fotmats) we don't receive any 
error because VB choose itself automatically the format to load.

Now, suppose that we have a icons library, or variuos icon files,
that embedded ONLY 32bpp images (that is Windows XP format and
Alpha Channel)?
VB refuse to load it! Therefore we can't use it.

-----------
SOLUTION
-----------
Naturally, with a lot of API functions we can obtain the
result. This source code show how to does it.

HOW TO:
-------
For more info about source code see the README.TXT.
However, source code is full commented, so you can know
better what is does.


======================================================================
2) How to create icons library (DLL) with 32bpp icons  (Alpha channel)
======================================================================

Another issue!

Visual Basic, since version 5.0, provide the Resource Editor which 
allow us to create resources libraries (DLL).

Although, even with this tool, if you try to load a icon which contains 
only image a 32bpp (with Alpha channel) the same error occur: 
"Invalid image.".

Also, if you load a icon with 'misc' format image (se above) you DON'T
receive any error, BUT the icon is loaded as 'CUSTOM' resource type!

-----------
SOLUTION
-----------
Fortunately, even this time, we can solve this annoying problem! ;-))

In this case we must create a DLL library by using the: 
  Resource Compiler (RC.EXE) 
provided with Microsoft Visual Basic.

HOW TO:
-------
Step 1: Create a RES file which contain your icons

   - In order to create a RES file you must create a RC script file.
     The RC script is a text file formatted as below:


	////////////	<- this is comments
	// Icons
	////////////

	101  ICON  MOVEABLE PURE   "cd rom drive.ico"
	102  ICON  MOVEABLE PURE   "harddisk drive.ico"
	103  ICON  MOVEABLE PURE   "my computer.ico"

	e so on...


   - To create a script RC file:
     a) Create a new text 
     b) Put a complete list of the icons you want to add
     c) save the file with RC extension (i.e. MyIcons.rc)
        Please note that this file MUST to be on the same 
        folder where icons is stored.

   - Now compile the script file:
     a) Open a DOS prompt session on the folder where is
        the icons and script file
     b) run your RC.EXE following by your script filename
        (i.e.:  rc myicons.rc)

     OK! Your myicons.RES is done.
     
     


Step 2: Create a DLL library

   - Start Visual Basic 6.0 environment
   - From NEW PROJECT dialog choose "DLL ActiveX"
     (now you have a project with a class module Class1 only)
   - Add the RES file to the project: 
     from 'Project' menu choose 'Inser file...' then navigate
     to folder which contain your RES file (myicons.res) and
     select it. Click OK to return to VB IDE.
   - Compile the project (from 'File' menu, 'Make... '

   OK! Your "myicons.dll" library is done.
     

   IMPORTANT!!!
   Don't to be curious! Don't open (and not re-save) the RES
   file from VB IDE.
   This may be corrupt your RES file (and you must create it
   again).
 


I hope that my work (and time!) if useful for you.

       Gibra
   Giorgio Brausi
vbcorner@vbcorner.net
