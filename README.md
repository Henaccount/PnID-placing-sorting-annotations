# PnID-setting-equipment-tags
Very simple example code for placing certain annotations for preselected items of the P&ID. The first annotation will be placed at a defined point of the drawing and the following beneath it with a defined distance. If annotations are used that build a table row, then the result can act as a bill of materials table.
Please refer to the more advanced example below if you are planning to parameterize the input like: basis point, offset, selection of drawing items, sorting of the table rows, ...

# PnID-sorting-equipment-tags
<pre>
Problem to fix: 
Equipment info tags have to be placed in a certain order gridwise on the bottom of the P&ID drawing. This is a time consuming task and the order has to be corrected during the drawing process. The command EQPANNOSORT shall do this ordering of already placed equipment annotations.

How to load the dll:
First rename pssCommands.txt in pssCommands.dll
type in _netload and choose the file or
put the following line in the acad2016doc.lsp for automatic loading on startup of Plant/P&ID e.g.:
(command "_netload" " C:/Program Files/Autodesk/AutoCAD 2015/PLNT3D/pssCommands.dll ")
Use slashes not backslashes!
In case of P&ID standalone application it will be the PNID folder.
Use the PLNT3D/PNID folder, because it is a trusted location, so you will not be prompted every time you start the software.
Prerequisites:
•	The annotations that you want to sort, needs to have a parameter “.Description.
•	The number of the tag field has the following structure, brackets mean optional: 
letter(letter)number1-99999(letter)

Commands:
command	EQPANNOSORT “originX=0,originY=0,originZ=0,xshift=100,yshift=100”

Choose appropriate parameters
action	Will sort the equipment tags:
1.	Sort by first letter plus number with trailing zeros if necessary (5 digits)
2.	Then sort by first letter plus number with trailing zeros if necessary (5 digits)
After this sorting, the script will list the sorted tags in rows and will create a new column if either the number+lastletters changes or when the firstchar changes.
CUI	You can put the command in your user interface like in the following screenshot:


 

Workflow:
•	When placing equipment, place also the equipment info annotation, but it can be anywhere
•	To sort the annotations, call the command (if set up like in the screenshot, rightclick in the drawing space and click the command)
</pre>
