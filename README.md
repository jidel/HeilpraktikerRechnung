# HeilpraktikerRechnung
Simples Rechnungsprogramm für Heilpraktiker. Erlaubt die Ausweisung von Leistung nach der Gebührenordnung für Ärzte (**GOÄ**), dem Leistungsverzeichnis klassischer Homöopathie (**LVKH**) und dem Gebührenverzeichnis für Heilpraktiker (**GebüH**). Leistungsverzeichnis muss selbst gepflegt werden.

![application_full](https://cloud.githubusercontent.com/assets/6048968/13897856/69c46fe0-edc0-11e5-8b52-1951d2faf560.PNG)

## Development Instructions
### Composing and Decomposing
Use the file ``compose.cmd`` (double click while holding ⇧ SHIFT) to reconstruct the original MS Access file from this repository.  
If you want ot save changes you made be sure to safe them and execute ``decompose.cmd`` (while holding ⇧ SHIFT).
Because of some strange MS Access workings, it could be that the file ``Rechnungsstellung_stub.accdb`` is recognized by Git as changed even though it was not. Be sure to only commit this file if it was changed.

### Backend
All data is located in a separat database file (``Datenbank.accdb``) in the folder ``Backend``.
More information soon!

## Deployment
The intended way to ship the application is to create an ``.accde`` file. After that, do not forget to change the applications backup location to fit your clients system.  
In order to run this newly created file, install the [MS Access 2007 Runtime Environment](https://www.microsoft.com/download/details.aspx?id=4438) on the clients machine.

### MS Access Trusted Locations
Use ``AddPath.exe`` to add the project to MS Access' trusted locations. This allows to start macros without user interaction. 

## Credits
Thanks to Oliver for his solution to compose and decompose the MS Access front-end: http://stackoverflow.com/a/211210

## License
**The MIT License (MIT)**  
Copyright (c) 2016 Johannes Idelhauser

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
