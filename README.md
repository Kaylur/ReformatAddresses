Reformat Addresses (Word VBA Macro)
A Microsoft Word macro that converts single-line addresses into a three-line format, with the suite or unit on its own line.
Example
Before:
```
114 Canal St Ste 503, Pooler, GA 31322
```
After:
```
114 Canal St
Suite 503
Pooler, GA 31322
```
If an address has no suite or unit, that line is left out:
```
200 Main St, Savannah, GA 31401
```
becomes
```
200 Main St
Savannah, GA 31401
```
Supported Unit Types
Input	Output
Ste, Ste., Suite	Suite
Apt, Apt., Apartment	Apartment
Unit	Unit
Bldg, Bldg., Building	Building
#	#
Matching is not case-sensitive. State abbreviations are converted to uppercase, and both 5-digit and ZIP+4 codes are supported.
Installation
Download `ReformatAddresses.bas`.
In Word, press Alt+F11 to open the VBA editor.
Go to File > Import File and select `ReformatAddresses.bas`.
To keep the macro available in all documents, import it into the `Normal` template.
Usage
Select the addresses you want to reformat, or select nothing to process the whole document.
Press Alt+F8, choose `ReformatAddresses`, and click Run.
A message shows how many addresses were reformatted.
All changes can be undone in one step with Ctrl+Z.
Line Break Setting
By default, the macro uses line breaks (Shift+Enter), which keeps the address in a single paragraph. To use separate paragraphs instead, change this line:
```vba
lineBreak = Chr(11)
```
to:
```vba
lineBreak = vbCr
```
Requirements
Microsoft Word 2010 or later for Windows
Not supported on Word for Mac, because the macro uses `VBScript.RegExp`, which is Windows-only
Limitations
Each address must be on its own line (paragraph or table cell). Addresses within a sentence are skipped.
The address must use commas between the street, city, and state, as in `Street, City, ST 12345`.
If an address has two unit types, such as `Bldg 2 Ste 100`, only the last one moves to its own line.
Unit types not listed above (for example, Fl or Rm) stay on the street line.
Review the results before use. The macro replaces text in place, so verify formatting against the original addresses.
