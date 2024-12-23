Late Binding: VBScript's late binding can cause runtime errors if an object or method doesn't exist.  This is especially problematic when working with COM objects or external libraries where version inconsistencies might lead to unexpected failures. 
Example:
```vbscript
Dim objExcel
Set objExcel = CreateObject("Excel.Application")
'Assume Excel is not installed or path is incorrect
 objExcel.Visible = True
```
This will throw an error if Excel is not properly installed or accessible. Early binding (declaring object types explicitly) helps mitigate this but is more complex.