<!-- ---
!-- Timestamp: 2025-03-24 21:36:05
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/README.md
!-- --- -->

# PySigMacro

A Python library for controlling SigmaPlot

## Install Python on Windows

https://www.python.org/ftp/python/3.13.2/python-3.13.2-amd64.exe

``` ps1
where.exe python
```

## Install PySigMacro via pip
``` ps1
cd C:\path\to\SigMacro\pysigmacro
python.exe -m pip install pip -U
python.exe -m pip install -r requirements.txt
python.exe -m pip install -e .
```

## (Optional) Work on WSL

``` bash
## Directory
mkdir -p ~/proj/
ln -s /mnt/c/Users/<YOUR-USER-NAME>/path/to/SigMacro/PySigMacro ~/proj/PySigMacro

## Executables
mkdir -p ~/.win-bin
ln -s /mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe ~/.win-bin/powershell.exe
ln -s /mnt/c/Program Files (x86)/SigmaPlot/SPW12/Spw.exe ~/.win-bin/sigmaplot.exe
export PATH:$PATH:~/.win-bin

## Aliases
alias 'kill-sigmaplot'='powershell.exe -File "$(wslpath -w /home/<YOUR-USER-NAME>/win/program_files_x86/<YOUR-USER-NAME>/kill-sigmaplot.ps1)"'
alias 'python.exe'='powershell.exe python.exe'
alias 'ipython.exe'='powershell.exe ipython.exe --no-autoindent'
```

<!-- ## Environmental Variables
 !-- ```powershell
 !-- $env:SIGMACRO_PATH = "C:/Users/<YOUR-USER-NAME>/Documents/SigmaPlot/SPW12/SigMacro.JNB"
 !-- ``` -->

```vba
Option Explicit

' Requires a reference to Microsoft Scripting Runtime (Tools > References in the VBA editor)

' ================================================================================
' Function: GetMacroArguments
' Purpose: Parses a string of arguments into a Dictionary object.
' Arguments should be in the format "key1=value1,key2=value2,...".
' Returns: A Scripting.Dictionary object containing the parsed arguments.
' ================================================================================
Function GetMacroArguments(argString As String) As Object
Dim args As Object ' As Scripting.Dictionary
Set args = CreateObject("Scripting.Dictionary")
Dim pairs As Variant
pairs = Split(argString, ",") ' Assuming comma-separated key=value pairs
Dim pair As Variant
Dim parts As Variant
For Each pair In pairs
parts = Split(Trim(pair), "=") ' Assuming key and value are separated by =
If UBound(parts) = 1 Then
args(Trim(parts(0))) = Trim(parts(1))
End If
Next pair
Set GetMacroArguments = args
End Function

' ================================================================================
' Sub: SetMacroVariablesFromWorksheet
' Purpose: Creates or clears a worksheet named "PythonMacroVars" and writes
' a dictionary of variables to it. Each key-value pair from the
' dictionary will be written to a new row, with the key in the first
' column and the value in the second.
' Parameters:
' varDict (Object): A Dictionary object passed from Python containing
' variable names as keys and their values.
' ================================================================================
Sub SetMacroVariablesFromWorksheet(varDict As Object) ' Expecting a Dictionary object passed from Python
Dim ws As Object ' As SigmaPlot.Worksheet
' Assuming the first worksheet in the active notebook will be used
On Error Resume Next
Set ws = ActiveDocument.NotebookItems("PythonMacroVars").Object ' Try to get an existing worksheet
On Error GoTo 0
If ws Is Nothing Then
Set ws = ActiveDocument.NotebookItems.Add(1) ' Add a new worksheet (1 = Worksheet)
ws.Name = "PythonMacroVars"
Else
ws.DataTable.Clear
End If

Dim key As Variant
Dim row As Long: row = 0
Dim col As Long

For Each key In varDict.Keys
col = 0
ws.DataTable.Cell(col, row) = key
col = col + 1
ws.DataTable.Cell(col, row) = varDict(key)
row = row + 1
Next key
End Sub

' ================================================================================
' Function: GetVariableFromWorksheet
' Purpose: Retrieves the value of a variable from the "PythonMacroVars" worksheet.
' Parameters:
' varName (String): The name of the variable to retrieve (case-insensitive).
' Returns: The value of the variable as a Variant, or an empty string if not found.
' ================================================================================
Function GetVariableFromWorksheet(varName As String) As Variant
Dim ws As Object ' As SigmaPlot.Worksheet
On Error Resume Next
Set ws = ActiveDocument.NotebookItems("PythonMacroVars").Object
On Error GoTo 0
If Not ws Is Nothing Then
Dim lastRow As Long
Dim lastCol As Long
ws.DataTable.GetMaxUsedSize lastCol, lastRow
Dim i As Long
For i = 0 To lastRow
If Trim(LCase(ws.DataTable.Cell(0, i))) = Trim(LCase(varName)) Then
GetVariableFromWorksheet = ws.DataTable.Cell(1, i)
Exit Function
End If
Next i
End If
' Return an empty string if the variable is not found
GetVariableFromWorksheet = ""
End Function

' ================================================================================
' Example Macro: MyVersatileMacro
' Purpose: Demonstrates how to use GetMacroArguments and GetVariableFromWorksheet.
' Parameters (passed as a string when running from Python):
' plot_type (String, optional): The type of plot to create (e.g., "scatter", "line").
' color (String, optional): The color of the plot (e.g., "red", "blue").
' Variables (passed from Python via worksheet):
' data_file (String): The path to the data file.
' threshold (Double): A threshold value for analysis.
' ================================================================================
Sub MyVersatileMacro(argString As String)
' Parse arguments
Dim args As Object
Set args = GetMacroArguments(argString)

' Retrieve variables from worksheet
Dim dataFile As String
dataFile = GetVariableFromWorksheet("data_file")
Dim threshold As Double
threshold = CDbl(GetVariableFromWorksheet("threshold"))

' Access arguments with default values
Dim plotType As String
If args.Exists("plot_type") Then
plotType = args("plot_type")
Else
plotType = "line" ' Default plot type
End If

Dim plotColor As String
If args.Exists("color") Then
plotColor = args("color")
Else
plotColor = "blue" ' Default color
End If

' Your macro logic here using the parsed arguments and variables
MsgBox "Running MyVersatileMacro..."
MsgBox "Data File from Python: " & dataFile
MsgBox "Threshold from Python: " & threshold
MsgBox "Plot Type Argument: " & plotType
MsgBox "Plot Color Argument: " & plotColor

' Example: You could use these variables and arguments to control
' how you import data, create plots, and perform analysis.

End Sub
```

**Python Utility Function to Pass Variables and Run Macro:**

```python
import win32com.client

def run_sigmaplot_macro(macro_name, variables=None, arguments=None):
"""
Runs a SigmaPlot macro, optionally passing variables and arguments.

Args:
macro_name (str): The name of the SigmaPlot macro to run.
variables (dict, optional): A dictionary of variables to pass to the macro.
These will be written to a worksheet named "PythonMacroVars".
Defaults to None.
arguments (dict, optional): A dictionary of arguments to pass to the macro as a
comma-separated string of key=value pairs. Defaults to None.
"""
try:
oSigmaPlot = win32com.client.Dispatch("SigmaPlot.Application")
oNotebook = oSigmaPlot.ActiveDocument

# Find the macro item
macro_item = None
for item in oNotebook.NotebookItems:
if item.Name == macro_name and item.ObjectType == 10: # 10 is the ObjectType for MacroItem
macro_item = item
break

if not macro_item:
print(f"Macro '{macro_name}' not found.")
return

# Pass variables via worksheet
if variables:
# Find or create the SetMacroVariablesFromWorksheet macro
set_vars_macro = None
for item in oNotebook.NotebookItems:
if item.Name == "SetMacroVariablesFromWorksheet" and item.ObjectType == 10:
set_vars_macro = item
break
if set_vars_macro:
# Prepare the dictionary for VBA
vba_dict = win32com.client.Dispatch("Scripting.Dictionary")
for key, value in variables.items():
vba_dict.Add(key, value)
set_vars_macro.Run(vba_dict)
else:
print("Warning: Macro 'SetMacroVariablesFromWorksheet' not found. Ensure it exists in your notebook.")

# Prepare the argument string
arg_string = ""
if arguments:
arg_list = [f"{key}={value}" for key, value in arguments.items()]
arg_string = ",".join(arg_list)

# Run the macro
if arg_string:
macro_item.Run(arg_string)
else:
macro_item.Run()

except Exception as e:
print(f"An error occurred: {e}")
finally:
# Ensure SigmaPlot is released (optional, depends on how you want to manage the application)
# oSigmaPlot = None
pass

# Example usage from Python:
if __name__ == "__main__":
macro_name_to_run = "MyVersatileMacro" # Replace with the actual name of your macro
python_variables = {"data_file": "C:\\path\\to\\your\\data.csv", "threshold": 0.05}
macro_arguments = {"plot_type": "scatter", "color": "red"}

# First, ensure the utility macros (GetMacroArguments, SetMacroVariablesFromWorksheet)
# are defined in your SigmaPlot notebook.

# Then, call the function to pass variables and run your macro.
run_sigmaplot_macro(macro_name_to_run, variables=python_variables, arguments=macro_arguments)
```

**Explanation:**

**VBA Macro Functions:**

1. **`GetMacroArguments(argString As String) As Object`**:
* Takes a string `argString` where arguments are formatted as `key1=value1,key2=value2,...`.
* Uses the `Split` function to break the string into key-value pairs based on the comma delimiter.
* For each pair, it further splits based on the equals sign to separate the key and the value.
* Stores these key-value pairs in a `Scripting.Dictionary` object (you need to add a reference to "Microsoft Scripting Runtime" in the VBA editor: Tools > References).
* Returns the `Dictionary` object.

2. **`SetMacroVariablesFromWorksheet(varDict As Object)`**:
* Takes a `Dictionary` object `varDict` as input (this will be passed from Python).
* Tries to get an existing worksheet named "PythonMacroVars". If it doesn't exist, it creates a new one. If it exists, it clears its data.
* Iterates through the keys and values of the `varDict`.
* Writes each key-value pair to a new row in the "PythonMacroVars" worksheet, with the key in the first column and the value in the second.

3. **`GetVariableFromWorksheet(varName As String) As Variant`**:
* Takes a `varName` (string) as input.
* Tries to access the "PythonMacroVars" worksheet.
* Iterates through the first column of the worksheet, looking for a cell that matches the `varName` (case-insensitive).
* If a match is found, it returns the value from the corresponding cell in the second column of the same row.
* If the `varName` is not found, it returns an empty string.

**Python Utility Function:**

1. **`run_sigmaplot_macro(macro_name, variables=None, arguments=None)`**:
* Takes the name of the SigmaPlot macro to run (`macro_name`).
* Optionally accepts a `variables` dictionary and an `arguments` dictionary.
* Uses `win32com.client.Dispatch("SigmaPlot.Application")` to connect to the active SigmaPlot application.
* Finds the `MacroItem` object in the active notebook based on its name.
* If `variables` are provided:
* It tries to find a macro named "SetMacroVariablesFromWorksheet" in the notebook.
* If found, it creates a `Scripting.Dictionary` object in VBA and populates it with the Python `variables`.
* It then runs the "SetMacroVariablesFromWorksheet" macro, passing the VBA dictionary as an argument.
* If `arguments` are provided:
* It formats the `arguments` dictionary into a comma-separated string of `key=value` pairs.
* Finally, it runs the specified `macro_name`. If `arguments` were provided, it passes the formatted argument string to the macro's `Run` method.

**How to Use:**

1. **In SigmaPlot:**
* Open your SigmaPlot notebook.
* Go to Tools > Macro > Visual Basic Editor.
* Insert a new module (Insert > Module).
* Paste the VBA code for `GetMacroArguments`, `SetMacroVariablesFromWorksheet`, `GetVariableFromWorksheet`, and your main macro function (e.g., `MyVersatileMacro`) into this module.
* Make sure you have a reference to "Microsoft Scripting Runtime" (Tools > References).

2. **In Python:**
* Save the Python code as a `.py` file.
* Replace `"MyVersatileMacro"` with the actual name of your SigmaPlot macro.
* Set the `python_variables` dictionary with the variables you want to pass to the macro.
* Set the `macro_arguments` dictionary with the keyword arguments you want to pass.
* Run the Python script.

The Python script will connect to SigmaPlot, write the variables to the "PythonMacroVars" worksheet, and then run your specified macro, passing the arguments as a string. Your SigmaPlot macro can then use the `GetMacroArguments` function to parse the arguments and the `GetVariableFromWorksheet` function to retrieve the variables from the worksheet.

<!-- EOF -->