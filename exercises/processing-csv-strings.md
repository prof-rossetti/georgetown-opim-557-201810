# "Processing CSV Strings" Exercise

## Learning Objectives

  1. Find practical applications for learning new programming concepts like loops, arrays, and string operations (i.e. splitting strings).
  2. Gain familiarity with processing machine-readable data in Comma-Separated Values (CSV) format.

## Instructions

Write VBA code that will process the provided Comma-Separated Values (CSV) string (i.e. the `MyStr` variable below) into a corresponding spreadsheet of cells.

Open a new workbook, rename the first sheet to "Interface, and create a new sheet called "Data".

In the "Interface" sheet, insert a command button, and inside the button's click event sub-procedure, paste the code below:

```vb
Dim MyStr As String

MyStr = "city,name,league" & vbNewLine & _
        "New York,Mets,Major" & vbNewLine & _
        "New York,Yankees,Major" & vbNewLine & _
        "Boston,Red Sox,Major" & vbNewLine & _
        "Washington,Nationals,Major" & vbNewLine & _
        "New Haven,Ravens,Minor"

MsgBox(MyStr)

' TODO: write some VBA code here!
```

Write code inside the command button's click event sub-procedure that will clear the contents of the "Data" sheet, write the desired spreadsheet output there (see below), and finally activate that sheet:

city | name | league
--- | --- | ---
New York | Mets | Major
New York | Yankees | Major
Boston | Red Sox | Major
Washington | Nationals | Major
New Haven | Ravens | Minor
