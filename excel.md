# Microsoft Excel Tips and Tricks

# referencing
- relative reference
- absolute reference [USING $ SIGN TO LOCK CELL]
- mixed reference
- cross reference
  `[WORKBOOK]SHEET!CELL`

# union cells reference
```
=SUM(A1:C3 B2:C5) // spacing between C3 and B2
// which equals
=SUM(B2:C3)
```

# arithmetics important order
```
// most important and first calculated 
// to less important and last calculate
-
%
^
*, /
+, -
&
=, <, >, <=, >=, <>
```

# show all formulas cells [Formulas>Show Formulas]

# paste values [without formatting]
`CTRL+v > CTRL > v`

# copying formula down [SELECT FILLING RANGE > F2 > CTRL+ENTER]
```
for selecting multiple, using 
goto [F5]>Special...>Blanks
= + UP_ARROW
CTRL+ENTER
```

# duplicate picture [CTRL + d]

# Track Changes
```
Review > Track Changes > Highlight Changes... >
[Checked] Track changes while editing. This also shares your workbook
[Checked] When: All
[Unchecked] Highlight changes on screen
Ok > OK > OK > OK
```

# Alt code reference sheet for excel, type special characters

# Unique Data with Advanced Filter
```
select range >
Data > 
Advanced >
Copy to another location >
[Checked] Unique records only >
Copy to >
OK
```

# wild card [*]

# multiple sheet sum [=SUM('FIRST_SHEET:LAST_SHEET'!FIRST_CELL:LAST_CELL)]

# Activate index, list all sheets [RIGHT_CLICK on left_button < > ]

# center image
- import image and set picture's name
- merge cells
- launch Microsoft Visual Basic for Applications [Alt + F11]
- insert module
  ```
  Sub CenterImages()
    With ActiveSheet.Shapes("Picture 1")  'Picture 1 is a name of image'
      .Top = Range("A1:E1").Top + (Range("A1:E1").Height - .Height) / 2 'A1:E1 is the merged cell range'
      .Left = Range("A1:E1").Left + (Range("A1:E1").Width - .Width) / 2
    End With
  End Sub
  ```
- run the module

# 