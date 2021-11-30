# The formula below will list files from a specific path to an excel sheet

## Select cell A1.

## Go to formula and then select define names

## Name your formula List_Of_Names in the Name area and then type in =FILES(Sheet1!$A$1) in the refers to area.

## copy the path of the folder and paste it in cell A1

## go to any cell and type the formula below

``` =INDEX(List_Of_Names,ROW(A1)) ```


Copy until you get REF!
