#r @".\bin\Debug\ExcelTypeProvider.dll"

   // to kill excels
   //get-process | where-object { $_.name -eq "excel" } | sort-object -property "Starttime" -descending | select-object -skip 1 | foreach { taskkill /pid $_.id }
   //get-process | where-object { $_.name -eq "excel" } | foreach { taskkill /F /pid $_.id }
open Samples.FSharp.ExcelProvider
type exc = ExcelFile<"C:\\temp\\Book1.xlsx",false>

let sheet1 = exc().Sheet1
let row1 = sheet1.Row3


