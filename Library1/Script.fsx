#r @".\bin\Debug\ExcelTypeProvider.dll"

   // to kill excels
   //get-process | where-object { $_.name -eq "excel" } | sort-object -property "Starttime" -descending | select-object -skip 1 | foreach { taskkill /pid $_.id }
   //get-process | where-object { $_.name -eq "excel" } | foreach { taskkill /F /pid $_.id }
open Samples.FSharp.ExcelProvider
type exc = ExcelFile<"C:\\temp\\BookTest.xlsx",true>

let sheet1 = exc().toto
let row1 = sheet1
for row in sheet1.Rows do
    printfn "%A" row.SEC
let row1col1 = sheet1.Row1.Col0

