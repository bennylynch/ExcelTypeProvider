#r @".\bin\Debug\ExcelTypeProvider.dll"

   // to kill excels
   //get-process | where-object { $_.name -eq "excel" } | sort-object -property "Starttime" -descending | select-object -skip 1 | foreach { taskkill /pid $_.id }
   //get-process | where-object { $_.name -eq "excel" } | foreach { taskkill /F /pid $_.id }
open Samples.FSharp.ExcelProvider
type exc = ExcelFile<"C:\\temp\\BookTest.xlsx",false>

let file = new exc(@"C:\\temp\\BookTest.xlsx")

for row in file.Sheet1.Rows do
    printfn "%s" row.BID
let sht1Row1Col20 = file.Sheet1.Row1.BID