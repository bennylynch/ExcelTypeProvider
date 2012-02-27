﻿// Copyright (c) Microsoft Corporation 2005-2011.
// This sample code is provided "as is" without warranty of any kind. 
// We disclaim all warranties, either express or implied, including the 
// warranties of merchantability and fitness for a particular purpose. 

namespace Samples.FSharpPreviewRelease2011.ExcelProvider

open System.Reflection
open System.IO
open System
open Samples.FSharpPreviewRelease2011.ProvidedTypes
open Microsoft.FSharp.Core.CompilerServices
open System.Text.RegularExpressions
open Microsoft.Office.Interop

// Simple type wrapping CSV data
type ExcelFile(filename) =
    // Cache the sequence of all data lines (all lines but the first)
    let data = 
        seq { for line in File.ReadAllLines(filename) |> Seq.skip 1 do
                yield line.Split(',') |> Array.map float }
        |> Seq.cache
    member __.Data = data

[<TypeProvider>]
type public MiniCsvProvider(cfg:TypeProviderConfig) as this =
    inherit TypeProviderForNamespaces()

    // Get the assembly and namespace used to house the provided types
    let asm = System.Reflection.Assembly.GetExecutingAssembly()
    let ns = "Samples.FSharpPreviewRelease2011.ExcelProvider"

    // Create the main provided type
    let excTy = ProvidedTypeDefinition(asm, ns, "Excel", Some(typeof<obj>))

    // Parameterize the type by the file to use as a template
    let filename = ProvidedStaticParameter("filename", typeof<string>)
    let forcestring = ProvidedStaticParameter("forcestring", typeof<bool>)

    do excTy.DefineStaticParameters([filename ; forcestring], fun tyName paramValues ->
        let (filename, forcestring) = match paramValues with
                                       | [| :? string  as filename;   :? bool as forcestring |] -> (filename, forcestring)
                                       | [| :? string  as filename|] -> (filename, false)
                                       | _ -> ("no file specified to type provider", true)

        // [| :? string as filename ,  :? bool  as forcestring |]
        // resolve the filename relative to the resolution folder
        let resolvedFilename = Path.Combine(cfg.ResolutionFolder, filename)
        
        let xlApp = new Excel.ApplicationClass()
        let xlWorkBookInput = xlApp.Workbooks.Open(resolvedFilename)
        let xlWorkSheetInput = xlWorkBookInput.Worksheets.["Sheet1"] :?> Excel.Worksheet

        let headerLine =  xlWorkSheetInput.Range(xlWorkSheetInput.Range("A1"), xlWorkSheetInput.Range("A1").End(Excel.XlDirection.xlToRight))
        let firstLine  =  xlWorkSheetInput.Range(xlWorkSheetInput.Range("B1"), xlWorkSheetInput.Range("B1").End(Excel.XlDirection.xlToRight))

        // define a provided type for each row, erasing to a float[]
        let rowTy = ProvidedTypeDefinition("Row", Some(typeof<float[]>))

        // add one property per Excel field
        for i in 0 .. headerLine.Columns.Count - 1 do
            let headerText = ((headerLine.Cells.Item(1,i) :?> Excel.Range).Value2).ToString()
            
            let valueType = 
               if  forcestring then
                  typeof<string>
               else
                  if xlApp.WorksheetFunction.IsText(firstLine.Cells.Item(1,i)) then
                     typeof<string>
                  elif  xlApp.WorksheetFunction.IsNumber(firstLine.Cells.Item(1,i)) then
                     typeof<float>
                  else
                     typeof<string>

            // try to decompose this header into a name and unit
            let fieldName, fieldTy =
                    headerText, valueType

                    //TODO
            let prop = ProvidedProperty(fieldName, fieldTy, GetterCode = fun [row] -> <@@ (%%row:float[]).[i] @@>)

            // Add metadata defining the property's location in the referenced file
            prop.AddDefinitionLocation(1, i, filename)
            rowTy.AddMember(prop)
                
        // define the provided type, erasing to excelFile
        let ty = ProvidedTypeDefinition(asm, ns, tyName, Some(typeof<ExcelFile>))

        // add a parameterless constructor which loads the file that was used to define the schema
        ty.AddMember(ProvidedConstructor([], InvokeCode = fun [] -> <@@ ExcelFile(resolvedFilename) @@>))

        // add a constructor taking the filename to load
        ty.AddMember(ProvidedConstructor([ProvidedParameter("filename", typeof<string>)], InvokeCode = fun [filename] -> <@@ ExcelFile(%%filename) @@>))
        
        // add a new, more strongly typed Data property (which uses the existing property at runtime)
        ty.AddMember(ProvidedProperty("Data", typedefof<seq<_>>.MakeGenericType(rowTy), GetterCode = fun [excFile] -> <@@ (%%excFile:ExcelFile).Data @@>))

        // add the row type as a nested type
        ty.AddMember(rowTy)
        ty)

    // add the type to the namespace
    do this.AddNamespace(ns, [excTy])

[<TypeProviderAssembly>]
do()