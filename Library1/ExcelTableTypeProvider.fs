namespace Samples.FSharp

module ExcelProvider =
   open ClosedXML.Excel
   open Microsoft.FSharp.Core.CompilerServices
   open Microsoft.Office.Interop
   open System.Reflection
   open System
   open System.Collections.Generic   
   open System.IO
   open Samples.FSharp.ProvidedTypes
   //open System.Xml.Linq

   open System.Diagnostics
   type  ExcelFileInternal(filename:string) =
        //TK: refactor to return correct types, as opposed to strings - DONE
        //TK: document
        let dict = new Dictionary<string,obj[][]>()
        let sheetsData = match Path.GetExtension(filename) with
                         |".xlsx"-> let wb = new XLWorkbook(filename) //doc is >= 2007
                                    let mysheets = wb.Worksheets
                                    let defnames = wb.NamedRanges
                                    let getData (rng:IXLRange) = 
                                                let sheetData = seq { for r in rng.RowsUsed() do
                                                                            yield seq {for c in r.Cells() do
                                                                                        yield c.GetValue()} |> Array.ofSeq
                                                                    } |> Array.ofSeq
                                                sheetData
                                    //add Sheets
                                    for sht in mysheets do
                                        let rng  = sht.RangeUsed()
                                        if rng <> null then
                                            let data = getData rng
                                            dict.Add(sht.Name,data)
                                    //add named ranges//TK: refactor for multiple ranges within a namedrange
                                    for namedrng in defnames do
                                        let rng  = namedrng.Ranges |> Seq.exactlyOne
                                        if rng <> null then
                                            let data = getData rng
                                            dict.Add(namedrng.Name,data)
                                    dict
                            |".xls"-> let xlApp = new Excel.ApplicationClass()//doc is < 2007, have to use offfice interop. Ho hum ...
                                      xlApp.Visible <- false
                                      xlApp.ScreenUpdating <- false
                                      xlApp.DisplayAlerts <- false;
                                      let xlWorkBookInput = xlApp.Workbooks.Open(filename)
                                      let mysheets = seq { for  sheet in xlWorkBookInput.Worksheets do yield sheet :?> Excel.Worksheet }
                                      let names = seq { for name in xlWorkBookInput.Names do yield name :?> Excel.Name}
                                      let getData (xlRangeInput:Excel.Range) = 
                                                                 let rows_data = seq { for row  in xlRangeInput.Rows do
                                                                                        yield row :?> Excel.Range }
                                                                 let res = 
                                                                   seq { for line_data in rows_data do 
                                                                         yield ( seq { for cell in line_data.Columns do
                                                                                        if (cell  :?> Excel.Range).Value2 <> null && (cell  :?> Excel.Range).Value2.ToString() <> String.Empty then
                                                                                             yield (cell  :?> Excel.Range).Value2}
                                                                                  |> Seq.filter (fun c -> c.ToString() <> String.Empty) |>Seq.toArray
                                                                               )
                                                                      }
                                                                      |> Seq.toArray |> Array.filter (fun r-> r.Length > 0)
                                                                 res
                                      for sht in mysheets do
                                            let xlRangeInput = sht.UsedRange
                                            if xlRangeInput <> null then
                                                let data = getData xlRangeInput
                                                if data.Length > 0 then
                                                    dict.Add(sht.Name,data)
                                      for rng in names do
                                            let xlRangeInput = rng.RefersToRange
                                            if xlRangeInput <> null then
                                                let data = getData xlRangeInput
                                                if data.Length > 0 then
                                                    dict.Add(rng.Name,data)
                                      xlWorkBookInput.Close()
                                      xlApp.Quit()
                                      dict
                            |_     -> failwithf "%s is not a valid path for a spreadsheet " filename
        member __.SheetAndRangeNames = dict.Keys |> Seq.map (fun k -> k) |> Array.ofSeq
        member __.SheetData(name:string)  = sheetsData.[name]
   [<TypeProvider>]
   type public ExcelProvider(cfg:TypeProviderConfig) as this =
      inherit TypeProviderForNamespaces()

      // Get the assembly and namespace used to house the provided types 
      let asm = System.Reflection.Assembly.GetExecutingAssembly()
      let ns = "Samples.FSharp.ExcelProvider"
      // Create the main provided type
      let excTy = ProvidedTypeDefinition(asm, ns, "ExcelFile", Some(typeof<ExcelFileInternal>))
      do excTy.AddXmlDoc("The main provided type - static parameters of filename:string, forcestring:bool. \n If forcestring, all data will be coerced to string")
      // Parameterize the type by the file to use as a template
      let filename = ProvidedStaticParameter("filename", typeof<string>)
      let forcestring = ProvidedStaticParameter("forecstring", typeof<bool>,false)
      let staticParams = [filename
                          forcestring]
      do excTy.DefineStaticParameters(staticParams, fun tyName paramValues ->
        let filename,forcestring = match paramValues with
                                   | [| :? string  as filename;   :? bool as forcestring |] -> (filename, forcestring)
                                   | [| :? string  as filename|] -> (filename, false)
                                   | _ -> ("no file specified to type provider",  true)
        let ex = ExcelFileInternal(filename)
        // define the provided type, erasing to excelFile
        let ty = ProvidedTypeDefinition(asm, ns, tyName, Some(typeof<ExcelFileInternal>))
        
        // add a parameterless constructor
        ty.AddMember(ProvidedConstructor([], InvokeCode = fun [] -> <@@  new ExcelFileInternal(filename) @@>))
        // TK: add second ctor with filename, forcestring
        //for each worksheet (with data), add a property of provided type shtTyp
        for sht in ex.SheetAndRangeNames do
            let shtTyp = if  forcestring then 
                            ProvidedTypeDefinition(sht,Some typeof<string[][]>,HideObjectMethods = true)
                         else
                            ProvidedTypeDefinition(sht,Some typeof<obj[][]>,HideObjectMethods = true)
            do shtTyp.AddXmlDoc(sprintf "Type for data in %s" sht)
            let data = ex.SheetData(sht)
            let rowTyp = ProvidedTypeDefinition("Row", 
                                                (if forcestring then 
                                                    Some typeof<string[]>
                                                else 
                                                    Some typeof<obj[]>), 
                                                HideObjectMethods = true)
            shtTyp.AddMember(rowTyp)
            let rowsProp = ProvidedProperty(propertyName = "Rows",
                                            propertyType = typedefof<seq<_>>.MakeGenericType(rowTyp),
                                            GetterCode = if forcestring then 
                                                            (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:string[][])
                                                                                                  |> Seq.skip 1 
                                                                                                  |> Array.ofSeq 
                                                                                                  |> Array.map ( fun row -> row |> Array.map (fun cel -> cel.ToString())) 
                                                                                                 @@>)
                                                         else
                                                            (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:obj[][])
                                                                                                  |> Seq.skip 1 
                                                                                                  |> Array.ofSeq 
                                                                                                 @@>)
                                                         )
            let colHdrs = data.[0]
            colHdrs |> Array.iteri (fun j col -> let propName = match col.ToString() with
                                                                |"" -> "Col" + j.ToString()
                                                                |_  ->  col.ToString()
                                                 let valueType, gettercode  = if forcestring then typeof<string>,(fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:string[])).[j] @@>)
                                                                              else
                                                                              match data.[1].[j] with
                                                                              | :? bool   -> typeof<bool>,(fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:obj[])   |> Array.map (fun o -> bool.Parse(o.ToString()))).[j] @@>)
                                                                              | :? string -> typeof<string>,(fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:obj[]) |> Array.map (sprintf "%A")).[j] @@>)
                                                                              | :? float  -> typeof<float>,(fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:obj[])  |> Array.map (fun o -> Double.Parse(o.ToString()))).[j] @@>)
                                                                              |_          -> typeof<obj>,(fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:obj[]).[j] @@>)
                                                 let colp = ProvidedProperty(propertyName = propName,
                                                                             propertyType = valueType,
                                                                             GetterCode= gettercode)
                                                 rowTyp.AddMember(colp))
            data |> Array.iteri (fun i r -> if i > 0 then //skip header col
                                                let rowTyp =  if  forcestring then
                                                                ProvidedTypeDefinition("Row" + i.ToString(),Some typeof<string[]>,HideObjectMethods = true)
                                                              else
                                                                ProvidedTypeDefinition("Row" + i.ToString(),Some typeof<obj[]>,HideObjectMethods = true)
                                                let getCode = if forcestring then
                                                                (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:string[][]).[i] @@>)
                                                              else
                                                                (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:obj[][]).[i] @@>)
                                                let rowp = ProvidedProperty(propertyName = "Row" + i.ToString(),
                                                                            propertyType = rowTyp,
                                                                            GetterCode = getCode
                                                                            )
                                                colHdrs |> Array.iteri (fun j col -> let propName = match col.ToString() with
                                                                                                    |"" -> "Col" + j.ToString()
                                                                                                    |_  ->  col.ToString()
                                                                                     let valueType, gettercode  = if forcestring then typeof<string>,(fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:string[])).[j] @@>)
                                                                                                                  else
                                                                                                                  match r.[j] with
                                                                                                                  | :? bool   -> typeof<bool>,(fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:obj[])   |> Array.map (fun o -> bool.Parse(o.ToString()))).[j] @@>)
                                                                                                                  | :? string -> typeof<string>,(fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:obj[]) |> Array.map (sprintf "%A")).[j] @@>)
                                                                                                                  | :? float  -> typeof<float>,(fun (args:Quotations.Expr list) -> <@@ ((%%args.[0]:obj[])  |> Array.map (fun o -> Double.Parse(o.ToString()))).[j] @@>)
                                                                                                                  |_          -> typeof<obj>,(fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:obj[] ).[j] @@>)
                                                                                     let colp = ProvidedProperty(propertyName = propName,
                                                                                                                 propertyType = valueType,
                                                                                                                 GetterCode= gettercode)
                                                                                     colp.AddXmlDoc(sprintf "Value for Cell in Col%d in Row%d in range %s" j i sht) 
                                                                                     rowTyp.AddMember(colp)
                                                                       )
                                                shtTyp.AddMember(rowTyp)
                                                rowp.AddXmlDoc(sprintf "Data for Row%d in range %s" i sht)
                                                shtTyp.AddMember(rowp)
                                )
            //data |> Array
            let shtGetCode = if forcestring then
                                (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:ExcelFileInternal).SheetData(sht) |> Array.map ( fun row -> row |> Array.map (fun cel -> cel.ToString())) @@>)
                             else
                                (fun (args:Quotations.Expr list) -> <@@ (%%args.[0]:ExcelFileInternal).SheetData(sht) @@>)
            let shtp = ProvidedProperty(propertyName = sht, 
                                        propertyType = shtTyp,
                                        GetterCode= shtGetCode
                                       )
            do shtp.AddXmlDoc(sprintf "Data in %s" sht)
            shtTyp.AddMember(rowsProp)
            ty.AddMember(shtTyp)
            ty.AddMember(shtp)
        ty
        )
      // add the type to the namespace
      do this.AddNamespace(ns, [excTy])
   [<TypeProviderAssembly>]
   do()