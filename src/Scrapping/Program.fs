// Learn more about F# at http://docs.microsoft.com/dotnet/fsharp

open System
open FSharp.Data
open System.IO
open OfficeOpenXml

type DryLandCharacteristic = {
    CropSpecies: string;
    IsDryHappened: bool;
}

type GminaDryLandCharacteristic = {
    id: string;
    name: string;
    characteristics: (string * DryLandCharacteristic list) list;
}

type PowiatDryLandCharacteristic = {
    id: string;
    name: string;
    gminas: GminaDryLandCharacteristic list;
}



//ttps://susza.iung.pulawy.pl/wykazy/2021,1001011/
type DryLandHtml = HtmlProvider<"https://susza.iung.pulawy.pl/wykazy/2021,1001062/", ResolutionFolder=__SOURCE_DIRECTORY__>

// Define a function to construct a message to print
let from whom =
    sprintf "from %s" whom

let getGminasInPowiat powId =
    DryLandHtml.Load(sprintf "https://susza.iung.pulawy.pl/wykazy/2021,%s/" powId).Lists.Html.CssSelect "select#sel-gmina > option" 
        |> Seq.map (fun cs -> (cs.Attribute("value").Value(), cs.InnerText()))
        |> Seq.filter (fun (id, _) -> id <> "-")
        
let getGminaDryLandCharacteristics ((gminaId, gminaName)) = 
    let take = 5
    printfn "Getting gmina: %s, %s" gminaId gminaName
    let html = DryLandHtml.Load(sprintf "https://susza.iung.pulawy.pl/wykazy/2021,%s/" gminaId)
    let t =
        html.Html.CssSelect("table.tab-gmina tbody")
        |> Seq.take 4
        |> Seq.map (fun tBody -> 
                let desc = tBody.CssSelect("tr") |> List.ofSeq
                
                desc
                |> List.map (fun tr -> 
                    let tds = tr.CssSelect("td")
                    let col = tr.CssSelect("td") |> List.skip take |> List.head |> (fun h -> h.InnerText())
                    { CropSpecies = tds.Head.InnerText(); IsDryHappened = col.Equals("+") })
            )
        |> List.ofSeq
        |> List.zip [ "Kategoria gleby I"; "Kategoria gleby II"; "Kategoria gleby III"; "Kategoria gleby IV" ]
        
    //printfn "%A" t
       
    {
        id = gminaId;
        name = gminaName;
        characteristics = t
    } 

let load () =
    let powiatsList = 
        DryLandHtml.Load("https://susza.iung.pulawy.pl/wykazy/2021,1014/").Lists.Html.CssSelect "select#sel-pow > option" 
        |> Seq.map (fun cs -> (cs.Attribute("value").Value(), cs.InnerText()))
        |> Seq.filter (fun (id, _) -> id <> "-")
    
    let allGminasPerPowiat = 
        powiatsList
        |> Seq.map (fun (powiatId, powiatName) -> { id = powiatId; name = powiatName; gminas = getGminasInPowiat powiatId |> Seq.map getGminaDryLandCharacteristics |> List.ofSeq })

    allGminasPerPowiat

let transformPowiat (powiats: PowiatDryLandCharacteristic list) =
    let folder map (powiat: PowiatDryLandCharacteristic) = 
        let f = powiat.gminas |> Seq.collect (fun gm -> gm.characteristics |> Seq.collect (fun (cat, chs) -> chs |> Seq.filter (fun ch -> ch.IsDryHappened) |> Seq.map (fun ch -> (cat, ch.CropSpecies, ch.IsDryHappened))))
        f |> Seq.fold (fun m (cat, crops, _) -> m |> Set.add (cat, crops)) map
    
    let addCategoriesWhichHappendedInAllGminas = powiats |> Seq.fold folder Set.empty
    
    (addCategoriesWhichHappendedInAllGminas, powiats)

    //printfn "%A" t2
    //()
    

let writeToFile (header: Set<string * string>, powiats: PowiatDryLandCharacteristic list) = 
    let fi = new FileInfo("./data.xlsx")
    use ep = new ExcelPackage(fi)
    let sheetName = "Dane"
    let wb = ep.Workbook
    let ws = 
        if wb.Worksheets |> Seq.map (fun s -> s.Name) |> Seq.contains sheetName then
            wb.Worksheets.[sheetName]       //if the sheet already exists, reference it
        else wb.Worksheets.Add(sheetName)   //otherwise, create it

    let startCol = 2
    //helper function to write data based on the start row and column
    let writeData iR iC data =
        data
            |> Seq.iteri (fun i row ->
                row
                |> Seq.iteri (fun k value ->
                    ws.Cells.[iR + i,iC + k].Value <- value))
    let mapHeaer = header |> Seq.mapi (fun i h -> (h, i)) |> Map.ofSeq

    header |> Seq.iteri (fun iC headVals -> 
        ws.Cells.[1, iC + startCol].Value <- (headVals |> snd) 
        ws.Cells.[2, iC + startCol].Value <- (headVals |> fst)) 
   
    let r = 
        powiats 
        |> Seq.collect (fun powiat -> 
                let gminasData =
                    powiat.gminas 
                    |> Seq.map (fun gm -> Seq.append [gm.name] (gm.characteristics |> Seq.collect (fun (cat, chs) -> chs |> Seq.filter (fun ch -> mapHeaer |> Map.containsKey (cat, ch.CropSpecies)) |> Seq.sortBy (fun ch -> mapHeaer |> Map.find (cat, ch.CropSpecies)) |> Seq.map (fun ch -> if ch.IsDryHappened then "+" else "-"))))
                
                Seq.append gminasData [[powiat.name]]
                )
    r |> printfn "rawData = %A"
    writeData 3 1 r
    //write headers and data
    // match headers with
    // | Some h ->
    //     h |> Seq.iteri (fun k value ->
    //         ws.Cells.[startRow,startCol + k].Value <- value) //write header
    //     writeData (startRow + 1) startCol       //skip 1 row due to headers and write table body
    // | None -> writeData startRow startCol       //write body without skipping 1 row

    //let hasHeaders = if Option.isSome headers then true else false

    // let tableN
    // let tableRange =
    //     ws.Cells.[
    //     startRow,   //first row
    //     startCol,   //first col
    //     data |> Seq.length |> (+) (if hasHeaders then startRow else startRow - 1), //last row
    //     data |> Seq.head |> Seq.length |> (+) startCol |> (+) -1] //last col: assumes no sparse data on 1º row
    
    // let tb = ws.Tables.Add(tableRange,tableName)
    //tb.ShowHeader <- hasHeaders
    
    ep.Save()
    


[<EntryPoint>]
let main argv =
    let message = from "F#" // Call the function
    printfn "Hello world %s" message
    //getGminaDryLandCharacteristics ("1001062", "Rusiec") |> ignore
    let (header, powiats) = 
        load() 
        //|> Seq.take 1 
        |> List.ofSeq 
        |> transformPowiat
    
    header |> printfn "%A"
    
    writeToFile (header, powiats)
    0 // return an integer exit code