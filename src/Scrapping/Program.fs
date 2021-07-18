// Learn more about F# at http://docs.microsoft.com/dotnet/fsharp

open System
open FSharp.Data
open System.IO
open OfficeOpenXml
open System.Drawing
open OfficeOpenXml.Style

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

//https://susza.iung.pulawy.pl/wykazy/2021,1001011/
type DryLandHtml = HtmlProvider<"https://susza.iung.pulawy.pl/wykazy/2021,1001062/", ResolutionFolder=__SOURCE_DIRECTORY__>


let getGminasInPowiat powId =
    async {
        let! html = DryLandHtml.AsyncLoad(sprintf "https://susza.iung.pulawy.pl/wykazy/2021,%s/" powId)
        return 
            html.Lists.Html.CssSelect "select#sel-gmina > option" 
            |> Seq.map (fun cs -> (cs.Attribute("value").Value(), cs.InnerText()))
            |> Seq.filter (fun (id, _) -> id <> "-")
    }

let getGminaHtml (gminaId, gminaName) = 
    async {
        let! html = DryLandHtml.AsyncLoad(sprintf "https://susza.iung.pulawy.pl/wykazy/2021,%s/" gminaId)
        return (gminaId, gminaName, html)
    }

let getGminaDryLandCharacteristics column dryLandHappened ((gminaId, gminaName, html: DryLandHtml)) = 
    printfn "Converting gmina: %s, %s" gminaId gminaName
    
    let convertGminaCharacteristics (categoryName, tBody : HtmlNode) set =
        let convertRow  (tr: HtmlNode) set =
            let tds = tr.CssSelect("td")
            let isDry = tds |> List.skip column |> List.head |> (fun h -> h.InnerText().Equals("+"))
            let char = { CropSpecies = tds.Head.InnerText(); IsDryHappened = isDry }
            
            if isDry then 
                (char, set |> Set.add (categoryName, char.CropSpecies))
            else 
                (char, set) 
        
        let (characteristics, populatedDryLandSet) = Seq.mapFoldBack convertRow (tBody.CssSelect("tr")) set
        ((categoryName, characteristics |> List.ofSeq), populatedDryLandSet)

    let zipped = 
        html.Html.CssSelect("table.tab-gmina tbody")
        |> Seq.zip [ "Kategoria gleby I"; "Kategoria gleby II"; "Kategoria gleby III"; "Kategoria gleby IV" ]

    let (characteristics, populatedDryLandSet) = Seq.mapFoldBack convertGminaCharacteristics zipped dryLandHappened
       
    (populatedDryLandSet, 
        {
            id = gminaId;
            name = gminaName;
            characteristics = characteristics |> List.ofSeq
        })

let loadPowiats column =
    let powiatsList = 
        async {
            let! html = DryLandHtml.AsyncLoad("https://susza.iung.pulawy.pl/wykazy/2021,1014/")
            
            return! html.Lists.Html.CssSelect "select#sel-pow > option" 
            |> Seq.map (fun cs -> (cs.Attribute("value").Value(), cs.InnerText()))
            |> Seq.filter (fun (id, _) -> id <> "-")
            |> Seq.map (fun (powId, powName) -> 
                async { 
                    let! gminasInPowiat = getGminasInPowiat powId
                    let! gminasHtml = gminasInPowiat |> Seq.map getGminaHtml |> Async.Parallel
                    return (powId, powName, gminasHtml)
                })
            |> Async.Parallel
        }
        |> Async.RunSynchronously
       
    let convertPowiat (powiatId, powiatName, gminas) dryHappened =
        let folder gm drySoFar = 
            let (newSet, gmChars) = getGminaDryLandCharacteristics column drySoFar gm
            (gmChars, newSet) 
            
        let (convertedGminas, fulFilledDryHappened) = 
            Seq.mapFoldBack folder gminas dryHappened
    
        ({ id = powiatId; name = powiatName; gminas = convertedGminas |> List.ofSeq; }, fulFilledDryHappened)
    
    let (convertedPowiats, populatedDryHappened)  = Seq.mapFoldBack convertPowiat powiatsList Set.empty
    (populatedDryHappened, convertedPowiats |> List.ofSeq)

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
                    let cell = ws.Cells.[iR + i,iC + k]
                    cell.Value <- value
                    if value.Equals("+") then
                        cell.Style.Fill.PatternType <- ExcelFillStyle.Solid
                        cell.Style.Fill.BackgroundColor.SetColor(Color.Red)
                        ()
                    else if row |> Seq.length = 1 then 
                        cell.Style.Font.Bold <- true;
                        cell.Style.Fill.PatternType <- ExcelFillStyle.Solid
                        cell.Style.Fill.BackgroundColor.SetColor(Color.LightSeaGreen)
                        
                    else 
                        ()))
    let mapHeaer = header |> Seq.mapi (fun i h -> (h, i)) |> Map.ofSeq

    header |> Seq.iteri (fun iC headVals -> 
        ws.Cells.[1, iC + startCol].Value <- (headVals |> snd) 
        ws.Cells.[2, iC + startCol].Value <- (headVals |> fst)) 
   
    let rawData = 
        powiats 
        |> Seq.collect (fun powiat -> 
                let gminasData =
                    powiat.gminas 
                    |> Seq.map (fun gm -> 
                        Seq.append [gm.name] (gm.characteristics 
                            |> Seq.collect (fun (cat, chs) -> 
                                chs 
                                |> Seq.filter (fun ch -> mapHeaer |> Map.containsKey (cat, ch.CropSpecies)) 
                                |> Seq.sortBy (fun ch -> mapHeaer |> Map.find (cat, ch.CropSpecies)) 
                                |> Seq.map (fun ch -> if ch.IsDryHappened then "+" else "-"))))
                
                Seq.append gminasData [[powiat.name]]
        )
    //rawData |> printfn "rawData = %A"
    writeData 3 1 rawData
    ep.Save()
    
let generateReport column = 
    let (header, powiats) = loadPowiats column
    header |> List.ofSeq |> printfn "%A"
    writeToFile (header, powiats)

let parseColumn (argv: string array) =
    if argv.Length > 0 then
        let columnstr = argv.[0]
        printfn "Columnstr: %s" columnstr
        Int32.Parse(columnstr)
    else
        raise (ArgumentException("Please provide 'column' argument to the process."))

[<EntryPoint>]
let main argv =
    printfn "Hello world" 
    printfn "%A" argv
    let column = parseColumn argv
    //getGminaDryLandCharacteristics 6 Set.empty ("1001062", "Rusiec") |> printfn "%A" 

    generateReport column 

    0 // return an integer exit code