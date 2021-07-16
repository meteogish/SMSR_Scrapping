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



//ttps://susza.iung.pulawy.pl/wykazy/2021,1001011/
type DryLandHtml = HtmlProvider<"https://susza.iung.pulawy.pl/wykazy/2021,1001062/", ResolutionFolder=__SOURCE_DIRECTORY__>

// Define a function to construct a message to print
let from whom =
    sprintf "from %s" whom

let getGminasInPowiat powId =
    DryLandHtml.Load(sprintf "https://susza.iung.pulawy.pl/wykazy/2021,%s/" powId).Lists.Html.CssSelect "select#sel-gmina > option" 
        |> Seq.map (fun cs -> (cs.Attribute("value").Value(), cs.InnerText()))
        |> Seq.filter (fun (id, _) -> id <> "-")
        
let getGminaDryLandCharacteristics column dryLandHappened ((gminaId, gminaName)) = 
    printfn "Getting gmina: %s, %s" gminaId gminaName
    let html = DryLandHtml.Load(sprintf "https://susza.iung.pulawy.pl/wykazy/2021,%s/" gminaId)
    
    let convertGminaCharacteristics (set, chSoFar) (categoryName, tBody : HtmlNode) =
        let convertRow (set, soFar) (tr: HtmlNode) =
            let tds = tr.CssSelect("td")
            let isDry = tds |> List.skip column |> List.head |> (fun h -> h.InnerText().Equals("+"))
            let char = { CropSpecies = tds.Head.InnerText(); IsDryHappened = isDry }
            
            if isDry then 
                (set |> Set.add (categoryName, char.CropSpecies), char :: soFar)
            else 
                (set, char :: soFar) 
        
        let (newSet, newChars) = tBody.CssSelect("tr") |> Seq.fold convertRow (set, [])
        (newSet, (categoryName, newChars) :: chSoFar)

    let zipped = 
        html.Html.CssSelect("table.tab-gmina tbody")
        |> Seq.zip [ "Kategoria gleby I"; "Kategoria gleby II"; "Kategoria gleby III"; "Kategoria gleby IV" ]

    let (newDryLandSet, chars) =
        zipped |> Seq.fold convertGminaCharacteristics (dryLandHappened, [])
       
    (newDryLandSet, 
        {
            id = gminaId;
            name = gminaName;
            characteristics = chars |> List.rev
        })

let loadPowiats column =
    let powiatsList = 
        DryLandHtml.Load("https://susza.iung.pulawy.pl/wykazy/2021,1014/").Lists.Html.CssSelect "select#sel-pow > option" 
        |> Seq.map (fun cs -> (cs.Attribute("value").Value(), cs.InnerText()))
        |> Seq.filter (fun (id, _) -> id <> "-")
       
    let convertPowiat (dryHappened, powiatsSoFar) (powiatId, powiatName) =
        let (fulFilledDryHappened, convertedGminas) = 
            getGminasInPowiat powiatId 
            |> Seq.fold (fun (drySoFar, convSoFar) gm ->  
                let (newSet, gmChars) = getGminaDryLandCharacteristics column drySoFar gm
                (newSet, gmChars :: convSoFar)) (dryHappened, [])
    
        (fulFilledDryHappened,
            { 
                id = powiatId; 
                name = powiatName; 
                gminas = convertedGminas |> List.rev;// |> List.sortBy (fun gm -> gm.name);
            } :: powiatsSoFar)
    
    let (set, convertedPowiats)  = powiatsList |> Seq.fold convertPowiat (Set.empty, [])
    (set, convertedPowiats |> List.rev)

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
    
let proces () = 
    let (header, powiats) = loadPowiats 6
    header |> List.ofSeq |> printfn "%A"
    writeToFile (header, powiats)

[<EntryPoint>]
let main argv =
    let message = from "F#" // Call the function
    printfn "Hello world %s" message
    //getGminaDryLandCharacteristics 6 Set.empty ("1001062", "Rusiec") |> printfn "%A" 

    proces ()

    0 // return an integer exit code