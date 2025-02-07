# Vereist Excel op het systeem
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Workbook = $Excel.Workbooks.Add()

# Definieer de bestandsnaam en bijbehorende tabbladnamen
$FileMapping = @{
    'Users'                        = 'Mastersheet'
    'Office365ActivationsUserDetail' = 'Activations'
    'Office365ActiveUserDetail'      = 'Products'
    'ProPlusUsageUserDetail'         = 'Activity'
    'ProductList'                    = 'Product List'
}

# Verkrijg de huidige map waarin het script draait
$CurrentDirectory = Get-Location

# Zoek alle CSV-bestanden in de map
$CsvFiles = Get-ChildItem -Path $CurrentDirectory -Filter "*.csv"

foreach ($File in $CsvFiles) {
    foreach ($Key in $FileMapping.Keys) {
        if ($File.Name -match "^$Key.*\.csv$") {
            $SheetName = $FileMapping[$Key]
            Write-Host "Verwerken: $($File.Name) -> Tabblad: $SheetName"
            
            # Maak een nieuw werkblad aan
            $Sheet = $Workbook.Sheets.Add()
            $Sheet.Name = $SheetName
            
            # Importeer CSV-inhoud
            $CsvContent = Import-Csv -Path $File.FullName
            
            # Schrijf de kolomnamen
            $ColIndex = 1
            foreach ($Column in $CsvContent[0].PSObject.Properties.Name) {
                $Sheet.Cells(1, $ColIndex) = $Column
                $ColIndex++
            }
            
            # Schrijf de data naar het werkblad
            $RowIndex = 2
            foreach ($Row in $CsvContent) {
                $ColIndex = 1
                foreach ($Value in $Row.PSObject.Properties.Value) {
                    $Sheet.Cells($RowIndex, $ColIndex) = $Value
                    $ColIndex++
                }
                $RowIndex++
            }
        }
    }
}

# Opslaan en afsluiten
$ExcelFilePath = "$CurrentDirectory\Samengevoegd.xlsx"
$Workbook.SaveAs($ExcelFilePath)
$Excel.Quit()

Write-Host "Excel-bestand opgeslagen als: $ExcelFilePath"
