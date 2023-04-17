###############################################################################################
####################################### FUNCTIONS #############################################
###############################################################################################

function Get-ParentData($file, $infoCSV, $CSVFile){

    $content = Get-Content -Path $file.FullName | Select-String "genealogy: "
        
    if ($content -ne $null){
        $contentline = $content.ToString()

        $jsonstart = $contentline.IndexOf("{")
        $jsonend = $contentline.length
        $jsonlength = $jsonend - $jsonstart
        $injson = $contentLine.Substring($jsonstart, $jsonlength) | ConvertFrom-Json

        $infoCSV.file_name = $file.Name
        $infoCSV.item_number = $injson.item_number
        $infoCSV.serial_number = $injson.serial_number
        $infoCSV.description = $injson.description
        
        $children = 1
        $childlength = $injson.children[0].children.Count

        $infoCSV.parent_item_number = $injson.children[0].item_number

        foreach($child in $injson.children[0].children){

            $infoCSV.childrens = "$childlength.$children"
            $infoCSV.child_item_number = $child.item_number
            $infoCSV.child_serial_number = $child.serial_number

            #Write-Host $infoCSV
            $infoCSV | Export-Csv $CSVFile -Append

            $children ++
        }
    }
}

function Get-CardsData($file, $infoCSV, $CSVFile){

    $content = Get-Content -Path $file.FullName | Select-String "test_station" | %{
        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $json = ($invcontentLine.Substring($invjsonstart, $invjsonlength)).replace("'", "`"").replace("None", "null").replace("True", "true") | ConvertFrom-Json
        if($json.sales_order -ne ""){

            $infoCSV.file_name = $file.Name
            $infoCSV.sales_order = $json.sales_order
            $infoCSV.test_station = $json.test_station
            $infoCSV.server_brand = $json.server_brand
            $infoCSV.serial_number = $json.serial_number
            $infoCSV.sales_line = $json.sales_line
            $infoCSV.start_time = $json.start_time
            $infoCSV.phase = $json.phase
        }
    }

    $invcontent = Get-Content -Path $file.FullName | Select-String "inventory: " | %{

        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $invjson = $invcontentLine.Substring($invjsonstart, $invjsonlength) | ConvertFrom-Json

        if($invjson.serial_number -ne ""){

            $infoCSV.inv_serial_number = $invjson.serial_number
            $infoCSV.inv_manufacturer = $invjson.manufacturer
            $infoCSV.inv_model = $invjson.model

            $partcount = 1
            $partlength = $invjson.cards.length

            if($partlength -gt 0){

                foreach($card in $invjson.cards){

                    $infoCSV.inv_card_count = "$partlength.$partcount"
                    $infoCSV.inv_card_vendor = $card.vendor
                    $infoCSV.inv_card_device = $card.device
                    $infoCSV.inv_card_functions = $card.functions -join ","
                    $infoCSV.inv_card_slot = $card.slot
                    $infoCSV.inv_card_phy_slot = $card.phy_slot
                    $infoCSV.inv_card_serial_number = $card.serial_number
                    $infoCSV.inv_card_dev_class = $card.dev_class

                    #Write-Host $infoCSV
                    $infoCSV | Export-Csv $CSVFile -Append

                    $partcount ++
                }
            }
        }
    }

}

function Get-CPUData($file, $infoCSV, $CSVFile){

    $content = Get-Content -Path $file.FullName | Select-String "test_station" | %{
        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $json = ($invcontentLine.Substring($invjsonstart, $invjsonlength)).replace("'", "`"").replace("None", "null").replace("True", "true") | ConvertFrom-Json
        if($json.sales_order -ne ""){

            $infoCSV.file_name = $file.Name
            $infoCSV.sales_order = $json.sales_order
            $infoCSV.test_station = $json.test_station
            $infoCSV.server_brand = $json.server_brand
            $infoCSV.serial_number = $json.serial_number
            $infoCSV.sales_line = $json.sales_line
            $infoCSV.start_time = $json.start_time
            $infoCSV.phase = $json.phase
        }
    }

    $invcontent = Get-Content -Path $file.FullName | Select-String "inventory: " | %{

        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $invjson = $invcontentLine.Substring($invjsonstart, $invjsonlength) | ConvertFrom-Json

        if($invjson.serial_number -ne ""){

            $infoCSV.inv_serial_number = $invjson.serial_number
            $infoCSV.inv_manufacturer = $invjson.manufacturer
            $infoCSV.inv_model = $invjson.model

            $partcount = 1
            $partlength = $invjson.cpu.length

            if($partlength -gt 0){

                foreach($cpu in $invjson.cpu){

                    $infoCSV.inv_cpu_count = "$partlength.$partcount"
                    $infoCSV.inv_cpu_location = $cpu[0]
                    $infoCSV.inv_cpu_manufacturer = $cpu[1]
                    $infoCSV.inv_cpu_model = $cpu[2]
                    $infoCSV.inv_cpu_cores = $cpu[3]
                    $infoCSV.inv_cpu_threads = $cpu[4]
                    $infoCSV.inv_cpu_health = $cpu[5]

                    $infoCSV | Export-Csv $CSVFile.FullName -Append

                    $partcount ++       
                }
            }
        }
    }
}

function Get-DriveData($file, $infoCSV, $CSVFile){

    $content = Get-Content -Path $file.FullName | Select-String "test_station" | %{
        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $json = ($invcontentLine.Substring($invjsonstart, $invjsonlength)).replace("'", "`"").replace("None", "null").replace("True", "true") | ConvertFrom-Json
        if($json.sales_order -ne ""){

            $infoCSV.file_name = $file.Name
            $infoCSV.sales_order = $json.sales_order
            $infoCSV.test_station = $json.test_station
            $infoCSV.server_brand = $json.server_brand
            $infoCSV.serial_number = $json.serial_number
            $infoCSV.sales_line = $json.sales_line
            $infoCSV.start_time = $json.start_time
            $infoCSV.phase = $json.phase
        }
    }

    $invcontent = Get-Content -Path $file.FullName | Select-String "inventory: " | %{

        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $invjson = $invcontentLine.Substring($invjsonstart, $invjsonlength) | ConvertFrom-Json

        if($invjson.serial_number -ne ""){

            $infoCSV.inv_serial_number = $invjson.serial_number
            $infoCSV.inv_manufacturer = $invjson.manufacturer
            $infoCSV.inv_model = $invjson.model

            $partcount = 1
            $partlength = $invjson.drives.length

            if($partlength -gt 0){

                foreach($drive in $invjson.drives){

                    $infoCSV.inv_drive_count = "$partlength.$partcount"
                    $infoCSV.inv_drive_name = $drive[0]
                    $infoCSV.inv_drive_manufacturer = $drive[1]
                    $infoCSV.inv_drive_serial_number = $drive[2]
                    $infoCSV.inv_drive_model = $drive[3]
                    $infoCSV.inv_drive_revision = $drive[4]
                    $infoCSV.inv_memory_mediatype = $drive[5]
                    $infoCSV.inv_memory_capacity = $drive[6]
                    $infoCSV.inv_memory_health = $drive[7]

                    $infoCSV | Export-Csv $CSVFile -Append

                    $partcount ++       
                }
            }
        }
    }
}

function Get-LicenseData($file, $infoCSV, $CSVFile){

    $content = Get-Content -Path $file.FullName | Select-String "test_station" 
    $invcontentLine = $content.ToString()
    $invjsonstart = $invcontentLine.IndexOf("{")
    $invjsonend = $invcontentLine.length
    $invjsonlength = $invjsonend - $invjsonstart
    $json = ($invcontentLine.Substring($invjsonstart, $invjsonlength)).replace("'", "`"").replace("None", "null").replace("True", "true") | ConvertFrom-Json
    if($json.sales_order -ne ""){

        $infoCSV.file_name = $file.Name
        $infoCSV.sales_order = $json.sales_order
        $infoCSV.test_station = $json.test_station
        $infoCSV.server_brand = $json.server_brand
        $infoCSV.serial_number = $json.serial_number
        $infoCSV.sales_line = $json.sales_line
        $infoCSV.start_time = $json.start_time
        $infoCSV.phase = $json.phase
    }
    

    $invcontent = Get-Content -Path $file.FullName | Select-String "Licenses: "

    $invcontentLine = $invcontent[0].ToString()
    $invjsonstart = $invcontentLine.IndexOf("{")
    $invjsonend = $invcontentLine.length
    $invjsonlength = $invjsonend - $invjsonstart
    $invjson = $invcontentLine.Substring($invjsonstart, $invjsonlength) | ConvertFrom-Json

    if($invjson.PSObject.Properties.name.Count -gt 0){

        $partcount = 1
        $partlength = $invjson.PSObject.Properties.name.Count

        foreach($license in $invjson.PSObject.Properties){

            $infoCSV.inv_licenses = "$partlength.$partcount"
            $infoCSV.inv_license_name = $license.name
            $infoCSV.inv_license_count = $license.value

            #Write-Host $infoCSV
            $infoCSV | Export-Csv $CSVFile -Append

            $partcount ++           
        }
    }
}

function Get-MemoryData($file, $infoCSV, $CSVFile){

$content = Get-Content -Path $file.FullName | Select-String "test_station" | %{
        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $json = ($invcontentLine.Substring($invjsonstart, $invjsonlength)).replace("'", "`"").replace("None", "null").replace("True", "true") | ConvertFrom-Json
        if($json.sales_order -ne ""){

            $infoCSV.file_name = $file.Name
            $infoCSV.sales_order = $json.sales_order
            $infoCSV.test_station = $json.test_station
            $infoCSV.server_brand = $json.server_brand
            $infoCSV.serial_number = $json.serial_number
            $infoCSV.sales_line = $json.sales_line
            $infoCSV.start_time = $json.start_time
            $infoCSV.phase = $json.phase
        }
    }

    $invcontent = Get-Content -Path $file.FullName | Select-String "inventory: " | %{

        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $invjson = $invcontentLine.Substring($invjsonstart, $invjsonlength) | ConvertFrom-Json

        if($invjson.serial_number -ne ""){

            $infoCSV.inv_serial_number = $invjson.serial_number
            $infoCSV.inv_manufacturer = $invjson.manufacturer
            $infoCSV.inv_model = $invjson.model

            $partcount = 1
            $partlength = $invjson.memory.length

            if($partlength -gt 0){

                foreach($memory in $invjson.memory){

                    $infoCSV.inv_memory_count = "$partlength.$partcount"
                    $infoCSV.inv_memory_location = $memory[0]
                    $infoCSV.inv_memory_state = $memory[1]
                    $infoCSV.inv_memory_size = $memory[2]
                    $infoCSV.inv_memory_type = $memory[3]
                    $infoCSV.inv_memory_speed = $memory[4]
                    $infoCSV.inv_memory_manufacturer = $memory[5]
                    $infoCSV.inv_memory_serial_number = $memory[6]
                    $infoCSV.inv_memory_part_number = $memory[7]
                    $infoCSV.inv_memory_health = $memory[8]


                    $infoCSV | Export-Csv $CSVFile -Append

                    $partcount ++
                    
                }
            }
        }
    }
}

function Get-PSUData($file, $infoCSV, $CSVFile){

    $content = Get-Content -Path $file.FullName | Select-String "test_station" | %{
        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $json = ($invcontentLine.Substring($invjsonstart, $invjsonlength)).replace("'", "`"").replace("None", "null").replace("True", "true") | ConvertFrom-Json
        if($json.sales_order -ne ""){

            $infoCSV.file_name = $file.Name
            $infoCSV.sales_order = $json.sales_order
            $infoCSV.test_station = $json.test_station
            $infoCSV.server_brand = $json.server_brand
            $infoCSV.serial_number = $json.serial_number
            $infoCSV.sales_line = $json.sales_line
            $infoCSV.start_time = $json.start_time
            $infoCSV.phase = $json.phase
        }
    }

    $invcontent = Get-Content -Path $file.FullName | Select-String "inventory: " | %{

        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $invjson = $invcontentLine.Substring($invjsonstart, $invjsonlength) | ConvertFrom-Json

        if($invjson.serial_number -ne ""){

            $infoCSV.inv_serial_number = $invjson.serial_number
            $infoCSV.inv_manufacturer = $invjson.manufacturer
            $infoCSV.inv_model = $invjson.model

            $partcount = 1
            $partlength = $invjson.psu.length

            if($partlength -gt 0){

                foreach($psu in $invjson.psu){

                    $infoCSV.inv_psu_count = "$partlength.$partcount"
                    $infoCSV.inv_psu_location = $psu[0]
                    $infoCSV.inv_psu_state = $psu[1]
                    $infoCSV.inv_psu_model = $psu[2]
                    $infoCSV.inv_psu_serial_number = $psu[3]
                    $infoCSV.inv_psu_firmware = $psu[4]
                    $infoCSV.inv_psu_power = $psu[5]
                    $infoCSV.inv_psu_health = $psu[6]


                    $infoCSV | Export-Csv $CSVFile -Append

                    $partcount ++
                    
                }
            }
        }
    }
}

function Get-StorageData($file, $infoCSV, $CSVFile){

$content = Get-Content -Path $file.FullName | Select-String "test_station" | %{
        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $json = ($invcontentLine.Substring($invjsonstart, $invjsonlength)).replace("'", "`"").replace("None", "null").replace("True", "true") | ConvertFrom-Json
        if($json.sales_order -ne ""){

            $infoCSV.file_name = $file.Name
            $infoCSV.sales_order = $json.sales_order
            $infoCSV.test_station = $json.test_station
            $infoCSV.server_brand = $json.server_brand
            $infoCSV.serial_number = $json.serial_number
            $infoCSV.sales_line = $json.sales_line
            $infoCSV.start_time = $json.start_time
            $infoCSV.phase = $json.phase
        }
    }

    $invcontent = Get-Content -Path $file.FullName | Select-String "inventory: " | %{

        $invcontentLine = $_.ToString()
        $invjsonstart = $invcontentLine.IndexOf("{")
        $invjsonend = $invcontentLine.length
        $invjsonlength = $invjsonend - $invjsonstart
        $invjson = $invcontentLine.Substring($invjsonstart, $invjsonlength) | ConvertFrom-Json

        if($invjson.serial_number -ne ""){

            $infoCSV.inv_serial_number = $invjson.serial_number
            $infoCSV.inv_manufacturer = $invjson.manufacturer
            $infoCSV.inv_model = $invjson.model

            if($invjson.storage[0].Length -gt 0) { 

                $partcount = 1
                if($invjson.storage[0][5].length -gt 0){

                    $partlength = $invjson.storage[0][5][0][2].Length

                    if($partlength -gt 0){

                        for($i = 0; $i -lt $partlength; $i++){

                            $infoCSV.inv_storage_disk_count = "$partlength.$partcount"

                            $infoCSV.inv_storage_controller_id = $invjson.storage[0][0]
                            $infoCSV.inv_storage_manufacturer = $invjson.storage[0][1]
                            $infoCSV.inv_storage_model = $invjson.storage[0][2]
                            $infoCSV.inv_storage_serialnumber = $invjson.storage[0][3]
                            $infoCSV.inv_storage_firmwareversion = $invjson.storage[0][4]
                            $infoCSV.inv_storage_health = $invjson.storage[0][6]

                            $infoCSV.inv_storage_volume_storagetype = $invjson.storage[0][5][0][0]
                            $infoCSV.inv_storage_volume_volumetype = $invjson.storage[0][5][0][1]
                            $infoCSV.inv_storage_volume_size = $invjson.storage[0][5][0][3]


                            $infoCSV.inv_storage_disk_name = $invjson.storage[0][5][0][2][$i][0]
                            $infoCSV.inv_storage_disk_manufacturer = $invjson.storage[0][5][0][2][$i][1]
                            $infoCSV.inv_storage_disk_serialnumber = $invjson.storage[0][5][0][2][$i][2]
                            $infoCSV.inv_storage_disk_model = $invjson.storage[0][5][0][2][$i][3]
                            $infoCSV.inv_storage_disk_revision = $invjson.storage[0][5][0][2][$i][4]
                            $infoCSV.inv_storage_disk_mediatype = $invjson.storage[0][5][0][2][$i][5]
                            $infoCSV.inv_storage_disk_capacity = $invjson.storage[0][5][0][2][$i][6]
                            $infoCSV.inv_storage_disk_health = $invjson.storage[0][5][0][2][$i][7]


                            #Write-Host $infoCSV 
                            $infoCSV | Export-Csv $CSVFile -Append

                            $partcount ++
                    
                        }
                    }
                }
            }
        }
    }
}

function Extract-JSON(){
    param($text)

    $text = $text.toString()
    $jsonstart = $text.IndexOf("{")
    $jsonend = $text.IndexOf("}")
    $jsonlength = $jsonend - $jsonstart + 1
    $json = $text.Substring($jsonstart, $jsonlength) | ConvertFrom-Json
    $json
}

function Merge-CSVs($targetFolder){

    Import-Module ImportExcel
    $csvs = Get-ChildItem -Path $TargetFolder -Filter *.csv
    $csvCount = $csvs.Count
    Write-Host "Detected the following CSV files: ($csvCount)"
    foreach ($csv in $csvs) {
        Write-Host " -"$csv.Name
    }

    $excelFileName = $(get-date -f yyyyMMdd_HHmmss) +"_Combined-Data.xlsx"
    Write-Host "Creating: $excelFileName"

    foreach ($csv in $csvs) {
        $csvPath = "$TargetFolder\" + $csv.Name
        $worksheetName = $csv.Name.Replace(".csv","")
        Write-Host " - Adding $worksheetName to $excelFileName"
        Import-Csv -Path $csvPath | Export-Excel -Path "$targetFolder\$excelFileName" -WorkSheetname $worksheetName
    }

}

###################################################################################################################
###################################################################################################################
###################################################################################################################
###################################################################################################################
###################################################################################################################


$Path = "C:\Users\spatoju\Downloads\AWS_Lucy\Logs\Include\v2"
$logPath = "C:\Users\spatoju\Downloads\AWS_Lucy\Logs\Include"
Set-Location $Path
$DateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$NewPath = New-Item -ItemType Directory -Path "$Path\$DateTime"
$files = Get-ChildItem $logPath *log
$JsonConfig = Get-Content "$logpath\ConfigJson.txt" | ConvertFrom-Json


$Keys = ($JsonConfig | Get-Member -MemberType NoteProperty).Name

## Parent ####################################################################################
$ParentCSVFile = New-Item -Path "$NewPath\Parent.csv" -ItemType File
$JsonConfig.parent | Export-Csv -Path $ParentCSVFile.FullName -NoTypeInformation
$ParentinfoCSV = Import-Csv $ParentCSVFile.FullName

## Cards #####################################################################################
$CardsCSVFile = New-Item -Path "$NewPath\Cards.csv" -ItemType File
$JsonConfig.cards | Export-Csv -Path $CardsCSVFile.FullName -NoTypeInformation
$CardsinfoCSV = Import-Csv $CardsCSVFile.FullName

## CPU #######################################################################################
$CPUCSVFile = New-Item -Path "$NewPath\CPU.csv" -ItemType File
$JsonConfig.cpu | Export-Csv -Path $CPUCSVFile.FullName -NoTypeInformation
$CPUinfoCSV = Import-Csv $CPUCSVFile.FullName

## Drive #####################################################################################
$DriveCSVFile = New-Item -Path "$NewPath\Drive.csv" -ItemType File
$JsonConfig.drive | Export-Csv -Path $DriveCSVFile.FullName -NoTypeInformation
$DriveinfoCSV = Import-Csv $DriveCSVFile.FullName

## License ###################################################################################
$LicenseCSVFile = New-Item -Path "$NewPath\License.csv" -ItemType File
$JsonConfig.license | Export-Csv -Path $LicenseCSVFile.FullName -NoTypeInformation
$LicenseinfoCSV = Import-Csv $LicenseCSVFile.FullName

## Memory ####################################################################################
$MemoryCSVFile = New-Item -Path "$NewPath\Memory.csv" -ItemType File
$JsonConfig.memory | Export-Csv -Path $MemoryCSVFile.FullName -NoTypeInformation
$MemoryinfoCSV = Import-Csv $MemoryCSVFile.FullName

## PSU #######################################################################################
$PSUCSVFile = New-Item -Path "$NewPath\PSU.csv" -ItemType File
$JsonConfig.psu | Export-Csv -Path $PSUCSVFile.FullName -NoTypeInformation
$PSUinfoCSV = Import-Csv $PSUCSVFile.FullName

## Storage ###################################################################################
$StorageCSVFile = New-Item -Path "$NewPath\Storage.csv" -ItemType File
$JsonConfig.storage | Export-Csv -Path $StorageCSVFile.FullName -NoTypeInformation
$StorageinfoCSV = Import-Csv $StorageCSVFile.FullName
$Count = 1


foreach($file in $files){

    Write-Host $Count -- $File
    
     Get-ParentData -file $file -infoCSV $ParentinfoCSV -CSVFile $ParentCSVFile
     Get-CardsData -file $file -infoCSV $CardsinfoCSV -CSVFile $CardsCSVFile
     Get-CPUData -file $file -infoCSV $CPUinfoCSV -CSVFile $CPUCSVFile
     Get-DriveData -file $file -infoCSV $DriveinfoCSV -CSVFile $DriveCSVFile
     Get-LicenseData -file $file -infoCSV $LicenseinfoCSV -CSVFile $LicenseCSVFile
     Get-MemoryData -file $file -infoCSV $MemoryinfoCSV -CSVFile $MemoryCSVFile
     Get-PSUData -file $file -infoCSV $PSUinfoCSV -CSVFile $PSUCSVFile
     Get-StorageData -file $file -infoCSV $StorageinfoCSV -CSVFile $StorageCSVFile

     $Count ++
}

Merge-CSVs -targetFolder $NewPath