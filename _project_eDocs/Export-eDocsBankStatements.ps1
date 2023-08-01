$importDirectory = "\\testcompany.com\Accounting\2020\"
$exportDirectory = "\\testcompany.com\eDocs"

# define hash table and bind all folders to IBANs
$foldersAndIBANs = @{
    "DSK-417" = "BG75STSA93001527083417"
    "DSK-461" = "BG40STSA93001527109461"
    "DSK-479" = "BG39STSA93001527109479"
    "DSK-505" = "BG16STSA93001527109505"
    "DSK-592" = "BG56STSA93001528411592"
    "РББ - 1012146462" = "BG60RZBB91551012146462"
    "РББ - 1012160331" = "BG60RZBB91551012160331"
    "РББ - 1012334449" = "BG33RZBB91551012334449"
    "РББ -1012146420" = "BG30RZBB91551012146420"
    "РББ -1012146668" = "BG27RZBB91551012146668"
}

# get all folders from import directory
$folders = Get-ChildItem $importDirectory | where-object { $_.PSIsContainer }

# loop throught all folders from import directory
foreach ($folder in $folders)
{
    $folderName = $folder.Name
    $folderIBAN = $foldersAndIBANs["$folderName"]

    if ($folderIBAN -ne $null)
    {
        $files = Get-ChildItem -Path $folder.FullName -Recurse
        foreach ($file in $files)
        {
            $fileName = "BankStatement" + "-" + $folderIBAN + "-" + $file.Name
            $fileDestination = $exportDirectory + "\" + $fileName
            Copy-Item -Path $file.FullName -Destination $fileDestination
        }
    }
    else 
    {
        Write-Warning "There is no IBAN configration for $folderName folder!"    
    }
}