# enums
enum JoinTypes 
{
    InnerJoin = 1
    LeftJoin = 2
    OuterJoin = 3
}

# functions
function Introduce-Script
{
    Write-Host -ForegroundColor DarkCyan ("`nJoin 2 or more CSV files based on related columns between them.`n" +
    "(Specified columns don't have to have the same name, but should have some matching records.)`n" +
    "Same concept as a SQL table join. See: https://www.w3schools.com/sql/sql_join.asp`n" +
    "Note that the names of each CSV will be added to the column headers.")
}

function Prompt-AllFiles
{
    $csvData = New-Object -TypeName System.Collections.Generic.List[CSVDataUnit]
    $i = 0
    $addMore = $true
    while ($addMore -eq $true)
    {
        $i++
        $file = Prompt-CSVFile
        $importedCSV = Import-CSV -Path $file 
        $joinColumn = Prompt-JoinColumn -importedCSV $importedCSV -fileIndex $i
        $csvDataUnit = New-Object -TypeName CSVDataUnit -ArgumentList $importedCSV, $file.BaseName, $joinColumn
        $csvData.Add($csvDataUnit)
        if ($i -lt 2) { continue }
        $addMore = Prompt-YesOrNo
    }
    return $csvData
}

function Prompt-CSVFile
{
    do
    {
        $filePath = Prompt-FilePath -fileIndex $i
        $file = Get-Item -Path $filePath -ErrorAction SilentlyContinue
        if ($null -eq $file)
        {
            Write-Warning "File not found. Please try again."
        }
    }
    while ($null -eq $file)

    return $file
}

function Prompt-FilePath($fileIndex)
{
    return Read-Host "Enter file path to CSV$fileIndex. (i.e C:\Users\Username\Desktop\myFile.csv)"
}

function Prompt-JoinColumn($importedCSV, $fileIndex)
{
    do
    {
        $columnName = Read-Host "Enter column name to join by from CSV$fileIndex."
        $columnExists = Test-ColumnExists -importedCSV $importedCSV -columnName $columnName
        if (!($columnExists))
        {
            Write-Warning "Column not found. Please try again."
        }
    }
    while (!($columnExists))

    return $columnName
}

function Test-ColumnExists($importedCSV, $columnName)
{
    $record1 = $importedCSV | Select-Object -First 1
    if ([bool]$record1.PSObject.Properties[$columnName]) # checks if object has specified property
    {
        return $true
    }
    return $false
}

function Prompt-YesOrNo
{
    do
    {
        $ans = Read-Host "Would you like to add more? y/n"
    } 
    while ($ans -inotmatch "\b[yn]\b")

    if ($ans -imatch "y") { return $true }
    return $false
}

function Prompt-JoinType
{
    Write-Host "Specify join type:"
    do
    {        
        Write-Host ("Enter 1 for Inner Join.`n" +
            "Enter 2 for Left Join.`n" +
            "Enter 3 for Outer Join.")
        try
        {
            [int]$userResponse = Read-Host
        }
        catch
        {
            continue
        }      
    }
    while ($userResponse -notmatch "^[1-3]$")

    return [JoinTypes].GetEnumName($userResponse)
}

function Test-FileExists($fileName, $csvData)
{
    foreach ($dataUnit in $csvData)
    {
        if ($dataUnit.Name -eq $fileName) { return $true }
    }
    return $false
}

function Join-CSVDataInner($csvData)
{
    Write-Host "Joining data..." -ForegroundColor DarkCyan
    $joinedData = New-Object -TypeName System.Collections.Generic.List[PSCustomObject]
    $smallestIndex = Get-SmallestCSVDataUnit $csvData
    $dataUnitsToHash = New-Object -TypeName System.Collections.Generic.List[CSVDataUnit]    

    for ($i = 0; $i -lt $csvData.Count; $i++)
    {
        if ($i -eq $smallestIndex) { continue }
        $dataUnitsToHash.Add($csvData[$i])
    }

    $hashTableCollection = New-HashTableCollection $dataUnitsToHash
    $baseCSV = $csvData[$smallestIndex].importedCSV
    $baseJoinColumn = $csvData[$smallestIndex].joinColumnName
    $baseName = $csvData[$smallestIndex].name

    :baseRow foreach ($baseRow in $baseCSV)
    {
        $stampedBaseRow = Stamp-Columns -rowObject $baseRow -csvName $baseName
        $joinedRow = $stampedBaseRow
        :hashTable foreach ($hashTable in $hashTableCollection)
        {
            if ($hashTable.Contains($baseRow.$baseJoinColumn))
            {
                $stampedOtherRow = Stamp-Columns -rowObject $hashTable[$baseRow.$baseJoinColumn] -csvName $hashTable["HashTableName"]
                $joinedRow = Join-Objects -object1 $joinedRow -object2 $stampedOtherRow
            }
            else
            {
                continue baseRow
            }
        }
        $joinedData.Add($joinedRow)
        Write-Progress -Activity "Joining data..." -Status "$($joinedData.Count) records joined."
    }
    return $joinedData
}

function Get-SmallestCSVDataUnit($csvData)
{
    $smallestIndex = 0
    for ($i = 1; $i -lt $csvData.Count; $i++)
    {
        if ($csvData[$i].importedCSV.Count -lt $csvData[$smallestIndex].importedCSV.Count)
        {
            $smallestIndex = $i
        }
    }
    return $smallestIndex
}

function New-HashTableCollection($csvData)
{
    $htCollection = New-Object -TypeName System.Collections.Generic.List[System.Collections.Hashtable]
    
    foreach ($csvDataUnit in $csvData)
    {
        $hashtable = Convert-CSVToHashTable $csvDataUnit
        $htCollection.Add($hashtable)
    }

    return $htCollection
}

function Convert-CSVToHashTable($csvDataUnit)
{
    $hashTable = @{}
    $importedCSV = $csvDataUnit.importedCSV
    $joinColumnName = $csvDataUnit.joinColumnName
    $csvName = $csvDataUnit.name
    $nullWarningSent = $false

    foreach ($row in $importedCSV)
    {
        $rowValueInJoinColumn = $row.$joinColumnName
        
        if ($null -eq $rowValueInJoinColumn)
        { 
            if ($nullWarningSent -eq $false)
            {
                Write-Warning "Found record(s) with null join column value in csv: $csvName."
                $nullWarningSent = $true
            }          
            continue
        }

        $trimmedValue = $rowValueinJoinColumn.Trim()

        if ($($hashTable.Contains($trimmedValue)))
        {
            Write-Warning ("Duplicate record found in csv: $csvName.`n" +            
                "Join column value is: $trimmedValue.`n" +
                "Record will be ignored.")
            continue
        }
        $hashTable.Add($trimmedValue, $row)
    }

    $hashTable.Add("HashTableName", $csvName)

    return $hashTable
}

function Stamp-Columns($rowObject, $csvName)
{
    $stampedRow = New-Object -TypeName PSCustomObject 
    foreach ($property in $rowObject.PSObject.Properties)
    {
        Add-Member -InputObject $stampedRow -MemberType NoteProperty -Name "$csvName - $($property.Name)" -Value $property.Value
    }
    return $stampedRow
}

function Join-Objects($object1, $object2)
{
    $joinedHashTable = [ordered]@{}

    foreach ($property in $object1.PSObject.Properties)
    {
        $joinedHashTable += @{$property.Name = $property.Value }
    }

    foreach ($property in $object2.PSObject.Properties)
    {
        $joinedHashTable += @{$property.Name = $property.Value }
    }

    return [PSCustomObject]$joinedHashTable
}

function Prompt-BaseFile($csvData)
{
    do
    {
        $baseFileName = Read-Host "Specify base file name. (Without file extension.)`nAll records in this file will be kept and compared against."
        $fileExists = Test-FileExists -fileName $baseFileName -csvData $csvData
        if ($fileExists -eq $false)
        {
            Write-Warning "File not found in imported CSV list. Please try again."
        }
    }
    while ($fileExists -eq $false)

    return $baseFileName
}

function Join-CSVDataLeft($csvData, $baseFileName)
{
    Write-Host "Joining data..." -ForegroundColor DarkCyan
    $joinedData = New-Object -TypeName System.Collections.Generic.List[PSCustomObject]
    $baseUnit = Get-BaseDataUnit -csvData $csvData -baseFileName $baseFileName    
    $dataUnitsToHash = New-Object -TypeName System.Collections.Generic.List[CSVDataUnit]

    for ($i = 0; $i -lt $csvData.Count; $i++)
    {
        if ($csvData[$i].name -eq $baseFileName) { continue }
        $dataUnitsToHash.Add($csvData[$i])
    }

    $hashTableCollection = New-HashTableCollection -csvData $dataUnitsToHash
    $baseCSV = $baseUnit.importedCSV
    $baseName = $baseUnit.name
    $baseJoinColumn = $baseUnit.joinColumnName

    foreach ($baseRow in $baseCSV)
    {
        Write-Progress -Activity "Joining data..." -Status "$($joinedData.Count) records joined."
        $stampedBaseRow = Stamp-Columns -rowObject $baseRow -csvName $baseName
        $moddedBaseRow = Add-Column -row $stampedBaseRow -columnName "In $($baseName)" -value $true
        $joinedRow = $moddedBaseRow
        foreach ($hashTable in $hashTableCollection)
        {
            if ($hashTable.Contains($baseRow.$baseJoinColumn))
            {
                $stampedOtherRow = Stamp-Columns -rowObject $hashTable[$baseRow.$baseJoinColumn] -csvName $hashTable["HashTableName"]
                $moddedOtherRow = Add-Column -row $stampedOtherRow -columnName "In $($hashTable["HashTableName"])" -value $true
                $joinedRow = Join-Objects -object1 $joinedRow -object2 $moddedOtherRow
            }
            else
            {
                $nullRow = Get-NullRowFromHashTable $hashTable
                $moddedNullRow = Add-Column -row $nullRow -columnName "In $($hashTable["HashTableName"])" -value $false
                $joinedRow = Join-Objects -object1 $joinedRow -object2 $moddedNullRow
            }
        }
        $joinedData.Add($joinedRow)
    }
    return $joinedData
}

function Get-BaseDataUnit($csvData, $baseFileName)
{
    foreach ($dataUnit in $csvData)
    {
        if ($dataUnit.Name -eq $baseFileName)
        {
            $baseUnit = $dataUnit
        }
    }
    if ($null -eq $baseUnit)
    {
        Write-Error "Base file is null."
    }
    return $baseUnit
}

function Add-Column($row, $columnName, $value)
{
    $moddedRow = $row
    return Add-Member -InputObject $moddedRow -MemberType NoteProperty -Name $columnName -Value $value -PassThru
}

function Get-NullRowFromHashTable($hashTable)
{
    $hashTableName = $hashTable["HashTableName"]
    $nullRow = New-Object -TypeName PSCustomObject

    foreach ($entry in $($hashTable.GetEnumerator()))
    {
        if ($entry.Key -eq "HashTableName") { continue }

        $exampleRow = $entry.Value
        break # we just get the first row and break
    }
    foreach ($property in $exampleRow.PSObject.Properties)
    {
        Add-Member -InputObject $nullRow -MemberType NoteProperty -Name "$hashTableName - $($property.Name)" -Value $null
    }
    return $nullRow
}

function Join-CSVDataOuter($csvData)
{
    Write-Host "Joining data..." -ForegroundColor DarkCyan
    $joinedData = New-Object -TypeName System.Collections.Generic.List[PSCustomObject]
    $hashTableCollection = New-HashTableCollection $csvData

    :leftTable for ($i = 0; $i -lt $hashTableCollection.Count; $i++)
    {
        $leftTable = $hashTableCollection[$i]

        :rowInLeftTable foreach ($entry in $($leftTable.GetEnumerator()))
        {
            if ($entry.Key -eq "HashTableName") { continue }
            Write-Progress -Activity "Joining data..." -Status "$($joinedData.Count) records joined."
            $leftTableRow = $entry.Value
            $stampedLeftRow = Stamp-Columns -rowObject $leftTableRow -csvName $leftTable["HashTableName"]
            $moddedLeftRow = Add-Column -row $stampedLeftRow -columnName "In $($leftTable["HashTableName"])" -value $true
            $joinedRow = $moddedLeftRow

            :rightTable for ($j = 0; $j -lt $hashTableCollection.Count; $j++)
            {
                if ($j -eq $i) { continue }

                $rightTable = $hashTableCollection[$j]

                if ($rightTable.Contains($entry.Key))
                {
                    if ($j -lt $i) { continue rowInLeftTable }

                    $rightTableRow = $rightTable[$entry.Key]
                    $stampedRightRow = Stamp-Columns -rowObject $rightTableRow -csvName $rightTable["HashTableName"]
                    $moddedRightRow = Add-Column -row $stampedRightRow -columnName "In $($rightTable["HashTableName"])" -value $true
                    $joinedRow = Join-Objects -object1 $joinedRow -object2 $moddedRightRow
                }
                else
                {
                    $nullRow = Get-NullRowFromHashTable $rightTable
                    $moddedNullRow = Add-Column -row $nullRow -columnName "In $($rightTable["HashTableName"])" -value $false
                    $joinedRow = Join-Objects -object1 $joinedRow -object2 $moddedNullRow
                }
            }
            $joinedData.Add($joinedRow)
        }
    }
    return $joinedData
}

function Export-JoinedCSV($joinedData)
{
    if ($null -eq $joinedData) 
    {
        Write-Error "Joined data is null."
        return
    }
    $dataToExport = $joinedData
    $path = New-Path
    $dataToExport | Export-Csv -Path $path -NoTypeInformation
    Write-Host "Finished exporting to $path." -ForegroundColor Green
}

function New-Path
{
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $timeStamp = New-TimeStamp
    return "$desktopPath\Joined CSV $timeStamp.csv"
}

function New-TimeStamp
{
    return (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
}

# classes
class CSVDataUnit
{
    [PSCustomObject]$importedCSV
    [string]$name
    [string]$joinColumnName

    CSVDataUnit($importedCSV, $name, $joinColumnName)
    {
        $this.importedCSV = $importedCSV
        $this.name = $name
        $this.joinColumnName = $joinColumnName
    }
}

# main program
Introduce-Script
$csvData = Prompt-AllFiles
$joinType = Prompt-JoinType
switch ($joinType)
{
    "InnerJoin"
    { 
        $joinedData = Join-CSVDataInner -csvData $csvData
    }
    "LeftJoin" 
    { 
        $baseFileName = Prompt-BaseFile -csvData $csvData
        $joinedData = Join-CSVDataLeft -csvData $csvData -baseFileName $baseFileName 
    }
    "OuterJoin" 
    { 
        $joinedData = Join-CSVDataOuter -csvData $csvData 
    }
}
Export-JoinedCSV $joinedData
Read-Host "Press Enter to exit."