[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$filePath = 'D:\coding\咨询师绩效审核\咨询师绩效-数据分析0326.xlsx'
$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$filePath;Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`";"
$conn = New-Object System.Data.OleDb.OleDbConnection $connStr
$conn.Open()

$schemaTable = $conn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
foreach($row in $schemaTable.Rows) {
    $tblName = $row['TABLE_NAME']
    if ($tblName -like '*已整理样本*') {
        Write-Host "Found table: $tblName"
        $yzlAdapter = New-Object System.Data.OleDb.OleDbDataAdapter ("SELECT * FROM [$tblName] WHERE [时间是否对应] = '√'", $conn)
        $yzlDs = New-Object System.Data.DataSet
        $yzlAdapter.Fill($yzlDs) | Out-Null
        $yzlTable = $yzlDs.Tables[0]
        Write-Host "Rows with checkmark: $($yzlTable.Rows.Count)"
        if ($yzlTable.Rows.Count -gt 0) {
            Write-Host "First ID:" $yzlTable.Rows[0]['病历号']
        }
    }
}
$conn.Close()