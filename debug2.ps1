[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionStop = $true
$filePath = 'D:\coding\咨询师绩效审核\咨询师绩效-数据分析0326.xlsx'
$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$filePath;Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`";"
$conn = New-Object System.Data.OleDb.OleDbConnection $connStr
$conn.Open()
$schemaTable = $conn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
$yzlTableName = $null
foreach($row in $schemaTable.Rows) { $t = $row['TABLE_NAME']; if ($t -like '*已整理样本*') { $yzlTableName = $t } }
$yzlAdapter = New-Object System.Data.OleDb.OleDbDataAdapter ("SELECT TOP 5 [病历号], [时间是否对应] FROM [$yzlTableName] WHERE [时间是否对应] IS NOT NULL", $conn)
$yzlDs = New-Object System.Data.DataSet; $yzlAdapter.Fill($yzlDs) | Out-Null; $yzlTable = $yzlDs.Tables[0]
foreach($row in $yzlTable.Rows) { $id = $row['病历号']; $f = $row['时间是否对应']; Write-Host ("ROW:" + $id + "|F:" + $f + "|FTYPE:" + $f.GetType().Name + "|FCODE:" + [int][char]$f) }
$conn.Close()