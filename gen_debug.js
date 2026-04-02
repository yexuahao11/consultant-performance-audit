const fs = require('fs');

const escapedPath = 'D:\\coding\\咨询师绩效审核\\咨询师绩效-数据分析0326.xlsx';

const script =
"[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\r\n" +
"$ErrorActionStop = $true\r\n" +
"$filePath = '" + escapedPath + "'\r\n" +
"$connStr = \"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$filePath;Extended Properties=`\"Excel 12.0 Xml;HDR=YES;IMEX=1`\";\"\r\n" +
"$conn = New-Object System.Data.OleDb.OleDbConnection $connStr\r\n" +
"$conn.Open()\r\n" +
"$schemaTable = $conn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)\r\n" +
"$yzlTableName = $null\r\n" +
"foreach($row in $schemaTable.Rows) { $t = $row['TABLE_NAME']; if ($t -like '*已整理样本*') { $yzlTableName = $t } }\r\n" +
"$yzlAdapter = New-Object System.Data.OleDb.OleDbDataAdapter (\"SELECT TOP 5 [病历号], [时间是否对应] FROM [$yzlTableName] WHERE [时间是否对应] IS NOT NULL\", $conn)\r\n" +
"$yzlDs = New-Object System.Data.DataSet; $yzlAdapter.Fill($yzlDs) | Out-Null; $yzlTable = $yzlDs.Tables[0]\r\n" +
"foreach($row in $yzlTable.Rows) { $id = $row['病历号']; $f = $row['时间是否对应']; Write-Host (\"ROW:\" + $id + \"|F:\" + $f + \"|FTYPE:\" + $f.GetType().Name + \"|FCODE:\" + [int][char]$f) }\r\n" +
"$conn.Close()";

const BOM = '\ufeff';
fs.writeFileSync('debug2.ps1', BOM + script, 'utf8');
console.log('Written');
