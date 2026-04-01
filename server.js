const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');

// PowerShell工具模块
const {
    executePowerShell,
    parsePowerShellJson,
    extractJson,
    cleanupFiles,
    cleanupUploadFile,
    copyToTemp,
    ensureTempDir,
    TEMP_DIR,
    CASH_HIGH_THRESHOLD,
    DAY_THRESHOLD
} = require('./lib/powershell-utils');

const app = express();
const PORT = 3000;

const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
}

const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadDir),
    filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});
const upload = multer({ storage, limits: { fileSize: 500 * 1024 * 1024 } });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// 新接口：双文件分析（绩效明细 + 业绩达成率）
app.post('/analyze2', upload.fields([{ name: 'perfFile' }, { name: 'rateFile' }]), async (req, res) => {
    const perfFile = req.files['perfFile']?.[0];
    const rateFile = req.files['rateFile']?.[0];
    const rangeInput = req.body.rangeInput || '';

    if (!perfFile || !rateFile) {
        return res.status(400).json({ error: '请上传两个Excel文件' });
    }

    // 解析可疑区间
    const rangeMatch = rangeInput.match(/(\d+(?:\.\d+)?)\s*%?\s*[-–]\s*(\d+(?:\.\d+)?)\s*%?/);
    if (!rangeMatch) {
        return res.status(400).json({ error: '请输入正确的区间格式，如 100%-105%' });
    }
    const rangeMin = parseFloat(rangeMatch[1]) / 100;
    const rangeMax = parseFloat(rangeMatch[2]) / 100;

    // Copy to simple paths to avoid ACE OLEDB Chinese path issues
    const simplePerfPath = copyToTemp(perfFile.path, 'perf.xlsx');
    const simpleRatePath = copyToTemp(rateFile.path, 'rate.xlsx');
    const escapedPerfPath = simplePerfPath.replace(/\\/g, '\\\\').replace(/'/g, "''");
    const escapedRatePath = simpleRatePath.replace(/\\/g, '\\\\').replace(/'/g, "''");

    const psLines = [
        '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8',
        '$ErrorActionStop = $true',
        '',
        '# === 连接绩效明细Excel ===',
        '$perfPath = \'' + escapedPerfPath + '\'',
        '$perfConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$perfPath;Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`""',
        '$perfConn = New-Object System.Data.OleDb.OleDbConnection $perfConnStr',
        '$perfConn.Open()',
        '$perfSchema = $perfConn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)',
        '$perfTableName = $null',
        'foreach($row in $perfSchema.Rows) { $t = $row[\'TABLE_NAME\']; if ($t -like \'*咨询师绩效*绩效明细*\') { $perfTableName = $t; break } }',
        'if (-not $perfTableName) { throw "Cannot find perf table" }',
        'Write-Host ("PERF_TABLE:" + $perfTableName)',
        '',
        '# === 连接业绩达成率Excel ===',
        '$ratePath = \'' + escapedRatePath + '\'',
        '$rateConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$ratePath;Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`""',
        '$rateConn = New-Object System.Data.OleDb.OleDbConnection $rateConnStr',
        '$rateConn.Open()',
        '$rateSchema = $rateConn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)',
        '$rateTableName = $null',
        'foreach($row in $rateSchema.Rows) { $t = $row[\'TABLE_NAME\']; if ($t -like \'*咨询师*业绩*\' -or $t -like \'*完成率*\') { $rateTableName = $t; break } }',
        'if (-not $rateTableName) { throw "Cannot find rate table" }',
        'Write-Host ("RATE_TABLE:" + $rateTableName)',
        '',
        '# === 从业绩达成率获取可疑咨询师 ===',
        '$rangeMin = ' + rangeMin,
        '$rangeMax = ' + rangeMax,
        '$rateQuery = "SELECT * FROM [$rateTableName]"',
        '$rateAdapter = New-Object System.Data.OleDb.OleDbDataAdapter $rateQuery, $rateConn',
        '$rateDs = New-Object System.Data.DataSet',
        '$rateAdapter.Fill($rateDs) | Out-Null',
        '$rateTable = $rateDs.Tables[0]',
        '',
        '# 建立两个映射: 1) 院区->咨询师列表  2) 咨询师->院区 (业绩完成率在可疑区间内的)',
        '$suspiciousConsultants = @{}',
        '$consultantToOrg = @{}',
        'foreach($row in $rateTable.Rows) {',
        '    $院区 = $row[\'院区\']',
        '    if ([string]::IsNullOrEmpty($院区)) { $院区 = $row[\'院区1\'] }',
        '    if ([string]::IsNullOrEmpty($院区)) { continue }',
        '    $咨询师 = $row[\'咨询师\']',
        '    if ([string]::IsNullOrEmpty($咨询师)) { continue }',
        '    $完成率 = [double]$row[\'业绩完成率\']',
        '    if ($完成率 -ge $rangeMin -and $完成率 -le $rangeMax) {',
        '        if (-not $suspiciousConsultants.ContainsKey($院区)) { $suspiciousConsultants[$院区] = @{} }',
        '        $suspiciousConsultants[$院区][$咨询师] = $true',
        '        $consultantToOrg[$咨询师] = $院区',
        '    }',
        '}',
        '',
        'Write-Host ("SUSPICIOUS_ORGS:" + $suspiciousConsultants.Count)',
        'Write-Host ("SUSPICIOUS_CONSULTANTS:" + ($suspiciousConsultants.Values | ForEach-Object { $_.Keys } | Measure-Object).Count)',
        '',
        '# === 从绩效明细查询存疑情况 ===',
        '# 查询: 同一院区 + 现金类实收>0 + 存在不同咨询师',
        '$perfQuery = "SELECT [病历号], [咨询师], [现金类实收], [收费机构], [收费自由项-院区], [收费时间] FROM [$perfTableName] WHERE [现金类实收] > 0"',
        '$perfAdapter = New-Object System.Data.OleDb.OleDbDataAdapter $perfQuery, $perfConn',
        '$perfDs = New-Object System.Data.DataSet',
        '$perfAdapter.Fill($perfDs) | Out-Null',
        '$perfTable = $perfDs.Tables[0]',
        '',
        '# 按 病历号+收费机构 分组，收集咨询师和现金',
        '$groups = @{}',
        'foreach($row in $perfTable.Rows) {',
        '    $病历号 = $row[\'病历号\']',
        '    $咨询师 = $row[\'咨询师\']',
        '    $收费机构 = $row[\'收费机构\']',
        '    $院区 = $row[\'收费自由项-院区\']',
        '    $现金Val = $row[\'现金类实收\']',
        '    $收费时间 = $row[\'收费时间\']',
        '    if ([string]::IsNullOrEmpty($现金Val)) { $现金Val = 0 }',
        '    $现金 = [double]$现金Val',
        '    if ([string]::IsNullOrEmpty($病历号) -or [string]::IsNullOrEmpty($咨询师)) { continue }',
        '    $key = $病历号 + "|" + $收费机构',
        '    if (-not $groups.ContainsKey($key)) {',
        '        $groups[$key] = @{ 病历号=$病历号; 收费机构=$收费机构; 院区=$院区; consultants=@{}; totalCash=0; 高度异常=$false }',
        '    }',
        '    $groups[$key].consultants[$咨询师] = $true',
        '    $groups[$key].totalCash += $现金',
        '    # 判断高度异常: 金额>8000 且 收费日期在25号之后',
        '    if ($现金 -gt 8000 -and $收费时间) {',
        '        try {',
        '            $datePart = [datetime]::Parse($收费时间.ToString())',
        '            if ($datePart.Day -ge 25) { $groups[$key].高度异常 = $true }',
        '        } catch {}',
        '    }',
        '}',
        '',
        '# 筛选: 同一院区在可疑列表中 且 涉及多个咨询师',
        '# 注意: 如果绩效明细中的院区为空，则使用咨询师在业绩达成率表中的院区',
        '$suspicious = @()',
        'foreach($key in $groups.Keys) {',
        '    $g = $groups[$key]',
        '    $org = $g.院区',
        '    # 如果绩效明细中的院区为空，尝试用咨询师在业绩达成率表中的院区',
        '    if ([string]::IsNullOrEmpty($org)) {',
        '        foreach($c in $g.consultants.Keys) {',
        '            if ($consultantToOrg.ContainsKey($c)) { $org = $consultantToOrg[$c]; break }',
        '        }',
        '    }',
        '    if (-not $suspiciousConsultants.ContainsKey($org)) { continue }',
        '    if ($g.consultants.Count -lt 2) { continue }',
        '    $consultantList = @($g.consultants.Keys)',
        '    $totalCash = $g.totalCash',
        '    $suspicious += [ordered]@{',
        '        病历号 = $g.病历号',
        '        收费机构 = $g.收费机构',
        '        院区 = $org',
        '        涉及咨询师数 = $g.consultants.Count',
        "        涉及咨询师 = ($consultantList | Sort-Object) -join \"; \"",
        '        现金实收总额 = [Math]::Round($totalCash, 2)',
        "        可疑咨询师 = (($consultantList | Where-Object { $suspiciousConsultants[$org].ContainsKey($_) }) -join \"; \")",
        '        高度异常 = if ($g.高度异常) { "⚠️高度异常" } else { "" }',
        '    }',
        '}',
        '',
        'Write-Host ("SUSPICIOUS_COUNT:" + $suspicious.Count)',
        '$suspicious = $suspicious | Sort-Object { [double]$_.现金实收总额 } -Descending',
        '$jsonOutput = $suspicious | ConvertTo-Json -Depth 10 -Compress',
        'Write-Host "JSON_START"',
        'Write-Host $jsonOutput',
        'Write-Host "JSON_END"',
        '',
        '$rateConn.Close()',
        '$perfConn.Close()',
    ];

    const psCommand = psLines.join('\r\n');

    try {
        const result = await executePowerShell(psLines);

        // 清理上传的文件
        cleanupFiles(perfFile.path, rateFile.path, simplePerfPath, simpleRatePath);

        if (result.error) {
            console.error('PS Error:', result.error);
            return res.status(500).json({ error: result.error });
        }

        const { data: suspicious, metadata } = parsePowerShellJson(result.output);
        const { suspiciousOrgCount = 0, suspiciousConsultantCount: suspiciousConsCount = 0 } = metadata;
        const totalCash = (suspicious || []).reduce((sum, s) => sum + (s.现金实收总额 || 0), 0);

        res.json({
            success: true,
            summary: {
                suspiciousOrgCount,
                suspiciousConsultantCount: suspiciousConsCount,
                suspiciousCount: suspicious.length,
                totalSuspiciousCash: Math.round(totalCash * 100) / 100,
                rangeMin: rangeMin,
                rangeMax: rangeMax
            },
            suspicious: suspicious.slice(0, 200)
        });

    } catch (error) {
        cleanupFiles(perfFile.path, rateFile.path, simplePerfPath, simpleRatePath);
        res.status(500).json({ error: error.message || 'Analysis failed' });
    }
});

app.post('/analyze', upload.single('file'), async (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
    }

    // Copy to simple path to avoid ACE OLEDB Chinese path issues
    const simplePath = copyToTemp(req.file.path, 'input.xlsx');
    const escapedPath = simplePath.replace(/\\/g, '\\\\').replace(/'/g, "''");

    const CHECKMARK = '\u221a';
    const psLines = [
        '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8',
        '$ErrorActionStop = $true',
        '$filePath = \'' + escapedPath + '\'',
        '$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$filePath;Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`""',
        '$conn = New-Object System.Data.OleDb.OleDbConnection $connStr',
        '$conn.Open()',
        '$schemaTable = $conn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)',
        '$perfTableName = $null',
        '$yzlTableName = $null',
        'foreach($row in $schemaTable.Rows) { $t = $row[\'TABLE_NAME\']; if ($t -like \'*咨询师绩效*绩效明细*\') { $perfTableName = $t } if ($t -like \'*已整理样本*\') { $yzlTableName = $t } }',
        'if (-not $perfTableName) { throw "Cannot find perf" }',
        'if (-not $yzlTableName) { throw "Cannot find yzl" }',
        '$yzlAdapter = New-Object System.Data.OleDb.OleDbDataAdapter ("SELECT [病历号] FROM [$yzlTableName] WHERE [时间是否对应] = \'' + CHECKMARK + '\'", $conn)',
        '$yzlDs = New-Object System.Data.DataSet',
        '$yzlAdapter.Fill($yzlDs) | Out-Null',
        '$yzlTable = $yzlDs.Tables[0]',
        '$checkedIDs = @()',
        'foreach($row in $yzlTable.Rows) { if ($row[\'病历号\']) { $checkedIDs += $row[\'病历号\'] } }',
        'Write-Host ("CHECKED_COUNT:" + $checkedIDs.Count)',
        'if ($checkedIDs.Count -eq 0) { throw "No checked" }',
        '$inClause = [char]39 + ($checkedIDs -join ([char]39 + [char]39)) + [char]39',
        '$tblName = $perfTableName.Trim([char]39)',
        "$tblName = `$tblName -replace '([`])', '`$1'",
        '$perfQuery = "SELECT [病历号], [咨询师], [现金类实收], [收费机构] FROM [" + $tblName + "] WHERE [病历号] IN (" + $inClause + ") AND [现金类实收] > 0"',
        'Write-Host ("PERF_QUERY:" + $perfQuery.Substring(0, [Math]::Min(200, $perfQuery.Length)))',
        'Write-Host ("PERF_QUERY_SAMPLE:" + $perfQuery.Substring(0, [Math]::Min(300, $perfQuery.Length)))',
        '$perfAdapter = New-Object System.Data.OleDb.OleDbDataAdapter ($perfQuery, $conn)',
        '$perfDs = New-Object System.Data.DataSet',
        '$perfAdapter.Fill($perfDs) | Out-Null',
        '$perfTable = $perfDs.Tables[0]',
        'Write-Host ("PERF_COUNT:" + $perfTable.Rows.Count)',
        '$groups = @{}',
        'foreach($row in $perfTable.Rows) { $key = $row[\'病历号\'] + "|" + $row[\'收费机构\']; if (-not $groups.ContainsKey($key)) { $groups[$key] = @{ 病历号 = $row[\'病历号\']; 收费机构 = $row[\'收费机构\']; records = @(); consultants = @{} } }; $c = $row[\'咨询师\']; if ($c) { $groups[$key].consultants[$c] = $true }; $groups[$key].records += @{ 咨询师 = $c; 现金 = [double]$row[\'现金类实收\'] } }',
        '$suspicious = @()',
        "foreach(`$key in `$groups.Keys) { `$g = `$groups[`$key]; if (`$g.consultants.Count -ge 2) { `$totalCash = (`$g.records | Measure-Object -Property 现金 -Sum).Sum; `$suspicious += [ordered]@{ 病历号 = `$g.病历号; 收费机构 = `$g.收费机构; 涉及咨询师数 = `$g.consultants.Count; 涉及咨询师 = (`$g.consultants.Keys | Sort-Object) -join \"; \"; 现金实收总额 = [Math]::Round(`$totalCash, 2) } } }",
        'Write-Host ("SUSPICIOUS_COUNT:" + $suspicious.Count)',
        '$suspicious = $suspicious | Sort-Object { [double]$_.现金实收总额 } -Descending',
        '$jsonOutput = $suspicious | ConvertTo-Json -Depth 10 -Compress',
        'Write-Host "JSON_START"',
        'Write-Host $jsonOutput',
        'Write-Host "JSON_END"',
        '$conn.Close()',
    ];
    try {
        const result = await executePowerShell(psLines);
        cleanupFiles(req.file.path, simplePath);

        if (result.error) {
            console.error('PS Error:', result.error);
            return res.status(500).json({ error: result.error });
        }

        const { data: suspicious, metadata } = parsePowerShellJson(result.output);
        const { checkedCount = 0, perfCount = 0 } = metadata;
        const totalCash = (suspicious || []).reduce((sum, s) => sum + (s.现金实收总额 || 0), 0);

        res.json({
            success: true,
            summary: {
                totalCheckedIDs: checkedCount,
                totalMatches: perfCount,
                suspiciousCount: suspicious.length,
                totalSuspiciousCash: Math.round(totalCash * 100) / 100
            },
            suspicious: suspicious.slice(0, 200)
        });

    } catch (error) {
        cleanupFiles(req.file.path, simplePath);
        res.status(500).json({ error: error.message || 'Analysis failed' });
    }
});

app.get('/download/:filename', (req, res) => {
    const filePath = path.join(uploadDir, req.params.filename);
    if (fs.existsSync(filePath)) res.download(filePath, req.params.filename);
    else res.status(404).json({ error: 'File not found' });
});

// 定期清理上传目录（使用异步操作避免阻塞）
setInterval(async () => {
    const now = Date.now();
    try {
        const files = await fs.promises.readdir(uploadDir);
        for (const file of files) {
            const filePath = path.join(uploadDir, file);
            const stats = await fs.promises.stat(filePath);
            if (now - stats.mtimeMs > 60 * 60 * 1000) {
                await fs.promises.unlink(filePath);
            }
        }
    } catch (e) {
        console.error('Cleanup error:', e.message);
    }
}, 10 * 60 * 1000);

// 详细报告接口
app.post('/getDetail', upload.single('perfFile'), async (req, res) => {
    console.log('=== /getDetail called ===');

    if (!req.file) {
        return res.status(400).json({ error: '请上传绩效明细文件' });
    }

    // 显式提取参数 - 避免Object.values()顺序问题
    const 病历号 = req.body['病历号'] || req.body['%E7%97%85%E5%8E%9F%E5%8F%B7'];
    const 可疑咨询师 = req.body['可疑咨询师'] || req.body['%E5%8F%AF%E7%96%91%E5%92%A8%E8%AF%9A%E5%B8%88'];

    console.log('病历号:', 病历号);
    console.log('可疑咨询师:', 可疑咨询师);

    if (!病历号 || !可疑咨询师) {
        return res.status(400).json({ error: '缺少参数: 病历号=' + 病历号 + ', 可疑咨询师=' + 可疑咨询师 });
    }

    // Copy to simple path
    const simplePath = copyToTemp(req.file.path, 'detail.xlsx');

    const psLines = [
        '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8',
        '$ErrorActionStop = $true',
        '',
        `$perfPath = '${simplePath.replace(/\\/g, '\\\\')}'`,
        '$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$perfPath;Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`""',
        '$conn = New-Object System.Data.OleDb.OleDbConnection $connStr',
        '$conn.Open()',
        '$schemaTable = $conn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)',
        '$perfTableName = $null',
        'foreach($row in $schemaTable.Rows) { $t = $row[\'TABLE_NAME\']; if ($t -like \'*绩效*\') { $perfTableName = $t; break } }',
        'if (-not $perfTableName) { throw "Cannot find perf table: " + ($schemaTable.Rows | Select-Object -ExpandProperty TABLE_NAME) }',
        '',
        '# 传入参数',
        `$targetId = '${病历号}'`,
        `$targetConsultant = '${可疑咨询师}'`,
        '',
        '# 1. 查询可疑咨询师在该病历号中的所有记录',
        '$query1 = "SELECT [病历号], [咨询师], [收费时间], [现金类实收], [收费机构] FROM [$perfTableName] WHERE [病历号] = \'" + $targetId + "\' AND [咨询师] = \'" + $targetConsultant + "\' ORDER BY [收费时间]"',
        '$adapter1 = New-Object System.Data.OleDb.OleDbDataAdapter $query1, $conn',
        '$ds1 = New-Object System.Data.DataSet',
        '$adapter1.Fill($ds1) | Out-Null',
        '$当前病历记录 = @()',
        'foreach($row in $ds1.Tables[0].Rows) {',
        '    $当前病历记录 += [ordered]@{',
        '        病历号 = $row[\'病历号\']',
        '        咨询师 = $row[\'咨询师\']',
        '        收费时间 = if($row[\'收费时间\']) { $row[\'收费时间\'].ToString(\'yyyy/MM/dd HH:mm:ss\') } else { \"\" }',
        '        现金类实收 = [double]$row[\'现金类实收\']',
        '        收费机构 = $row[\'收费机构\']',
        '    }',
        '}',
        '',
        '# 2. 找到现金最高的记录日期',
        '$maxRecord = $ds1.Tables[0] | Sort-Object { [double]$_[\'现金类实收\'] } -Descending | Select-Object -First 1',
        '$maxDate = if($maxRecord[\'收费时间\']) { [datetime]$maxRecord[\'收费时间\'] } else { $null }',
        '$maxDateStr = if($maxDate) { $maxDate.ToString(\'yyyy-MM-dd\') } else { \"\" }',
        '$月 = if($maxDate) { $maxDate.Month } else { 0 }',
        '',
        'Write-Host "MAX_DATE:" $maxDateStr',
        'Write-Host "MONTH:" $月',
        'Write-Host "TARGET_CONSULTANT:" $targetConsultant',
        'Write-Host "TABLE_NAME:" $perfTableName',
        'Write-Host "MAX_DATE_STR:" $maxDateStr',
        '',
        '# 3. 查询该咨询师在当月高点日期的记录',
        '$当月记录 = @()',
        'if ($maxDate) {',
        '    $startOfMonth = Get-Date -Date ($maxDate.Year, $maxDate.Month, 1) -Format "yyyy-MM-dd"',
        '    $nextDay = $maxDate.AddDays(1)',
        '    $nextDayStr = $nextDay.ToString(\'yyyy-MM-dd\')',
        '    Write-Host "DEBUG: startOfMonth=" $startOfMonth " nextDayStr=" $nextDayStr',
        '    $query2 = "SELECT [病历号], [咨询师], [收费时间], [现金类实收], [收费机构] FROM [$perfTableName] WHERE [咨询师] = \'" + $targetConsultant + "\' AND [收费时间] >= #" + $startOfMonth + "# AND [收费时间] < #" + $nextDayStr + "# ORDER BY [收费时间]"',
        '    Write-Host "QUERY2:" $query2',
        '    try {',
        '        $adapter2 = New-Object System.Data.OleDb.OleDbDataAdapter $query2, $conn',
        '        $ds2 = New-Object System.Data.DataSet',
        '        $adapter2.Fill($ds2) | Out-Null',
        '        Write-Host "QUERY2_OK: rows=" $ds2.Tables[0].Rows.Count',
        '        foreach($row in $ds2.Tables[0].Rows) {',
        '            $当月记录 += [ordered]@{',
        '                病历号 = $row[\'病历号\']',
        '                咨询师 = $row[\'咨询师\']',
        '                收费时间 = if($row[\'收费时间\']) { $row[\'收费时间\'].ToString(\'yyyy/MM/dd HH:mm:ss\') } else { \"\" }',
        '                现金类实收 = [double]$row[\'现金类实收\']',
        '                收费机构 = $row[\'收费机构\']',
        '            }',
        '        }',
        '    } catch {',
        '        Write-Host "QUERY2_ERR:" $_.Exception.Message',
        '    }',
        '}',
        'Write-Host "AFTER_Q2: records=" $当月记录.Count',
        '',
        '# 4. 查询当月1日至高点日期之前的记录',
        '$当月前期记录 = @()',
        'if ($maxDate -and $maxDate.Day -gt 1) {',
        '    $startOfMonth = Get-Date -Date ($maxDate.Year, $maxDate.Month, 1) -Format "yyyy-MM-dd"',
        '    Write-Host "DEBUG_Q3: startOfMonth=" $startOfMonth " maxDateStr=" $maxDateStr',
        '    $query3 = "SELECT [病历号], [咨询师], [收费时间], [现金类实收], [收费机构] FROM [$perfTableName] WHERE [咨询师] = \'" + $targetConsultant + "\' AND [收费时间] >= #" + $startOfMonth + "# AND [收费时间] < #" + $maxDateStr + "# ORDER BY [收费时间]"',
        '    Write-Host "QUERY3:" $query3',
        '    try {',
        '        $adapter3 = New-Object System.Data.OleDb.OleDbDataAdapter $query3, $conn',
        '        $ds3 = New-Object System.Data.DataSet',
        '        $adapter3.Fill($ds3) | Out-Null',
        '        Write-Host "QUERY3_OK: rows=" $ds3.Tables[0].Rows.Count',
        '        foreach($row in $ds3.Tables[0].Rows) {',
        '            $当月前期记录 += [ordered]@{',
        '                病历号 = $row[\'病历号\']',
        '                咨询师 = $row[\'咨询师\']',
        '                收费时间 = if($row[\'收费时间\']) { $row[\'收费时间\'].ToString(\'yyyy/MM/dd HH:mm:ss\') } else { \"\" }',
        '                现金类实收 = [double]$row[\'现金类实收\']',
        '                收费机构 = $row[\'收费机构\']',
        '            }',
        '        }',
        '    } catch {',
        '        Write-Host "QUERY3_ERR:" $_.Exception.Message',
        '    }',
        '}',
        'Write-Host "AFTER_Q3: before=" $当月前期记录.Count',
        '',
        '$conn.Close()',
        '',
        '# Output simple debug string first',
        '$debugStr = "DEBUG: consultant=" + $targetConsultant + " maxDate=" + $maxDateStr + " ds1=" + $ds1.Tables[0].Rows.Count + " records=" + $当月记录.Count + " before=" + $当月前期记录.Count',
        'Write-Host $debugStr',
        '',
        '# Output JSON with debug',
        'try {',
        '    $result = @{',
        '        当前病历记录 = $当前病历记录',
        '        当月 = $月',
        '        高点日期 = $maxDateStr',
        '        当月高点日期记录 = $当月记录',
        '        当月前期记录 = $当月前期记录',
        '        _debug = @{',
        '            targetConsultant = $targetConsultant',
        '            maxDateStr = $maxDateStr',
        '            ds1Count = $ds1.Tables[0].Rows.Count',
        '            recordsCount = $当月记录.Count',
        '            beforeCount = $当月前期记录.Count',
        '        }',
        '    }',
        '    $jsonOutput = $result | ConvertTo-Json -Depth 10',
        '    Write-Host "JSON_START"',
        '    Write-Host $jsonOutput',
        '    Write-Host "JSON_END"',
        '} catch {',
        '    Write-Host "JSON_ERROR:" $_.Exception.Message',
        '}',
    ];

    const psCommand = psLines.join('\r\n');

    try {
        const result = await executePowerShell(psLines);
        cleanupFiles(req.file.path, simplePath);

        if (result.error) {
            console.error('PS Error:', result.error);
            return res.status(500).json({ error: result.error });
        }

        const output = result.output;
        console.log('PS Output:', output);

        // Extract all debug lines
        const debugMatches = output.match(/DEBUG:.*/g);
        const debugStr = debugMatches ? debugMatches.join('\n') : 'No debug info';

        // Extract JSON using helper function
        const { data, rawJson } = extractJson(output);

        if (!data) {
            return res.json({ error: 'PowerShell输出格式错误', debug: debugStr, output: output.substring(0, 500) });
        }

        if (rawJson.includes('JSON_ERROR')) {
            return res.json({ error: 'PowerShell错误: ' + rawJson, debug: debugStr });
        }

        data._debugStr = debugStr;
        res.json(data);

    } catch (error) {
        cleanupFiles(req.file.path, simplePath);
        res.status(500).json({ error: error.message || '获取详细失败' });
    }
});

// 模块2：排班休息日异常分析
app.post('/analyzeSchedule', upload.fields([
    { name: 'perfFile', maxCount: 1 },
    { name: 'scheduleFile', maxCount: 1 }
]), async (req, res) => {
    console.log('=== /analyzeSchedule called ===');

    const perfFile = req.files?.['perfFile']?.[0];
    const scheduleFile = req.files?.['scheduleFile']?.[0];

    if (!perfFile || !scheduleFile) {
        return res.status(400).json({ error: '请上传绩效分析和排班表两个文件' });
    }

    const perfPath = copyToTemp(perfFile.path, 'perf.xlsx');
    const schedulePath = copyToTemp(scheduleFile.path, 'schedule.xlsx');

    const psLines = [
        '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8',
        '$ErrorActionStop = $true',
        '',
        `$perfPath = '${perfPath.replace(/\\/g, '\\\\')}'`,
        '$schedulePath = \'' + schedulePath.replace(/\\/g, '\\\\') + '\'',
        '',
        '# 连接绩效文件',
        '$perfConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$perfPath;Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`""',
        '$perfConn = New-Object System.Data.OleDb.OleDbConnection $perfConnStr',
        '$perfConn.Open()',
        '',
        '# 获取绩效表名',
        '$perfSchema = $perfConn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)',
        '$perfTableName = $null',
        'foreach($row in $perfSchema.Rows) { $t = $row[\'TABLE_NAME\']; if ($t -like \'*绩效*\' -or $t -like \'*绩效明细*\') { $perfTableName = $t; break } }',
        'if (-not $perfTableName) { $perfTableName = $perfSchema.Rows[0][\'TABLE_NAME\'] }',
        'Write-Host "PERF_TABLE:" $perfTableName',
        '',
        '# 连接排班文件',
        '$schedConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$schedulePath;Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`""',
        '$schedConn = New-Object System.Data.OleDb.OleDbConnection $schedConnStr',
        '$schedConn.Open()',
        '',
        '# 获取排班表名',
        '$schedSchema = $schedConn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)',
        '$schedTableName = $null',
        'foreach($row in $schedSchema.Rows) { $t = $row[\'TABLE_NAME\']; if ($t) { $schedTableName = $t; break } }',
        'if (-not $schedTableName) { $schedTableName = $schedSchema.Rows[0][\'TABLE_NAME\'] }',
        '$debug = @()',
        '$debug += "SCHED_TABLE:$schedTableName"',
        '',
        '# 读取排班表所有数据',
        '$schedFullQuery = "SELECT * FROM [$schedTableName]"',
        '$schedFullAdapter = New-Object System.Data.OleDb.OleDbDataAdapter $schedFullQuery, $schedConn',
        '$schedFullDs = New-Object System.Data.DataSet',
        '$schedFullAdapter.Fill($schedFullDs) | Out-Null',
        '$schedTable = $schedFullDs.Tables[0]',
        '$debug += "SCHED_ROWS:$($schedTable.Rows.Count)"',
        '',
        '# 姓名列在B列(索引1)，数据从D列(索引3)开始',
        '$nameColIndex = 1',
        '$dataStartColIndex = 3',
        '',
        '# 构建日期映射：列索引 -> 日期（新格式：2026-02-01(星期日)）',
        '$colDateMap = @{}',
        '$headerRow = $schedTable.Rows[0]',
        'for($i = $dataStartColIndex; $i -lt $schedTable.Columns.Count; $i++) {',
        '    $colName = $headerRow[$i].ToString()',
        '    # 解析 "2026-02-01(星期日)" 格式',
        '    if ($colName -match "^(\\d{4})-(\\d{2})-(\\d{2})") {',
        '        $year = [int]$matches[1]',
        '        $month = [int]$matches[2]',
        '        $day = [int]$matches[3]',
        '        try {',
        '            $date = Get-Date -Year $year -Month $month -Day $day',
        '            $colDateMap[$i] = $date',
        '        } catch { }',
        '    }',
        '}',
        '$debug += "COL_MAP_COUNT:$($colDateMap.Count)"',
        '',
        '# 构建工作日字典：(咨询师, 日期) -> true',
        '# 如果单元格包含"上班"或"加班"，则该天为工作日',
        '$workDays = @{}',
        '$allConsultants = @{}',
        '',
        'foreach($row in $schedTable.Rows) {',
        '    $name = $row[$nameColIndex].ToString().Trim()',
        '    if (-not $name) { continue }',
        '    $allConsultants[$name] = $true',
        '    ',
        '    for($i = $dataStartColIndex; $i -lt $schedTable.Columns.Count; $i++) {',
        '        $cell = $row[$i].ToString()',
        '        # 包含"上班"或"加班"即为工作日',
        '        if ($cell -match "上班" -or $cell -match "加班") {',
        '            if ($colDateMap.ContainsKey($i)) {',
        '                $dateStr = $colDateMap[$i].ToString("yyyy-MM-dd")',
        '                $key = $name + "|" + $dateStr',
        '                $workDays[$key] = $true',
        '            }',
        '        }',
        '    }',
        '}',
        '$debug += "CONSULTANTS:$($allConsultants.Count)"',
        '$debug += "WORK_DAYS:$($workDays.Count)"',
        '',
        '# 读取绩效现金记录',
        '$cashQuery = "SELECT [病历号], [咨询师], [收费时间], [现金类实收], [收费机构] FROM [$perfTableName] WHERE [现金类实收] <> 0"',
        '$cashAdapter = New-Object System.Data.OleDb.OleDbDataAdapter $cashQuery, $perfConn',
        '$cashDs = New-Object System.Data.DataSet',
        '$cashAdapter.Fill($cashDs) | Out-Null',
        '$debug += "CASH_ROWS:$($cashDs.Tables[0].Rows.Count)"',
        '',
        '# 查找异常：现金收入日期不在工作日中',
        '$异常记录 = New-Object System.Collections.ArrayList',
        'foreach($row in $cashDs.Tables[0].Rows) {',
        '    $name = $row[\'咨询师\'].ToString().Trim()',
        '    $dt = $row[\'收费时间\']',
        '    if ($dt -is [datetime]) {',
        '        $dateStr = $dt.ToString("yyyy-MM-dd")',
        '        $key = $name + "|" + $dateStr',
        '        # 只在不在工作日时标记为异常',
        '        if (-not $workDays.ContainsKey($key)) {',
        '            $rec = [ordered]@{',
        '                咨询师 = $name',
        '                异常日期 = $dateStr',
        '                排班状态 = \'休息\'',
        '                病历号 = $row[\'病历号\']',
        '                收费时间 = $dt.ToString(\'yyyy/MM/dd HH:mm:ss\')',
        '                现金类实收 = [double]$row[\'现金类实收\']',
        '                收费机构 = $row[\'收费机构\']',
        '            }',
        '            $异常记录.Add($rec) | Out-Null',
        '        }',
        '    }',
        '}',
        '$debug += "ANOMALIES:$($异常记录.Count)"',
        '',
        '$perfConn.Close()',
        '$schedConn.Close()',
        '',
        '# 返回',
        '$uniqueConsultants = @{}',
        '$totalCashAmount = 0.0',
        'foreach($r in $异常记录) {',
        '    $uniqueConsultants[$r[\'咨询师\']] = $true',
        '    $totalCashAmount += $r[\'现金类实收\']',
        '}',
        '$result = @{',
        '    records = @($异常记录)',
        '    suspiciousRecords = $异常记录.Count',
        '    suspiciousConsultants = $uniqueConsultants.Count',
        '    totalCash = $totalCashAmount',
        '    debug = @{',
        '        info = $debug',
        '        schedRows = $schedTable.Rows.Count',
        '        consultants = $allConsultants.Count',
        '        workDays = $workDays.Count',
        '        cashRows = $cashDs.Tables[0].Rows.Count',
        '    }',
        '}',
        '$jsonOutput = $result | ConvertTo-Json -Depth 10',
        'Write-Host "JSON_START"',
        'Write-Host $jsonOutput',
        'Write-Host "JSON_END"',
    ];

    const psCommand = psLines.join('\r\n');

    try {
        const result = await executePowerShell(psLines);
        cleanupFiles(perfFile.path, scheduleFile.path, perfPath, schedulePath);

        if (result.error) {
            console.error('PS Error:', result.error);
            return res.status(500).json({ error: result.error, psOutput: result.output });
        }

        const output = result.output;
        console.log('PS Output:', output);

        const { data } = extractJson(output);

        if (!data) {
            return res.status(500).json({ error: 'PowerShell输出格式错误', psOutput: output.substring(0, 1000) });
        }

        if (!data.debug) {
            data.psOutput = output.substring(0, 1000);
        }
        res.json(data);

    } catch (error) {
        cleanupFiles(perfFile.path, scheduleFile.path, perfPath, schedulePath);
        res.status(500).json({ error: error.message || '分析失败' });
    }
});

app.listen(PORT, () => {
    console.log('Server running at http://localhost:' + PORT);
    console.log('Upload directory: ' + uploadDir);
});
