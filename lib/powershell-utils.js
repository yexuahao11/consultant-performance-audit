/**
 * PowerShell Utilities for Consultant Performance Analysis
 * 公共PowerShell工具模块 - 消除重复代码
 */

const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

// ==================== 常量 ====================

const TEMP_DIR = 'C:\\temp_analysis';
const OLEDB_PROVIDER = 'Microsoft.ACE.OLEDB.12.0';
const EXCEL_EXTENDED_PROPS = 'Excel 12.0 Xml;HDR=YES;IMEX=1';
const CASH_HIGH_THRESHOLD = 8000;
const DAY_THRESHOLD = 25;
const WARNING_FLAG = '\u26a0\ufe0f\u9ad8\u5ea6\u5f02\u5e38'; // ⚠️高度异常

// ==================== 目录管理 ====================

/**
 * 确保临时目录存在
 */
function ensureTempDir() {
    if (!fs.existsSync(TEMP_DIR)) {
        fs.mkdirSync(TEMP_DIR, { recursive: true });
    }
}

/**
 * 复制文件到临时目录
 * @param {string} sourcePath - 源文件路径
 * @param {string} destName - 目标文件名
 * @returns {string} 目标文件完整路径
 */
function copyToTemp(sourcePath, destName) {
    ensureTempDir();
    const destPath = path.join(TEMP_DIR, destName);
    fs.copyFileSync(sourcePath, destPath);
    return destPath;
}

// ==================== 文件清理 ====================

/**
 * 清理多个文件，忽略错误
 * @param {...string} paths - 要清理的文件路径
 */
function cleanupFiles(...paths) {
    for (const p of paths) {
        try {
            if (p && fs.existsSync(p)) {
                fs.unlinkSync(p);
            }
        } catch (e) {
            console.error('Cleanup failed:', p, e.message);
        }
    }
}

/**
 * 从请求文件对象清理上传的临时文件
 * @param {Object} file - multer文件对象
 */
function cleanupUploadFile(file) {
    if (file && file.path) {
        cleanupFiles(file.path);
    }
}

// ==================== OLEDB连接 ====================

/**
 * 构建OLEDB连接字符串
 * @param {string} filePath - Excel文件路径
 * @returns {string} 连接字符串
 */
function buildOleDbConnection(filePath) {
    const escapedPath = filePath.replace(/\\/g, '\\\\').replace(/'/g, "''");
    return `Provider=${OLEDB_PROVIDER};Data Source=${escapedPath};Extended Properties="${EXCEL_EXTENDED_PROPS}"`;
}

/**
 * 获取Excel表名
 * @param {Object} conn - OleDbConnection对象
 * @param {string} tablePattern - 表名匹配模式（如 '*咨询师绩效*'）
 * @returns {string} 表名
 */
function buildFindTableScript(connVarName, tablePattern) {
    return `
    $schema = ${connVarName}.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
    $tableName = $null
    foreach($row in $schema.Rows) {
        $t = $row['TABLE_NAME']
        if ($t -like '${tablePattern}') { $tableName = $t; break }
    }
    if (-not $tableName) { throw "Cannot find table matching: ${tablePattern}" }
    Write-Host "TABLE_NAME:" $tableName
`;
}

// ==================== PowerShell执行 ====================

/**
 * 执行PowerShell脚本并返回结果
 * @param {string[]} psLines - PowerShell脚本行数组
 * @returns {Promise<{output: string, error: string|null, code: number}>}
 */
function executePowerShell(psLines) {
    return new Promise((resolve) => {
        // 创建临时脚本文件，使用UTF-8 BOM
        const tmpFile = path.join(__dirname, '..', 'uploads', 'tmp_' + Date.now() + '.ps1');
        const BOM = '\ufeff';
        fs.writeFileSync(tmpFile, BOM + psLines.join('\r\n'), 'utf8');

        const ps = spawn('powershell.exe', [
            '-ExecutionPolicy', 'Bypass',
            '-NoProfile',
            '-File', tmpFile
        ], {
            windowsHide: true
        });

        let stdout = '', stderr = '';
        ps.stdout.on('data', (data) => { stdout += data.toString(); });
        ps.stderr.on('data', (data) => { stderr += data.toString(); });
        ps.on('close', (code) => {
            try { fs.unlinkSync(tmpFile); } catch (e) {}
            resolve({ output: stdout, error: code !== 0 ? stderr : null, code });
        });
        ps.on('error', (err) => {
            try { fs.unlinkSync(tmpFile); } catch (e) {}
            resolve({ output: '', error: err.message, code: 1 });
        });
    });
}

// ==================== JSON解析 ====================

/**
 * 从PowerShell输出中解析JSON（支持JSON_START/END标记）
 * @param {string} output - PowerShell标准输出
 * @returns {{data: Object|Array, metadata: Object}}
 */
function parsePowerShellJson(output) {
    const lines = output.split('\n');
    const metadata = {};
    let jsonData = '';

    for (const line of lines) {
        // 解析元数据标记
        if (line.startsWith('SUSPICIOUS_ORGS:')) {
            metadata.suspiciousOrgCount = parseInt(line.replace('SUSPICIOUS_ORGS:', ''));
        } else if (line.startsWith('SUSPICIOUS_CONSULTANTS:')) {
            metadata.suspiciousConsultantCount = parseInt(line.replace('SUSPICIOUS_CONSULTANTS:', ''));
        } else if (line.startsWith('SUSPICIOUS_COUNT:')) {
            metadata.suspiciousCount = parseInt(line.replace('SUSPICIOUS_COUNT:', ''));
        } else if (line.startsWith('CHECKED_COUNT:')) {
            metadata.checkedCount = parseInt(line.replace('CHECKED_COUNT:', ''));
        } else if (line.startsWith('PERF_COUNT:')) {
            metadata.perfCount = parseInt(line.replace('PERF_COUNT:', ''));
        } else if (line.startsWith('TABLE_NAME:')) {
            metadata.tableName = line.replace('TABLE_NAME:', '').trim();
        } else if (line.startsWith('JSON_START')) {
            continue;
        } else if (line.startsWith('JSON_END')) {
            continue;
        } else if (line.startsWith('{') || line.startsWith('[')) {
            jsonData = line;
        }
    }

    let data = null;
    if (jsonData) {
        try {
            data = JSON.parse(jsonData);
            if (!Array.isArray(data)) data = [data];
        } catch (e) {
            console.error('JSON parse error:', e.message);
        }
    }

    return { data, metadata };
}

/**
 * 从PowerShell输出中提取JSON（使用indexOf方法）
 * @param {string} output - PowerShell标准输出
 * @returns {{data: Object|Array, rawJson: string}}
 */
function extractJson(output) {
    const startIdx = output.indexOf('JSON_START');
    const endIdx = output.indexOf('JSON_END');

    if (startIdx === -1 || endIdx === -1) {
        return { data: null, rawJson: '' };
    }

    const jsonStr = output.substring(startIdx + 'JSON_START'.length, endIdx).trim();

    let data = null;
    try {
        data = JSON.parse(jsonStr);
    } catch (e) {
        console.error('JSON parse error:', e.message);
    }

    return { data, rawJson: jsonStr };
}

// ==================== PowerShell脚本构建器 ====================

/**
 * 构建基础PowerShell脚本头
 * @returns {string[]}
 */
function buildScriptHeader() {
    return [
        '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8',
        '$ErrorActionStop = $true'
    ];
}

/**
 * 构建JSON输出包装器
 * @param {string} resultVarName - 结果变量名
 * @returns {string[]}
 */
function buildJsonOutput(resultVarName) {
    return [
        '$jsonOutput = ' + resultVarName + ' | ConvertTo-Json -Depth 10 -Compress',
        'Write-Host "JSON_START"',
        'Write-Host $jsonOutput',
        'Write-Host "JSON_END"'
    ];
}

/**
 * 构建OLEDB连接脚本
 * @param {string} connVarName - 连接变量名
 * @param {string} filePathVar - 文件路径变量名
 * @returns {string[]}
 */
function buildConnectionScript(connVarName, filePathVar) {
    return [
        `$${connVarName}ConnStr = "Provider=${OLEDB_PROVIDER};Data Source=$${filePathVar};Extended Properties=\"${EXCEL_EXTENDED_PROPS}\""`,
        `$${connVarName}Conn = New-Object System.Data.OleDb.OleDbConnection $${connVarName}ConnStr`,
        `$${connVarName}Conn.Open()`
    ];
}

/**
 * 关闭OLEDB连接
 * @param {...string} connVarNames - 连接变量名
 * @returns {string[]}
 */
function buildCloseConnections(...connVarNames) {
    return connVarNames.map(name => `$${name}Conn.Close()`);
}

// ==================== 导出 ====================

module.exports = {
    // 常量
    TEMP_DIR,
    OLEDB_PROVIDER,
    EXCEL_EXTENDED_PROPS,
    CASH_HIGH_THRESHOLD,
    DAY_THRESHOLD,
    WARNING_FLAG,

    // 目录管理
    ensureTempDir,
    copyToTemp,

    // 文件清理
    cleanupFiles,
    cleanupUploadFile,

    // OLEDB
    buildOleDbConnection,
    buildFindTableScript,

    // PowerShell执行
    executePowerShell,

    // JSON解析
    parsePowerShellJson,
    extractJson,

    // 脚本构建器
    buildScriptHeader,
    buildJsonOutput,
    buildConnectionScript,
    buildCloseConnections
};
