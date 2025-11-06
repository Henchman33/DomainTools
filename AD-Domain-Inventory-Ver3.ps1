# Active Directory Infrastructure Documentation Script - Enhanced Version
# Run this on a Domain Controller with appropriate permissions

# Set output directory to user's desktop
$OutputPath = [Environment]::GetFolderPath("Desktop") + "\AD_Documentation"
if (!(Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = "$OutputPath\AD_Infrastructure_Report_$Timestamp.html"
$ExcelFile = "$OutputPath\AD_Infrastructure_Report_$Timestamp.xlsx"

# Import required module
Import-Module ActiveDirectory

Write-Host "Starting Active Directory Infrastructure Documentation..." -ForegroundColor Green
Write-Host "Output directory: $OutputPath" -ForegroundColor Yellow

# Initialize collections for Excel export
$ExcelData = @{
    DomainControllers = @()
    ReplicationHealth = @()
    Sites = @()
    DNSZones = @()
    DHCPScopes = @()
    Tier0Accounts = @()
    Tier1Accounts = @()
    Tier2Accounts = @()
    ServiceAccounts = @()
    ExchangeServers = @()
    GroupPolicies = @()
}

# Initialize HTML report with collapsible sections
$HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Active Directory Infrastructure Documentation</title>
    <title2>Author: Stephen McKee IGTPLC</title2>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        h1 { color: #003366; border-bottom: 3px solid #003366; padding-bottom: 10px; }
        h2 { color: #003366; margin-top: 30px; border-bottom: 2px solid #003366; padding-bottom: 5px; cursor: pointer; user-select: none; }
        h2:hover { background-color: #e6f3ff; }
        h2::before { content: '▼ '; font-size: 0.8em; }
        h2.collapsed::before { content: '▶ '; }
        h3 { color: #003366; margin-top: 20px; }
        .section-content { margin-left: 20px; }
        .collapsed-content { display: none; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; background-color: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        th { background-color: #003366; color: white; padding: 12px; text-align: left; font-weight: bold; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #e6f0f5; }
        .info-box { background-color: white; padding: 15px; margin: 15px 0; border-left: 4px solid #003366; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .warning { border-left-color: #ff9900; }
        .success { border-left-color: #00cc66; }
        .error { border-left-color: #cc0000; }
        .timestamp { color: #666; font-size: 0.9em; }
        .critical { color: #cc0000; font-weight: bold; }
        .healthy { color: #00cc00; font-weight: bold; }
        .tier0 { background-color: #ffcccc; }
        .tier1 { background-color: #ffffcc; }
        .tier2 { background-color: #ccffcc; }
        .toggle-all { margin: 20px 0; padding: 10px 20px; background-color: #003366; color: white; border: none; cursor: pointer; font-size: 14px; border-radius: 4px; }
        .toggle-all:hover { background-color: #002244; }
        /* Search bar styles */
        .search-bar { margin: 15px 0; display: flex; gap: 8px; align-items: center; }
        .search-bar input[type="text"] { flex: 1; padding: 10px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }
        .search-bar button { padding: 10px 14px; background-color: #003366; color: white; border: none; border-radius: 4px; cursor: pointer; }
        .search-bar button:hover { background-color: #002244; }
        .highlight { background-color: #fff176; padding: 2px 0; border-radius: 2px; }
    </style>
    <script>
        function toggleSection(element) {
            const content = element.nextElementSibling;
            const isCollapsed = content.classList.contains('collapsed-content');
            
            if (isCollapsed) {
                content.classList.remove('collapsed-content');
                element.classList.remove('collapsed');
            } else {
                content.classList.add('collapsed-content');
                element.classList.add('collapsed');
            }
        }
        
        function toggleAll() {
            const headers = document.querySelectorAll('h2');
            const firstHeader = headers[0];
            const firstContent = firstHeader.nextElementSibling;
            const shouldExpand = firstContent.classList.contains('collapsed-content');
            
            headers.forEach(header => {
                const content = header.nextElementSibling;
                if (shouldExpand) {
                    content.classList.remove('collapsed-content');
                    header.classList.remove('collapsed');
                } else {
                    content.classList.add('collapsed-content');
                    header.classList.add('collapsed');
                }
            });
        }
        
          // ...
        // Utility to escape regex special characters
        function escapeRegExp(string) {
            // Use \x24 and \x7B/\x7D so there is no literal ${&} sequence in the here-string.
            return string.replace(/[.*+?^\x24\x7B\x7D()|[\]\\]/g, '\\`$&');
        }
        
        }

        // Remove existing highlight spans
        function clearHighlightsWithin(root) {
            const highlights = (root || document).querySelectorAll('.highlight');
            highlights.forEach(span => {
                const parent = span.parentNode;
                parent.replaceChild(document.createTextNode(span.textContent), span);
                parent.normalize();
            });
        }

        // Highlight matches inside an element
        function highlight(element, query) {
            if (!query) { return; }
            clearHighlightsWithin(element);
            const regex = new RegExp('(' + escapeRegExp(query) + ')', 'ig');
            // Use innerHTML replacement - acceptable for generated report
            element.innerHTML = element.innerHTML.replace(regex, '<span class="highlight">$1</span>');
        }

        // Perform a search across section headers and contents; expand matching sections and collapse non-matching ones
        function performSearch() {
            const query = document.getElementById('searchInput').value.trim().toLowerCase();
            const headers = document.querySelectorAll('h2');
            if (!query) {
                // If query empty, clear highlights and expand all sections
                headers.forEach(header => {
                    const content = header.nextElementSibling;
                    content.classList.remove('collapsed-content');
                    header.classList.remove('collapsed');
                    clearHighlightsWithin(content);
                });
                return;
            }

            headers.forEach(header => {
                const content = header.nextElementSibling;
                // Combine header text and content text for searching
                const combinedText = (header.textContent + ' ' + content.textContent).toLowerCase();
                if (combinedText.indexOf(query) !== -1) {
                    content.classList.remove('collapsed-content');
                    header.classList.remove('collapsed');
                    highlight(content, query);
                } else {
                    // Collapse and clear highlights
                    content.classList.add('collapsed-content');
                    header.classList.add('collapsed');
                    clearHighlightsWithin(content);
                }
            });
        }

        function clearSearch() {
            document.getElementById('searchInput').value = '';
            performSearch(); // will clear highlights and expand all
        }
        
        window.onload = function() {
            document.querySelectorAll('h2').forEach(header => {
                header.addEventListener('click', function() { toggleSection(this); });
            });
        };
    </script>
</head>
<body>
    <h1>Active Directory Infrastructure Documentation</h1>
    <p class="timestamp">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
    <p class="timestamp">Output Location: $OutputPath</p>
    <button class="toggle-all" onclick="toggleAll()">Expand/Collapse All Sections</button>
"@

# 1. FOREST INFORMATION
Write-Host "Gathering Forest Information..." -ForegroundColor Cyan
# Insert search bar just above the Forest Overview section
$HTML += @"
    <div class='search-bar'>
        <input type='text' id='searchInput' placeholder='Search the report (e.g., domain controller, replication, DNS, GPO)...' onkeypress='if(event.key === \"Enter\") performSearch();' />
        <button onclick='performSearch()'>Search</button>
        <button onclick='clearSearch()'>Clear</button>
    </div>

    <h2>1. Forest Overview</h2>
    <div class="section-content">
    <div class="info-box success">
        <table>
            <tr><th>Property</th><th>Value</th></tr>
"@

$Forest = Get-ADForest
$HTML += @"
            <tr><td>Forest Name</td><td>$($Forest.Name)</td></tr>
            <tr><td>Forest Functional Level</td><td>$($Forest.ForestMode)</td></tr>
            <tr><td>Schema Master</td><td>$($Forest.SchemaMaster)</td></tr>
            <tr><td>Domain Naming Master</td><td>$($Forest.DomainNamingMaster)</td></tr>
            <tr><td>Root Domain</td><td>$($Forest.RootDomain)</td></tr>
            <tr><td>Total Domains</td><td>$($Forest.Domains.Count)</td></tr>
        </table>
    </div>
    <h3>Domains in Forest</h3>
    <ul>
"@
foreach ($domain in $Forest.Domains) {
    $HTML += "        <li>$domain</li>`n"
}
$HTML += "    </ul>`n</div>`n"

# 2. DOMAIN CONTROLLERS
Write-Host "Gathering Domain Controller Information..." -ForegroundColor Cyan
$DCs = Get-ADDomainController -Filter *
$HTML += @"
    <h2>2. Domain Controllers (Total: $($DCs.Count))</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>Hostname</th>
            <th>Site</th>
            <th>IP Address</th>
            <th>OS Version</th>
            <th>Global Catalog</th>
            <th>FSMO Roles</th>
            <th>DNS Service</th>
        </tr>
"@

foreach ($DC in $DCs) {
    $FSMORoles = @()
    if ($DC.OperationMasterRoles) {
        $FSMORoles = $DC.OperationMasterRoles -join ", "
    } else {
        $FSMORoles = "None"
    }
    
    $IsGC = if ($DC.IsGlobalCatalog) { "Yes" } else { "No" }
    $IsGCHTML = if ($DC.IsGlobalCatalog) { "<span class='healthy'>Yes</span>" } else { "No" }
    
    # Check DNS service
    $DNSStatus = "Unknown"
    try {
        $DNSService = Get-Service -ComputerName $DC.HostName -Name DNS -ErrorAction SilentlyContinue
        if ($DNSService) {
            $DNSStatus = $DNSService.Status
        }
    } catch {
        $DNSStatus = "Unable to query"
    }
    
    $DNSStatusHTML = if ($DNSStatus -eq "Running") { "<span class='healthy'>Running</span>" } else { "<span class='critical'>$DNSStatus</span>" }
    
    # Add to Excel data
    $ExcelData.DomainControllers += [PSCustomObject]@{
        Hostname = $DC.HostName
        Site = $DC.Site
        IPAddress = $DC.IPv4Address
        OSVersion = $DC.OperatingSystem
        GlobalCatalog = $IsGC
        FSMORoles = $FSMORoles
        DNSService = $DNSStatus
    }
    
    $HTML += @"
        <tr>
            <td>$($DC.HostName)</td>
            <td>$($DC.Site)</td>
            <td>$($DC.IPv4Address)</td>
            <td>$($DC.OperatingSystem)</td>
            <td>$IsGCHTML</td>
            <td>$FSMORoles</td>
            <td>$DNSStatusHTML</td>
        </tr>
"@
}
$HTML += "    </table>`n</div>`n"

# ... rest of the script unchanged ...
# (Keep the rest of the script exactly as in the original file)
# Save HTML report
$HTML | Out-File -FilePath $ReportFile -Encoding UTF8

Write-Host "`nHTML Documentation Complete!" -ForegroundColor Green
Write-Host "Report saved to: $ReportFile" -ForegroundColor Yellow

# Export to Excel using ImportExcel module or COM object
Write-Host "`nExporting data to Excel..." -ForegroundColor Cyan

# Try using ImportExcel module first (if available)
if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel
    
    # Export each dataset to a separate worksheet
    $ExcelData.DomainControllers | Export-Excel -Path $ExcelFile -WorksheetName "Domain Controllers" -AutoSize -TableName "DomainControllers" -TableStyle Medium2
    $ExcelData.ReplicationHealth | Export-Excel -Path $ExcelFile -WorksheetName "Replication Health" -AutoSize -TableName "ReplicationHealth" -TableStyle Medium2
    $ExcelData.Sites | Export-Excel -Path $ExcelFile -WorksheetName "Sites and Subnets" -AutoSize -TableName "Sites" -TableStyle Medium2
    $ExcelData.DNSZones | Export-Excel -Path $ExcelFile -WorksheetName "DNS Zones" -AutoSize -TableName "DNSZones" -TableStyle Medium2
    $ExcelData.DHCPScopes | Export-Excel -Path $ExcelFile -WorksheetName "DHCP Scopes" -AutoSize -TableName "DHCPScopes" -TableStyle Medium2
    $ExcelData.Tier0Accounts | Export-Excel -Path $ExcelFile -WorksheetName "Tier 0 Accounts" -AutoSize -TableName "Tier0" -TableStyle Medium2
    $ExcelData.Tier1Accounts | Export-Excel -Path $ExcelFile -WorksheetName "Tier 1 Accounts" -AutoSize -TableName "Tier1" -TableStyle Medium2
    $ExcelData.Tier2Accounts | Export-Excel -Path $ExcelFile -WorksheetName "Tier 2 Accounts" -AutoSize -TableName "Tier2" -TableStyle Medium2
    $ExcelData.ServiceAccounts | Export-Excel -Path $ExcelFile -WorksheetName "Service Accounts" -AutoSize -TableName "ServiceAccounts" -TableStyle Medium2
    $ExcelData.ExchangeServers | Export-Excel -Path $ExcelFile -WorksheetName "Exchange Servers" -AutoSize -TableName "ExchangeServers" -TableStyle Medium2
    $ExcelData.GroupPolicies | Export-Excel -Path $ExcelFile -WorksheetName "Group Policies" -AutoSize -TableName "GroupPolicies" -TableStyle Medium2
    
    Write-Host "Excel file created successfully using ImportExcel module!" -ForegroundColor Green
    Write-Host "Excel file saved to: $ExcelFile" -ForegroundColor Yellow
} else {
    # Fallback to CSV exports
    Write-Host "ImportExcel module not found. Exporting to CSV files instead..." -ForegroundColor Yellow
    Write-Host "To create Excel files, install the module with: Install-Module -Name ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    
    $ExcelData.DomainControllers | Export-Csv "$OutputPath\DomainControllers_$Timestamp.csv" -NoTypeInformation
    $ExcelData.ReplicationHealth | Export-Csv "$OutputPath\ReplicationHealth_$Timestamp.csv" -NoTypeInformation
    $ExcelData.Sites | Export-Csv "$OutputPath\Sites_$Timestamp.csv" -NoTypeInformation
    $ExcelData.DNSZones | Export-Csv "$OutputPath\DNSZones_$Timestamp.csv" -NoTypeInformation
    $ExcelData.DHCPScopes | Export-Csv "$OutputPath\DHCPScopes_$Timestamp.csv" -NoTypeInformation
    $ExcelData.Tier0Accounts | Export-Csv "$OutputPath\Tier0Accounts_$Timestamp.csv" -NoTypeInformation
    $ExcelData.Tier1Accounts | Export-Csv "$OutputPath\Tier1Accounts_$Timestamp.csv" -NoTypeInformation
    $ExcelData.Tier2Accounts | Export-Csv "$OutputPath\Tier2Accounts_$Timestamp.csv" -NoTypeInformation
    $ExcelData.ServiceAccounts | Export-Csv "$OutputPath\ServiceAccounts_$Timestamp.csv" -NoTypeInformation
    $ExcelData.ExchangeServers | Export-Csv "$OutputPath\ExchangeServers_$Timestamp.csv" -NoTypeInformation
    $ExcelData.GroupPolicies | Export-Csv "$OutputPath\GroupPolicies_$Timestamp.csv" -NoTypeInformation
    
    Write-Host "CSV files created successfully!" -ForegroundColor Green
    Write-Host "CSV files saved to: $OutputPath" -ForegroundColor Yellow
}

Write-Host "`nOpening HTML report in default browser..." -ForegroundColor Cyan
Start-Process $ReportFile

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "Documentation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "All files saved to: $OutputPath" -ForegroundColor Yellow
