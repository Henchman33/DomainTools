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
$Forest = Get-ADForest
$HTML += @"
    <h2>1. Forest Overview</h2>
    <div class="section-content">
    <div class="info-box success">
        <table>
            <tr><th>Property</th><th>Value</th></tr>
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

# 3. REPLICATION HEALTH
Write-Host "Checking Replication Health..." -ForegroundColor Cyan
$HTML += @"
    <h2>3. Active Directory Replication Health</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>Source DC</th>
            <th>Destination DC</th>
            <th>Last Replication</th>
            <th>Status</th>
            <th>Failures</th>
        </tr>
"@

foreach ($DC in $DCs) {
    try {
        $ReplPartners = Get-ADReplicationPartnerMetadata -Target $DC.HostName -ErrorAction SilentlyContinue
        foreach ($Partner in $ReplPartners) {
            $LastRepl = $Partner.LastReplicationSuccess
            $TimeSince = (Get-Date) - $LastRepl
            $Status = if ($TimeSince.TotalHours -lt 24) { "Healthy" } else { "Warning" }
            $StatusHTML = if ($TimeSince.TotalHours -lt 24) { "<span class='healthy'>Healthy</span>" } else { "<span class='critical'>Warning</span>" }
            $Failures = $Partner.ConsecutiveReplicationFailures
            $FailuresHTML = if ($Partner.ConsecutiveReplicationFailures -eq 0) { "<span class='healthy'>0</span>" } else { "<span class='critical'>$($Partner.ConsecutiveReplicationFailures)</span>" }
            
            # Add to Excel data
            $ExcelData.ReplicationHealth += [PSCustomObject]@{
                SourceDC = $Partner.Partner
                DestinationDC = $DC.HostName
                LastReplication = $LastRepl.ToString("yyyy-MM-dd HH:mm:ss")
                Status = $Status
                Failures = $Failures
            }
            
            $HTML += @"
        <tr>
            <td>$($Partner.Partner)</td>
            <td>$($DC.HostName)</td>
            <td>$($LastRepl.ToString("yyyy-MM-dd HH:mm:ss"))</td>
            <td>$StatusHTML</td>
            <td>$FailuresHTML</td>
        </tr>
"@
        }
    } catch {
        $HTML += @"
        <tr>
            <td colspan="5"><span class='critical'>Unable to query replication data for $($DC.HostName)</span></td>
        </tr>
"@
    }
}
$HTML += "    </table>`n</div>`n"

# 4. SITES AND SUBNETS
Write-Host "Gathering Sites and Subnets..." -ForegroundColor Cyan
$Sites = Get-ADReplicationSite -Filter *
$HTML += @"
    <h2>4. Active Directory Sites (Total: $($Sites.Count))</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>Site Name</th>
            <th>Description</th>
            <th>Subnets</th>
            <th>Domain Controllers</th>
        </tr>
"@

foreach ($Site in $Sites) {
    $Subnets = Get-ADReplicationSubnet -Filter "Site -eq '$($Site.DistinguishedName)'" | Select-Object -ExpandProperty Name
    $SubnetList = if ($Subnets) { ($Subnets -join ", ") } else { "None configured" }
    $SubnetListHTML = if ($Subnets) { ($Subnets -join "<br>") } else { "None configured" }
    
    $SiteDCs = $DCs | Where-Object { $_.Site -eq $Site.Name } | Select-Object -ExpandProperty HostName
    $DCList = if ($SiteDCs) { ($SiteDCs -join ", ") } else { "None" }
    $DCListHTML = if ($SiteDCs) { ($SiteDCs -join "<br>") } else { "None" }
    
    # Add to Excel data
    $ExcelData.Sites += [PSCustomObject]@{
        SiteName = $Site.Name
        Description = $Site.Description
        Subnets = $SubnetList
        DomainControllers = $DCList
    }
    
    $HTML += @"
        <tr>
            <td>$($Site.Name)</td>
            <td>$($Site.Description)</td>
            <td>$SubnetListHTML</td>
            <td>$DCListHTML</td>
        </tr>
"@
}
$HTML += "    </table>`n</div>`n"

# 5. DNS SERVERS
Write-Host "Gathering DNS Information..." -ForegroundColor Cyan
$HTML += @"
    <h2>5. DNS Servers on Domain Controllers</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>Server</th>
            <th>DNS Zones</th>
            <th>Zone Type</th>
            <th>Dynamic Updates</th>
        </tr>
"@

foreach ($DC in $DCs) {
    try {
        $DNSZones = Get-DnsServerZone -ComputerName $DC.HostName -ErrorAction SilentlyContinue
        if ($DNSZones) {
            foreach ($Zone in $DNSZones | Where-Object { $_.ZoneType -ne "Cache" }) {
                # Add to Excel data
                $ExcelData.DNSZones += [PSCustomObject]@{
                    Server = $DC.HostName
                    ZoneName = $Zone.ZoneName
                    ZoneType = $Zone.ZoneType
                    DynamicUpdate = $Zone.DynamicUpdate
                }
                
                $HTML += @"
        <tr>
            <td>$($DC.HostName)</td>
            <td>$($Zone.ZoneName)</td>
            <td>$($Zone.ZoneType)</td>
            <td>$($Zone.DynamicUpdate)</td>
        </tr>
"@
            }
        }
    } catch {
        $HTML += @"
        <tr>
            <td>$($DC.HostName)</td>
            <td colspan="3"><span class='critical'>Unable to query DNS zones</span></td>
        </tr>
"@
    }
}
$HTML += "    </table>`n</div>`n"

# 6. DHCP SERVERS
Write-Host "Gathering DHCP Server Information..." -ForegroundColor Cyan
try {
    $DHCPServers = Get-DhcpServerInDC -ErrorAction Stop
    $HTML += @"
    <h2>6. DHCP Servers (Total: $($DHCPServers.Count))</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>Server Name</th>
            <th>IP Address</th>
            <th>Scopes</th>
            <th>Scope Range</th>
            <th>Scope State</th>
        </tr>
"@
    
    foreach ($DHCPServer in $DHCPServers) {
        try {
            $Scopes = Get-DhcpServerv4Scope -ComputerName $DHCPServer.DnsName -ErrorAction SilentlyContinue
            if ($Scopes) {
                foreach ($Scope in $Scopes) {
                    $ScopeState = $Scope.State
                    $ScopeStateHTML = if ($Scope.State -eq "Active") { "<span class='healthy'>Active</span>" } else { "<span class='critical'>$($Scope.State)</span>" }
                    
                    # Add to Excel data
                    $ExcelData.DHCPScopes += [PSCustomObject]@{
                        ServerName = $DHCPServer.DnsName
                        IPAddress = $DHCPServer.IPAddress
                        ScopeName = $Scope.Name
                        StartRange = $Scope.StartRange
                        EndRange = $Scope.EndRange
                        State = $ScopeState
                    }
                    
                    $HTML += @"
        <tr>
            <td>$($DHCPServer.DnsName)</td>
            <td>$($DHCPServer.IPAddress)</td>
            <td>$($Scope.Name)</td>
            <td>$($Scope.StartRange) - $($Scope.EndRange)</td>
            <td>$ScopeStateHTML</td>
        </tr>
"@
                }
            } else {
                $HTML += @"
        <tr>
            <td>$($DHCPServer.DnsName)</td>
            <td>$($DHCPServer.IPAddress)</td>
            <td colspan="3">No scopes configured or unable to query</td>
        </tr>
"@
            }
        } catch {
            $HTML += @"
        <tr>
            <td>$($DHCPServer.DnsName)</td>
            <td>$($DHCPServer.IPAddress)</td>
            <td colspan="3"><span class='critical'>Unable to query scopes</span></td>
        </tr>
"@
        }
    }
    $HTML += "    </table>`n</div>`n"
} catch {
    $HTML += @"
    <h2>6. DHCP Servers</h2>
    <div class="section-content">
    <div class="info-box warning">
        <p><span class='critical'>Unable to retrieve DHCP servers from Active Directory.</span></p>
        <p>This may be because no DHCP servers are authorized or you lack permissions.</p>
    </div>
    </div>
"@
}

# 7. PRIVILEGED ACCOUNTS - TIER 0
Write-Host "Gathering Tier 0 Privileged Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>7. Tier 0 Accounts (Highest Privilege - Domain/Enterprise/Schema Admins)</h2>
    <div class="section-content">
"@

$Tier0Groups = @(
    "Domain Admins",
    "Enterprise Admins",
    "Schema Admins",
    "Administrators"
)

$HTML += "<table><tr><th>Group</th><th>Member Name</th><th>Account Type</th><th>Enabled</th><th>Last Logon</th><th>Password Last Set</th></tr>"

foreach ($GroupName in $Tier0Groups) {
    try {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
        if ($Group) {
            $Members = Get-ADGroupMember -Identity $Group -ErrorAction SilentlyContinue
            foreach ($Member in $Members) {
                try {
                    if ($Member.objectClass -eq "user") {
                        $User = Get-ADUser -Identity $Member.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                        $EnabledStatus = $User.Enabled
                        $EnabledStatusHTML = if ($User.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
                        $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                        $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                        
                        # Add to Excel data
                        $ExcelData.Tier0Accounts += [PSCustomObject]@{
                            Group = $GroupName
                            MemberName = $User.Name
                            AccountType = "User"
                            Enabled = $EnabledStatus
                            LastLogon = $LastLogon
                            PasswordLastSet = $PwdLastSet
                        }
                        
                        $HTML += "<tr class='tier0'><td>$GroupName</td><td>$($User.Name)</td><td>User</td><td>$EnabledStatusHTML</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                    } else {
                        $ExcelData.Tier0Accounts += [PSCustomObject]@{
                            Group = $GroupName
                            MemberName = $Member.Name
                            AccountType = $Member.objectClass
                            Enabled = "N/A"
                            LastLogon = "N/A"
                            PasswordLastSet = "N/A"
                        }
                        
                        $HTML += "<tr class='tier0'><td>$GroupName</td><td>$($Member.Name)</td><td>$($Member.objectClass)</td><td>N/A</td><td>N/A</td><td>N/A</td></tr>"
                    }
                } catch {
                    $HTML += "<tr class='tier0'><td>$GroupName</td><td>$($Member.Name)</td><td>Unknown</td><td colspan='3'>Error querying</td></tr>"
                }
            }
        }
    } catch {
        $HTML += "<tr class='tier0'><td>$GroupName</td><td colspan='5'><span class='critical'>Unable to query group</span></td></tr>"
    }
}
$HTML += "</table></div>"

# 8. TIER 1 ACCOUNTS
Write-Host "Gathering Tier 1 Privileged Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>8. Tier 1 Accounts (Server/Infrastructure Management)</h2>
    <div class="section-content">
"@

$Tier1Groups = @(
    "Server Operators",
    "Backup Operators",
    "Account Operators",
    "Print Operators"
)

$HTML += "<table><tr><th>Group</th><th>Member Name</th><th>Account Type</th><th>Enabled</th><th>Last Logon</th><th>Password Last Set</th></tr>"

foreach ($GroupName in $Tier1Groups) {
    try {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
        if ($Group) {
            $Members = Get-ADGroupMember -Identity $Group -ErrorAction SilentlyContinue
            if ($Members) {
                foreach ($Member in $Members) {
                    try {
                        if ($Member.objectClass -eq "user") {
                            $User = Get-ADUser -Identity $Member.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                            $EnabledStatus = $User.Enabled
                            $EnabledStatusHTML = if ($User.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
                            $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                            $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                            
                            # Add to Excel data
                            $ExcelData.Tier1Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $User.Name
                                AccountType = "User"
                                Enabled = $EnabledStatus
                                LastLogon = $LastLogon
                                PasswordLastSet = $PwdLastSet
                            }
                            
                            $HTML += "<tr class='tier1'><td>$GroupName</td><td>$($User.Name)</td><td>User</td><td>$EnabledStatusHTML</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                        } else {
                            $ExcelData.Tier1Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $Member.Name
                                AccountType = $Member.objectClass
                                Enabled = "N/A"
                                LastLogon = "N/A"
                                PasswordLastSet = "N/A"
                            }
                            
                            $HTML += "<tr class='tier1'><td>$GroupName</td><td>$($Member.Name)</td><td>$($Member.objectClass)</td><td>N/A</td><td>N/A</td><td>N/A</td></tr>"
                        }
                    } catch {
                        $HTML += "<tr class='tier1'><td>$GroupName</td><td>$($Member.Name)</td><td>Unknown</td><td colspan='3'>Error querying</td></tr>"
                    }
                }
            } else {
                $HTML += "<tr class='tier1'><td>$GroupName</td><td colspan='5'>No members</td></tr>"
            }
        }
    } catch {
        $HTML += "<tr class='tier1'><td>$GroupName</td><td colspan='5'><span class='critical'>Unable to query group</span></td></tr>"
    }
}
$HTML += "</table></div>"

# 9. TIER 2 ACCOUNTS
Write-Host "Gathering Tier 2 Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>9. Tier 2 Accounts (User/Workstation Management)</h2>
    <div class="section-content">
    <div class="info-box">
        <p><strong>Note:</strong> Tier 2 typically includes help desk and user support groups. Common groups are listed below. Customize the script to add organization-specific Tier 2 groups.</p>
    </div>
"@

$Tier2Groups = @(
    "Help Desk",
    "Helpdesk Operators",
    "Desktop Support",
    "Remote Desktop Users"
)

$HTML += "<table><tr><th>Group</th><th>Member Name</th><th>Account Type</th><th>Enabled</th><th>Last Logon</th><th>Password Last Set</th></tr>"

$Tier2Found = $false
foreach ($GroupName in $Tier2Groups) {
    try {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
        if ($Group) {
            $Tier2Found = $true
            $Members = Get-ADGroupMember -Identity $Group -ErrorAction SilentlyContinue
            if ($Members) {
                foreach ($Member in $Members) {
                    try {
                        if ($Member.objectClass -eq "user") {
                            $User = Get-ADUser -Identity $Member.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                            $EnabledStatus = $User.Enabled
                            $EnabledStatusHTML = if ($User.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
                            $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                            $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                            
                            # Add to Excel data
                            $ExcelData.Tier2Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $User.Name
                                AccountType = "User"
                                Enabled = $EnabledStatus
                                LastLogon = $LastLogon
                                PasswordLastSet = $PwdLastSet
                            }
                            
                            $HTML += "<tr class='tier2'><td>$GroupName</td><td>$($User.Name)</td><td>User</td><td>$EnabledStatusHTML</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                        } else {
                            $ExcelData.Tier2Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $Member.Name
                                AccountType = $Member.objectClass
                                Enabled = "N/A"
                                LastLogon = "N/A"
                                PasswordLastSet = "N/A"
                            }
                            
                            $HTML += "<tr class='tier2'><td>$GroupName</td><td>$($Member.Name)</td><td>$($Member.objectClass)</td><td>N/A</td><td>N/A</td><td>N/A</td></tr>"
                        }
                    } catch {
                        $HTML += "<tr class='tier2'><td>$GroupName</td><td>$($Member.Name)</td><td>Unknown</td><td colspan='3'>Error querying</td></tr>"
                    }
                }
            } else {
                $HTML += "<tr class='tier2'><td>$GroupName</td><td colspan='5'>No members</td></tr>"
            }
        }
    } catch {
        continue
    }
}

if (-not $Tier2Found) {
    $HTML += "<tr class='tier2'><td colspan='6'>No standard Tier 2 groups found. Please customize the script with your organization's Tier 2 group names.</td></tr>"
}

$HTML += "</table></div>"

# 10. SERVICE ACCOUNTS
Write-Host "Gathering Service Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>10. Service Accounts</h2>
    <div class="section-content">
"@

$ServiceAccounts = Get-ADUser -Filter * -Properties ServicePrincipalName, Description, PasswordLastSet, LastLogonDate, Enabled | 
    Where-Object { 
        ($_.ServicePrincipalName -ne $null) -or 
        ($_.SamAccountName -like "svc_*") -or 
        ($_.SamAccountName -like "svc-*") -or
        ($_.SamAccountName -like "*service*") -or
        ($_.Description -like "*service account*")
    }

$HTML += @"
    <table>
        <tr>
            <th>Account Name</th>
            <th>Description</th>
            <th>Enabled</th>
            <th>Password Last Set</th>
            <th>Last Logon</th>
            <th>SPN Count</th>
        </tr>
"@

foreach ($SvcAcct in $ServiceAccounts) {
    $EnabledStatus = $SvcAcct.Enabled
    $EnabledStatusHTML = if ($SvcAcct.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
    $PwdLastSet = if ($SvcAcct.PasswordLastSet) { $SvcAcct.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
    $LastLogon = if ($SvcAcct.LastLogonDate) { $SvcAcct.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
    $SPNCount = if ($SvcAcct.ServicePrincipalName) { $SvcAcct.ServicePrincipalName.Count } else { 0 }
    
    # Add to Excel data
    $ExcelData.ServiceAccounts += [PSCustomObject]@{
        AccountName = $SvcAcct.SamAccountName
        Description = $SvcAcct.Description
        Enabled = $EnabledStatus
        PasswordLastSet = $PwdLastSet
        LastLogon = $LastLogon
        SPNCount = $SPNCount
    }
    
    $HTML += @"
        <tr>
            <td>$($SvcAcct.SamAccountName)</td>
            <td>$($SvcAcct.Description)</td>
            <td>$EnabledStatusHTML</td>
            <td>$PwdLastSet</td>
            <td>$LastLogon</td>
            <td>$SPNCount</td>
        </tr>
"@
}

$HTML += "</table></div>"

# 11. EXCHANGE SERVERS
Write-Host "Gathering Exchange Server Information..." -ForegroundColor Cyan
$HTML += @"
    <h2>11. Exchange Servers</h2>
    <div class="section-content">
"@

try {
    $ConfigNC = (Get-ADRootDSE).configurationNamingContext
    $ExchangeServers = Get-ADObject -Filter {objectClass -eq "msExchExchangeServer"} -SearchBase $ConfigNC -Properties Name, msExchServerSite, serialNumber, versionNumber, msExchCurrentServerRoles, networkAddress, whenCreated
    
    if ($ExchangeServers) {
        $HTML += @"
        <table>
            <tr>
                <th>Server Name</th>
                <th>Site</th>
                <th>Roles</th>
                <th>Version</th>
                <th>FQDN</th>
                <th>Created</th>
            </tr>
"@
        
        foreach ($ExchServer in $ExchangeServers) {
            $Roles = switch ($ExchServer.msExchCurrentServerRoles) {
                2 { "Mailbox" }
                4 { "Client Access" }
                16 { "Unified Messaging" }
                32 { "Hub Transport" }
                64 { "Edge Transport" }
                54 { "Mailbox, Client Access, Hub Transport" }
                default { $ExchServer.msExchCurrentServerRoles }
            }
            
            $SiteName = if ($ExchServer.msExchServerSite) {
                ($ExchServer.msExchServerSite -split ",")[0] -replace "CN=", ""
            } else {
                "Unknown"
            }
            
            $FQDN = "N/A"
            if ($ExchServer.networkAddress) {
                $FQDN = ($ExchServer.networkAddress | Where-Object { $_ -like "ncacn_ip_tcp:*" }) -replace "ncacn_ip_tcp:", ""
            }
            
            $Version = "Unknown"
            if ($ExchServer.serialNumber) {
                $VersionNumber = $ExchServer.serialNumber
                if ($VersionNumber -like "Version 15.2*") { $Version = "Exchange 2019" }
                elseif ($VersionNumber -like "Version 15.1*") { $Version = "Exchange 2016" }
                elseif ($VersionNumber -like "Version 15.0*") { $Version = "Exchange 2013" }
                elseif ($VersionNumber -like "Version 14.*") { $Version = "Exchange 2010" }
                else { $Version = $VersionNumber }
            }
            
            # Add to Excel data
            $ExcelData.ExchangeServers += [PSCustomObject]@{
                ServerName = $ExchServer.Name
                Site = $SiteName
                Roles = $Roles
                Version = $Version
                FQDN = $FQDN
                Created = $ExchServer.whenCreated.ToString("yyyy-MM-dd")
            }
            
            $HTML += @"
            <tr>
                <td>$($ExchServer.Name)</td>
                <td>$SiteName</td>
                <td>$Roles</td>
                <td>$Version</td>
                <td>$FQDN</td>
                <td>$($ExchServer.whenCreated.ToString("yyyy-MM-dd"))</td>
            </tr>
"@
        }
        
        $HTML += "</table>"
    } else {
        $HTML += @"
        <div class="info-box">
            <p>No Exchange servers found in Active Directory.</p>
        </div>
"@
    }
} catch {
    $HTML += @"
    <div class="info-box warning">
        <p><span class='critical'>Unable to query Exchange servers from Active Directory.</span></p>
        <p>Error: $($_.Exception.Message)</p>
    </div>
"@
}

$HTML += "</div>"

# 12. GROUP POLICIES
Write-Host "Gathering Group Policy Information..." -ForegroundColor Cyan
$HTML += @"
    <h2>12. Group Policy Objects</h2>
    <div class="section-content">
"@

try {
    $GPOs = Get-GPO -All
    $HTML += @"
    <table>
        <tr>
            <th>GPO Name</th>
            <th>Status</th>
            <th>Created</th>
            <th>Modified</th>
            <th>User Version</th>
            <th>Computer Version</th>
        </tr>
"@
    
    foreach ($GPO in $GPOs) {
        $Status = $GPO.GpoStatus
        $StatusHTML = if ($Status -eq "AllSettingsEnabled") { 
            "<span class='healthy'>Enabled</span>" 
        } elseif ($Status -like "*Disabled") { 
            "<span class='critical'>$Status</span>" 
        } else { 
            $Status 
        }
        
        # Add to Excel data
        $ExcelData.GroupPolicies += [PSCustomObject]@{
            GPOName = $GPO.DisplayName
            Status = $Status
            Created = $GPO.CreationTime.ToString("yyyy-MM-dd HH:mm")
            Modified = $GPO.ModificationTime.ToString("yyyy-MM-dd HH:mm")
            UserVersion = $GPO.User.DSVersion
            ComputerVersion = $GPO.Computer.DSVersion
        }
        
        $HTML += @"
        <tr>
            <td>$($GPO.DisplayName)</td>
            <td>$StatusHTML</td>
            <td>$($GPO.CreationTime.ToString("yyyy-MM-dd HH:mm"))</td>
            <td>$($GPO.ModificationTime.ToString("yyyy-MM-dd HH:mm"))</td>
            <td>$($GPO.User.DSVersion)</td>
            <td>$($GPO.Computer.DSVersion)</td>
        </tr>
"@
    }
    
    $HTML += "</table>"
} catch {
    $HTML += @"
    <div class="info-box warning">
        <p><span class='critical'>Unable to retrieve Group Policy Objects.</span></p>
        <p>Error: $($_.Exception.Message)</p>
    </div>
"@
}

$HTML += "</div>"

# 13. FSMO ROLES SUMMARY
Write-Host "Generating FSMO Roles Summary..." -ForegroundColor Cyan
$Domain = Get-ADDomain
$HTML += @"
    <h2>13. FSMO Roles Summary</h2>
    <div class="section-content">
    <div class="info-box">
        <h3>Forest-Wide Roles</h3>
        <table>
            <tr><th>Role</th><th>Holder</th></tr>
            <tr><td>Schema Master</td><td>$($Forest.SchemaMaster)</td></tr>
            <tr><td>Domain Naming Master</td><td>$($Forest.DomainNamingMaster)</td></tr>
        </table>
        
        <h3>Domain-Wide Roles</h3>
        <table>
            <tr><th>Role</th><th>Holder</th></tr>
            <tr><td>PDC Emulator</td><td>$($Domain.PDCEmulator)</td></tr>
            <tr><td>RID Master</td><td>$($Domain.RIDMaster)</td></tr>
            <tr><td>Infrastructure Master</td><td>$($Domain.InfrastructureMaster)</td></tr>
        </table>
    </div>
    </div>
"@

# Close HTML
$HTML += @"
    <div class="info-box" style="margin-top: 40px;">
        <p><strong>Documentation Complete</strong></p>
        <p>This report provides a comprehensive overview of your Active Directory infrastructure.</p>
        <p><strong>Files Generated:</strong></p>
        <ul>
            <li>HTML Report: $ReportFile</li>
            <li>Excel Report: $ExcelFile</li>
        </ul>
    </div>
</body>
</html>
"@

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
