#
# Quickbase Report.ps1 - Quickbase API for Reports
#


$Log_MaskableKeys = @(
    # Put a comma-separated list of attribute names here, whose value should be masked before 
    'Password',
    "proxy_password"
)

#
# System functions
#
function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log info "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"

    if ($Connection) {
        @(
            @{
                name = 'hostname'
                type = 'textbox'
                label = 'Hostname'
                description = 'Hostname for Web Services'
                value = 'customer.quickbase.com'
            }
            @{
                name = 'report_id'
                type = 'textbox'
                label = 'Report ID'
                description = 'ID for the report to retrieve'
                value = ''
            }
            @{
                name = 'app_id'
                type = 'textbox'
                label = 'App ID'
                description = 'App ID to target'
                value = ''
            }
            @{
                name = 'token'
                type = 'textbox'
                password = $true
                label = 'Token'
                label_indent = $true
                description = 'API Token'
                value = ''
            }
            @{
                name = 'use_proxy'
                type = 'checkbox'
                label = 'Use Proxy'
                description = 'Use Proxy server for requets'
                value = $false                  # Default value of checkbox item
            }
            @{
                name = 'proxy_address'
                type = 'textbox'
                label = 'Proxy Address'
                description = 'Address of the proxy server'
                value = 'http://localhost:8888'
                disabled = '!use_proxy'
                hidden = '!use_proxy'
            }
            @{
                name = 'use_proxy_credentials'
                type = 'checkbox'
                label = 'Use Proxy'
                description = 'Use Proxy server for requets'
                value = $false
                disabled = '!use_proxy'
                hidden = '!use_proxy'
            }
            @{
                name = 'proxy_username'
                type = 'textbox'
                label = 'Proxy Username'
                label_indent = $true
                description = 'Username account'
                value = ''
                disabled = '!use_proxy_credentials'
                hidden = '!use_proxy_credentials'
            }
            @{
                name = 'proxy_password'
                type = 'textbox'
                password = $true
                label = 'Proxy Password'
                label_indent = $true
                description = 'User account password'
                value = ''
                disabled = '!use_proxy_credentials'
                hidden = '!use_proxy_credentials'
            }
            @{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                description = ''
                value = 1
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                description = ''
                value = 1
            }
        )
    }

    if ($TestConnection) {
        
    }

    if ($Configuration) {
        @()
    }

    Log info "Done"
}

function Idm-OnUnload {
}

#
# Object CRUD functions
#

$ColumnInfoCache = @{}
$TableInfoCache = @{}


function Idm-Dispatcher {
    param (
        # Optional Class/Operation
        [string] $Class,
        [string] $Operation,
        # Mode
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )
    
    Log info "-Class='$Class' -Operation='$Operation' -GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    $ConnectionParams = ConvertFrom-Json2 $SystemParams

    if ($Class -eq '') {

        if ($GetMeta) {
            Get-QuickbaseTableData -ConnectionParams $ConnectionParams
        }
        else {
            # Purposely no-operation.
        }

    }
    else {
        
        if ($GetMeta) {
            $Global:ColumnInfoCache[$Class] = Get-QuickbaseColumnData -ConnectionParams $ConnectionParams -Class $Class
            $Global:ColumnInfoCache[$Class]

        }
        else {
            if (! $Global:TableInfoCache[$Class]) {
                Get-QuickbaseTableData -ConnectionParams $ConnectionParams > $null
            }

            $classInfo = $Global:TableInfoCache[$Class]
            $results = [System.Collections.ArrayList]@()

            while($true) {
                $result = Execute-QuickbaseRequest -SystemParams $ConnectionParams -Method "POST" -EndpointUri "/v1/reports/$($classInfo.ID)/run?tableId=$($classInfo.TableID)" -Body $null

                foreach($item in $result.data) {
                    [void]$results.Add($item)
                }

                if($result.metadata.totalRecords -eq $null -or  $results.count -ge $result.metadata.totalRecords ) {
                    LogIO info "Retrieved [$($results.count)] records"
                    break
                }
                LogIO info "Retrieved [$($results.count)] records, getting next page" -f $results.count
            }

            $columns = $result.fields | Group-Object -Property id
            $hash_table = [ordered]@{}
            $columnMap = [ordered]@{}

            foreach ($column in $columns) {
                $label = $column.Group.label -replace '\s+', ''
                $hash_table[$label] = ""
                $columnMap[$column.Name] = $label
            }

            foreach($row in $result.data) {
                $obj = New-Object -TypeName PSObject -Property $hash_table    
                
                foreach ($property in $row.PSObject.Properties) {
                    $obj.($columnMap[$property.Name]) = $property.Value.value
                }
                
                $obj
            }
        }
    }

    Log info "Done"
}

function Get-QuickbaseColumnData {
    param(
        [hashtable] $ConnectionParams,
        [string] $Class
    )
    
    if (! $Global:TableInfoCache[$Class]) {
        Get-QuickbaseTableData -ConnectionParams $ConnectionParams
    }

    $classInfo = $Global:TableInfoCache[$Class]

    if($classInfo.Type -eq 'Table') {
        $columns = Execute-QuickbaseRequest -SystemParams $ConnectionParams -EndpointUri "/v1/fields" -Body @{ "tableId"= $classInfo.ID }
    
        @(
            @{
                name = 'selected_columns'
                type = 'grid'
                label = 'Include columns'
                tooltip = 'Selected columns'
                table = @{
                    rows = @($columns | ForEach-Object {
                        @{
                            name = $_.label
                            config = @(
                                if ($_.properties.primaryKey) { 'Primary key' }
                            ) -join ' | '
                        }
                    })
                    settings_grid = @{
                        selection = 'multiple'
                        key_column = 'name'
                        checkbox = $true
                        filter = $true
                        columns = @(
                            @{
                                name = 'name'
                                display_name = 'Name'
                            }
                            @{
                                name = 'config'
                                display_name = 'Configuration'
                            }
                        )
                    }
                }
                value = @($columns | ForEach-Object { $_.name })
            }
        )
    } elseif($classInfo.Type -eq 'Report') {
        $columns = (Execute-QuickbaseRequest -SystemParams $ConnectionParams -Method "POST" -EndpointUri "/v1/reports/$($classInfo.ID)/run?top=1&tableId=$($classInfo.TableID)" -Body $null).fields
    
        @(
            @{
                name = 'selected_columns'
                type = 'grid'
                label = 'Include columns'
                tooltip = 'Selected columns'
                table = @{
                    rows = @($columns | ForEach-Object {
                        @{
                            id = $_.id
                            name = $_.label
                            config = @(
                            ) -join ' | '
                        }
                    })
                    settings_grid = @{
                        selection = 'multiple'
                        key_column = 'name'
                        checkbox = $true
                        filter = $true
                        columns = @(
                            @{
                                name = 'name'
                                display_name = 'Name'
                            }
                            @{
                                name = 'config'
                                display_name = 'Configuration'
                            }
                        )
                    }
                }
                value = @($columns | ForEach-Object { $_.name })
            }
        )
    }
}

function Get-QuickbaseTableData {
    param(
        [hashtable] $ConnectionParams
    )
    $tables = Execute-QuickbaseRequest -SystemParams $ConnectionParams -EndpointUri "/v1/tables" -Body @{ "appId"= $ConnectionParams.app_id }

    foreach($row in $tables) {

        <# Table access not available in development environment, will provide in future version
        $table_item = [ordered]@{
            Class = $row.name
            Operation = 'Read'
            'Type' = 'Table'
            'Key' = $row.keyFieldId
            'ID' = $row.id
        }
        $Global:TableInfoCache[$row.Name] = $table_item
        $table_item
        #>

        $reports = Execute-QuickbaseRequest -SystemParams $ConnectionParams -EndpointUri "/v1/reports" -Body @{ "tableId"= $row.id }

        foreach($report in $reports) {
            $className = "{0}_{1}" -f ($row.name -replace "\s" ,""), ($report.name -replace "\s","")

            $report_item = [ordered]@{
                Class = $className
                Operation = 'Read'
                'Type' = 'Report'
                'Key' = ''
                'ID' = $report.id
                'TableID' = $row.id
            }
            $Global:TableInfoCache[$className] = $report_item
            $report_item
        }

    }
}

function Execute-QuickbaseRequest {
    param (
        [hashtable] $SystemParams,
        [string] $EndpointUri,
        [hashtable] $Body,
        [string] $Method = "GET"
    )
    $uri = "https://api.quickbase.com{0}" -f $EndpointUri

    $headers = @{
        "Content-Type" = "application/json"
        "Authorization" = "QB-USER-TOKEN {0}" -f $SystemParams.token
        "qb-realm-hostname" = $SystemParams.hostname
    }

    try {
		$splat = @{
            Method = $Method
            Uri = $uri
            Headers = $headers
            Body = $Body
        }
        
        if($SystemParams.use_proxy)
        {
            Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
            $splat["Proxy"] = $SystemParams.proxy_address

            if($SystemParams.use_proxy_credentials)
            {
                $splat["proxyCredential"] = New-Object System.Management.Automation.PSCredential ($SystemParams.proxy_username, (ConvertTo-SecureString $SystemParams.proxy_password -AsPlainText -Force) )
            }
        }

        $result = Invoke-RestMethod @splat -ErrorAction Stop
        
	}
	catch [System.Net.WebException] {
        $message = "Error : $($_)"
        Log error $message
        Write-Error $_
	}
    catch {
        $message = "Error : $($_)"
        Log error $message
        Write-Error $_
    }
    finally {
        Write-Output $result
    }
}
