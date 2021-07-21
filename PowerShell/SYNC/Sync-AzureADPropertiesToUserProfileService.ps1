#requries -Modules PnP.PowerShell
#requries -Modules Microsoft.Graph.Users

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

function Remove-PropertiesWithNullValue
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][Object[]]$Identity
    )

    foreach( $object in $Identity )
    {
        $newObject = [PSCustomObject] @{}

        foreach( $propertyName in $object.PSObject.Properties | Select-Object -ExpandProperty Name )
        {
            if( $object.$propertyName -ne $null )
            {
                $newObject | Add-Member -MemberType NoteProperty -Name $propertyName -Value $object.$propertyName
            }
        }

        $newObject
    }

}

function Get-UserAttributes
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][Object[]]$Identity,
        [Parameter(Mandatory=$true)][Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]$ImportProfilePropertiesUserIdType,
        [Parameter(Mandatory=$true)][string[]]$Attributes,
        [Parameter(Mandatory=$false)][string[]]$OnPremisesExtensionAttributes
    )

    begin
    {
    }
    process
    {
        foreach( $user in $Identity )
        {
            $object = [PSCustomObject] @{}

            switch( $ImportProfilePropertiesUserIdType )
            {
                ([Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::CloudId)
                {
                    $object | Add-Member -MemberType NoteProperty -Name "Id" -Value $user.Id
                }
                ([Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::Email)
                {
                    $object | Add-Member -MemberType NoteProperty -Name "mail" -Value $user.mail
                }
                ([Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::PrincipalName)
                {
                    $object | Add-Member -MemberType NoteProperty -Name "userPrincipalName" -Value $user.userPrincipalName
                }
            }

            foreach( $attribute in $Attributes )
            {
                if( -not ($object.psobject.Properties | Where-Object -Property Name -eq $attribute) )
                {
                    $object | Add-Member -MemberType NoteProperty -Name $attribute -Value $user.$attribute
                }
            }

            foreach( $onPremisesExtensionAttribute in $OnPremisesExtensionAttributes )
            {
                if( -not ($object.psobject.properties | Where-Object -Property Name -eq $onPremisesExtensionAttribute) )
                {
                    $object | Add-Member -MemberType NoteProperty -Name $onPremisesExtensionAttribute -Value $user.OnPremisesExtensionAttributes.$onPremisesExtensionAttribute
                }
            }

            $object
        }
    }
    end
    {
    }
}

function ConvertTo-UserProfileBulkImportJson
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][Object[]]$Identity
    )
    
    begin
    {
        <#
            REQUIRED JSON FORMAT
            {
                "value": [
                    {
                        "IdName":    "vesaj@contoso.com",
                        "Property1": "Helsinki",
                        "Property2": "Viper"
                    },
                    {
                        "IdName":    "bjansen@contoso.com",
                        "Property1": "Brussels",
                        "Property2": "Beetle"
                    }
                ]
            }
        #>

        $maxPropertiesPerJsonObject = 500000

        $mappings   = New-Object System.Collections.ArrayList
        $jsonChunks = New-Object System.Collections.ArrayList

        $counter = 0
    }
    process
    {
        foreach( $mapping in $Identity )
        {
            $mappingPropertyCountBefore = ($mapping.PSObject.Properties | Select-Object -ExpandProperty Name).Count

            $mapping = Remove-PropertiesWithNullValue -Identity $mapping

            $mappingPropertyCount = ($mapping.PSObject.Properties | Select-Object -ExpandProperty Name).Count

            if( $counter + $mappingPropertyCount -gt $maxPropertiesPerJsonObject )
            {
                # reached max properties per JSON file, convert rows and reset

                $mappings = Remove-PropertiesWithNullValue -Identity $mappings

                $jsonChunks += [PSCustomObject] @{ value = $mappings } | ConvertTo-Json -Depth 100

                $counter = 0
                
                $mappings.Clear()
            }
            else
            {
                # keep adding rows

                $counter += $mappingPropertyCount

                $null = $mappings.Add($mapping)
            }
        }
        
        $mappings = Remove-PropertiesWithNullValue -Identity $mappings

        # convert the remaining chunk
        $jsonChunks += [PSCustomObject] @{ value = $mappings } | ConvertTo-Json -Depth 100

        # return chunks of JSON
        return $jsonChunks
    }
    end
    {
    }
}

function Save-JsonFile
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string[]]$Json,
        [Parameter(Mandatory=$true)][string]$WebUrl,
        [Parameter(Mandatory=$true)][string]$FolderPath,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][string]$Thumbprint,
        [Parameter(Mandatory=$true)][string]$Tenant

    )
    
    begin
    {
        $fileNameFormat = "UserProfilePropertyBulkImportJson_{0}.txt"
    }
    process
    {
        $connection = Connect-PnPOnline -Url $WebUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant "$Tenant.onmicrosoft.com" -ReturnConnection

        $web = Get-PnPWeb -Includes Url -Connection $connection
        
        $uri = [System.Uri]::new($web.Url)

        foreach( $jsonString in $Json )
        {
            $fileName = $fileNameFormat -f (Get-Date -Format FileDateTime)
            
            $bytes = [System.Text.Encoding]::UTF8.GetPreamble() + [System.Text.Encoding]::UTF8.GetBytes( $jsonString )    
            
            $stream = New-Object System.IO.MemoryStream( @(,$bytes) )

            $file = Add-PnPFile -Stream $stream -Folder $FolderPath -FileName $fileName -Values @{} -Connection $connection

            # output full URL of new file
            "{0}://{1}{2}" -f $uri.Scheme, $uri.Host, $file.ServerRelativeUrl

            if( $stream )
            {
                $stream.Close()
                $stream.Dispose()
            }
        }

        Disconnect-PnPOnline -Connection $connection
    }
    end
    {
    }
}

function New-UserProfilePropertyImportJob
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]$ImportProfilePropertiesUserIdType,
        [Parameter(Mandatory=$true)][string[]]$SourceUri,
        [Parameter(Mandatory=$true)][System.Collections.Generic.IDictionary[string,string]]$PropertyMap,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][string]$Thumbprint,
        [Parameter(Mandatory=$true)][string]$Tenant

    )
    
    begin
    {
        $sourceDataIdProperty = switch ($ImportProfilePropertiesUserIdType)
                                {
                                    ([Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::CloudId)
                                    {
                                        "id"
                                    }
                                    ([Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::Email)
                                    {
                                        "mail"
                                    }
                                    ([Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::PrincipalName)
                                    {
                                        "userPrincipalName"
                                    }
                                }
    }
    process
    {
        $connection = Connect-PnPOnline -Url "https://$Tenant-admin.sharepoint.com" -ClientId $ClientId -Thumbprint $Thumbprint -Tenant "$Tenant.onmicrosoft.com" -ReturnConnection

        $o365Tenant = New-Object Microsoft.Online.SharePoint.TenantManagement.Office365Tenant($connection.Context)

        $connection.Context.Load($o365Tenant)

        Invoke-PnPQuery

        foreach( $uri in $SourceUri )
        {
            $o365Tenant.QueueImportProfileProperties( $ImportProfilePropertiesUserIdType, $sourceDataIdProperty, $PropertyMap, $SourceUri )

            Invoke-PnPQuery
        }

        Disconnect-PnPOnline -Connection $connection
    }
    end
    {
    }
}

function Get-UserProfilePropertyImportJob
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][System.Guid[]]$JobId,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][string]$Thumbprint,
        [Parameter(Mandatory=$true)][string]$Tenant

    )
    
    begin
    {
    }
    process
    {
        $connection = Connect-PnPOnline -Url "https://$Tenant-admin.sharepoint.com" -ClientId $ClientId -Thumbprint $Thumbprint -Tenant "$Tenant.onmicrosoft.com" -ReturnConnection

        $o365Tenant = New-Object Microsoft.Online.SharePoint.TenantManagement.Office365Tenant($connection.Context)

        $connection.Context.Load($o365Tenant)

        Invoke-PnPQuery

        foreach( $id in $JobId )
        {
            $job = $o365Tenant.GetImportProfilePropertyJob($id)

            $connection.Context.Load($job)

            Invoke-PnPQuery

            $job
        }

        Disconnect-PnPOnline -Connection $connection
    }
    end
    {
    }
}




# Azure AD App Principal Connection Settings

    # requries Microsoft Graph > User.Read.All
    #          SharePoint      > Sites.FullControl 
    #          Sharepoint      > Users.ReadWrite.All

    $clientId   = $env:O365_CLIENTID
    $thumbprint = $env:O365_THUMBPRINT
    $tenant     = $env:O365_TENANT


# primary identifier
    $importProfilePropertiesUserIdType = [Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::PrincipalName


# property mappings

    # Key   == JSON file property name 
    # Value == User Profile Service prpperty name
    # keep in sync the Attributes and OnPremisesExtensionAttributes parameters requested by Get-UserAttributes

    $propertyMappings = [System.Collections.Generic.Dictionary[string,string]]::new()
    $propertyMappings.Id                   = "EmployeeNumber"
    $propertyMappings.City                 = "WorkLocation"
    #$propertyMappings.extensionAttribute9  = "CostCenter"
    #$propertyMappings.extensionAttribute10 = "OrganizationCode"



# connect to Graph API

    Connect-MgGraph -ClientId $clientId -CertificateThumbprint $thumbprint -TenantId "$tenant.onmicrosoft.com" | Out-Null


# get non-guests accounts from graph API. Keep requested properties in sync with $propertyMappings and call to Get-UserAttributes 

    $users = Get-MgUser -All -Property Id, City, OnPremisesExtensionAttributes -Filter "UserType eq 'Member'"


# get the user profile properties from Graph API

    $userAttributes = Get-UserAttributes `
                        -Identity $users `
                        -ImportProfilePropertiesUserIdType $importProfilePropertiesUserIdType `
                        -Attributes Id, City `
                        -OnPremisesExtensionAttributes "extensionAttribute9", "extensionAttribute10"


# convert the properties to json file(s) 

    $jsonChunks = ConvertTo-UserProfileBulkImportJson -Identity $userAttributes


# upload the files to a folder in SPO

    $files = Save-JsonFile `
                -Json       $jsonChunks `
                -WebUrl     "https://$tenant.sharepoint.com/sites/teamsite" `
                -FolderPath "/sites/teamsite/shared documents/UPAImport" `
                -ClientId   $clientId `
                -Thumbprint $thumbprint `
                -Tenant     $tenant


# create import jobs for each file we uploaded

    $importJobs = New-UserProfilePropertyImportJob `
                    -ImportProfilePropertiesUserIdType $importProfilePropertiesUserIdType `
                    -PropertyMap  $propertyMappings `
                    -SourceUri    $files `
                    -ClientId     $clientId `
                    -Thumbprint   $thumbprint `
                    -Tenant       $tenant

    $importJobIds = @($importJobs | Select-Object -ExpandProperty Value)

    if( $importJobIds.Count -eq 0 -or $importJobIds[0] -eq [System.Guid]::Empty ) { return }


# wait for jobs to complete
    
    do
    {

        $importJob = Get-UserProfilePropertyImportJob -JobId $importJobIds -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenant
        
        $importJob | SELECT @{ N="Timestamp"; E={ Get-Date }}, JobId, State, SourceUri, LogFolderUri, Error, ErrorMessage 

        Start-Sleep -Seconds 30
    }
    while( $importJob.State -ne "Succeeded" -and $importJob.State -ne "Error" -and $importJob.State -ne "Unknown" )
