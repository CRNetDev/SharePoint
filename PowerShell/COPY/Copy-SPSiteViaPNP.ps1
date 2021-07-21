# https://pnp.github.io/powershell/cmdlets/Get-PnPProvisioningTemplate.html
# https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/introducing-the-pnp-provisioning-engine

<#

This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
## Please note: None of the conditions outlined in the disclaimer above will supersede the terms and conditions contained within the Premier Customer Services Description.
#>

function Copy-DocumentLibraryItems{
    param([parameter(Mandatory=$true)][string]$LibraryName)
        # Get the Folders and Items in the Library
        $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $LibraryName -ItemType Folder
        $items =  Get-PnPFolderItem -FolderSiteRelativeUrl $LibraryName -ItemType File
    
        #Copy Folders
        foreach($folder in $folders){
            #Exclude the Forms Folder
            if($folder.Name -eq "Forms"){
                continue;
            }
    
            write-host "Copying folder: $($folder.ServerRelativeUrl)"
            Copy-PnPFile -SourceUrl $folder.ServerRelativeUrl -TargetUrl "/sites/TemplateDestination/Shared Documents" -Overwrite
        }
    
        #Copy List Items
        write-host "Copying Files in the Root Folder";
        foreach($item in $items){
            Copy-PnPFile -SourceUrl $item.ServerRelativeUrl -TargetUrl "/sites/TemplateDestination/Shared Documents" -Overwrite
        }
    }
    
    #Save the Source Site Schema XML to a file
    $sourceSiteURL = "https://tenant.sharepoint.com/sites/Name" 
    $SchemaPath = "C:\Temp\Schema_M365LP.pnp"
    $destinationSiteURL = "https://tenant.sharepoint.com/sites/TemplateDestination"
    
    
    #Connect to the source site via PnP
    Connect-PnPOnline -Url $sourceSiteURL -Interactive
     
    #Get Site Schema
    Get-PnPSiteTemplate -Out ($SchemaPath) -IncludeAllPages -PersistBrandingFiles
    
    # Specifically extract the List Handlers
    Get-PnPSiteTemplate -Out $SchemaPath -Handlers Lists -ListsToExtract "Issue tracker" 
    Add-PnPDataRowsToSiteTemplate -Path $SchemaPath -List 'Issue tracker' 
    
    
    #Create a new destination site if needed
     
    #Connect to PnP Online
    Connect-PnPOnline -Url $destinationSiteURL -Interactive
     
    #Apply Pnp Provisioning Template
    Invoke-PnPSiteTemplate -Path $SchemaPath -ClearNavigation
    
    #Apply Lists - Use if there is a separate schema file for lists
    #Invoke-PnPSiteTemplate -Path $listSchemaPath -Handlers Lists
    
    #Reconnect to the Source Site to read the Library Data
    Connect-PnPOnline -Url $sourceSiteURL -Interactive
    Copy-DocumentLibraryItems -LibraryName "Shared Documents"