<#
.DESCRIPTION
    Load Entities or Attributes to Erwin data model using API.
.PARAMETER Update
    Specifies update action to perform. ('Entity', 'Attribute').
.EXAMPLE
    Update-Erwin.ps1 -Update Attribute
.INPUTS
    System.String.
.OUTPUTS
    none
.NOTES
    Created by Dave Welden
.LINK
    https://support.erwin.com/hc/en-us/articles/115000243748-erwin-DM-populating-Entity-and-Attribute-Definitions-and-Datatypes-using-the-API-
#>
Param (
    [Parameter(Mandatory = $true)]
    [ValidateSet('Entity', 'Attribute')]
    [string]$Update
)

<#
.DESCRIPTION
    Select and open file using System.Windows.Forms.OpenFileDialog
.PARAMETER Title
    Specifies dialog window title.
.PARAMETER Filter
    Specifies file type filter.
.EXAMPLE
    Open-File -Title "Select Erwin model file" -Filter "Erwin DM (*.erwin)|*.erwin"
.LINK
    https://gallery.technet.microsoft.com/scriptcenter/GUI-popup-FileOpenDialog-babd911d
#>
Function Open-File {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter(Mandatory = $true)]
        [string]$Filter
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()   
    $FileBrowser.title = $Title
    $FileBrowser.filter = $Filter
    $FileBrowser.ShowHelp = $False
    $result = $FileBrowser.ShowDialog()
    if ($result -eq "Cancel") {
        Write-Output "No file selected, exiting"
        Exit
    }
    $FileName = $FileBrowser.FileName
    return $FileName
}

<#
.DESCRIPTION
    Update Erwin data model Entities.
#>
Function Update-Entity {
    $ErrorActionPreference = 'SilentlyContinue'
    
    # Select and open Erwin data model
    $erwinFile = Open-File -Title "Select Erwin model file" -Filter "Erwin DM (*.erwin)|*.erwin"
    $erwinFile = "erwin://" + $erwinFile
    
    # Select and open updates file
    $updatesFile = Open-File -Title "Select Entity updates file" -Filter "CSV file (*.csv)|*.csv"
    $csv = Import-Csv -Path $updatesFile

    # Intialize Erwin API and begin transaction
    $ERwin = New-Object -ComObject erwin9.SCAPI
    $PersistenceUnit = $ERwin.PersistenceUnits.Add($erwinFile)
    $Session = $ERwin.Sessions.Add()
    $Session.Open($PersistenceUnit)
    $RootObj = $Session.ModelObjects.Collect($Session.ModelObjects.Root,"Entity")
    $TxId = $Session.BeginTransaction()
    
    # Read updates file and apply changes to Erwin data model
    ForEach ($line in $csv) {
        $Entity = $line.Entity
        $Definition = $line.Definition
        $Comment = $line.Comment
        
        # Search for Entity, update if found or create new
        $EntityObject = $null
        $EntityObject = $RootObj.Item($Entity, "Entity")
        if ($null -eq $EntityObject) {
            Write-Output "$Entity not found, creating Entity"
            $EntityObject = $Session.ModelObjects.Add("Entity")
            $EntityObject.Properties("Name").Value = $Entity
        }
        $EntityObject.Properties("Definition").Value = $Definition
        $EntityObject.Properties("Comment").Value = $Comment
    }
    
    # Commit transaction and shutdown API
    $Session.CommitTransaction($TxId)
    $PersistenceUnit.Save()
    $Session.Close()
    $ERwin.Sessions.Clear()
    $PersistenceUnit = $null
    $ERwin = $null
    }

<#
.DESCRIPTION
    Update Erwin data model Entity Attributes.
#>
Function Update-Attribute {
    $ErrorActionPreference = 'SilentlyContinue'

    # Select and open Erwin data model
    $erwinFile = Open-File -Title "Select Erwin model file" -Filter "Erwin DM (*.erwin)|*.erwin"
    $erwinFile = "erwin://" + $erwinFile
    
    # Select and open updates file
    $updatesFile = Open-File -Title "Select Attribute updates file" -Filter "CSV file (*.csv)|*.csv"
    $csv = Import-Csv -Path $updatesFile

    # Intialize Erwin API and begin transaction
    $ERwin = New-Object -ComObject erwin9.SCAPI
    $PersistenceUnit = $ERwin.PersistenceUnits.Add($erwinFile)
    $Session = $ERwin.Sessions.Add()
    $Session.Open($PersistenceUnit)
    $RootObj = $Session.ModelObjects.Collect($Session.ModelObjects.Root,"Entity")
    $TxId = $Session.BeginTransaction()
    
    # Read updates file and apply changes to Erwin data model
    ForEach ($line in $csv) {
        $Entity = $line.Entity
        $Attribute = $line.Attribute
        $LogicalDataType = $line.Logical_Data_Type
        $Definition = $line.Definition
        $Comment = $line.Comment
        
        # Search for Entity, update if found or create new
        $EntityObject = $null
        $EntityObject = $Session.ModelObjects.Item($Entity, "Entity")
        if ($null -eq $EntityObject) {
            Write-Output "$Entity not found, creating Entity"
            $EntityObject = $Session.ModelObjects.Add("Entity")
            $EntityObject.Properties("Name").Value = $Entity
        }

        # Search for Entity Attribute, update if found or create new
        $EntityAttributes = $Session.ModelObjects.Collect($EntityObject, "Attribute")
        $AttributeObject = $null
        $AttributeObject = $EntityAttributes.Item($Attribute, "Attribute")
        if ($null -eq $AttributeObject) {
            Write-Output "$Attribute not found, creating Attribute"
            $AttributeObject = $EntityAttributes.Add("Attribute")
            $AttributeObject.Properties("Name").Value = $Attribute
            $AttributeObject.Properties("Type").Value = 100
        }
        $AttributeObject.Properties("Logical_Data_Type").Value = $LogicalDataType
        $AttributeObject.Properties("Definition").Value = $Definition
        $AttributeObject.Properties("Comment").Value = $Comment
    }
    
    # Commit transaction and shutdown API
    $Session.CommitTransaction($TxId)
    $PersistenceUnit.Save()
    $Session.Close()
    $ERwin.Sessions.Clear()
    $PersistenceUnit = $null
    $ERwin = $null
    }
    
    Switch ($Update.ToLower()) {
    entity {
        Write-Output "Entity updates selected"
        Update-Entity
    }
    attribute {
        Write-Output "Attribute updates selected"
        Update-Attribute
    }
}
