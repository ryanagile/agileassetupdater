<#
.SYNOPSIS
Autotask Asset Updater runtime script
#>

$CredFile = "C:\AgileICT\AssetUpdater\creds.dat"

# Read and decode Base64 credentials
$Encoded = Get-Content $CredFile -Raw
$Json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Encoded))
$Cred = $Json | ConvertFrom-Json

$CompanyId = $Cred.CompanyId

# Autotask REST headers
$Headers = @{

    ApiIntegrationCode = $Cred.ApiCode
    UserName           = $Cred.Username
    Secret             = $Cred.Password
    "Content-Type"     = "application/json"

}

$Asset = @{

    referenceTitle    = $env:COMPUTERNAME
    serialNumber      = (Get-CimInstance Win32_BIOS).SerialNumber
    isActive          = $true
    userDefinedFields = @(
    @{
        name = "Operating System"
        value = (Get-CimInstance Win32_OperatingSystem).Caption
      },
      @{
        name = "Disks"
        value = (Get-PhysicalDisk |Select-Object FriendlyName, MediaType, @{Name="SizeGB";Expression={[math]::Round($_.Size / 1GB,2)}} | ForEach-Object { "$($_.FriendlyName) $($_.MediaType) $($_.SizeGB)GB" }) -join ", "
      },
      @{
        name = "AssetUpdater Date"
        value = Get-Date
      },
      @{
        name = "RAM"
        value = "$([math]::Ceiling((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB))GB"
      }
      ,
      @{
        name = "Processor"
        value = (Get-CimInstance Win32_Processor | Select Name).Name
      }
    )

}

$BaseUri = "https://webservices16.autotask.net/atservicesrest/v1.0"

$Query = @{
    filter = @(@{field="serialNumber"; op="eq"; value=$Asset.serialNumber})
} | ConvertTo-Json

$Result = Invoke-RestMethod -Uri "$BaseUri/ConfigurationItems/query" -Method POST -Headers $Headers -Body $Query

if ($Result.items.Count -gt 0) {

    $Asset.id = $Result.items[0].id
    Invoke-RestMethod -Uri "$BaseUri/ConfigurationItems" -Method PATCH -Headers $Headers -Body ($Asset | ConvertTo-Json -Depth 5)

} else { 

    # Create a new asset
    $Asset.companyId = $CompanyId
    $Asset.configurationItemType = 1

    $ChassisTypes = (Get-CimInstance Win32_SystemEnclosure).ChassisTypes
    $IsLaptop = $ChassisTypes | Where-Object { $_ -in 8,9,10,11,12,14,18,21 }
    $ComputerType = if ($IsLaptop) { "Laptop" } else { "Desktop" }

    $Make  = (Get-CimInstance Win32_ComputerSystem).Manufacturer
    $Model = (Get-CimInstance Win32_ComputerSystem).Model

    $ProductName = "$Make $Model $ComputerType Computer"

    $ProductQuery = @{
        filter = @(
            @{
                field = "name"
                op    = "eq"
                value = $ProductName
            }
        )
    } | ConvertTo-Json

    $ProductResult = Invoke-RestMethod `
        -Uri "$BaseUri/Products/query" `
        -Method POST `
        -Headers $Headers `
        -Body $ProductQuery

    if ($ProductResult.items.Count -gt 0) {
        $ProductId = $ProductResult.items[0].id
    }
    else {
        $NewProduct = @{
            name        = $ProductName
            isActive    = $true
            description = "Auto-created by Asset Updater script"
            isSerialized = $false
            productBillingCodeID = 29683550
            defaultInstalledProductCategoryID = 3
            manufacturerName = $Make
            manufacturerProductName = $Model
            productCategory = 14
        }

        $CreatedProduct = Invoke-RestMethod `
            -Uri "$BaseUri/Products" `
            -Method POST `
            -Headers $Headers `
            -Body ($NewProduct | ConvertTo-Json -Depth 5)

        $ProductId = $CreatedProduct.itemId
    }

    $Asset.productID = $ProductId
    
    Invoke-RestMethod -Uri "$BaseUri/ConfigurationItems" -Method POST -Headers $Headers -Body ($Asset | ConvertTo-Json -Depth 5)
}