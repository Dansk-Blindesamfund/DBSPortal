param(
    [string]$ResourceGroup = "DBS",
    [string]$Location = "westeurope",
    [string]$ContainerAppName = "dbs-itsupport-portal",
    [string]$AcrName = "dbsitsupportacr",
    [string]$ContainerEnvName = "dbs-itsupport-env",
    [string]$ImageTag = "latest"
)

$ErrorActionPreference = "Stop"

Write-Host "Validerer Azure login..."
az account show --output table | Out-Host

Write-Host "Registrerer noedvendige resource providers (kan tage et par minutter)..."
az provider register --namespace Microsoft.ContainerRegistry --wait | Out-Null
az provider register --namespace Microsoft.App --wait | Out-Null
az provider register --namespace Microsoft.OperationalInsights --wait | Out-Null

Write-Host "Sikrer resource group..."
az group create --name $ResourceGroup --location $Location --output table | Out-Host

Write-Host "Sikrer Azure Container Registry..."
$acrExists = az acr show --name $AcrName --resource-group $ResourceGroup --query name -o tsv 2>$null
if (-not $acrExists) {
    az acr create --name $AcrName --resource-group $ResourceGroup --sku Basic --admin-enabled true --output table | Out-Host
}

$acrLoginServer = az acr show --name $AcrName --resource-group $ResourceGroup --query loginServer -o tsv
$acrUsername = az acr credential show --name $AcrName --resource-group $ResourceGroup --query username -o tsv
$acrPassword = az acr credential show --name $AcrName --resource-group $ResourceGroup --query passwords[0].value -o tsv

if (-not $acrLoginServer -or -not $acrUsername -or -not $acrPassword) {
    throw "Kunne ikke hente ACR loginoplysninger."
}

$imageName = "$acrLoginServer/itsupport-portal:$ImageTag"

Write-Host "Bygger container image i ACR..."
az acr build --registry $AcrName --image "itsupport-portal:$ImageTag" --file "deploy/portal.Dockerfile" . | Out-Host

Write-Host "Sikrer Container Apps environment..."
$envId = az containerapp env show --name $ContainerEnvName --resource-group $ResourceGroup --query id -o tsv 2>$null
if (-not $envId) {
    az containerapp env create --name $ContainerEnvName --resource-group $ResourceGroup --location $Location --output table | Out-Host
}

Write-Host "Deployer Container App..."
$appExists = az containerapp show --name $ContainerAppName --resource-group $ResourceGroup --query name -o tsv 2>$null
if (-not $appExists) {
    az containerapp create `
      --name $ContainerAppName `
      --resource-group $ResourceGroup `
      --environment $ContainerEnvName `
      --image $imageName `
      --target-port 8080 `
      --ingress external `
      --registry-server $acrLoginServer `
    --registry-username $acrUsername `
    --registry-password $acrPassword `
      --cpu 1.0 `
      --memory 2.0Gi `
      --min-replicas 1 `
      --max-replicas 1 `
      --output table | Out-Host
} else {
    az containerapp update `
      --name $ContainerAppName `
      --resource-group $ResourceGroup `
      --image $imageName `
            --registry-server $acrLoginServer `
            --registry-username $acrUsername `
            --registry-password $acrPassword `
      --output table | Out-Host
}

$portalUrl = az containerapp show --name $ContainerAppName --resource-group $ResourceGroup --query properties.configuration.ingress.fqdn -o tsv
Write-Host "Portal URL: https://$portalUrl"
Write-Host "Husk at saette hemmeligheder sikkert (Graph credentials / storage), fx med containerapp secrets og env vars."
