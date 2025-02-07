# Importar el módulo PnP PowerShell
Import-Module PnP.PowerShell

# Definir rutas
$outputCsvPath = "C:\Path\To\Output\RecycleBinReport.csv"
$logFolderPath = "C:\Path\To\Output\Logs"

# Crear la carpeta de logs si no existe
if (-not (Test-Path $logFolderPath)) {
    New-Item -ItemType Directory -Path $logFolderPath | Out-Null
}

# Inicializar un array para almacenar los resultados
$results = @()

# Definir Client ID y Thumbprint para la conexión PnP
$clientId = "your-client-id"
$thumbprint = "your-thumbprint"

# Conectar al tenant de SharePoint (script principal)
$adminSiteUrl = "https://<your-tenant>-admin.sharepoint.com"
Connect-PnPOnline -Url $adminSiteUrl -ClientId $clientId -Thumbprint $thumbprint

# Obtener todos los sitios de SharePoint en el tenant
$sites = Get-PnPTenantSite

# Definir el tamaño del lote (por ejemplo, 100 sitios por lote)
$batchSize = 100
$totalSites = $sites.Count
$batches = [math]::Ceiling($totalSites / $batchSize)

# Definir el límite de jobs simultáneos
$maxJobs = 10

# Registrar el inicio del procesamiento
"Procesamiento iniciado el $(Get-Date). Total de sitios: $totalSites, Tamaño del lote: $batchSize, Total de lotes: $batches" | Out-File -FilePath "$logFolderPath\OverallLog.txt" -Append

# Función para obtener todos los elementos de la papelera de reciclaje con paginación
function Get-AllRecycleBinItems {
    param ($stage)

    $allItems = @()
    $rowLimit = 5000
    do {
        $currentItems = Get-PnPRecycleBinItem -Stage $stage -RowLimit $rowLimit
        $allItems += $currentItems
    } while ($currentItems.Count -eq $rowLimit)

    return $allItems
}

# Función para procesar un sitio
function Process-Site {
    param ($site, $batchNumber, $siteIndex, $totalSites, $logFilePath, $clientId, $thumbprint)

    # Inicializar un objeto de resultado
    $result = [PSCustomObject]@{
        SiteUrl                              = $site.Url
        StorageUsedGB                        = "Error"
        StorageCapacityGB                    = "Error"
        FirstStageTotalItems                 = "Error"
        FirstStageSystemItemsCount           = "Error"
        FirstStageSystemItemsSizeGB          = "Error"
        SecondStageTotalItems                = "Error"
        SecondStageSystemItemsCount          = "Error"
        SecondStageSystemItemsSizeGB         = "Error"
        TotalItemsBothStages                 = "Error"
        TotalSystemItemsBothStages           = "Error"
        TotalSystemItemsSizeBothStagesGB     = "Error"
        TotalRecycleBinSizeGB                = "Error"
        Error                                = ""
    }

    try {
        # Registrar el progreso
        $logMessage = "Lote $batchNumber, Sitio $siteIndex de $totalSites, URL: $($site.Url)"
        $logMessage | Out-File -FilePath $logFilePath -Append

        # Conectar al sitio usando Client ID y Thumbprint
        Write-Host "Conectando al sitio: $($site.Url)"
        Connect-PnPOnline -Url $site.Url -ClientId $clientId -Thumbprint $thumbprint -ErrorAction Stop

        # Obtener información de almacenamiento del sitio
        $siteStorage = Get-PnPSite -Includes StorageUsage, StorageQuota
        $storageUsed = $siteStorage.StorageUsage
        $storageCapacity = $siteStorage.StorageQuota

        # Obtener todos los elementos de la papelera de reciclaje (primera etapa)
        try {
            $firstStageRecycleBin = Get-AllRecycleBinItems -stage "FirstStage"
            $firstStageItemsDeletedBySystem = $firstStageRecycleBin | Where-Object { $_.DeletedByEmail -eq "System Account" }
            $firstStageTotalItems = $firstStageRecycleBin.Count
            $firstStageSystemItemsCount = $firstStageItemsDeletedBySystem.Count
            $firstStageSystemItemsSize = ($firstStageItemsDeletedBySystem | Measure-Object -Property Size -Sum).Sum
        }
        catch {
            $result.Error = "Error: No se pudo acceder a la papelera (primera etapa)"
            throw
        }

        # Obtener todos los elementos de la papelera de reciclaje (segunda etapa)
        try {
            $secondStageRecycleBin = Get-AllRecycleBinItems -stage "SecondStage"
            $secondStageItemsDeletedBySystem = $secondStageRecycleBin | Where-Object { $_.DeletedByEmail -eq "System Account" }
            $secondStageTotalItems = $secondStageRecycleBin.Count
            $secondStageSystemItemsCount = $secondStageItemsDeletedBySystem.Count
            $secondStageSystemItemsSize = ($secondStageItemsDeletedBySystem | Measure-Object -Property Size -Sum).Sum
        }
        catch {
            $result.Error = "Error: No se pudo acceder a la papelera (segunda etapa)"
            throw
        }

        # Calcular totales combinados
        $totalItemsBothStages = $firstStageTotalItems + $secondStageTotalItems
        $totalSystemItemsBothStages = $firstStageSystemItemsCount + $secondStageSystemItemsCount
        $totalSystemItemsSizeBothStages = $firstStageSystemItemsSize + $secondStageSystemItemsSize
        $totalRecycleBinSize = ($firstStageRecycleBin | Measure-Object -Property Size -Sum).Sum + ($secondStageRecycleBin | Measure-Object -Property Size -Sum).Sum

        # Actualizar el objeto de resultado
        $result.StorageUsedGB = [math]::Round($storageUsed / 1GB, 2)
        $result.StorageCapacityGB = [math]::Round($storageCapacity / 1GB, 2)
        $result.FirstStageTotalItems = $firstStageTotalItems
        $result.FirstStageSystemItemsCount = $firstStageSystemItemsCount
        $result.FirstStageSystemItemsSizeGB = [math]::Round($firstStageSystemItemsSize / 1GB, 2)
        $result.SecondStageTotalItems = $secondStageTotalItems
        $result.SecondStageSystemItemsCount = $secondStageSystemItemsCount
        $result.SecondStageSystemItemsSizeGB = [math]::Round($secondStageSystemItemsSize / 1GB, 2)
        $result.TotalItemsBothStages = $totalItemsBothStages
        $result.TotalSystemItemsBothStages = $totalSystemItemsBothStages
        $result.TotalSystemItemsSizeBothStagesGB = [math]::Round($totalSystemItemsSizeBothStages / 1GB, 2)
        $result.TotalRecycleBinSizeGB = [math]::Round($totalRecycleBinSize / 1GB, 2)
    }
    catch {
        # Registrar errores
        if (-not $result.Error) {
            $result.Error = "Error: No se pudo conectar al sitio"
        }
        $errorMessage = "Error procesando el sitio $($site.Url): $_"
        $errorMessage | Out-File -FilePath $logFilePath -Append
        Write-Host $errorMessage -ForegroundColor Red
    }

    # Devolver el objeto de resultado
    return $result
}

# Procesar todos los lotes
for ($batchNumber = 1; $batchNumber -le $batches; $batchNumber++) {
    $batchStart = ($batchNumber - 1) * $batchSize
    $batchEnd = [math]::Min($batchStart + $batchSize - 1, $totalSites - 1)
    $batchSites = $sites[$batchStart..$batchEnd]

    # Crear un archivo de log para este lote
    $batchLogFilePath = "$logFolderPath\Batch_$batchNumber.log"
    "Iniciando lote $batchNumber (sitios $batchStart a $batchEnd)" | Out-File -FilePath $batchLogFilePath -Append

    # Inicializar una lista de jobs
    $jobs = @()

    # Procesar sitios en paralelo (hasta $maxJobs simultáneos)
    foreach ($site in $batchSites) {
        $siteIndex = $batchSites.IndexOf($site) + $batchStart + 1

        # Si hay $maxJobs en ejecución, esperar a que terminen
        while ((Get-Job -State Running).Count -ge $maxJobs) {
            Start-Sleep -Seconds 1
        }

        # Crear un job para el sitio
        $job = Start-Job -ScriptBlock {
            # Importar el módulo PnP dentro del job
            Import-Module PnP.PowerShell

            # Definir las funciones dentro del job
            function Get-AllRecycleBinItems {
                param ($stage)

                $allItems = @()
                $rowLimit = 5000
                do {
                    $currentItems = Get-PnPRecycleBinItem -Stage $stage -RowLimit $rowLimit
                    $allItems += $currentItems
                } while ($currentItems.Count -eq $rowLimit)

                return $allItems
            }

            function Process-Site {
                param ($site, $batchNumber, $siteIndex, $totalSites, $logFilePath, $clientId, $thumbprint)

                # Inicializar un objeto de resultado
                $result = [PSCustomObject]@{
                    SiteUrl                              = $site.Url
                    StorageUsedGB                        = "Error"
                    StorageCapacityGB                    = "Error"
                    FirstStageTotalItems                 = "Error"
                    FirstStageSystemItemsCount           = "Error"
                    FirstStageSystemItemsSizeGB          = "Error"
                    SecondStageTotalItems                = "Error"
                    SecondStageSystemItemsCount          = "Error"
                    SecondStageSystemItemsSizeGB         = "Error"
                    TotalItemsBothStages                 = "Error"
                    TotalSystemItemsBothStages           = "Error"
                    TotalSystemItemsSizeBothStagesGB     = "Error"
                    TotalRecycleBinSizeGB                = "Error"
                    Error                                = ""
                }

                try {
                    # Registrar el progreso
                    $logMessage = "Lote $batchNumber, Sitio $siteIndex de $totalSites, URL: $($site.Url)"
                    $logMessage | Out-File -FilePath $logFilePath -Append

                    # Conectar al sitio usando Client ID y Thumbprint
                    Write-Host "Conectando al sitio: $($site.Url)"
                    Connect-PnPOnline -Url $site.Url -ClientId $clientId -Thumbprint $thumbprint -ErrorAction Stop

                    # Obtener información de almacenamiento del sitio
                    $siteStorage = Get-PnPSite -Includes StorageUsage, StorageQuota
                    $storageUsed = $siteStorage.StorageUsage
                    $storageCapacity = $siteStorage.StorageQuota

                    # Obtener todos los elementos de la papelera de reciclaje (primera etapa)
                    try {
                        $firstStageRecycleBin = Get-AllRecycleBinItems -stage "FirstStage"
                        $firstStageItemsDeletedBySystem = $firstStageRecycleBin | Where-Object { $_.DeletedByEmail -eq "System Account" }
                        $firstStageTotalItems = $firstStageRecycleBin.Count
                        $firstStageSystemItemsCount = $firstStageItemsDeletedBySystem.Count
                        $firstStageSystemItemsSize = ($firstStageItemsDeletedBySystem | Measure-Object -Property Size -Sum).Sum
                    }
                    catch {
                        $result.Error = "Error: No se pudo acceder a la papelera (primera etapa)"
                        throw
                    }

                    # Obtener todos los elementos de la papelera de reciclaje (segunda etapa)
                    try {
                        $secondStageRecycleBin = Get-AllRecycleBinItems -stage "SecondStage"
                        $secondStageItemsDeletedBySystem = $secondStageRecycleBin | Where-Object { $_.DeletedByEmail -eq "System Account" }
                        $secondStageTotalItems = $secondStageRecycleBin.Count
                        $secondStageSystemItemsCount = $secondStageItemsDeletedBySystem.Count
                        $secondStageSystemItemsSize = ($secondStageItemsDeletedBySystem | Measure-Object -Property Size -Sum).Sum
                    }
                    catch {
                        $result.Error = "Error: No se pudo acceder a la papelera (segunda etapa)"
                        throw
                    }

                    # Calcular totales combinados
                    $totalItemsBothStages = $firstStageTotalItems + $secondStageTotalItems
                    $totalSystemItemsBothStages = $firstStageSystemItemsCount + $secondStageSystemItemsCount
                    $totalSystemItemsSizeBothStages = $firstStageSystemItemsSize + $secondStageSystemItemsSize
                    $totalRecycleBinSize = ($firstStageRecycleBin | Measure-Object -Property Size -Sum).Sum + ($secondStageRecycleBin | Measure-Object -Property Size -Sum).Sum

                    # Actualizar el objeto de resultado
                    $result.StorageUsedGB = [math]::Round($storageUsed / 1GB, 2)
                    $result.StorageCapacityGB = [math]::Round($storageCapacity / 1GB, 2)
                    $result.FirstStageTotalItems = $firstStageTotalItems
                    $result.FirstStageSystemItemsCount = $firstStageSystemItemsCount
                    $result.FirstStageSystemItemsSizeGB = [math]::Round($firstStageSystemItemsSize / 1GB, 2)
                    $result.SecondStageTotalItems = $secondStageTotalItems
                    $result.SecondStageSystemItemsCount = $secondStageSystemItemsCount
                    $result.SecondStageSystemItemsSizeGB = [math]::Round($secondStageSystemItemsSize / 1GB, 2)
                    $result.TotalItemsBothStages = $totalItemsBothStages
                    $result.TotalSystemItemsBothStages = $totalSystemItemsBothStages
                    $result.TotalSystemItemsSizeBothStagesGB = [math]::Round($totalSystemItemsSizeBothStages / 1GB, 2)
                    $result.TotalRecycleBinSizeGB = [math]::Round($totalRecycleBinSize / 1GB, 2)
                }
                catch {
                    # Registrar errores
                    if (-not $result.Error) {
                        $result.Error = "Error: No se pudo conectar al sitio"
                    }
                    $errorMessage = "Error procesando el sitio $($site.Url): $_"
                    $errorMessage | Out-File -FilePath $logFilePath -Append
                    Write-Host $errorMessage -ForegroundColor Red
                }

                # Devolver el objeto de resultado
                return $result
            }

            # Llamar a la función Process-Site dentro del job
            Process-Site -site $using:site -batchNumber $using:batchNumber -siteIndex $using:siteIndex -totalSites $using:totalSites -logFilePath $using:batchLogFilePath -clientId $using:clientId -thumbprint $using:thumbprint
        }

        # Agregar el job a la lista
        $jobs += $job
    }

    # Esperar a que todos los jobs del lote terminen
    $jobs | Wait-Job

    # Recopilar resultados de los jobs
    foreach ($job in $jobs) {
        $result = Receive-Job -Job $job
        if ($result) {
            $results += $result
        }
    }

    # Limpiar los jobs
    $jobs | Remove-Job

    # Registrar la finalización del lote
    "Lote $batchNumber completado" | Out-File -FilePath $batchLogFilePath -Append
}

# Exportar los resultados a un archivo CSV
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation

Write-Host "Reporte exportado a $outputCsvPath"
Write-Host "Archivos de log guardados en $logFolderPath"
