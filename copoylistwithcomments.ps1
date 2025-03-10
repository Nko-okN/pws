# Script para migrar lista de SharePoint incluyendo comentarios, adjuntos y versiones
# Prerequisitos: Módulos PnP.PowerShell instalados
# Install-Module -Name "PnP.PowerShell" -Scope CurrentUser -Force

# Parámetros de conexión
$sourceUrl = "https://tudominio.sharepoint.com/sites/SitioOrigen"
$destUrl = "https://tudominio.sharepoint.com/sites/SitioDestino" # URL del sitio que contiene el canal privado
$sourceListName = "NombreDeTuListaOrigen"
$destListName = "NombreDeTuListaDestino"
$logFilePath = "C:\Temp\MigracionSharePoint_$(Get-Date -Format 'yyyy-MM-dd_HH-mm').log"

# Función para escribir en el log
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logFilePath -Value $logMessage
    Write-Host $logMessage
}

try {
    Write-Log "Iniciando proceso de migración de lista con comentarios"
    
    # Crear conexiones independientes usando PnP PowerShell y conexiones con contexto
    Write-Log "Conectando al sitio origen..."
    $sourcePnpConnection = Connect-PnPOnline -Url $sourceUrl -Interactive -ReturnConnection
    
    Write-Log "Conectando al sitio destino..."
    $destPnpConnection = Connect-PnPOnline -Url $destUrl -Interactive -ReturnConnection
    
    # Obtener lista origen y sus campos
    Write-Log "Obteniendo lista origen y sus campos..."
    $sourceList = Get-PnPList -Identity $sourceListName -Connection $sourcePnpConnection
    $sourceFields = Get-PnPField -List $sourceList -Connection $sourcePnpConnection | Where-Object { 
        -not $_.Hidden -and -not $_.ReadOnlyField -and $_.InternalName -ne "ContentType" 
    }
    
    # Obtener lista destino
    Write-Log "Obteniendo lista destino..."
    $destList = Get-PnPList -Identity $destListName -Connection $destPnpConnection
    
    # Obtener todos los elementos de origen
    Write-Log "Obteniendo elementos de la lista origen..."
    $sourceItems = Get-PnPListItem -List $sourceList -PageSize 100 -Connection $sourcePnpConnection
    Write-Log "Se encontraron $($sourceItems.Count) elementos para migrar"
    
    # Crear mapeo para seguimiento de IDs viejo->nuevo
    $idMapping = @{}
    
    # Migrar cada elemento
    foreach ($item in $sourceItems) {
        $itemId = $item.Id
        Write-Log "Procesando elemento ID: $itemId - Título: $($item['Title'])"
        
        # Preparar propiedades para el nuevo elemento
        $newItemProperties = @{}
        
        # Copiar todas las propiedades relevantes
        foreach ($field in $sourceFields) {
            $fieldName = $field.InternalName
            if ($item[$fieldName] -ne $null) {
                # Manejar personas/usuarios - convertir de FieldUserValue a email
                if ($field.TypeAsString -eq "User" -or $field.TypeAsString -eq "UserMulti") {
                    if ($item[$fieldName] -is [Microsoft.SharePoint.Client.FieldUserValue]) {
                        # Usuario único
                        $userValue = $item[$fieldName]
                        if ($userValue.Email) {
                            $newItemProperties[$fieldName] = $userValue.Email
                        }
                    }
                    elseif ($item[$fieldName] -is [Microsoft.SharePoint.Client.FieldUserValue[]]) {
                        # Múltiples usuarios
                        $userEmails = @()
                        foreach ($userValue in $item[$fieldName]) {
                            if ($userValue.Email) {
                                $userEmails += $userValue.Email
                            }
                        }
                        $newItemProperties[$fieldName] = $userEmails
                    }
                }
                # Manejar campos de elección múltiple
                elseif ($field.TypeAsString -eq "MultiChoice") {
                    $newItemProperties[$fieldName] = $item[$fieldName]
                }
                # Manejar campos de fecha
                elseif ($field.TypeAsString -eq "DateTime") {
                    $newItemProperties[$fieldName] = $item[$fieldName]
                }
                # Campos normales
                else {
                    $newItemProperties[$fieldName] = $item[$fieldName]
                }
            }
        }
        
        # Crear nuevo elemento en el destino
        Write-Log "Creando nuevo elemento en destino..."
        $newItem = Add-PnPListItem -List $destList -Values $newItemProperties -Connection $destPnpConnection
        $idMapping[$itemId] = $newItem.Id
        Write-Log "Elemento creado con ID: $($newItem.Id)"
        
        # Migrar versiones (si es necesario)
        Write-Log "Obteniendo versiones del elemento..."
        $sourceCtx = Get-PnPContext -Connection $sourcePnpConnection
        $versions = Get-PnPProperty -ClientObject $item -Property Versions -Connection $sourcePnpConnection
        $versionsCount = $versions.Count
        
        if ($versionsCount -gt 1) {
            Write-Log "Se encontraron $versionsCount versiones para migrar"
            # Nota: La migración de versiones requiere acceso directo al historial de versiones
            # Esta funcionalidad requiere un enfoque más complejo con CSOM
        }
        
        # Migrar archivos adjuntos
        Write-Log "Obteniendo archivos adjuntos..."
        $attachmentFiles = Get-PnPProperty -ClientObject $item -Property AttachmentFiles -Connection $sourcePnpConnection
        
        foreach ($attachment in $attachmentFiles) {
            try {
                Write-Log "Procesando adjunto: $($attachment.FileName)"
                
                # Descargar archivo adjunto a temporal
                $tempFile = [System.IO.Path]::GetTempFileName()
                $tempFile = [System.IO.Path]::ChangeExtension($tempFile, [System.IO.Path]::GetExtension($attachment.FileName))
                
                # Obtener contextos específicos
                $sourceCtx = Get-PnPContext -Connection $sourcePnpConnection
                
                # Usar el contexto de origen para descargar el archivo
                $fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($sourceCtx, $attachment.ServerRelativeUrl)
                $fileStream = [System.IO.FileStream]::new($tempFile, [System.IO.FileMode]::Create)
                $fileInfo.Stream.CopyTo($fileStream)
                $fileStream.Close()
                
                # Subir archivo adjunto al nuevo elemento usando la conexión de destino
                Add-PnPAttachment -List $destList -Identity $newItem.Id -Path $tempFile -FileName $attachment.FileName -Connection $destPnpConnection
                
                # Limpiar archivos temporales
                Remove-Item -Path $tempFile -Force
                
                Write-Log "Adjunto migrado correctamente: $($attachment.FileName)"
            }
            catch {
                Write-Log "Error al migrar adjunto $($attachment.FileName): $_" -Level "ERROR"
            }
        }
        
        # Migrar comentarios
        Write-Log "Obteniendo comentarios del elemento..."
        $comments = Get-PnPComment -List $sourceList -ItemId $itemId -Connection $sourcePnpConnection
        
        if ($comments.Count -gt 0) {
            Write-Log "Se encontraron $($comments.Count) comentarios para migrar"
            
            foreach ($comment in $comments) {
                try {
                    # Extraer email del autor del comentario
                    $commentAuthorEmail = $comment.Author.Email
                    
                    # Migrar comentario preservando autor
                    if ([string]::IsNullOrEmpty($commentAuthorEmail)) {
                        # Si no se puede obtener el email, usar texto plano
                        Add-PnPComment -List $destList -ItemId $newItem.Id -Text "$($comment.Text) [Comentario original de: $($comment.Author.LoginName) - $($comment.CreatedDate)]" -Connection $destPnpConnection
                    }
                    else {
                        # Si tenemos email, preservar el autor
                        Add-PnPComment -List $destList -ItemId $newItem.Id -Text $comment.Text -Author $commentAuthorEmail -Connection $destPnpConnection
                    }
                    
                    Write-Log "Comentario migrado correctamente del autor: $($comment.Author.LoginName)"
                    
                    # Migrar respuestas a comentarios (si existen)
                    $replies = Get-PnPCommentReply -List $sourceList -ItemId $itemId -CommentId $comment.Id -Connection $sourcePnpConnection
                    
                    if ($replies -and $replies.Count -gt 0) {
                        Write-Log "Migrando $($replies.Count) respuestas al comentario"
                        
                        # Necesitamos primero obtener el comentario recién creado para poder agregarle respuestas
                        $newComments = Get-PnPComment -List $destList -ItemId $newItem.Id -Connection $destPnpConnection
                        $newCommentId = $newComments | Where-Object { $_.Text -eq $comment.Text } | Select-Object -First 1 -ExpandProperty Id
                        
                        if ($newCommentId) {
                            foreach ($reply in $replies) {
                                $replyAuthorEmail = $reply.Author.Email
                                
                                if ([string]::IsNullOrEmpty($replyAuthorEmail)) {
                                    Add-PnPCommentReply -List $destList -ItemId $newItem.Id -CommentId $newCommentId -Text "$($reply.Text) [Respuesta original de: $($reply.Author.LoginName) - $($reply.CreatedDate)]" -Connection $destPnpConnection
                                }
                                else {
                                    Add-PnPCommentReply -List $destList -ItemId $newItem.Id -CommentId $newCommentId -Text $reply.Text -Author $replyAuthorEmail -Connection $destPnpConnection
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-Log "Error al migrar comentario: $_" -Level "ERROR"
                }
            }
        }
        else {
            Write-Log "No se encontraron comentarios para este elemento"
        }
    }
    
    Write-Log "Migración completada exitosamente!"
}
catch {
    Write-Log "Error en el proceso de migración: $_" -Level "ERROR"
}
finally {
    # Desconectar ambas conexiones
    if ($sourcePnpConnection) {
        Disconnect-PnPOnline -Connection $sourcePnpConnection
    }
    if ($destPnpConnection) {
        Disconnect-PnPOnline -Connection $destPnpConnection
    }
    
    Write-Log "Proceso de migración finalizado. Revise el log para más detalles."
}
