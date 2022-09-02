function ImportFiles($localPath, $spPath, $webUrl) {
    Get-ChildItem -Path $localPath | ForEach-Object {
        $isFolder = (Get-Item $_.FullName) -is [System.IO.DirectoryInfo]
        
        if ($isFolder) {
            Write-host -ForegroundColor Yellow "Creating folder '$($_.Name)'"
            $folder = m365 spo folder get --webUrl $webUrl `
                --folderUrl "$spPath/$($_.Name)" `
                | ConvertFrom-Json

            if ($null -eq $folder) {
                m365 spo folder add --webUrl $webUrl `
                    --parentFolderUrl $spPath `
                    --name $_.Name
                if (!$?) {
                    throw "It has not been possible to create the folder: $($_.Name)"
                }
            }

            Write-host -ForegroundColor Yellow "Importing folder '$($_.Name)'"
            ImportFiles -localPath $_.FullName `
                -spPath "$spPath/$($_.Name)" `
                -webUrl $webUrl
        }
        else {
            $pageName = "$($_.Name.Replace($_.Extension, '')).aspx"
            Write-host -ForegroundColor Yellow "Creating page '$($pageName)'"
            $pagePath = $spPath -eq '' ? $pageName : "$($spPath)/$($pageName)"
            $page = m365 spo page get --webUrl $webUrl --name $pagePath

            if ($null -ne $page) {
                m365 spo page remove --name $pageName --webUrl $webUrl --confirm
            }

            m365 spo page add --name $pagePath `
                --webUrl $webUrl `
                --publish
            
            $pageContent = Get-Content $_.FullName | Out-String

            $apiUrl = "https://api.github.com/markdown/raw"
            $htmlContent = $(Invoke-WebRequest -Method POST -Uri $apiUrl -Body $pageContent -ContentType "text/plain").Content
            
            $webpartData = @{
                "id" = "1ef5ed11-ce7b-44be-bc5e-4abd55101d16"
                "instanceId" = "9a477494-5aa5-4cc9-838d-bf98896adfc6"
                "title" = "Markdown"
                "description" = "Use Markdown language to add and format text."
                "audiences" = @()
                "serverProcessedContent" = @{
                    "htmlStrings" = @{
                        "html" = $htmlContent
                    }
                    "searchablePlainTexts" = @{
                        "code" = $pageContent
                    }
                    "imageSources" = @{}
                    "links" = @{}
                }
                "dataVersion" = "2.0"
                "properties" = @{
                    "displayPreview" = $true
                    "lineWrapping" = $true
                    "miniMap" = @{
                        "enabled" = $false
                    }
                    "previewState" = "Show"
                    "theme" = "Monokai"
                }
            } | ConvertTo-Json -Compress -EscapeHandling EscapeHtml

            $webpartData = $webpartData.Replace('"','""')
            m365 spo page clientsidewebpart add --webUrl $webUrl `
                --pageName $pageName `
                --webPartId 1ef5ed11-ce7b-44be-bc5e-4abd55101d16 `
                --webPartData $webpartData

            $fileUrl = "/sites/$($webUrl.Split('/sites/').TrimEnd('/')[1])/sitepages/$pagePath"
            m365 spo file checkout --webUrl $webUrl `
                --fileUrl $fileUrl
            m365 spo file checkin --webUrl $webUrl `
                --fileUrl $fileUrl `
                --type Major
            if (!$?) {
                throw "It has not been possible to import file: $($_.Name)"
            }
        }
    }
}

$ErrorActionPreference = "Stop"
try {
    Write-Host -ForegroundColor Yellow "Configuring documentation..."

    m365 login --authType certificate `
        --certificateBase64Encoded $env:CERTIFICATE_BASE64 `
        --appId $env:appId `
        --tenant $env:tenantId

    $webUrl = $env:siteUrl

    $scriptPath = $PSScriptRoot
    $directorySeparator = [System.IO.Path]::DirectorySeparatorChar
    $docsPath = "$($scriptPath)$($directorySeparator)..$($directorySeparator)docs"

    ImportFiles -localPath $docsPath `
        -spPath '' `
        -webUrl $webUrl

    Write-Host -ForegroundColor Green "Done."
}
catch {
    Write-Error "Error configuring the doc library on '$webUrl': $($_.Exception.Message)"
}
$ErrorActionPreference = "Continue"
