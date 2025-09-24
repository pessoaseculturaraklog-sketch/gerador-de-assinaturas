[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$currentUser = $env:USERNAME

$signaturePath = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Signatures"

$networkPath = "\\srvrad\usuarios\bkp\Assinatura\arquivos_de_assinatura"

$backupPath = "C:\Users\$($currentUser)\assinaturas_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

$nome = "NOME_AQUI"
$cargo = "CARGO_AQUI"
$telefone = "TELEFONE_AQUI"

if (!(Test-Path -Path $signaturePath)) {
    New-Item -ItemType Directory -Path $signaturePath
}

$htmlFiles = Get-ChildItem -Path $signaturePath -Filter "*.htm*"
if ($htmlFiles) {
    $htmlFileName = $htmlFiles[0].Name
    $assinaturaNome = [System.IO.Path]::GetFileNameWithoutExtension($htmlFileName)
    $assinaturaArquivosNome = "$assinaturaNome" + "_arquivos"
} else {
    Write-Host "Nenhum arquivo HTML encontrado na pasta de assinaturas. Usando o nome padrão 'assinatura001'." -ForegroundColor Yellow
    $assinaturaNome = "assinatura001"
    $assinaturaArquivosNome = "assinatura001_arquivos"
}

New-Item -ItemType Directory -Path $backupPath

if (Test-Path -Path "$($signaturePath)\*") {
    try {
        Move-Item -Path "$($signaturePath)\*" -Destination $backupPath -Force -ErrorAction Stop
    } catch {
        Write-Host "Erro ao mover arquivos para o backup: $($_.Exception.Message)" -ForegroundColor Red
    }
}

if (Test-Path -Path "$($signaturePath)\*") {
    try {
        Remove-Item -Path "$($signaturePath)\*" -Force -ErrorAction Stop
    } catch {
        Write-Host "Erro ao remover arquivos da pasta de assinaturas: $($_.Exception.Message)" -ForegroundColor Red
    }
}

if (Test-Path -Path "$($networkPath)\*") {
    try {
        Copy-Item -Path "$($networkPath)\*" -Destination $signaturePath -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Host "Erro ao copiar arquivos. Verifique se o caminho da rede está acessível: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

Get-ChildItem -Path $signaturePath -Filter "mudarnome.*" | ForEach-Object {
    $newFileName = "$assinaturaNome$($_.Extension)"
    Rename-Item -Path $_.FullName -NewName $newFileName
}

$pastaArquivosTemp = Get-ChildItem -Path $signaturePath -Directory | Where-Object {$_.Name -eq "mudarnome_arquivos"}
if ($pastaArquivosTemp) {
    Rename-Item -Path $pastaArquivosTemp.FullName -NewName $assinaturaArquivosNome
}

$filelistPath = Join-Path -Path $signaturePath -ChildPath "$assinaturaArquivosNome\filelist.xml"
if (Test-Path -Path $filelistPath) {
    try {
        $filelistContent = Get-Content -Path $filelistPath -Raw
        $filelistContent = $filelistContent -replace 'TEMP_HTML_FILE_PLACEHOLDER', "$assinaturaNome"
        Set-Content -Path $filelistPath -Value $filelistContent -Encoding UTF8 -Force
    } catch {
        Write-Host "Erro ao atualizar o filelist.xml: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Get-ChildItem -Path $signaturePath | ForEach-Object {
    $file = $_
    $tempFile = Join-Path -Path $signaturePath -ChildPath "temp.$($file.Extension)"

    try {
        if ($file.Extension -eq ".htm") {
            $content = Get-Content -Path $file.FullName -Encoding UTF8
            $headTag = "<head>"
            $metaTag = '<meta charset="UTF-8">'
            $content = $content.Replace($headTag, "$headTag`n$metaTag")
            $content = $content -replace 'NOME_PLACEHOLDER', [System.Text.RegularExpressions.Regex]::Unescape($nome)
            $content = $content -replace 'CARGO_PLACEHOLDER', [System.Text.RegularExpressions.Regex]::Unescape($cargo)
            $content = $content -replace 'TELEFONE_PLACEHOLDER', [System.Text.RegularExpressions.Regex]::Unescape($telefone)
            $content = $content -replace 'TEMP_ASSINATURA_ARQUIVOS_PLACEHOLDER', $assinaturaArquivosNome
            Set-Content -Path $tempFile -Value $content -Encoding UTF8 -Force -ErrorAction Stop
        } elseif ($file.Extension -eq ".rtf") {
            $rtfContent = Get-Content -Path $file.FullName -Raw
            $rtfContent = $rtfContent -replace 'NOME_PLACEHOLDER', [System.Text.RegularExpressions.Regex]::Unescape($nome)
            $rtfContent = $rtfContent -replace 'CARGO_PLACEHOLDER', [System.Text.RegularExpressions.Regex]::Unescape($cargo)
            $rtfContent = $rtfContent -replace 'TELEFONE_PLACEHOLDER', [System.Text.RegularExpressions.Regex]::Unescape($telefone)
            $rtfContent = $rtfContent -replace 'TEMP_ASSINATURA_ARQUIVOS_PLACEHOLDER', $assinaturaArquivosNome
            Set-Content -Path $tempFile -Value $rtfContent -Encoding Default -Force -ErrorAction Stop
        } elseif ($file.Extension -eq ".txt") {
            $content = Get-Content -Path $file.FullName -Encoding UTF8
            $content = $content -replace 'NOME_PLACEHOLDER', [System.Text.RegularExpressions.Regex]::Unescape($nome)
            $content = $content -replace 'CARGO_PLACEHOLDER', [System.Text.RegularExpressions.Regex]::Unescape($cargo)
            $content = $content -replace 'TELEFONE_PLACEHOLDER', [System.Text.RegularExpressions.Regex]::Unescape($telefone)
            $content = $content -replace 'TEMP_ASSINATURA_ARQUIVOS_PLACEHOLDER', $assinaturaArquivosNome
            Set-Content -Path $tempFile -Value $content -Encoding UTF8 -Force -ErrorAction Stop
        }

        if (Test-Path $tempFile) {
            Copy-Item -Path $tempFile -Destination $file.FullName -Force -ErrorAction Stop
            Remove-Item -Path $tempFile -Force -ErrorAction Stop
        }
    } catch {
        Write-Host "Erro ao processar o arquivo $($file.FullName): $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "Assinatura atualizada com sucesso!" -ForegroundColor Green
