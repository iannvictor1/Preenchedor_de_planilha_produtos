param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,

    [string]$ProjectDir = "C:\sistema_produtos",
    [string]$BackendService = "",
    [string]$FrontendService = "",
    [string]$BackendTask = "",
    [string]$FrontendTask = "",
    [string]$ProjectTask = "Sistema Produtos",
    [string]$UserName = "",
    [switch]$StashLocalChanges,
    [switch]$RestoreLatestStash
)

$ErrorActionPreference = "Stop"

$credential = $null
if (![string]::IsNullOrWhiteSpace($UserName)) {
    $credential = Get-Credential -UserName $UserName -Message "Credenciais da maquina de producao"
}

$winRmPort = 5985
$connection = Test-NetConnection -ComputerName $ComputerName -Port $winRmPort -WarningAction SilentlyContinue
if (!$connection.TcpTestSucceeded) {
    throw "Nao foi possivel acessar $ComputerName na porta WinRM $winRmPort. Verifique se a maquina esta ligada, se o IP esta correto, se ela esta acessivel na rede e se o WinRM/firewall estao habilitados no computador de producao."
}

$script = {
    param($RemoteProjectDir, $RemoteBackendService, $RemoteFrontendService, $RemoteBackendTask, $RemoteFrontendTask, $RemoteProjectTask, $RemoteStashLocalChanges, $RemoteRestoreLatestStash)

    function Stop-RemoteScheduledTask($Name) {
        if ([string]::IsNullOrWhiteSpace($Name)) {
            return $false
        }

        $taskName = if ($Name.StartsWith("\")) { $Name } else { "\$Name" }
        Write-Host "Parando tarefa antes da atualizacao: $taskName" -ForegroundColor Yellow
        $output = schtasks.exe /End /TN $taskName 2>&1
        if ($LASTEXITCODE -eq 0) {
            return $true
        }

        $message = ($output | Out-String).Trim()
        if ($message -match "nao esta em execucao|not currently running") {
            return $false
        }
        if ($message -match "nao encontrada|cannot find|nao existe") {
            Write-Host "Tarefa '$taskName' nao encontrada; continuando." -ForegroundColor Yellow
            return $false
        }

        throw "Nao foi possivel parar a tarefa '$taskName'. Detalhe: $message"
    }

    function Start-RemoteScheduledTask($Name) {
        if ([string]::IsNullOrWhiteSpace($Name)) {
            return
        }

        $taskName = if ($Name.StartsWith("\")) { $Name } else { "\$Name" }
        Write-Host "Religando tarefa apos falha: $taskName" -ForegroundColor Yellow
        $output = schtasks.exe /Run /TN $taskName 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Nao foi possivel religar '$taskName': $(($output | Out-String).Trim())" -ForegroundColor Red
        }
    }

    function Stop-RemoteProjectProcesses() {
        function Stop-RemoteProcessTree($ProcessId, $ProcessTable) {
            $children = @($ProcessTable | Where-Object { $_.ParentProcessId -eq $ProcessId })
            foreach ($child in $children) {
                Stop-RemoteProcessTree -ProcessId $child.ProcessId -ProcessTable $ProcessTable
            }

            $process = $ProcessTable | Where-Object { $_.ProcessId -eq $ProcessId } | Select-Object -First 1
            if ($process) {
                Write-Host "Encerrando processo do sistema: $($process.Name) PID $ProcessId" -ForegroundColor Yellow
                Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
            }
        }

        $resolvedProjectDir = (Resolve-Path -LiteralPath $RemoteProjectDir).Path
        $processTable = @(Get-CimInstance Win32_Process)
        $targetIds = [System.Collections.Generic.HashSet[int]]::new()

        foreach ($process in $processTable) {
            if ($process.ProcessId -eq $PID) {
                continue
            }
            $commandLine = [string]$process.CommandLine
            $executablePath = [string]$process.ExecutablePath
            if (
                $commandLine -like "*$resolvedProjectDir*" -or
                $executablePath -like "$resolvedProjectDir*"
            ) {
                [void]$targetIds.Add([int]$process.ProcessId)
            }
        }

        foreach ($port in @(8000, 5173, 4173)) {
            Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction SilentlyContinue |
                ForEach-Object { [void]$targetIds.Add([int]$_.OwningProcess) }
        }

        foreach ($processId in @($targetIds)) {
            if (Get-Process -Id $processId -ErrorAction SilentlyContinue) {
                Stop-RemoteProcessTree -ProcessId $processId -ProcessTable $processTable
            }
        }

        Start-Sleep -Seconds 2
    }

    function Restore-LatestStashForRetry() {
        if (!$RemoteRestoreLatestStash) {
            return
        }

        $latest = & git stash list --format="%gd" -n 1 2>&1
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($latest)) {
            throw "Nenhum stash remoto foi encontrado para restaurar."
        }

        Write-Host "Restaurando stash da tentativa anterior: $latest" -ForegroundColor Yellow
        $output = & git stash pop $latest 2>&1
        if ($output) {
            $output | ForEach-Object { Write-Host $_ }
        }
        if ($LASTEXITCODE -ne 0) {
            throw "Nao foi possivel restaurar o stash $latest. Resolva o conflito diretamente na producao."
        }
    }

    function Get-ProtectedLocalFiles() {
        return @(
            "start-backend.cmd",
            "start-frontend.cmd",
            "start-sistema-produtos.cmd",
            "start-sistema-produtos-hidden.vbs",
            "admin.local.json",
            ".session_secret"
        )
    }

    function Recover-MissingProtectedFilesFromStash() {
        $stashRefs = @(& git stash list --format="%gd" 2>$null)
        foreach ($relativePath in Get-ProtectedLocalFiles) {
            $target = Join-Path $RemoteProjectDir $relativePath
            if (Test-Path -LiteralPath $target) {
                continue
            }

            foreach ($stashRef in $stashRefs) {
                $untrackedCommit = "$stashRef^3"
                & git cat-file -e "${untrackedCommit}:$relativePath" 2>$null
                if ($LASTEXITCODE -ne 0) {
                    continue
                }

                Write-Host "Recuperando arquivo local do stash $stashRef`: $relativePath" -ForegroundColor Yellow
                & git checkout $untrackedCommit -- $relativePath 2>&1 | Out-Null
                if ($LASTEXITCODE -ne 0) {
                    throw "Nao foi possivel recuperar '$relativePath' do stash $stashRef."
                }
                & git reset -- $relativePath 2>&1 | Out-Null
                break
            }
        }
    }

    function Backup-ProtectedLocalFiles() {
        $backupDir = Join-Path $env:TEMP "sistema-produtos-local-$(Get-Date -Format 'yyyyMMdd-HHmmssfff')"
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

        foreach ($relativePath in Get-ProtectedLocalFiles) {
            $source = Join-Path $RemoteProjectDir $relativePath
            if (Test-Path -LiteralPath $source) {
                Copy-Item -LiteralPath $source -Destination (Join-Path $backupDir $relativePath) -Force
            }
        }
        return $backupDir
    }

    function Restore-ProtectedLocalFiles($BackupDir) {
        if ([string]::IsNullOrWhiteSpace($BackupDir) -or !(Test-Path -LiteralPath $BackupDir)) {
            return
        }

        foreach ($relativePath in Get-ProtectedLocalFiles) {
            $source = Join-Path $BackupDir $relativePath
            if (Test-Path -LiteralPath $source) {
                Copy-Item -LiteralPath $source -Destination (Join-Path $RemoteProjectDir $relativePath) -Force
            }
        }
    }

    function Remove-ProtectedFilesBackup($BackupDir) {
        if (![string]::IsNullOrWhiteSpace($BackupDir) -and (Test-Path -LiteralPath $BackupDir)) {
            Remove-Item -LiteralPath $BackupDir -Recurse -Force
        }
    }

    function Backup-LocalChanges() {
        $status = & git status --porcelain --untracked-files=normal 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Nao foi possivel verificar alteracoes locais na producao."
        }

        if (!$status) {
            return $null
        }

        if (!$RemoteStashLocalChanges) {
            Write-Host "Alteracoes locais encontradas na producao:" -ForegroundColor Yellow
            $status | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
            throw "A producao possui alteracoes locais. Revise-as ou execute novamente com -StashLocalChanges para guarda-las em backup antes do pull."
        }

        $stashName = "backup-atualizacao-remota-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        Write-Host "Guardando alteracoes locais em git stash: $stashName" -ForegroundColor Yellow
        $stashOutput = & git stash push --include-untracked -m $stashName 2>&1
        $stashExitCode = $LASTEXITCODE
        if ($stashOutput) {
            $stashOutput | ForEach-Object { Write-Host $_ }
        }
        if ($stashExitCode -ne 0) {
            throw "Nao foi possivel criar o backup das alteracoes locais."
        }

        return "stash@{0}"
    }

    function Restore-LocalChanges($StashRef) {
        if ([string]::IsNullOrWhiteSpace($StashRef)) {
            return
        }

        Write-Host "Restaurando alteracoes locais apos falha: $StashRef" -ForegroundColor Yellow
        $output = & git stash pop $StashRef 2>&1
        if ($output) {
            $output | ForEach-Object { Write-Host $_ }
        }
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Nao foi possivel restaurar automaticamente o stash. Ele continua preservado em $StashRef." -ForegroundColor Red
        }
    }

    function Assert-FileNotLocked($RelativePath) {
        $path = Join-Path $RemoteProjectDir $RelativePath
        if (!(Test-Path -LiteralPath $path)) {
            return
        }

        try {
            $stream = [System.IO.File]::Open(
                $path,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None
            )
            $stream.Dispose()
        }
        catch {
            throw "O arquivo '$RelativePath' esta aberto ou bloqueado na producao. Feche o Excel e qualquer processo que esteja usando a planilha antes de atualizar."
        }
    }

    function Invoke-RemoteGitPull() {
        $previousErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = "Continue"

        try {
            $output = & git pull 2>&1
            $exitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $previousErrorActionPreference
        }

        if ($output) {
            $output | ForEach-Object { Write-Host $_ }
        }

        if ($exitCode -ne 0) {
            throw "git pull remoto falhou com codigo $exitCode."
        }
    }

    Set-Location -LiteralPath $RemoteProjectDir
    Restore-LatestStashForRetry
    Recover-MissingProtectedFilesFromStash

    $stoppedTasks = @()
    $createdStash = $null
    $protectedFilesBackup = Backup-ProtectedLocalFiles
    try {
        foreach ($taskName in @($RemoteProjectTask, $RemoteBackendTask, $RemoteFrontendTask) | Select-Object -Unique) {
            if (Stop-RemoteScheduledTask -Name $taskName) {
                $stoppedTasks += $taskName
            }
        }

        Stop-RemoteProjectProcesses
        Start-Sleep -Seconds 2
        Assert-FileNotLocked -RelativePath "Cópia de FICHA CADASTRO PRODUTO C&M.xlsx"
        $createdStash = Backup-LocalChanges

        Write-Host ""
        Write-Host "==> Sincronizando script de producao pelo Git" -ForegroundColor Cyan
        Invoke-RemoteGitPull
        Restore-ProtectedLocalFiles -BackupDir $protectedFilesBackup
    }
    catch {
        Restore-LocalChanges -StashRef $createdStash
        Restore-ProtectedLocalFiles -BackupDir $protectedFilesBackup
        foreach ($taskName in $stoppedTasks) {
            Start-RemoteScheduledTask -Name $taskName
        }
        throw
    }
    finally {
        Remove-ProtectedFilesBackup -BackupDir $protectedFilesBackup
    }

    & .\scripts\atualizar_producao.ps1 `
        -ProjectDir $RemoteProjectDir `
        -BackendService $RemoteBackendService `
        -FrontendService $RemoteFrontendService `
        -BackendTask $RemoteBackendTask `
        -FrontendTask $RemoteFrontendTask `
        -ProjectTask $RemoteProjectTask
}

$params = @{
    ComputerName = $ComputerName
    ScriptBlock = $script
    ArgumentList = @($ProjectDir, $BackendService, $FrontendService, $BackendTask, $FrontendTask, $ProjectTask, [bool]$StashLocalChanges, [bool]$RestoreLatestStash)
}

if ($credential) {
    $params.Credential = $credential
}

Invoke-Command @params
