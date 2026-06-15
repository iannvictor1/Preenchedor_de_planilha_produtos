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
    [switch]$StashLocalChanges
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
    param($RemoteProjectDir, $RemoteBackendService, $RemoteFrontendService, $RemoteBackendTask, $RemoteFrontendTask, $RemoteProjectTask, $RemoteStashLocalChanges)

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

    function Backup-LocalChanges() {
        $status = & git status --porcelain --untracked-files=normal 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Nao foi possivel verificar alteracoes locais na producao."
        }

        if (!$status) {
            return
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

    $stoppedTasks = @()
    try {
        foreach ($taskName in @($RemoteProjectTask, $RemoteBackendTask, $RemoteFrontendTask) | Select-Object -Unique) {
            if (Stop-RemoteScheduledTask -Name $taskName) {
                $stoppedTasks += $taskName
            }
        }

        Start-Sleep -Seconds 2
        Backup-LocalChanges

        Write-Host ""
        Write-Host "==> Sincronizando script de producao pelo Git" -ForegroundColor Cyan
        Invoke-RemoteGitPull
    }
    catch {
        foreach ($taskName in $stoppedTasks) {
            Start-RemoteScheduledTask -Name $taskName
        }
        throw
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
    ArgumentList = @($ProjectDir, $BackendService, $FrontendService, $BackendTask, $FrontendTask, $ProjectTask, [bool]$StashLocalChanges)
}

if ($credential) {
    $params.Credential = $credential
}

Invoke-Command @params
