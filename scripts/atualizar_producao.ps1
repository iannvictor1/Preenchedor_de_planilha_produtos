param(
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot),
    [string]$BackendService = "",
    [string]$FrontendService = "",
    [string]$BackendTask = "",
    [string]$FrontendTask = "",
    [string]$ProjectTask = "Sistema Produtos",
    [switch]$NoAutoRestartTasks
)

$ErrorActionPreference = "Stop"

function Write-Step($Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Restart-OptionalService($Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return
    }

    Write-Step "Reiniciando servico: $Name"
    Restart-Service -Name $Name -Force
}

function Restart-OptionalScheduledTask($Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return
    }

    Write-Step "Reiniciando tarefa agendada: $Name"
    $taskName = if ($Name.StartsWith("\")) { $Name } else { "\$Name" }

    $queryOutput = schtasks.exe /Query /TN $taskName 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = ($queryOutput | Out-String).Trim()
        if ($message -match "Acesso negado|Access is denied") {
            throw "Acesso negado ao consultar a tarefa '$taskName'. Abra o PowerShell como Administrador e rode o update novamente."
        }

        throw "Tarefa agendada '$taskName' nao encontrada. Detalhe: $message"
    }

    $endOutput = schtasks.exe /End /TN $taskName 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = ($endOutput | Out-String).Trim()
        if ($message -notmatch "nao esta em execucao|not currently running") {
            if ($message -match "Acesso negado|Access is denied") {
                throw "Acesso negado ao parar a tarefa '$taskName'. Abra o PowerShell como Administrador e rode o update novamente."
            }

            throw "Nao foi possivel parar a tarefa agendada '$taskName'. Detalhe: $message"
        }
    }

    Start-Sleep -Seconds 2

    $runOutput = schtasks.exe /Run /TN $taskName 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = ($runOutput | Out-String).Trim()
        if ($message -match "Acesso negado|Access is denied") {
            throw "Acesso negado ao iniciar a tarefa '$taskName'. Abra o PowerShell como Administrador e rode o update novamente."
        }

        throw "Nao foi possivel iniciar a tarefa agendada '$taskName'. Detalhe: $message"
    }
}

function Get-NormalizedTaskName($Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    return $Name.Replace("/", "\").TrimStart("\").ToLowerInvariant()
}

function Get-TaskFullName($Task) {
    if ($Task.TaskPath.EndsWith("\")) {
        return "$($Task.TaskPath)$($Task.TaskName)"
    }

    return "$($Task.TaskPath)\$($Task.TaskName)"
}

function Restart-ProjectScheduledTasks($ProjectDir, $ExcludedNames) {
    if ($NoAutoRestartTasks) {
        return
    }

    $resolvedProjectDir = (Resolve-Path -LiteralPath $ProjectDir).Path
    $projectLeaf = Split-Path -Leaf $resolvedProjectDir
    $needles = @(
        $resolvedProjectDir,
        $resolvedProjectDir.Replace("\", "/"),
        $projectLeaf
    ) | Where-Object { ![string]::IsNullOrWhiteSpace($_) }

    $excluded = @($ExcludedNames | ForEach-Object { Get-NormalizedTaskName $_ })
    $matchingTasks = @()

    try {
        $matchingTasks = Get-ScheduledTask | Where-Object {
            $fullName = Get-TaskFullName $_
            if ($excluded -contains (Get-NormalizedTaskName $fullName) -or $excluded -contains (Get-NormalizedTaskName $_.TaskName)) {
                return $false
            }

            $actionText = ($_.Actions | ForEach-Object {
                "$($_.Execute) $($_.Arguments) $($_.WorkingDirectory)"
            }) -join " "

            $matched = $false
            foreach ($needle in $needles) {
                if ($actionText -like "*$needle*") {
                    $matched = $true
                    break
                }
            }

            return $matched
        }
    }
    catch {
        Write-Host "Nao foi possivel detectar tarefas agendadas automaticamente: $($_.Exception.Message)" -ForegroundColor Yellow
        return
    }

    if (!$matchingTasks -or $matchingTasks.Count -eq 0) {
        Write-Host ""
        Write-Host "Nenhuma tarefa agendada vinculada a '$resolvedProjectDir' foi encontrada para reinicio automatico." -ForegroundColor Yellow
        return
    }

    foreach ($task in $matchingTasks) {
        Restart-OptionalScheduledTask -Name (Get-TaskFullName $task)
    }
}

function Invoke-GitPull() {
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
        throw "git pull falhou com codigo $exitCode."
    }
}

function Invoke-NpmCommand($Arguments, $FailureMessage) {
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    try {
        $output = & npm @Arguments 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    if ($output) {
        $output | ForEach-Object { Write-Host $_ }
    }

    if ($exitCode -ne 0) {
        throw "$FailureMessage Codigo: $exitCode."
    }
}

function Get-SystemPython() {
    $candidates = @(
        @{ Command = "py"; Arguments = @("-3.12") },
        @{ Command = "py"; Arguments = @("-3") },
        @{ Command = "python"; Arguments = @() },
        @{ Command = "python3"; Arguments = @() },
        @{ Command = "$env:LocalAppData\Programs\Python\Python312\python.exe"; Arguments = @() },
        @{ Command = "$env:ProgramFiles\Python312\python.exe"; Arguments = @() },
        @{ Command = "${env:ProgramFiles(x86)}\Python312\python.exe"; Arguments = @() }
    )

    foreach ($candidate in $candidates) {
        if ($candidate.Command -like "*.exe" -and !(Test-Path -LiteralPath $candidate.Command)) {
            continue
        }

        $previousErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try {
            $output = & $candidate.Command @($candidate.Arguments) -c "import sys; print(sys.executable)" 2>$null
            $exitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $previousErrorActionPreference
        }

        if ($exitCode -eq 0 -and ![string]::IsNullOrWhiteSpace($output)) {
            return @{
                Command = $candidate.Command
                Arguments = $candidate.Arguments
            }
        }
    }

    throw "Python 3 nao foi encontrado na producao. Instale Python 3.12 em https://www.python.org/downloads/windows/ marcando 'Add python.exe to PATH', ou instale pelo winget: winget install -e --id Python.Python.3.12"
}

Write-Step "Entrando no projeto"
Set-Location -LiteralPath $ProjectDir

Write-Step "Atualizando codigo pelo Git"
Invoke-GitPull

Write-Step "Atualizando dependencias Python"
if (!(Test-Path -LiteralPath ".\.venv\Scripts\python.exe")) {
    $python = Get-SystemPython
    & $python.Command @($python.Arguments) -m venv .venv
}

.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

Write-Step "Atualizando frontend React"
Set-Location -LiteralPath (Join-Path $ProjectDir "frontend-react")
Invoke-NpmCommand @("install") "npm install falhou."
Invoke-NpmCommand @("run", "build") "npm run build falhou."

Set-Location -LiteralPath $ProjectDir

Restart-OptionalService -Name $BackendService
Restart-OptionalService -Name $FrontendService
Restart-OptionalScheduledTask -Name $BackendTask
Restart-OptionalScheduledTask -Name $FrontendTask
Restart-OptionalScheduledTask -Name $ProjectTask

if ([string]::IsNullOrWhiteSpace($ProjectTask)) {
    Restart-ProjectScheduledTasks -ProjectDir $ProjectDir -ExcludedNames @($BackendTask, $FrontendTask)
}

Write-Host ""
Write-Host "Atualizacao concluida." -ForegroundColor Green
