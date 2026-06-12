param(
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot),
    [string]$BackendService = "",
    [string]$FrontendService = "",
    [string]$BackendTask = "",
    [string]$FrontendTask = ""
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

Write-Step "Entrando no projeto"
Set-Location -LiteralPath $ProjectDir

Write-Step "Atualizando codigo pelo Git"
Invoke-GitPull

Write-Step "Atualizando dependencias Python"
if (!(Test-Path -LiteralPath ".\.venv\Scripts\python.exe")) {
    python -m venv .venv
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

Write-Host ""
Write-Host "Atualizacao concluida." -ForegroundColor Green
