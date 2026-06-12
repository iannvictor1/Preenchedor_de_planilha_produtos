param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,

    [string]$ProjectDir = "C:\sistema_produtos",
    [string]$BackendService = "",
    [string]$FrontendService = "",
    [string]$BackendTask = "",
    [string]$FrontendTask = "",
    [string]$ProjectTask = "Sistema Produtos",
    [string]$UserName = ""
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
    param($RemoteProjectDir, $RemoteBackendService, $RemoteFrontendService, $RemoteBackendTask, $RemoteFrontendTask, $RemoteProjectTask)

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

    Write-Host ""
    Write-Host "==> Sincronizando script de producao pelo Git" -ForegroundColor Cyan
    Invoke-RemoteGitPull

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
    ArgumentList = @($ProjectDir, $BackendService, $FrontendService, $BackendTask, $FrontendTask, $ProjectTask)
}

if ($credential) {
    $params.Credential = $credential
}

Invoke-Command @params
