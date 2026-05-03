param(
  [switch]$SkipNodeInstall
)

$ErrorActionPreference = "Stop"

$NpmExe = "npm.cmd"
$VercelExe = "vercel.cmd"

try {
  [Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
  [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
} catch {}

function Write-Banner {
  Clear-Host
  Write-Host "==============================================" -ForegroundColor Cyan
  Write-Host "         XHTTPRelayECO Windows Installer      " -ForegroundColor Cyan
  Write-Host "==============================================" -ForegroundColor Cyan
  Write-Host ""
}

function Write-Step([string]$Text) {
  Write-Host ""
  Write-Host ">> $Text" -ForegroundColor Yellow
}

function Read-Default([string]$Prompt, [string]$DefaultValue) {
  $raw = Read-Host "$Prompt [$DefaultValue]"
  if ([string]::IsNullOrWhiteSpace($raw)) { return $DefaultValue }
  return $raw.Trim()
}

function Read-Optional([string]$Prompt) {
  $raw = Read-Host $Prompt
  if ([string]::IsNullOrWhiteSpace($raw)) { return "" }
  return $raw.Trim()
}

function Read-Required([string]$Prompt) {
  while ($true) {
    $raw = Read-Host $Prompt
    if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw.Trim() }
    Write-Host "Value is required." -ForegroundColor Red
  }
}

function Refresh-Path {
  $machine = [Environment]::GetEnvironmentVariable("Path", "Machine")
  $user = [Environment]::GetEnvironmentVariable("Path", "User")
  $env:Path = "$machine;$user"
}

function Get-TokenStorePath([string]$ProjectRoot) {
  return (Join-Path $ProjectRoot ".vercel-token.dpapi")
}

function Get-ScopeStorePath([string]$ProjectRoot) {
  return (Join-Path $ProjectRoot ".vercel-scope.txt")
}

function Save-Scope([string]$Scope, [string]$Path) {
  $value = ""
  if ($null -ne $Scope) { $value = $Scope.Trim() }
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($Path, $value, $utf8NoBom)
}

function Load-Scope([string]$Path) {
  if (-not (Test-Path $Path)) { return "" }
  try {
    $raw = Get-Content $Path -Raw
    if ($null -eq $raw) { return "" }
    return $raw.Trim()
  } catch {
    return ""
  }
}

function Save-TokenSecure([string]$Token, [string]$Path) {
  $secure = ConvertTo-SecureString -String $Token -AsPlainText -Force
  $text = ConvertFrom-SecureString -SecureString $secure
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($Path, $text, $utf8NoBom)
}

function Load-TokenSecure([string]$Path) {
  if (-not (Test-Path $Path)) { return "" }
  try {
    $text = (Get-Content $Path -Raw).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $secure = ConvertTo-SecureString -String $text
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
      return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
      [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
  } catch {
    return ""
  }
}

function Invoke-NativeSafe {
  param(
    [Parameter(Mandatory = $true)][string]$FilePath,
    [Parameter(Mandatory = $true)][string[]]$Arguments
  )

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $FilePath
  $psi.WorkingDirectory = (Get-Location).Path
  $psi.UseShellExecute = $false
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true

  # PowerShell 5.1-compatible argument handling
  $escaped = $Arguments | ForEach-Object {
    if ($_ -match '[\s"]') {
      '"' + ($_ -replace '"', '\"') + '"'
    } else {
      $_
    }
  }
  $psi.Arguments = ($escaped -join ' ')

  $proc = New-Object System.Diagnostics.Process
  $proc.StartInfo = $psi
  [void]$proc.Start()

  $stdout = $proc.StandardOutput.ReadToEnd()
  $stderr = $proc.StandardError.ReadToEnd()
  $proc.WaitForExit()

  $lines = @()
  if (-not [string]::IsNullOrWhiteSpace($stdout)) {
    $lines += ($stdout -split "`r?`n" | Where-Object { $_ -ne "" })
  }
  if (-not [string]::IsNullOrWhiteSpace($stderr)) {
    $lines += ($stderr -split "`r?`n" | Where-Object { $_ -ne "" })
  }

  return @{
    Output = @($lines)
    ExitCode = $proc.ExitCode
  }
}

function New-RandomProjectName {
  $chars = "abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray()
  $suffix = -join (1..8 | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
  return "relay-$suffix"
}

function Ensure-NodeAndNpm {
  if (Get-Command $NpmExe -ErrorAction SilentlyContinue) {
    Write-Host "npm already installed." -ForegroundColor Green
    return
  }

  if ($SkipNodeInstall) {
    throw "npm is missing and -SkipNodeInstall was used."
  }

  if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
    throw "winget not found. Install Node.js LTS manually and run again."
  }

  Write-Step "Installing Node.js LTS (npm included) via winget..."
  winget install --id OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements
  Refresh-Path

  if (-not (Get-Command $NpmExe -ErrorAction SilentlyContinue)) {
    throw "Node.js installation finished but npm is still not detected. Re-open PowerShell and retry."
  }
}

function Ensure-VercelCli {
  if (Get-Command $VercelExe -ErrorAction SilentlyContinue) {
    Write-Host "Vercel CLI already installed." -ForegroundColor Green
    return
  }

  Write-Step "Installing Vercel CLI..."
  & $NpmExe i -g vercel | Out-Host
  Refresh-Path
  if (-not (Get-Command $VercelExe -ErrorAction SilentlyContinue)) {
    throw "vercel command not found after installation."
  }
}

function Start-VercelOobLogin([string]$OutputDir) {
  Write-Host "Starting manual device login (no auto browser)..." -ForegroundColor Yellow
  $prevBrowser = $env:BROWSER
  $env:BROWSER = "none"
  try {
    $loginResult = Invoke-NativeSafe -FilePath $VercelExe -Arguments @("login", "--oob", "--no-color")
  } finally {
    if ($null -eq $prevBrowser) {
      Remove-Item Env:\BROWSER -ErrorAction SilentlyContinue
    } else {
      $env:BROWSER = $prevBrowser
    }
  }
  $loginResult.Output | Out-Host
  if ($loginResult.ExitCode -ne 0) {
    throw "vercel login failed."
  }

  $urls = @()
  foreach ($line in $loginResult.Output) {
    $matches = [regex]::Matches($line, 'https?://[^\s\)\]]+')
    foreach ($m in $matches) {
      if ($m.Value -match 'vercel\.com/(oauth/device|device)') {
        $urls += $m.Value
      }
    }
  }
  $urls = $urls | Select-Object -Unique

  if ($urls.Count -gt 0) {
    $txtPath = Join-Path $OutputDir "vercel-login-link.txt"
    $content = @(
      "Vercel Login Links"
      "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
      ""
    ) + $urls
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($txtPath, $content, $utf8NoBom)

    Write-Host ""
    Write-Host "Manual login URL(s):" -ForegroundColor Cyan
    $urls | ForEach-Object { Write-Host "  $_" -ForegroundColor Cyan }
    Write-Host "Saved to: $txtPath" -ForegroundColor Cyan
  } else {
    Write-Host "No explicit login URL found in output. If prompted, use: https://vercel.com/device" -ForegroundColor DarkYellow
  }
}

function Ensure-VercelLogin([string]$OutputDir, [string]$TokenStorePath) {
  Write-Step "Checking Vercel login..."

  Write-Host "Choose auth mode:" -ForegroundColor Cyan
  Write-Host "[1] Use existing login session (default)"
  Write-Host "[2] Token mode (recommended: never opens browser, supports secure save)"
  $authMode = Read-Default "Select auth mode" "1"

  if ($authMode -eq "2" -or $authMode -eq "3") {
    $token = ""
    $saved = Load-TokenSecure -Path $TokenStorePath
    if (-not [string]::IsNullOrWhiteSpace($saved)) {
      $useSaved = Read-Default "Use saved encrypted token from project folder? (Y/n)" "y"
      if ($useSaved.ToLowerInvariant() -ne "n") {
        $token = $saved
        Write-Host "Using saved encrypted token." -ForegroundColor Green
      }
    }

    if ([string]::IsNullOrWhiteSpace($token)) {
      $token = Read-Required "Paste Vercel token (create from Vercel dashboard -> Settings -> Tokens)"
      $saveNow = Read-Default "Save token encrypted in this project folder? (Y/n)" "y"
      if ($saveNow.ToLowerInvariant() -ne "n") {
        Save-TokenSecure -Token $token -Path $TokenStorePath
        Write-Host "Token saved securely: $TokenStorePath" -ForegroundColor Green
      }
    }

    $env:VERCEL_TOKEN = $token
    $tokenWhoami = Invoke-NativeSafe -FilePath $VercelExe -Arguments @("whoami", "--token", $token)
    if ($tokenWhoami.ExitCode -ne 0) {
      throw "Token auth failed. Check token and retry."
    }
    $tokenWhoami.Output | Out-Host
    return
  }

  $whoamiResult = Invoke-NativeSafe -FilePath $VercelExe -Arguments @("whoami")
  $loggedIn = $whoamiResult.ExitCode -eq 0

  if ($authMode -eq "1" -and $loggedIn) {
    $whoamiResult.Output | Out-Host
    $useCurrent = Read-Default "Use current logged-in session? (Y/n)" "y"
    if ($useCurrent.ToLowerInvariant() -eq "n") {
      Write-Step "Logging out and creating a fresh login link..."
      $logoutResult = Invoke-NativeSafe -FilePath $VercelExe -Arguments @("logout")
      $logoutResult.Output | Out-Host
      Start-VercelOobLogin -OutputDir $OutputDir
    }
  } else {
    Start-VercelOobLogin -OutputDir $OutputDir
  }

  & $VercelExe whoami | Out-Host
}

function Resolve-SessionScope([string]$ScopeStorePath) {
  $savedScope = Load-Scope -Path $ScopeStorePath
  if (-not [string]::IsNullOrWhiteSpace($savedScope)) {
    $useSaved = Read-Default ("Use saved scope/team '{0}'? (Y/n)" -f $savedScope) "y"
    if ($useSaved.ToLowerInvariant() -ne "n") {
      return $savedScope
    }
  }

  Write-Host ""
  Write-Host "Scope note: enter your Vercel team slug to avoid wrong CLI context." -ForegroundColor DarkYellow
  $scope = Read-Optional "Scope slug/team (optional, press Enter for personal account)"
  if (-not [string]::IsNullOrWhiteSpace($scope)) {
    Save-Scope -Scope $scope -Path $ScopeStorePath
  }
  return $scope
}

function Ensure-VercelProject([string]$ProjectName, [string]$Scope) {
  Write-Step "Creating project (or reusing if already exists)..."
  $args = @("project", "add", $ProjectName)
  if (-not [string]::IsNullOrWhiteSpace($Scope)) { $args += @("--scope", $Scope) }

  $result = Invoke-NativeSafe -FilePath $VercelExe -Arguments $args
  $output = $result.Output
  $text = $output | Out-String
  if ($result.ExitCode -ne 0 -and ($text -notmatch "already exists")) {
    throw "vercel project add failed: $text"
  }
  $output | Out-Host
}

function Link-VercelProject([string]$ProjectName, [string]$Scope) {
  Write-Step "Linking local folder to Vercel project..."
  $args = @("link", "--yes", "--project", $ProjectName)
  if (-not [string]::IsNullOrWhiteSpace($Scope)) { $args += @("--scope", $Scope) }
  & $VercelExe @args | Out-Host
  if ($LASTEXITCODE -ne 0) {
    throw "vercel link failed."
  }
}

function Set-VercelEnv([string]$Name, [string]$Value, [string]$Target, [string]$Scope) {
  $args = @("env", "add", $Name, $Target, "--value", $Value, "--force", "--yes", "--no-sensitive")
  if (-not [string]::IsNullOrWhiteSpace($Scope)) { $args += @("--scope", $Scope) }
  & $VercelExe @args | Out-Host
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to set env var $Name for $Target."
  }
}

function New-RandomPackageName {
  $chars = "abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray()
  $suffix = -join (1..10 | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
  return "host-$suffix"
}

function New-RandomPackageVersion {
  $major = Get-Random -Minimum 1 -Maximum 4
  $minor = Get-Random -Minimum 0 -Maximum 20
  $patch = Get-Random -Minimum 0 -Maximum 30
  return "$major.$minor.$patch"
}

function New-RandomPackageDescription {
  $descriptions = @(
    "Lightweight hosting edge relay for low-bandwidth delivery",
    "Optimized download gateway for shared hosting workloads",
    "Traffic-shaped relay runtime for static and media hosting",
    "Resource-friendly transfer bridge for multi-tenant hosting",
    "Adaptive download routing layer for budget hosting plans",
    "Low-overhead HTTP delivery relay for content hosting",
    "Bandwidth-governed relay node for file delivery services",
    "Edge proxy core for controlled-speed hosting and downloads"
  )
  return ($descriptions | Get-Random)
}

function New-RandomVercelConfigName {
  $chars = "abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray()
  $suffix = -join (1..10 | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
  return "edge-$suffix"
}

function Prepare-RandomizedPackageMetadataForDeploy {
  $pkgPath = Join-Path $scriptDir "package.json"
  if (-not (Test-Path $pkgPath)) {
    return @{
      Modified = $false
      PackagePath = $pkgPath
      OriginalContent = ""
    }
  }

  $original = Get-Content -Path $pkgPath -Raw
  try {
    $obj = $original | ConvertFrom-Json
  } catch {
    Write-Host "package.json parse failed; deploying without metadata randomization." -ForegroundColor DarkYellow
    return @{
      Modified = $false
      PackagePath = $pkgPath
      OriginalContent = $original
    }
  }

  $randomName = New-RandomPackageName
  $randomVersion = New-RandomPackageVersion
  $randomDesc = New-RandomPackageDescription

  if ($obj.PSObject.Properties.Name -contains "name") {
    $obj.name = $randomName
  } else {
    $obj | Add-Member -NotePropertyName "name" -NotePropertyValue $randomName
  }
  if ($obj.PSObject.Properties.Name -contains "version") {
    $obj.version = $randomVersion
  } else {
    $obj | Add-Member -NotePropertyName "version" -NotePropertyValue $randomVersion
  }
  if ($obj.PSObject.Properties.Name -contains "description") {
    $obj.description = $randomDesc
  } else {
    $obj | Add-Member -NotePropertyName "description" -NotePropertyValue $randomDesc
  }

  $json = $obj | ConvertTo-Json -Depth 50
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($pkgPath, $json, $utf8NoBom)

  Write-Host ("Randomized package metadata for this deploy: name={0}, version={1}" -f $randomName, $randomVersion) -ForegroundColor DarkGray
  Write-Host ("Description: {0}" -f $randomDesc) -ForegroundColor DarkGray

  return @{
    Modified = $true
    PackagePath = $pkgPath
    OriginalContent = $original
  }
}

function Restore-PackageMetadataAfterDeploy($state) {
  if ($null -eq $state) { return }
  if (-not $state.Modified) { return }
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($state.PackagePath, [string]$state.OriginalContent, $utf8NoBom)
  Write-Host "package.json restored to original local content." -ForegroundColor DarkGray
}

function Prepare-RandomizedVercelConfigForDeploy {
  $vercelPath = Join-Path $scriptDir "vercel.json"
  if (-not (Test-Path $vercelPath)) {
    return @{
      Modified = $false
      ConfigPath = $vercelPath
      OriginalContent = ""
    }
  }

  $original = Get-Content -Path $vercelPath -Raw
  try {
    $obj = $original | ConvertFrom-Json
  } catch {
    Write-Host "vercel.json parse failed; deploying without vercel.json randomization." -ForegroundColor DarkYellow
    return @{
      Modified = $false
      ConfigPath = $vercelPath
      OriginalContent = $original
    }
  }

  $randomName = New-RandomVercelConfigName
  if ($obj.PSObject.Properties.Name -contains "name") {
    $obj.name = $randomName
  } else {
    $obj | Add-Member -NotePropertyName "name" -NotePropertyValue $randomName
  }

  $json = $obj | ConvertTo-Json -Depth 50
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($vercelPath, $json, $utf8NoBom)

  Write-Host ("Randomized vercel.json name for this deploy: {0}" -f $randomName) -ForegroundColor DarkGray

  return @{
    Modified = $true
    ConfigPath = $vercelPath
    OriginalContent = $original
  }
}

function Restore-VercelConfigAfterDeploy($state) {
  if ($null -eq $state) { return }
  if (-not $state.Modified) { return }
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($state.ConfigPath, [string]$state.OriginalContent, $utf8NoBom)
  Write-Host "vercel.json restored to original local content." -ForegroundColor DarkGray
}

function Deploy-Production([string]$Scope) {
  Write-Step "Deploying to production..."
  $randomizeState = Prepare-RandomizedPackageMetadataForDeploy
  $vercelRandomizeState = Prepare-RandomizedVercelConfigForDeploy
  try {
    $args = @("deploy", "--prod", "--yes")
    if (-not [string]::IsNullOrWhiteSpace($Scope)) { $args += @("--scope", $Scope) }

    $result = Invoke-NativeSafe -FilePath $VercelExe -Arguments $args
    $lines = $result.Output
    $lines | Out-Host
    if ($result.ExitCode -ne 0) {
      throw "vercel deploy failed."
    }

    $alias = ""
    $prod = ""
    foreach ($line in $lines) {
      if ($line -match "Aliased:\s*(https://\S+)") { $alias = $Matches[1] }
      if ($line -match "Production:\s*(https://\S+)") { $prod = $Matches[1] }
    }

    return @{
      Alias = $alias
      Production = $prod
    }
  } finally {
    Restore-VercelConfigAfterDeploy -state $vercelRandomizeState
    Restore-PackageMetadataAfterDeploy -state $randomizeState
  }
}

function Get-LinkedProjectInfo([string]$ProjectRoot) {
  $projectFile = Join-Path $ProjectRoot ".vercel\project.json"
  if (-not (Test-Path $projectFile)) {
    return @{
      IsLinked = $false
      ProjectName = ""
      ProjectId = ""
      Scope = ""
    }
  }

  try {
    $obj = Get-Content $projectFile -Raw | ConvertFrom-Json
  } catch {
    return @{
      IsLinked = $false
      ProjectName = ""
      ProjectId = ""
      Scope = ""
    }
  }

  $name = ""
  if ($obj.PSObject.Properties.Name -contains "projectName" -and $obj.projectName) { $name = [string]$obj.projectName }
  $projectId = ""
  if ($obj.PSObject.Properties.Name -contains "projectId" -and $obj.projectId) { $projectId = [string]$obj.projectId }
  $scope = ""
  if ($obj.PSObject.Properties.Name -contains "orgId" -and $obj.orgId) { $scope = [string]$obj.orgId }

  return @{
    IsLinked = $true
    ProjectName = $name
    ProjectId = $projectId
    Scope = $scope
  }
}

function Show-DeploySummary($deployInfo) {
  Write-Host ""
  Write-Host "==============================================" -ForegroundColor Green
  Write-Host "Deployment complete." -ForegroundColor Green
  if ($deployInfo.Production) { Write-Host "Production: $($deployInfo.Production)" -ForegroundColor Green }
  if ($deployInfo.Alias) { Write-Host "Aliased:    $($deployInfo.Alias)" -ForegroundColor Green }
  Write-Host "==============================================" -ForegroundColor Green
  Write-Host ""
}

function Collect-NewDeploymentConfig([string]$DefaultScope) {
  Write-Step "Collecting config values..."
  $projectNameInput = Read-Optional "Project name on Vercel (leave empty for random)"
  $projectName = if ([string]::IsNullOrWhiteSpace($projectNameInput)) { New-RandomProjectName } else { $projectNameInput }
  if ([string]::IsNullOrWhiteSpace($DefaultScope)) {
    $scope = Read-Host "Scope slug/team (optional, press Enter to skip)"
  } else {
    $scope = Read-Default "Scope slug/team" $DefaultScope
  }
  $scope = $scope.Trim()
  $targetDomain = Read-Required "TARGET_DOMAIN (example: https://your-upstream-domain:443)"
  $relayPath = Read-Default "RELAY_PATH (MUST be EXACT inbound path on your foreign server, e.g. /api or /freedom)" "/api"
  $publicRelayPath = Read-Default "PUBLIC_RELAY_PATH (public endpoint on this domain)" "/api"
  $landingTemplate = Read-Optional "LANDING_TEMPLATE (optional template folder name; empty = random each build)"
  $maxInflight = Read-Default "MAX_INFLIGHT" "128"
  $maxUpBps = Read-Default "MAX_UP_BPS" "2621440"
  $maxDownBps = Read-Default "MAX_DOWN_BPS" "2621440"
  $upstreamTimeoutMs = Read-Default "UPSTREAM_TIMEOUT_MS" "50000"
  $successLogSampleRate = Read-Default "SUCCESS_LOG_SAMPLE_RATE" "0"
  $successLogMinDurationMs = Read-Default "SUCCESS_LOG_MIN_DURATION_MS" "3000"
  $errorLogMinIntervalMs = Read-Default "ERROR_LOG_MIN_INTERVAL_MS" "5000"

  if (-not $relayPath.StartsWith("/")) { $relayPath = "/$relayPath" }
  if (-not $publicRelayPath.StartsWith("/")) { $publicRelayPath = "/$publicRelayPath" }

  Write-Step "Environment values selected:"
  Write-Host "TARGET_DOMAIN = $targetDomain"
  Write-Host "PROJECT_NAME  = $projectName"
  Write-Host "RELAY_PATH    = $relayPath"
  Write-Host "PUBLIC_RELAY_PATH = $publicRelayPath"
  if (-not [string]::IsNullOrWhiteSpace($landingTemplate)) { Write-Host "LANDING_TEMPLATE = $landingTemplate" } else { Write-Host "LANDING_TEMPLATE = (random)" }
  Write-Host "MAX_INFLIGHT  = $maxInflight"
  Write-Host "MAX_UP_BPS    = $maxUpBps"
  Write-Host "MAX_DOWN_BPS  = $maxDownBps"
  Write-Host "UPSTREAM_TIMEOUT_MS        = $upstreamTimeoutMs"
  Write-Host "SUCCESS_LOG_SAMPLE_RATE    = $successLogSampleRate"
  Write-Host "SUCCESS_LOG_MIN_DURATION_MS= $successLogMinDurationMs"
  Write-Host "ERROR_LOG_MIN_INTERVAL_MS  = $errorLogMinIntervalMs"

  return @{
    ProjectName = $projectName
    Scope = $scope
    TargetDomain = $targetDomain
    RelayPath = $relayPath
    PublicRelayPath = $publicRelayPath
    LandingTemplate = $landingTemplate
    MaxInflight = $maxInflight
    MaxUpBps = $maxUpBps
    MaxDownBps = $maxDownBps
    UpstreamTimeoutMs = $upstreamTimeoutMs
    SuccessLogSampleRate = $successLogSampleRate
    SuccessLogMinDurationMs = $successLogMinDurationMs
    ErrorLogMinIntervalMs = $errorLogMinIntervalMs
  }
}

function Apply-ProductionEnv($cfg) {
  Write-Step "Setting environment variables for production..."
  Set-VercelEnv -Name "TARGET_DOMAIN" -Value $cfg.TargetDomain -Target "production" -Scope $cfg.Scope
  Set-VercelEnv -Name "RELAY_PATH" -Value $cfg.RelayPath -Target "production" -Scope $cfg.Scope
  Set-VercelEnv -Name "PUBLIC_RELAY_PATH" -Value $cfg.PublicRelayPath -Target "production" -Scope $cfg.Scope
  if (-not [string]::IsNullOrWhiteSpace($cfg.LandingTemplate)) {
    Set-VercelEnv -Name "LANDING_TEMPLATE" -Value $cfg.LandingTemplate -Target "production" -Scope $cfg.Scope
  }
  Set-VercelEnv -Name "MAX_INFLIGHT" -Value $cfg.MaxInflight -Target "production" -Scope $cfg.Scope
  Set-VercelEnv -Name "MAX_UP_BPS" -Value $cfg.MaxUpBps -Target "production" -Scope $cfg.Scope
  Set-VercelEnv -Name "MAX_DOWN_BPS" -Value $cfg.MaxDownBps -Target "production" -Scope $cfg.Scope
  Set-VercelEnv -Name "UPSTREAM_TIMEOUT_MS" -Value $cfg.UpstreamTimeoutMs -Target "production" -Scope $cfg.Scope
  Set-VercelEnv -Name "SUCCESS_LOG_SAMPLE_RATE" -Value $cfg.SuccessLogSampleRate -Target "production" -Scope $cfg.Scope
  Set-VercelEnv -Name "SUCCESS_LOG_MIN_DURATION_MS" -Value $cfg.SuccessLogMinDurationMs -Target "production" -Scope $cfg.Scope
  Set-VercelEnv -Name "ERROR_LOG_MIN_INTERVAL_MS" -Value $cfg.ErrorLogMinIntervalMs -Target "production" -Scope $cfg.Scope
}

function Run-NewDeploymentFlow([string]$DefaultScope) {
  $cfg = Collect-NewDeploymentConfig -DefaultScope $DefaultScope
  Ensure-VercelProject -ProjectName $cfg.ProjectName -Scope $cfg.Scope
  Link-VercelProject -ProjectName $cfg.ProjectName -Scope $cfg.Scope
  Apply-ProductionEnv -cfg $cfg
  $deployInfo = Deploy-Production -Scope $cfg.Scope
  Show-DeploySummary $deployInfo
  Write-Host "Done."
}

function Run-UpdateEnvFlow([string]$Scope) {
  Write-Step "Update production env vars (required + economic defaults)..."
  $targetDomain = Read-Required "TARGET_DOMAIN"
  $relayPath = Read-Required "RELAY_PATH (inbound path on foreign server)"
  $publicRelayPath = Read-Default "PUBLIC_RELAY_PATH (public endpoint on this domain)" "/api"
  $landingTemplate = Read-Optional "LANDING_TEMPLATE (optional template folder name; empty keeps previous/random)"
  $maxInflight = Read-Default "MAX_INFLIGHT" "128"
  $maxUpBps = Read-Default "MAX_UP_BPS" "2621440"
  $maxDownBps = Read-Default "MAX_DOWN_BPS" "2621440"
  $upstreamTimeoutMs = Read-Default "UPSTREAM_TIMEOUT_MS" "50000"
  $successLogSampleRate = Read-Default "SUCCESS_LOG_SAMPLE_RATE" "0"
  $successLogMinDurationMs = Read-Default "SUCCESS_LOG_MIN_DURATION_MS" "3000"
  $errorLogMinIntervalMs = Read-Default "ERROR_LOG_MIN_INTERVAL_MS" "5000"

  if (-not $relayPath.StartsWith("/")) { $relayPath = "/$relayPath" }
  if (-not $publicRelayPath.StartsWith("/")) { $publicRelayPath = "/$publicRelayPath" }

  Set-VercelEnv -Name "TARGET_DOMAIN" -Value $targetDomain -Target "production" -Scope $Scope
  Set-VercelEnv -Name "RELAY_PATH" -Value $relayPath -Target "production" -Scope $Scope
  Set-VercelEnv -Name "PUBLIC_RELAY_PATH" -Value $publicRelayPath -Target "production" -Scope $Scope
  if (-not [string]::IsNullOrWhiteSpace($landingTemplate)) {
    Set-VercelEnv -Name "LANDING_TEMPLATE" -Value $landingTemplate -Target "production" -Scope $Scope
  }
  Set-VercelEnv -Name "MAX_INFLIGHT" -Value $maxInflight -Target "production" -Scope $Scope
  Set-VercelEnv -Name "MAX_UP_BPS" -Value $maxUpBps -Target "production" -Scope $Scope
  Set-VercelEnv -Name "MAX_DOWN_BPS" -Value $maxDownBps -Target "production" -Scope $Scope
  Set-VercelEnv -Name "UPSTREAM_TIMEOUT_MS" -Value $upstreamTimeoutMs -Target "production" -Scope $Scope
  Set-VercelEnv -Name "SUCCESS_LOG_SAMPLE_RATE" -Value $successLogSampleRate -Target "production" -Scope $Scope
  Set-VercelEnv -Name "SUCCESS_LOG_MIN_DURATION_MS" -Value $successLogMinDurationMs -Target "production" -Scope $Scope
  Set-VercelEnv -Name "ERROR_LOG_MIN_INTERVAL_MS" -Value $errorLogMinIntervalMs -Target "production" -Scope $Scope

  $redeployNow = Read-Default "Redeploy now? (Y/n)" "y"
  if ($redeployNow.ToLowerInvariant() -eq "y") {
    $deployInfo = Deploy-Production -Scope $Scope
    Show-DeploySummary $deployInfo
  }
}

function Show-DeploymentList([string]$ProjectName, [string]$Scope) {
  Write-Step "Recent deployments..."
  $args = @("list")
  if (-not [string]::IsNullOrWhiteSpace($ProjectName)) { $args += @($ProjectName) }
  if (-not [string]::IsNullOrWhiteSpace($Scope)) { $args += @("--scope", $Scope) }
  $result = Invoke-NativeSafe -FilePath $VercelExe -Arguments $args
  $result.Output | Out-Host
  if ($result.ExitCode -ne 0) {
    Write-Host "Could not list deployments with scoped project. Trying generic list..." -ForegroundColor DarkYellow
    $fallback = Invoke-NativeSafe -FilePath $VercelExe -Arguments @("list")
    $fallback.Output | Out-Host
  }
}

function Ensure-LinkedToProject([string]$ProjectName, [string]$Scope) {
  if ([string]::IsNullOrWhiteSpace($ProjectName)) {
    throw "Project name is required."
  }
  Link-VercelProject -ProjectName $ProjectName -Scope $Scope
}

function Parse-ProjectListText([string[]]$Lines) {
  $projects = @()
  foreach ($line in $Lines) {
    if ($null -eq $line) { continue }
    $clean = [regex]::Replace([string]$line, '\x1B\[[0-9;]*[A-Za-z]', '')
    $trim = $clean.Trim()
    if ([string]::IsNullOrWhiteSpace($trim)) { continue }
    if ($trim.StartsWith(">")) { continue }
    if ($trim.StartsWith("-")) { continue }
    if ($trim -match '^(Projects|Name|Updated|ID|Inspect|No projects|Fetching|Retrieving|Error:|Warning:|Visit )') { continue }
    if ($trim -match '^https?://') { continue }

    $name = ($trim -split '\s+')[0]
    if ($name -match '^[a-zA-Z0-9][a-zA-Z0-9\-_\.]+$') {
      $projects += [PSCustomObject]@{
        Name = $name
        Id = ""
      }
    }
  }
  return @($projects | Sort-Object Name -Unique)
}

function Get-ProjectsFromVercel([string]$Scope) {
  $args = @("project", "list", "--format", "json", "--no-color")
  if (-not [string]::IsNullOrWhiteSpace($Scope)) { $args += @("--scope", $Scope) }
  $result = Invoke-NativeSafe -FilePath $VercelExe -Arguments $args

  if ($result.ExitCode -eq 0) {
    $raw = ($result.Output -join "`n")
    $raw = $raw -replace "`0", ""
    $raw = $raw.Trim()

    if (-not [string]::IsNullOrWhiteSpace($raw)) {
      # Try to isolate JSON payload even if CLI prints extra lines before/after.
      $jsonCandidate = $raw
      $firstArray = $raw.IndexOf("[")
      $lastArray = $raw.LastIndexOf("]")
      $firstObject = $raw.IndexOf("{")
      $lastObject = $raw.LastIndexOf("}")

      if ($firstArray -ge 0 -and $lastArray -gt $firstArray) {
        $jsonCandidate = $raw.Substring($firstArray, $lastArray - $firstArray + 1)
      } elseif ($firstObject -ge 0 -and $lastObject -gt $firstObject) {
        $jsonCandidate = $raw.Substring($firstObject, $lastObject - $firstObject + 1)
      }

      try {
        $parsed = $jsonCandidate | ConvertFrom-Json

        $items = @()
        if ($parsed -is [System.Array]) {
          $items = $parsed
        } elseif ($parsed.PSObject.Properties.Name -contains "projects") {
          $items = @($parsed.projects)
        } else {
          $items = @($parsed)
        }

        $projects = @()
        foreach ($item in $items) {
          $name = ""
          if ($item.PSObject.Properties.Name -contains "name" -and $item.name) { $name = [string]$item.name }
          if ([string]::IsNullOrWhiteSpace($name)) { continue }

          $projectId = ""
          if ($item.PSObject.Properties.Name -contains "id" -and $item.id) { $projectId = [string]$item.id }
          $projects += [PSCustomObject]@{
            Name = $name
            Id = $projectId
          }
        }

        if ($projects.Count -gt 0) {
          return @($projects | Sort-Object Name -Unique)
        }
      } catch {
        # ignore and continue to text parsing fallbacks below
      }
    }

    $parsedText = Parse-ProjectListText -Lines $result.Output
    if ($parsedText.Count -gt 0) { return $parsedText }
  }

  # Fallback for CLI variants/versions where JSON format is unavailable.
  $fallbackCommands = @(
    @("project", "list", "--no-color"),
    @("projects", "list", "--no-color"),
    @("project", "ls", "--no-color"),
    @("projects", "ls", "--no-color")
  )

  foreach ($cmd in $fallbackCommands) {
    $fallbackArgs = @($cmd)
    if (-not [string]::IsNullOrWhiteSpace($Scope)) { $fallbackArgs += @("--scope", $Scope) }
    $fallback = Invoke-NativeSafe -FilePath $VercelExe -Arguments $fallbackArgs
    if ($fallback.ExitCode -ne 0) { continue }

    $parsedText = Parse-ProjectListText -Lines $fallback.Output
    if ($parsedText.Count -gt 0) { return $parsedText }
  }

  throw "Could not parse Vercel project list. Try auth mode 2 (token) or set a valid scope."
}

function Select-ProjectFromList([string]$Scope) {
  Write-Step "Loading projects from Vercel..."
  try {
    $projects = Get-ProjectsFromVercel -Scope $Scope
  } catch {
    Write-Host "Could not load project list: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Tip: continue with option 5 (Deploy as NEW project), or retry with token auth." -ForegroundColor DarkYellow
    return $null
  }
  if ($projects.Count -eq 0) {
    Write-Host "No projects found in this scope." -ForegroundColor DarkYellow
    return $null
  }

  Write-Host ""
  Write-Host "Projects:" -ForegroundColor Cyan
  for ($i = 0; $i -lt $projects.Count; $i++) {
    Write-Host ("[{0}] {1}" -f ($i + 1), $projects[$i].Name)
  }
  Write-Host "[0] Cancel"

  while ($true) {
    $choiceRaw = Read-Default "Select project number" "0"
    $n = 0
    if (-not [int]::TryParse($choiceRaw, [ref]$n)) {
      Write-Host "Invalid number." -ForegroundColor Red
      continue
    }
    if ($n -eq 0) { return $null }
    if ($n -ge 1 -and $n -le $projects.Count) {
      return $projects[$n - 1]
    }
    Write-Host "Out of range." -ForegroundColor Red
  }
}

function Select-ProjectOrNewForFirstRun([string]$Scope) {
  Write-Step "No linked project found. Loading your Vercel projects..."
  try {
    $projects = Get-ProjectsFromVercel -Scope $Scope
  } catch {
    Write-Host "Could not load projects list. Continuing with NEW-project flow." -ForegroundColor DarkYellow
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor DarkGray
    $projects = @()
  }

  Write-Host ""
  Write-Host "Choose a target to continue:" -ForegroundColor Cyan
  if ($projects.Count -gt 0) {
    for ($i = 0; $i -lt $projects.Count; $i++) {
      Write-Host ("[{0}] Use existing project: {1}" -f ($i + 1), $projects[$i].Name)
    }
  } else {
    Write-Host "No existing projects found in current scope/account." -ForegroundColor DarkYellow
  }

  $newIndex = $projects.Count + 1
  Write-Host ("[{0}] Deploy as NEW project" -f $newIndex)

  while ($true) {
    $choiceRaw = Read-Host "Select one option"
    $n = 0
    if (-not [int]::TryParse($choiceRaw, [ref]$n)) {
      Write-Host "Invalid number." -ForegroundColor Red
      continue
    }

    if ($n -eq $newIndex) {
      return @{
        Mode = "new"
        ProjectName = ""
      }
    }

    if ($n -ge 1 -and $n -le $projects.Count) {
      return @{
        Mode = "existing"
        ProjectName = $projects[$n - 1].Name
      }
    }

    Write-Host "Out of range. Choose one of the listed options." -ForegroundColor Red
  }
}

function Show-ManageMenu($selectedProjectName, $scope) {
  Write-Host ""
  Write-Host "Current target project:" -ForegroundColor Cyan
  if ($selectedProjectName) { Write-Host "Project: $selectedProjectName" } else { Write-Host "Project: (not selected)" }
  if ($scope) { Write-Host "Scope:   $scope" } else { Write-Host "Scope:   (default)" }
  Write-Host ""
  Write-Host "[1] Select project from Vercel list"
  Write-Host "[2] Redeploy selected project"
  Write-Host "[3] Update production env vars (selected project)"
  Write-Host "[4] List recent deployments (selected project)"
  Write-Host "[5] Deploy as NEW project"
  Write-Host "[6] Exit"
  return (Read-Default "Choose action" "1")
}

function Run-ManagementLoop([string]$InitialScope) {
  $link = Get-LinkedProjectInfo -ProjectRoot $scriptDir
  $scope = if (-not [string]::IsNullOrWhiteSpace($link.Scope)) { $link.Scope } else { $InitialScope }
  $selectedProjectName = $link.ProjectName

  if ([string]::IsNullOrWhiteSpace($selectedProjectName)) {
    $firstChoice = Select-ProjectOrNewForFirstRun -Scope $scope
    if ($firstChoice.Mode -eq "existing") {
      $selectedProjectName = $firstChoice.ProjectName
      Ensure-LinkedToProject -ProjectName $selectedProjectName -Scope $scope
      Write-Host "Selected project: $selectedProjectName" -ForegroundColor Green
    } else {
      Run-NewDeploymentFlow -DefaultScope $scope
      $link = Get-LinkedProjectInfo -ProjectRoot $scriptDir
      if ($link.ProjectName) { $selectedProjectName = $link.ProjectName }
      if ($link.Scope) { $scope = $link.Scope }
    }
  }

  while ($true) {
    if ([string]::IsNullOrWhiteSpace($selectedProjectName)) {
      Write-Host "No target project selected yet." -ForegroundColor DarkYellow
    }

    $choice = Show-ManageMenu -selectedProjectName $selectedProjectName -scope $scope
    if ($choice -eq "6") {
      Write-Host "Exit."
      break
    }

    try {
      switch ($choice) {
        "1" {
          $selected = Select-ProjectFromList -Scope $scope
          if ($null -ne $selected) {
            $selectedProjectName = $selected.Name
            Write-Host "Selected project: $selectedProjectName" -ForegroundColor Green
            Ensure-LinkedToProject -ProjectName $selectedProjectName -Scope $scope
          }
        }
        "2" {
          if ([string]::IsNullOrWhiteSpace($selectedProjectName)) {
            Write-Host "Select a project first (option 1)." -ForegroundColor Red
            break
          }
          Ensure-LinkedToProject -ProjectName $selectedProjectName -Scope $scope
          $deployInfo = Deploy-Production -Scope $scope
          Show-DeploySummary $deployInfo
          Write-Host "Done."
        }
        "3" {
          if ([string]::IsNullOrWhiteSpace($selectedProjectName)) {
            Write-Host "Select a project first (option 1)." -ForegroundColor Red
            break
          }
          Ensure-LinkedToProject -ProjectName $selectedProjectName -Scope $scope
          Run-UpdateEnvFlow -Scope $scope
          Write-Host "Done."
        }
        "4" {
          if ([string]::IsNullOrWhiteSpace($selectedProjectName)) {
            Write-Host "Select a project first (option 1)." -ForegroundColor Red
            break
          }
          Show-DeploymentList -ProjectName $selectedProjectName -Scope $scope
          Write-Host "Done."
        }
        "5" {
          Run-NewDeploymentFlow -DefaultScope $scope
          $newLink = Get-LinkedProjectInfo -ProjectRoot $scriptDir
          if ($newLink.ProjectName) { $selectedProjectName = $newLink.ProjectName }
          if ($newLink.Scope) { $scope = $newLink.Scope }
        }
        default {
          Write-Host "Invalid option." -ForegroundColor Red
        }
      }
    } catch {
      Write-Host ""
      Write-Host "Action failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    Read-Host "Press Enter to return to main menu (or Ctrl+C to exit)"
  }
}

Write-Banner
Write-Host "Important: connect your VPN in TUN Mode before continuing." -ForegroundColor Magenta
Read-Host "Press Enter to continue"
Write-Host "Tip: Press Ctrl+C at any step to stop/exit." -ForegroundColor DarkYellow

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir
$tokenStorePath = Get-TokenStorePath -ProjectRoot $scriptDir
$scopeStorePath = Get-ScopeStorePath -ProjectRoot $scriptDir

if (-not (Test-Path (Join-Path $scriptDir "api\index.js"))) {
  throw "api/index.js not found. Run this script from project root."
}
if (-not (Test-Path (Join-Path $scriptDir "vercel.json"))) {
  throw "vercel.json not found. Run this script from project root."
}

Ensure-NodeAndNpm
Ensure-VercelCli
Ensure-VercelLogin -OutputDir $scriptDir -TokenStorePath $tokenStorePath
$sessionScope = Resolve-SessionScope -ScopeStorePath $scopeStorePath
Write-Host "Deploy path: $scriptDir" -ForegroundColor DarkGray

Run-ManagementLoop -InitialScope $sessionScope
