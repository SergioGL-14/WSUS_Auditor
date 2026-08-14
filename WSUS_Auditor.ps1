# SCRIPT WSUS + AD + SOLUCIONES AUTOMÁTICAS 

param(
    [string]$ProfilesPath,
    [string]$ProfileName,
    [switch]$SelfTest,
    [string]$MockDataPath
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Net

# ----------------------------------------
# VARIABLES GLOBALES Y CONFIGURACIÓN
# ----------------------------------------
$script:appTitle = "WSUS Auditor"
$script:wsusConnected = $false
$script:wsusConnectionMsg = ""
$script:wsus = $null
$script:resultadosGlobal = @()
$script:resultadosFiltrados = @()
$script:wsusLookup = @{}  # Cache global de equipos WSUS
$script:wsusOnlyComputers = @()  # Equipos solo en WSUS (sin AD)
$script:environmentConfig = $null
$script:ouBase = ""
$script:topLevelOUFilter = @()
$script:domainDnsRoot = ""
$script:profilesFilePath = if ($ProfilesPath) { $ProfilesPath } else { Join-Path -Path $PSScriptRoot -ChildPath "wsus_auditor.profiles.json" }
$script:defaultMockDataPath = if ($MockDataPath) { $MockDataPath } else { Join-Path -Path $PSScriptRoot -ChildPath "wsus_auditor.mock.json" }
$script:profilesStore = $null
$script:useMockMode = $false
$script:mockData = $null

function Convert-DomainToDistinguishedName {
    param([string]$domainName)

    if ([string]::IsNullOrWhiteSpace($domainName)) { return "" }
    return (($domainName.Trim().Split(".")) | ForEach-Object { "DC=$_" }) -join ","
}

function Convert-DistinguishedNameToDomain {
    param([string]$distinguishedName)

    if ([string]::IsNullOrWhiteSpace($distinguishedName)) { return "" }
    $domainParts = @()
    foreach ($part in $distinguishedName.Split(",")) {
        $trimmed = $part.Trim()
        if ($trimmed -like "DC=*") {
            $domainParts += $trimmed.Substring(3)
        }
    }
    return ($domainParts -join ".")
}

function Get-DefaultNamingContext {
    try {
        $rootDse = [ADSI]"LDAP://RootDSE"
        return [string]$rootDse.defaultNamingContext
    } catch {
        return ""
    }
}

function Split-ListInput {
    param([string]$text)

    if ([string]::IsNullOrWhiteSpace($text)) { return @() }

    $items = @()
    foreach ($line in ($text -split "(`r`n|`n|`r)")) {
        foreach ($item in ($line -split "[,;]")) {
            $value = $item.Trim()
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $items += $value
            }
        }
    }
    return $items | Select-Object -Unique
}

function Resolve-AppPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ([System.IO.Path]::IsPathRooted($Path)) { return $Path }
    return Join-Path -Path $PSScriptRoot -ChildPath $Path
}

function Get-ComputerFqdn {
    param(
        $adObj,
        [string]$computerName,
        $wsusObj
    )

    try {
        if ($adObj -and $adObj.DNSHostName) { return [string]$adObj.DNSHostName }
    } catch {}

    try {
        if ($wsusObj -and $wsusObj.FullDomainName) { return [string]$wsusObj.FullDomainName }
    } catch {}

    if (-not [string]::IsNullOrWhiteSpace($computerName) -and -not [string]::IsNullOrWhiteSpace($script:domainDnsRoot)) {
        return "$computerName.$($script:domainDnsRoot)"
    }

    return $computerName
}

function Get-ScopedTopLevelOUs {
    $ous = Get-DirectoryOrganizationalUnits -SearchBase $script:ouBase -SearchScope OneLevel -Filter * | Sort-Object Name

    if ($script:topLevelOUFilter.Count -gt 0) {
        $allowedNames = $script:topLevelOUFilter
        $ous = $ous | Where-Object { $allowedNames -contains $_.Name }
    }

    return @($ous)
}

function New-EmptyProfilesStore {
    return [PSCustomObject]@{
        Version = 1
        LastProfile = ""
        Profiles = @()
    }
}

function Normalize-EnvironmentProfile {
    param($profile)

    if (-not $profile) { return $null }

    return [PSCustomObject]@{
        Name = [string]$profile.Name
        WsusServer = [string]$profile.WsusServer
        UseSsl = [bool]$profile.UseSsl
        WsusPort = if ($profile.WsusPort) { [int]$profile.WsusPort } else { 8530 }
        DomainDnsRoot = [string]$profile.DomainDnsRoot
        BaseOU = [string]$profile.BaseOU
        TopLevelOUFilter = @($profile.TopLevelOUFilter)
        UseMockMode = [bool]$profile.UseMockMode
        MockDataPath = if ($profile.MockDataPath) { [string]$profile.MockDataPath } else { $script:defaultMockDataPath }
    }
}

function Get-EnvironmentProfileByName {
    param(
        $store,
        [string]$name
    )

    if (-not $store -or [string]::IsNullOrWhiteSpace($name)) { return $null }
    return @($store.Profiles | Where-Object { $_.Name -eq $name } | Select-Object -First 1)[0]
}

function Load-EnvironmentProfiles {
    param([string]$path = $script:profilesFilePath)

    if (-not (Test-Path -LiteralPath $path)) {
        return New-EmptyProfilesStore
    }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
        $profiles = @()
        foreach ($profile in @($raw.Profiles)) {
            $normalized = Normalize-EnvironmentProfile -profile $profile
            if ($normalized) { $profiles += $normalized }
        }

        return [PSCustomObject]@{
            Version = 1
            LastProfile = [string]$raw.LastProfile
            Profiles = $profiles
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "⚠️ No se pudo leer el fichero de perfiles JSON:`n$path`n`n$($_.Exception.Message)",
            "Perfiles",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return New-EmptyProfilesStore
    }
}

function Save-EnvironmentProfiles {
    param(
        $store,
        [string]$path = $script:profilesFilePath
    )

    $directory = Split-Path -Path $path -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $payload = [PSCustomObject]@{
        Version = 1
        LastProfile = [string]$store.LastProfile
        Profiles = @(
            @($store.Profiles) | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.Name
                    WsusServer = $_.WsusServer
                    UseSsl = [bool]$_.UseSsl
                    WsusPort = [int]$_.WsusPort
                    DomainDnsRoot = $_.DomainDnsRoot
                    BaseOU = $_.BaseOU
                    TopLevelOUFilter = @($_.TopLevelOUFilter)
                    UseMockMode = [bool]$_.UseMockMode
                    MockDataPath = $_.MockDataPath
                }
            }
        )
    }

    $json = $payload | ConvertTo-Json -Depth 6
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8
}

function Upsert-EnvironmentProfile {
    param(
        $store,
        $profile
    )

    $profile = Normalize-EnvironmentProfile -profile $profile
    if (-not $profile) { return $store }

    $profiles = @()
    $updated = $false
    foreach ($existing in @($store.Profiles)) {
        if ($existing.Name -eq $profile.Name) {
            $profiles += $profile
            $updated = $true
        } else {
            $profiles += $existing
        }
    }
    if (-not $updated) {
        $profiles += $profile
    }

    return [PSCustomObject]@{
        Version = 1
        LastProfile = [string]$store.LastProfile
        Profiles = $profiles
    }
}

function Remove-EnvironmentProfile {
    param(
        $store,
        [string]$name
    )

    return [PSCustomObject]@{
        Version = 1
        LastProfile = if ($store.LastProfile -eq $name) { "" } else { [string]$store.LastProfile }
        Profiles = @($store.Profiles | Where-Object { $_.Name -ne $name })
    }
}

function Get-ParentDistinguishedName {
    param([string]$distinguishedName)

    if ([string]::IsNullOrWhiteSpace($distinguishedName)) { return "" }
    $parts = $distinguishedName.Split(",", 2)
    if ($parts.Count -lt 2) { return "" }
    return $parts[1]
}

function Get-DirectoryOrganizationalUnits {
    param(
        [string]$SearchBase,
        [string]$SearchScope = "OneLevel",
        [string]$Filter = "*"
    )

    if (-not $script:useMockMode) {
        return @(Get-ADOrganizationalUnit -SearchBase $SearchBase -SearchScope $SearchScope -Filter $Filter)
    }

    $ous = @($script:mockData.OrganizationalUnits)
    switch ($SearchScope) {
        "OneLevel" {
            $ous = $ous | Where-Object { $_.ParentDistinguishedName -eq $SearchBase }
        }
        "Subtree" {
            $ous = $ous | Where-Object { $_.DistinguishedName -eq $SearchBase -or $_.DistinguishedName -like "*,$SearchBase" }
        }
    }

    if ($Filter -and $Filter -ne "*") {
        $nameMatches = [regex]::Matches($Filter, "Name -eq '([^']+)'")
        if ($nameMatches.Count -gt 0) {
            $allowedNames = @($nameMatches | ForEach-Object { $_.Groups[1].Value })
            $ous = $ous | Where-Object { $allowedNames -contains $_.Name }
        }
    }

    return @($ous)
}

function Get-DirectoryComputers {
    param(
        [string]$SearchBase,
        [string]$SearchScope = "Subtree",
        [string]$Filter = "*",
        [string[]]$Properties = @()
    )

    if (-not $script:useMockMode) {
        return @(Get-ADComputer -SearchBase $SearchBase -SearchScope $SearchScope -Filter $Filter -Properties $Properties)
    }

    $computers = @($script:mockData.ADComputers)
    switch ($SearchScope) {
        "OneLevel" {
            $computers = $computers | Where-Object { $_.ParentDistinguishedName -eq $SearchBase }
        }
        "Subtree" {
            $computers = $computers | Where-Object { $_.DistinguishedName -like "*,$SearchBase" }
        }
    }

    if ($Filter -and $Filter -ne "*") {
        $nameMatches = [regex]::Matches($Filter, "Name -eq '([^']+)'")
        if ($nameMatches.Count -gt 0) {
            $allowedNames = @($nameMatches | ForEach-Object { $_.Groups[1].Value.ToLower() })
            $computers = $computers | Where-Object { $allowedNames -contains $_.Name.ToLower() }
        }
    }

    return @($computers)
}

function Get-DirectoryObjectByIdentity {
    param([string]$Identity)

    if (-not $script:useMockMode) {
        return Get-ADObject -Identity $Identity -ErrorAction Stop
    }

    $matches = @($script:mockData.OrganizationalUnits + $script:mockData.ADComputers) | Where-Object {
        $_.DistinguishedName -eq $Identity -or $_.Name -eq $Identity
    }

    if (-not $matches -or $matches.Count -eq 0) {
        throw "No se encontró el objeto '$Identity' en los datos mock."
    }

    return $matches[0]
}

function Initialize-MockEnvironment {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "No se encontró el fichero de datos mock: $Path"
    }

    $data = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop

    $ous = @()
    foreach ($ou in @($data.OrganizationalUnits)) {
        $ous += [PSCustomObject]@{
            Name = [string]$ou.Name
            DistinguishedName = [string]$ou.DistinguishedName
            ParentDistinguishedName = [string]$ou.ParentDistinguishedName
        }
    }

    $adComputers = @()
    foreach ($computer in @($data.ADComputers)) {
        $dn = [string]$computer.DistinguishedName
        $adComputers += [PSCustomObject]@{
            Name = [string]$computer.Name
            SamAccountName = [string]$computer.SamAccountName
            DNSHostName = [string]$computer.DNSHostName
            IPv4Address = [string]$computer.IPv4Address
            OperatingSystem = [string]$computer.OperatingSystem
            OperatingSystemVersion = [string]$computer.OperatingSystemVersion
            Manufacturer = [string]$computer.Manufacturer
            Model = [string]$computer.Model
            DistinguishedName = $dn
            ParentDistinguishedName = Get-ParentDistinguishedName -distinguishedName $dn
        }
    }

    $updateLookup = @{}
    foreach ($update in @($data.WSUSUpdates)) {
        $updateLookup[[string]$update.UpdateId] = [PSCustomObject]@{
            UpdateId = [string]$update.UpdateId
            Title = [string]$update.Title
        }
    }

    $wsusTargets = @()
    foreach ($wsusComputer in @($data.WSUSComputers)) {
        $updates = @()
        foreach ($updateState in @($wsusComputer.Updates)) {
            $updates += [PSCustomObject]@{
                UpdateId = [string]$updateState.UpdateId
                UpdateInstallationState = [string]$updateState.State
            }
        }

        $target = [PSCustomObject]@{
            FullDomainName = [string]$wsusComputer.FullDomainName
            IpAddress = [string]$wsusComputer.IpAddress
            OSDescription = [string]$wsusComputer.OSDescription
            Make = [string]$wsusComputer.Make
            Model = [string]$wsusComputer.Model
            LastSyncTime = [datetime]$wsusComputer.LastSyncTime
            LastReportedStatusTime = [datetime]$wsusComputer.LastReportedStatusTime
            __Updates = $updates
        }
        $target | Add-Member -MemberType ScriptMethod -Name GetUpdateInstallationInfoPerUpdate -Value {
            return @($this.__Updates)
        }
        $wsusTargets += $target
    }

    $script:wsus = [PSCustomObject]@{
        __Targets = $wsusTargets
        __UpdateLookup = $updateLookup
    }
    $script:wsus | Add-Member -MemberType ScriptMethod -Name GetComputerTargets -Value {
        return @($this.__Targets)
    }
    $script:wsus | Add-Member -MemberType ScriptMethod -Name GetUpdate -Value {
        param($updateId)
        return $this.__UpdateLookup[[string]$updateId]
    }

    $script:mockData = [PSCustomObject]@{
        OrganizationalUnits = $ous
        ADComputers = $adComputers
        WsusTargets = $wsusTargets
        RootInfo = $data.RootInfo
    }
}

function Show-EnvironmentConfigurationDialog {
    param(
        [string]$DefaultWsusServer,
        [bool]$DefaultUseSsl,
        [int]$DefaultWsusPort,
        [string]$DefaultDomainDnsRoot,
        [string]$DefaultBaseOU,
        [string[]]$DefaultTopLevelOUFilter,
        [bool]$DefaultUseMockMode,
        [string]$DefaultMockDataPath,
        $ProfilesStore,
        [string]$PreselectedProfileName
    )

    $localProfilesStore = if ($ProfilesStore) { $ProfilesStore } else { New-EmptyProfilesStore }

    $configForm = New-Object System.Windows.Forms.Form
    $configForm.Text = "Configuración inicial - $($script:appTitle)"
    $configForm.StartPosition = "CenterScreen"
    $configForm.Size = New-Object System.Drawing.Size(700, 645)
    $configForm.FormBorderStyle = "FixedDialog"
    $configForm.MaximizeBox = $false
    $configForm.MinimizeBox = $false
    $configForm.BackColor = [System.Drawing.Color]::White
    $configForm.Font = New-Object System.Drawing.Font("Segoe UI",9)

    $lblIntro = New-Object System.Windows.Forms.Label
    $lblIntro.Text = "Configure el entorno de Active Directory y WSUS antes de abrir la herramienta."
    $lblIntro.Location = New-Object System.Drawing.Point(15, 15)
    $lblIntro.Size = New-Object System.Drawing.Size(640, 20)
    $configForm.Controls.Add($lblIntro)

    $lblProfiles = New-Object System.Windows.Forms.Label
    $lblProfiles.Text = "Perfil guardado:"
    $lblProfiles.Location = New-Object System.Drawing.Point(20, 48)
    $lblProfiles.Size = New-Object System.Drawing.Size(120, 20)
    $configForm.Controls.Add($lblProfiles)

    $cmbProfiles = New-Object System.Windows.Forms.ComboBox
    $cmbProfiles.Location = New-Object System.Drawing.Point(145, 45)
    $cmbProfiles.Size = New-Object System.Drawing.Size(260, 25)
    $cmbProfiles.DropDownStyle = "DropDownList"
    $configForm.Controls.Add($cmbProfiles)

    $lblProfileName = New-Object System.Windows.Forms.Label
    $lblProfileName.Text = "Nombre del perfil:"
    $lblProfileName.Location = New-Object System.Drawing.Point(20, 82)
    $lblProfileName.Size = New-Object System.Drawing.Size(120, 20)
    $configForm.Controls.Add($lblProfileName)

    $txtProfileName = New-Object System.Windows.Forms.TextBox
    $txtProfileName.Location = New-Object System.Drawing.Point(145, 80)
    $txtProfileName.Size = New-Object System.Drawing.Size(260, 23)
    $configForm.Controls.Add($txtProfileName)

    $btnSaveProfile = New-Object System.Windows.Forms.Button
    $btnSaveProfile.Text = "Guardar perfil"
    $btnSaveProfile.Location = New-Object System.Drawing.Point(425, 44)
    $btnSaveProfile.Size = New-Object System.Drawing.Size(110, 28)
    $configForm.Controls.Add($btnSaveProfile)

    $btnDeleteProfile = New-Object System.Windows.Forms.Button
    $btnDeleteProfile.Text = "Eliminar"
    $btnDeleteProfile.Location = New-Object System.Drawing.Point(545, 44)
    $btnDeleteProfile.Size = New-Object System.Drawing.Size(90, 28)
    $configForm.Controls.Add($btnDeleteProfile)

    $lblWsusServer = New-Object System.Windows.Forms.Label
    $lblWsusServer.Text = "Servidor WSUS:"
    $lblWsusServer.Location = New-Object System.Drawing.Point(20, 128)
    $lblWsusServer.Size = New-Object System.Drawing.Size(120, 20)
    $configForm.Controls.Add($lblWsusServer)

    $txtWsusServer = New-Object System.Windows.Forms.TextBox
    $txtWsusServer.Location = New-Object System.Drawing.Point(145, 126)
    $txtWsusServer.Size = New-Object System.Drawing.Size(260, 23)
    $txtWsusServer.Text = $DefaultWsusServer
    $configForm.Controls.Add($txtWsusServer)

    $chkUseSsl = New-Object System.Windows.Forms.CheckBox
    $chkUseSsl.Text = "Usar SSL"
    $chkUseSsl.Location = New-Object System.Drawing.Point(425, 126)
    $chkUseSsl.Size = New-Object System.Drawing.Size(90, 23)
    $chkUseSsl.Checked = $DefaultUseSsl
    $configForm.Controls.Add($chkUseSsl)

    $lblWsusPort = New-Object System.Windows.Forms.Label
    $lblWsusPort.Text = "Puerto WSUS:"
    $lblWsusPort.Location = New-Object System.Drawing.Point(20, 160)
    $lblWsusPort.Size = New-Object System.Drawing.Size(120, 20)
    $configForm.Controls.Add($lblWsusPort)

    $numWsusPort = New-Object System.Windows.Forms.NumericUpDown
    $numWsusPort.Location = New-Object System.Drawing.Point(145, 158)
    $numWsusPort.Size = New-Object System.Drawing.Size(110, 23)
    $numWsusPort.Minimum = 1
    $numWsusPort.Maximum = 65535
    $numWsusPort.Value = $DefaultWsusPort
    $configForm.Controls.Add($numWsusPort)

    $chkUseSsl.Add_CheckedChanged({
        if ($chkUseSsl.Checked) {
            if ($numWsusPort.Value -eq 8530) { $numWsusPort.Value = 8531 }
        } else {
            if ($numWsusPort.Value -eq 8531) { $numWsusPort.Value = 8530 }
        }
    })

    $lblDomain = New-Object System.Windows.Forms.Label
    $lblDomain.Text = "Dominio DNS/FQDN:"
    $lblDomain.Location = New-Object System.Drawing.Point(20, 192)
    $lblDomain.Size = New-Object System.Drawing.Size(120, 20)
    $configForm.Controls.Add($lblDomain)

    $txtDomain = New-Object System.Windows.Forms.TextBox
    $txtDomain.Location = New-Object System.Drawing.Point(145, 190)
    $txtDomain.Size = New-Object System.Drawing.Size(350, 23)
    $txtDomain.Text = $DefaultDomainDnsRoot
    $configForm.Controls.Add($txtDomain)

    $lblDomainHelp = New-Object System.Windows.Forms.Label
    $lblDomainHelp.Text = "Se usa para completar FQDNs cuando AD no expone DNSHostName."
    $lblDomainHelp.Location = New-Object System.Drawing.Point(145, 214)
    $lblDomainHelp.Size = New-Object System.Drawing.Size(450, 18)
    $lblDomainHelp.ForeColor = [System.Drawing.Color]::DimGray
    $configForm.Controls.Add($lblDomainHelp)

    $lblBaseOu = New-Object System.Windows.Forms.Label
    $lblBaseOu.Text = "OU base / DN base:"
    $lblBaseOu.Location = New-Object System.Drawing.Point(20, 246)
    $lblBaseOu.Size = New-Object System.Drawing.Size(120, 20)
    $configForm.Controls.Add($lblBaseOu)

    $txtBaseOu = New-Object System.Windows.Forms.TextBox
    $txtBaseOu.Location = New-Object System.Drawing.Point(145, 244)
    $txtBaseOu.Size = New-Object System.Drawing.Size(520, 23)
    $txtBaseOu.Text = $DefaultBaseOU
    $configForm.Controls.Add($txtBaseOu)

    $lblBaseOuHelp = New-Object System.Windows.Forms.Label
    $lblBaseOuHelp.Text = "Ejemplo: OU=Sedes,DC=contoso,DC=local. Puede ser también la raíz del dominio."
    $lblBaseOuHelp.Location = New-Object System.Drawing.Point(145, 268)
    $lblBaseOuHelp.Size = New-Object System.Drawing.Size(500, 18)
    $lblBaseOuHelp.ForeColor = [System.Drawing.Color]::DimGray
    $configForm.Controls.Add($lblBaseOuHelp)

    $lblTopLevelFilter = New-Object System.Windows.Forms.Label
    $lblTopLevelFilter.Text = "Filtro de OUs de primer nivel (opcional):"
    $lblTopLevelFilter.Location = New-Object System.Drawing.Point(20, 302)
    $lblTopLevelFilter.Size = New-Object System.Drawing.Size(260, 20)
    $configForm.Controls.Add($lblTopLevelFilter)

    $txtTopLevelFilter = New-Object System.Windows.Forms.TextBox
    $txtTopLevelFilter.Location = New-Object System.Drawing.Point(24, 327)
    $txtTopLevelFilter.Size = New-Object System.Drawing.Size(640, 90)
    $txtTopLevelFilter.Multiline = $true
    $txtTopLevelFilter.ScrollBars = "Vertical"
    $txtTopLevelFilter.Text = (($DefaultTopLevelOUFilter | Where-Object { $_ }) -join [Environment]::NewLine)
    $configForm.Controls.Add($txtTopLevelFilter)

    $lblTopLevelHelp = New-Object System.Windows.Forms.Label
    $lblTopLevelHelp.Text = "Déjelo vacío para mostrar todas las OUs hijas directas de la base. Puede separarlas por coma, punto y coma o saltos de línea."
    $lblTopLevelHelp.Location = New-Object System.Drawing.Point(24, 420)
    $lblTopLevelHelp.Size = New-Object System.Drawing.Size(640, 34)
    $lblTopLevelHelp.ForeColor = [System.Drawing.Color]::DimGray
    $configForm.Controls.Add($lblTopLevelHelp)

    $chkUseMockMode = New-Object System.Windows.Forms.CheckBox
    $chkUseMockMode.Text = "Modo simulación local (sin AD/WSUS)"
    $chkUseMockMode.Location = New-Object System.Drawing.Point(24, 468)
    $chkUseMockMode.Size = New-Object System.Drawing.Size(260, 23)
    $chkUseMockMode.Checked = $DefaultUseMockMode
    $configForm.Controls.Add($chkUseMockMode)

    $lblMockPath = New-Object System.Windows.Forms.Label
    $lblMockPath.Text = "Fichero de datos mock:"
    $lblMockPath.Location = New-Object System.Drawing.Point(20, 500)
    $lblMockPath.Size = New-Object System.Drawing.Size(120, 20)
    $configForm.Controls.Add($lblMockPath)

    $txtMockPath = New-Object System.Windows.Forms.TextBox
    $txtMockPath.Location = New-Object System.Drawing.Point(145, 498)
    $txtMockPath.Size = New-Object System.Drawing.Size(430, 23)
    $txtMockPath.Text = $DefaultMockDataPath
    $configForm.Controls.Add($txtMockPath)

    $btnBrowseMockPath = New-Object System.Windows.Forms.Button
    $btnBrowseMockPath.Text = "..."
    $btnBrowseMockPath.Location = New-Object System.Drawing.Point(585, 497)
    $btnBrowseMockPath.Size = New-Object System.Drawing.Size(35, 25)
    $configForm.Controls.Add($btnBrowseMockPath)

    $btnBrowseMockPath.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = "Seleccionar fichero de datos mock"
        $dialog.Filter = "JSON (*.json)|*.json|Todos los archivos (*.*)|*.*"
        $dialog.FileName = $txtMockPath.Text
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtMockPath.Text = $dialog.FileName
        }
    })

    $btnAceptar = New-Object System.Windows.Forms.Button
    $btnAceptar.Text = "Abrir herramienta"
    $btnAceptar.Location = New-Object System.Drawing.Point(365, 553)
    $btnAceptar.Size = New-Object System.Drawing.Size(135, 34)
    $btnAceptar.BackColor = [System.Drawing.Color]::FromArgb(19,90,145)
    $btnAceptar.ForeColor = [System.Drawing.Color]::White
    $btnAceptar.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $configForm.Controls.Add($btnAceptar)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Location = New-Object System.Drawing.Point(515, 553)
    $btnCancelar.Size = New-Object System.Drawing.Size(110, 34)
    $btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $configForm.Controls.Add($btnCancelar)

    $configForm.AcceptButton = $btnAceptar
    $configForm.CancelButton = $btnCancelar

    function New-ConfigFromForm {
        return [PSCustomObject]@{
            Name = $txtProfileName.Text.Trim()
            WsusServer = $txtWsusServer.Text.Trim()
            UseSsl = $chkUseSsl.Checked
            WsusPort = [int]$numWsusPort.Value
            DomainDnsRoot = $txtDomain.Text.Trim()
            BaseOU = $txtBaseOu.Text.Trim()
            TopLevelOUFilter = @(Split-ListInput -text $txtTopLevelFilter.Text)
            UseMockMode = $chkUseMockMode.Checked
            MockDataPath = $txtMockPath.Text.Trim()
        }
    }

    function Apply-ConfigToForm {
        param($config)
        if (-not $config) { return }

        $txtProfileName.Text = [string]$config.Name
        $txtWsusServer.Text = [string]$config.WsusServer
        $chkUseSsl.Checked = [bool]$config.UseSsl
        $port = if ($config.WsusPort) { [int]$config.WsusPort } else { 8530 }
        if ($port -lt $numWsusPort.Minimum) { $port = [int]$numWsusPort.Minimum }
        if ($port -gt $numWsusPort.Maximum) { $port = [int]$numWsusPort.Maximum }
        $numWsusPort.Value = $port
        $txtDomain.Text = [string]$config.DomainDnsRoot
        $txtBaseOu.Text = [string]$config.BaseOU
        $txtTopLevelFilter.Text = (@($config.TopLevelOUFilter | Where-Object { $_ }) -join [Environment]::NewLine)
        $chkUseMockMode.Checked = [bool]$config.UseMockMode
        $txtMockPath.Text = if ($config.MockDataPath) { [string]$config.MockDataPath } else { $script:defaultMockDataPath }
    }

    function Refresh-ProfilesCombo {
        param([string]$selectedName)
        $cmbProfiles.Items.Clear()
        foreach ($profile in @($localProfilesStore.Profiles | Sort-Object Name)) {
            [void]$cmbProfiles.Items.Add($profile.Name)
        }
        if ($selectedName -and $cmbProfiles.Items.Contains($selectedName)) {
            $cmbProfiles.SelectedItem = $selectedName
        } elseif ($cmbProfiles.Items.Count -gt 0) {
            $cmbProfiles.SelectedIndex = 0
        }
    }

    $cmbProfiles.Add_SelectedIndexChanged({
        if (-not $cmbProfiles.SelectedItem) { return }
        $selected = Get-EnvironmentProfileByName -store $localProfilesStore -name ([string]$cmbProfiles.SelectedItem)
        if ($selected) {
            Apply-ConfigToForm -config $selected
        }
    })

    $btnSaveProfile.Add_Click({
        $current = New-ConfigFromForm
        if ([string]::IsNullOrWhiteSpace($current.Name)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Indique un nombre para guardar el perfil.",
                "Perfiles",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        $localProfilesStore = Upsert-EnvironmentProfile -store $localProfilesStore -profile $current
        $localProfilesStore.LastProfile = $current.Name
        Save-EnvironmentProfiles -store $localProfilesStore
        Refresh-ProfilesCombo -selectedName $current.Name
        [System.Windows.Forms.MessageBox]::Show(
            "Perfil guardado correctamente en:`n$($script:profilesFilePath)",
            "Perfiles",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })

    $btnDeleteProfile.Add_Click({
        $name = $txtProfileName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { return }

        $resp = [System.Windows.Forms.MessageBox]::Show(
            "¿Desea eliminar el perfil '$name'?",
            "Perfiles",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $localProfilesStore = Remove-EnvironmentProfile -store $localProfilesStore -name $name
        Save-EnvironmentProfiles -store $localProfilesStore
        Refresh-ProfilesCombo -selectedName ""
        $txtProfileName.Text = ""
    })

    $btnAceptar.Add_Click({
        $config = New-ConfigFromForm

        if (-not $config.UseMockMode -and [string]::IsNullOrWhiteSpace($config.WsusServer)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Debe indicar un servidor WSUS.",
                "Configuración",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        if ([string]::IsNullOrWhiteSpace($config.BaseOU)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Debe indicar una OU base o un DN base para navegar AD.",
                "Configuración",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        if ($config.UseMockMode -and [string]::IsNullOrWhiteSpace($config.MockDataPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Debe indicar un fichero de datos mock para el modo simulación.",
                "Configuración",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        if (-not [string]::IsNullOrWhiteSpace($config.Name)) {
            $localProfilesStore = Upsert-EnvironmentProfile -store $localProfilesStore -profile $config
            $localProfilesStore.LastProfile = $config.Name
            Save-EnvironmentProfiles -store $localProfilesStore
        }

        $configForm.Tag = $config
        $configForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $configForm.Close()
    })

    $defaultConfig = [PSCustomObject]@{
        Name = ""
        WsusServer = $DefaultWsusServer
        UseSsl = $DefaultUseSsl
        WsusPort = $DefaultWsusPort
        DomainDnsRoot = $DefaultDomainDnsRoot
        BaseOU = $DefaultBaseOU
        TopLevelOUFilter = @($DefaultTopLevelOUFilter)
        UseMockMode = $DefaultUseMockMode
        MockDataPath = $DefaultMockDataPath
    }

    $selectedProfile = if ($PreselectedProfileName) { Get-EnvironmentProfileByName -store $localProfilesStore -name $PreselectedProfileName } else { $null }
    Apply-ConfigToForm -config $(if ($selectedProfile) { $selectedProfile } else { $defaultConfig })
    Refresh-ProfilesCombo -selectedName $(if ($selectedProfile) { $selectedProfile.Name } else { $null })

    if ($configForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $configForm.Tag
    }

    return $null
}

function Initialize-EnvironmentContext {
    if ($script:environmentConfig.UseMockMode) {
        $resolvedMockPath = Resolve-AppPath -Path $script:environmentConfig.MockDataPath
        Initialize-MockEnvironment -Path $resolvedMockPath
        $script:environmentConfig.MockDataPath = $resolvedMockPath
        $script:useMockMode = $true
        $script:wsusConnected = $true
        $script:wsusConnectionMsg = "✅ Modo simulación cargado desde: $($script:environmentConfig.MockDataPath)"
        return
    }

    $script:useMockMode = $false

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show("❌ Error al importar módulo ActiveDirectory: $($_.Exception.Message)",
            "Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        throw
    }

    try {
        Get-DirectoryObjectByIdentity -Identity $script:ouBase | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "❌ La OU base/DN base configurada no es válida o no es accesible:`n$($script:ouBase)`n`n$($_.Exception.Message)",
            "Configuración de AD",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        throw
    }

    try {
        [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
        $script:wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer(
            $script:environmentConfig.WsusServer,
            $script:environmentConfig.UseSsl,
            [int]$script:environmentConfig.WsusPort
        )
        $script:wsusConnected = $true
        $protocol = if ($script:environmentConfig.UseSsl) { "https" } else { "http" }
        $script:wsusConnectionMsg = "✅ Conectado a WSUS: {0}://{1}:{2}" -f $protocol, $script:environmentConfig.WsusServer, $script:environmentConfig.WsusPort
    } catch {
        $script:wsusConnectionMsg = "❌ Error al conectar con WSUS '$($script:environmentConfig.WsusServer)': $($_.Exception.Message)"
    }
}

$defaultNamingContext = Get-DefaultNamingContext
$defaultDomainDnsRoot = if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN.Trim() } else { Convert-DistinguishedNameToDomain -distinguishedName $defaultNamingContext }
$defaultBaseOu = if ($defaultNamingContext) { $defaultNamingContext } elseif ($defaultDomainDnsRoot) { Convert-DomainToDistinguishedName -domainName $defaultDomainDnsRoot } else { "" }
$script:profilesStore = Load-EnvironmentProfiles

$selectedProfile = if ($ProfileName) {
    Get-EnvironmentProfileByName -store $script:profilesStore -name $ProfileName
} elseif ($script:profilesStore.LastProfile) {
    Get-EnvironmentProfileByName -store $script:profilesStore -name $script:profilesStore.LastProfile
} else {
    $null
}

$defaultConfig = if ($selectedProfile) {
    $selectedProfile
} else {
    [PSCustomObject]@{
        Name = ""
        WsusServer = $env:COMPUTERNAME
        UseSsl = $false
        WsusPort = 8530
        DomainDnsRoot = $defaultDomainDnsRoot
        BaseOU = $defaultBaseOu
        TopLevelOUFilter = @()
        UseMockMode = $false
        MockDataPath = $script:defaultMockDataPath
    }
}

if ($SelfTest) {
    if (-not $selectedProfile) {
        [System.Windows.Forms.MessageBox]::Show(
            "No se encontró el perfil solicitado para SelfTest.",
            "SelfTest",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    $script:environmentConfig = $selectedProfile
} else {
    $script:environmentConfig = Show-EnvironmentConfigurationDialog `
        -DefaultWsusServer $defaultConfig.WsusServer `
        -DefaultUseSsl $defaultConfig.UseSsl `
        -DefaultWsusPort $defaultConfig.WsusPort `
        -DefaultDomainDnsRoot $defaultConfig.DomainDnsRoot `
        -DefaultBaseOU $defaultConfig.BaseOU `
        -DefaultTopLevelOUFilter @($defaultConfig.TopLevelOUFilter) `
        -DefaultUseMockMode $defaultConfig.UseMockMode `
        -DefaultMockDataPath $defaultConfig.MockDataPath `
        -ProfilesStore $script:profilesStore `
        -PreselectedProfileName $defaultConfig.Name
}

if (-not $script:environmentConfig) { return }

$script:ouBase = $script:environmentConfig.BaseOU
$script:topLevelOUFilter = @($script:environmentConfig.TopLevelOUFilter)
$script:domainDnsRoot = $script:environmentConfig.DomainDnsRoot
try {
    Initialize-EnvironmentContext
} catch {
    return
}

# ----------------------------------------
# FUNCIONES AUXILIARES
# ----------------------------------------
function Seleccionar-RutaCSV {
    param([string]$titulo = "Guardar reporte CSV")
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = $titulo
    $dialog.Filter = "Archivo CSV (*.csv)|*.csv"
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    } else {
        [System.Windows.Forms.MessageBox]::Show("❌ Exportación cancelada por el usuario.","Aviso",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        return $null
    }
}

# ----------------------------------------
# FORMULARIO PRINCIPAL (ESTILO MODERNO - LAYOUT 3 PANELES)
# ----------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "$($script:appTitle) - $($script:environmentConfig.WsusServer)"
$form.StartPosition = "CenterScreen"
$form.Size        = New-Object System.Drawing.Size(1600, 950)
$form.MinimumSize = New-Object System.Drawing.Size(1200, 700)
$form.WindowState = 'Maximized'
$form.BackColor = [System.Drawing.Color]::FromArgb(240,242,245)
$form.Font      = New-Object System.Drawing.Font("Segoe UI",10)

# ========================================
# HEADER SUPERIOR (Título) - Posición fija
# ========================================
$header = New-Object System.Windows.Forms.Panel
$header.Location = New-Object System.Drawing.Point(0, 0)
$header.Size = New-Object System.Drawing.Size(1600, 50)
$header.Anchor = 'Top, Left, Right'
$header.BackColor = [System.Drawing.Color]::FromArgb(19,90,145)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "  $($script:appTitle)"
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI",16,[System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(10, 10)
$lblTitle.AutoSize = $true
$header.Controls.Add($lblTitle)
$form.Controls.Add($header)

# ========================================
# LAYOUT FIJO - SIN SPLITCONTAINER
# ========================================

# PANEL IZQUIERDO - Control (posición y tamaño fijos)
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Location = New-Object System.Drawing.Point(10, 60)
$leftPanel.Size = New-Object System.Drawing.Size(350, 700)
$leftPanel.Anchor = 'Top, Bottom, Left'
$leftPanel.BackColor = [System.Drawing.Color]::White
$leftPanel.BorderStyle = 'FixedSingle'
$form.Controls.Add($leftPanel)

# Label TreeView
$lblTree = New-Object System.Windows.Forms.Label
$lblTree.Text = "Jerarquia AD"
$lblTree.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$lblTree.ForeColor = [System.Drawing.Color]::FromArgb(19,90,145)
$lblTree.Location = New-Object System.Drawing.Point(10, 10)
$lblTree.AutoSize = $true
$leftPanel.Controls.Add($lblTree)

# TreeView (tamaño fijo)
$tree = New-Object System.Windows.Forms.TreeView
$tree.CheckBoxes = $true
$tree.Location = New-Object System.Drawing.Point(10, 35)
$tree.Size = New-Object System.Drawing.Size(325, 430)
$tree.Anchor = 'Top, Left, Right'
$tree.HideSelection = $false
$tree.ShowLines = $true
$tree.Font = New-Object System.Drawing.Font("Segoe UI",9)
$tree.BorderStyle = 'FixedSingle'
$leftPanel.Controls.Add($tree)

# Función para crear botones con estilo
function Crear-Boton($texto, $y, $colorBase){
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $texto
    $btn.Size = New-Object System.Drawing.Size(325, 30)
    $btn.Location = New-Object System.Drawing.Point(10, $y)
    $btn.BackColor = $colorBase
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

# Colores para botones
$colorPrimary = [System.Drawing.Color]::FromArgb(19,90,145)
$colorSuccess = [System.Drawing.Color]::FromArgb(40,167,69)
$colorWarning = [System.Drawing.Color]::FromArgb(255,152,0)
$colorDanger = [System.Drawing.Color]::FromArgb(220,53,69)
$colorInfo = [System.Drawing.Color]::FromArgb(23,162,184)

# Botones (posiciones fijas, más abajo)
$btnListar = Crear-Boton "Listar Equipos Seleccionados" 480 $colorPrimary
$btnErrores = Crear-Boton "Filtrar Solo Errores" 515 $colorWarning
$btnExportCSV = Crear-Boton "Exportar a CSV" 550 $colorSuccess
$btnSoluciones = Crear-Boton "Aplicar Soluciones" 585 $colorDanger
$btnRefreshTree = Crear-Boton "Actualizar TreeView" 620 $colorInfo

$leftPanel.Controls.AddRange(@($btnListar,$btnErrores,$btnExportCSV,$btnSoluciones,$btnRefreshTree))

# Campo de búsqueda (más abajo)
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Buscar en resultados:"
$lblSearch.Location = New-Object System.Drawing.Point(10, 660)
$lblSearch.AutoSize = $true
$lblSearch.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$leftPanel.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Size = New-Object System.Drawing.Size(325, 25)
$txtSearch.Location = New-Object System.Drawing.Point(10, 683)
$txtSearch.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtSearch.BorderStyle = 'FixedSingle'
$leftPanel.Controls.Add($txtSearch)

# ========================================
# PANEL DERECHO - RESULTADOS (posición fija)
# ========================================
$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Location = New-Object System.Drawing.Point(370, 60)
$rightPanel.Size = New-Object System.Drawing.Size(1210, 700)
$rightPanel.Anchor = 'Top, Bottom, Left, Right'
$rightPanel.BackColor = [System.Drawing.Color]::White
$rightPanel.BorderStyle = 'FixedSingle'
$form.Controls.Add($rightPanel)

# Label resultados
$lblResults = New-Object System.Windows.Forms.Label
$lblResults.Text = "Resultados - Equipos"
$lblResults.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$lblResults.ForeColor = [System.Drawing.Color]::FromArgb(19,90,145)
$lblResults.Location = New-Object System.Drawing.Point(10, 5)
$lblResults.AutoSize = $true
$rightPanel.Controls.Add($lblResults)

# Barra de estado de resultados
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Selecciona equipos del arbol y pulsa 'Listar Equipos'"
$lblStatus.Location = New-Object System.Drawing.Point(180, 7)
$lblStatus.Size = New-Object System.Drawing.Size(500, 18)
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI",9)
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$rightPanel.Controls.Add($lblStatus)

# ProgressBar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(5, 28)
$progressBar.Size = New-Object System.Drawing.Size(1195, 12)
$progressBar.Anchor = 'Top, Left, Right'
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$rightPanel.Controls.Add($progressBar)

# DataGridView para resultados
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(5, 45)
$grid.Size = New-Object System.Drawing.Size(1195, 645)
$grid.Anchor = 'Top, Bottom, Left, Right'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.MultiSelect = $true
$grid.RowHeadersVisible = $false
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = 'None'
$grid.GridColor = [System.Drawing.Color]::FromArgb(220,220,220)

# Estilo de cabeceras
$grid.ColumnHeadersVisible = $true
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(19,90,145)
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersHeight = 30

# Estilo de filas
$grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI",9)
$grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(19,90,145)
$grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$grid.RowTemplate.Height = 26

# Alternar colores de filas
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248,249,250)

# Autoajuste
$grid.AutoSizeColumnsMode = 'Fill'
$grid.AutoGenerateColumns = $true

# Anti-flicker (scroll suave)
$pi = [System.Windows.Forms.DataGridView].GetProperty('DoubleBuffered', [System.Reflection.BindingFlags] 'NonPublic, Instance')
if ($pi) { $pi.SetValue($grid, $true, $null) }

$rightPanel.Controls.Add($grid)

# ========================================
# PANEL INFERIOR - CONSOLA DE LOGS (en el borde inferior)
# ========================================
$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Location = New-Object System.Drawing.Point(10, 770)
$logPanel.Size = New-Object System.Drawing.Size(1570, 130)
$logPanel.Anchor = 'Bottom, Left, Right'
$logPanel.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
$logPanel.BorderStyle = 'FixedSingle'
$form.Controls.Add($logPanel)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Consola de Logs"
$lblLog.ForeColor = [System.Drawing.Color]::FromArgb(100,200,255)
$lblLog.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$lblLog.Location = New-Object System.Drawing.Point(10, 5)
$lblLog.AutoSize = $true
$logPanel.Controls.Add($lblLog)

$txtSalida = New-Object System.Windows.Forms.RichTextBox
$txtSalida.Location = New-Object System.Drawing.Point(10, 25)
$txtSalida.Size = New-Object System.Drawing.Size(1545, 95)
$txtSalida.Anchor = 'Top, Bottom, Left, Right'
$txtSalida.ReadOnly = $true
$txtSalida.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
$txtSalida.ForeColor = [System.Drawing.Color]::FromArgb(200,255,200)
$txtSalida.Font = New-Object System.Drawing.Font("Consolas",9)
$txtSalida.BorderStyle = 'None'
$logPanel.Controls.Add($txtSalida)

# ========================================
# FUNCIÓN PARA MOSTRAR MENSAJES EN LOG
# ========================================
function Mostrar-Mensaje($msg){
    $timestamp = Get-Date -Format 'HH:mm:ss'
    $txtSalida.SelectionStart = $txtSalida.TextLength
    $txtSalida.SelectionColor = [System.Drawing.Color]::FromArgb(100,200,255)
    $txtSalida.AppendText("[$timestamp] ")
    $txtSalida.SelectionColor = [System.Drawing.Color]::FromArgb(200,255,200)
    $txtSalida.AppendText("$msg`r`n")
    $txtSalida.ScrollToCaret()
}

# Función para actualizar la barra de estado
function Actualizar-Estado($texto, $color) {
    $lblStatus.Text = "$texto"
    $lblStatus.ForeColor = $color
}

# ----------------------------------------
# CARGAR TREEVIEW DESDE AD (SOLO CentOS de Salud) Y COMPARAR CON WSUS
# ----------------------------------------
function Get-ADTreeAndComputers {
    # Devuelve un hashtable con OUs y equipos
    $result = @{}
    $ous = Get-DirectoryOrganizationalUnits -SearchBase $script:ouBase -SearchScope OneLevel -Filter * | Sort-Object Name
    foreach ($ou in $ous) {
        # Obtener sub-OUs recursivamente y equipos
        $result[$ou.DistinguishedName] = @{
            Name = $ou.Name
            DistinguishedName = $ou.DistinguishedName
            Computers = Get-DirectoryComputers -SearchBase $ou.DistinguishedName -SearchScope OneLevel -Filter * -Properties @("DNSHostName") | Sort-Object Name
            SubOUs = Get-DirectoryOrganizationalUnits -SearchBase $ou.DistinguishedName -SearchScope OneLevel -Filter * | Sort-Object Name
        }
    }
    return $result
}

# ----------------------------------------
# CARGA PEREZOSA (LAZY LOADING) DEL TREEVIEW
# ----------------------------------------

# Nodo placeholder para indicar que hay hijos por cargar
$script:PLACEHOLDER_TEXT = "__LOADING_PLACEHOLDER__"

function Build-TreeView {
    try {
        $tree.Nodes.Clear()
        $tree.BeginUpdate()  # Evitar parpadeo durante la carga

        $script:wsusLookup = @{}

        if ($script:wsusConnected -and $script:wsus) {
            Mostrar-Mensaje "Cargando cache WSUS (una sola vez)..."
            $wsusComputers = $script:wsus.GetComputerTargets()
            foreach ($w in $wsusComputers) {
                if ($w.FullDomainName) {
                    $script:wsusLookup[$w.FullDomainName.ToLower()] = $w
                }
            }
            Mostrar-Mensaje "✅ Cache WSUS cargado: $($script:wsusLookup.Count) equipos"
        } else {
            Mostrar-Mensaje "⚠️ WSUS no conectado. Se cargará solo la jerarquía de AD."
        }

        Mostrar-Mensaje "Cargando OUs de primer nivel (carga diferida activada)..."

        $firstLevelOUs = Get-ScopedTopLevelOUs

        foreach ($ou in $firstLevelOUs) {
            $nodeOU = New-Object System.Windows.Forms.TreeNode($ou.Name)
            $nodeOU.Tag = @{ Type = "OU"; DN = $ou.DistinguishedName; Loaded = $false }
            $nodeOU.ForeColor = [System.Drawing.Color]::FromArgb(19,90,145)
            
            # Añadir placeholder para mostrar el [+] y permitir expansión
            $placeholder = New-Object System.Windows.Forms.TreeNode($script:PLACEHOLDER_TEXT)
            $placeholder.Tag = @{ Type = "Placeholder" }
            [void]$nodeOU.Nodes.Add($placeholder)
            
            [void]$tree.Nodes.Add($nodeOU)
        }
        
        # Añadir nodo WSUS (sin AD) - también con carga diferida
        $nodeWSUSOnly = New-Object System.Windows.Forms.TreeNode("WSUS (sin AD)")
        $nodeWSUSOnly.Tag = @{ Type = "WSUSOnly"; Loaded = $false }
        $nodeWSUSOnly.ForeColor = [System.Drawing.Color]::DarkRed
        $placeholderWsus = New-Object System.Windows.Forms.TreeNode($script:PLACEHOLDER_TEXT)
        $placeholderWsus.Tag = @{ Type = "Placeholder" }
        [void]$nodeWSUSOnly.Nodes.Add($placeholderWsus)
        [void]$tree.Nodes.Add($nodeWSUSOnly)
        
        $tree.EndUpdate()
        Mostrar-Mensaje "✅ TreeView inicializado con $($firstLevelOUs.Count) OUs de primer nivel + nodo WSUS. Expande un nodo para cargar su contenido."
    } catch {
        $tree.EndUpdate()
        Mostrar-Mensaje "❌ Error cargando TreeView: $($_.Exception.Message)"
    }
}

# Función para cargar hijos de un nodo OU cuando se expande
function Load-NodeChildren {
    param(
        [System.Windows.Forms.TreeNode]$node
    )
    
    if ($node.Tag -eq $null) { return }
    if ($node.Tag.Loaded -eq $true) { return }  # Ya cargado
    
    $tree.BeginUpdate()
    
    try {
        # Eliminar placeholder
        $node.Nodes.Clear()
        
        if ($node.Tag.Type -eq "OU") {
            $ouDN = $node.Tag.DN
            Mostrar-Mensaje "Cargando contenido de: $($node.Text)..."
            
            # Cargar equipos de esta OU (solo primer nivel)
            $computers = Get-DirectoryComputers -SearchBase $ouDN -SearchScope OneLevel -Filter * -Properties @("DNSHostName") | Sort-Object Name
            foreach ($c in $computers) {
                $fqdn = Get-ComputerFqdn -adObj $c -computerName $c.Name
                $childNode = New-Object System.Windows.Forms.TreeNode($fqdn)
                $childNode.Tag = @{ Type = "Computer"; SamAccountName = $c.SamAccountName; ADObject = $c; FQDN = $fqdn }

                if ($script:wsusLookup.ContainsKey($fqdn.ToLower())) {
                    $childNode.ForeColor = [System.Drawing.Color]::DarkGreen
                } else {
                    $childNode.ForeColor = [System.Drawing.Color]::DarkRed
                }
                [void]$node.Nodes.Add($childNode)
            }
            
            # Cargar sub-OUs con placeholder (carga diferida también)
            $subOUs = Get-DirectoryOrganizationalUnits -SearchBase $ouDN -SearchScope OneLevel -Filter * | Sort-Object Name
            foreach ($s in $subOUs) {
                $subNode = New-Object System.Windows.Forms.TreeNode($s.Name)
                $subNode.Tag = @{ Type = "OU"; DN = $s.DistinguishedName; Loaded = $false }
                $subNode.ForeColor = [System.Drawing.Color]::FromArgb(19,90,145)
                
                # Placeholder para sub-OU
                $placeholder = New-Object System.Windows.Forms.TreeNode($script:PLACEHOLDER_TEXT)
                $placeholder.Tag = @{ Type = "Placeholder" }
                [void]$subNode.Nodes.Add($placeholder)
                
                [void]$node.Nodes.Add($subNode)
            }
            
            Mostrar-Mensaje "✅ Cargado: $($computers.Count) equipos, $($subOUs.Count) sub-OUs en '$($node.Text)'"
        }
        elseif ($node.Tag.Type -eq "WSUSOnly") {
            Mostrar-Mensaje "Cargando equipos WSUS sin AD..."
            
            $adComputers = @{}

            if ($script:topLevelOUFilter.Count -gt 0) {
                foreach ($topLevelOU in (Get-ScopedTopLevelOUs)) {
                    try {
                        Get-DirectoryComputers -SearchBase $topLevelOU.DistinguishedName -SearchScope Subtree -Filter * -Properties @("DNSHostName") | ForEach-Object {
                            $fqdn = Get-ComputerFqdn -adObj $_ -computerName $_.Name
                            if ($fqdn) { $adComputers[$fqdn.ToLower()] = $true }
                        }
                    } catch { }
                }
            } else {
                try {
                    Get-DirectoryComputers -SearchBase $script:ouBase -SearchScope Subtree -Filter * -Properties @("DNSHostName") | ForEach-Object {
                        $fqdn = Get-ComputerFqdn -adObj $_ -computerName $_.Name
                        if ($fqdn) { $adComputers[$fqdn.ToLower()] = $true }
                    }
                } catch { }
            }
            
            # Filtrar equipos WSUS que no están en AD
            $wsusOnly = $script:wsusLookup.Keys | Where-Object { -not $adComputers.ContainsKey($_) } | Sort-Object
            
            foreach ($fqdn in $wsusOnly) {
                $child = New-Object System.Windows.Forms.TreeNode($fqdn)
                $child.Tag = @{ Type = "WSUSOnly"; FQDN = $fqdn }
                $child.ForeColor = [System.Drawing.Color]::DarkRed
                [void]$node.Nodes.Add($child)
            }
            
            Mostrar-Mensaje "✅ Cargado: $($wsusOnly.Count) equipos WSUS sin AD"
        }
        
        # Marcar como cargado
        $node.Tag.Loaded = $true
        
    } catch {
        Mostrar-Mensaje "❌ Error cargando nodo: $($_.Exception.Message)"
    }
    
    $tree.EndUpdate()
}

# Evento BeforeExpand para carga perezosa
$tree.Add_BeforeExpand({
    param($sender, $e)
    $node = $e.Node
    
    # Si tiene placeholder, cargar hijos reales
    if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq $script:PLACEHOLDER_TEXT) {
        Load-NodeChildren -node $node
    }
})

# Construir TreeView inicialmente
Build-TreeView

# Botón para refrescar el TreeView
$btnRefreshTree.Add_Click({
    Mostrar-Mensaje "Refrescando TreeView..."
    Build-TreeView
})

# ----------------------------------------
# SELECCIÓN JERÁRQUICA AUTOMÁTICA (CHECK/UNCHECK)
# ----------------------------------------
# Cuando se marca/desmarca un nodo, propagar a hijos y actualizar padres
# IMPORTANTE: Si el nodo no está cargado (tiene placeholder), cargar primero

$tree.Add_AfterCheck({
    param($sender,$e)
    # Evitar recursión si se dispara por nosotros mismos (toggle)
    if ($e.Action -ne [System.Windows.Forms.TreeViewAction]::Unknown) {
        
        # Función para cargar recursivamente todos los hijos de un nodo
        function Cargar-NodoCompleto {
            param($node)
            
            # Si tiene placeholder, cargar los hijos reales
            if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq $script:PLACEHOLDER_TEXT) {
                Load-NodeChildren -node $node
            }
            
            # Recursivamente cargar todos los sub-nodos
            foreach ($child in $node.Nodes) {
                if ($child.Tag -and $child.Tag.Type -eq "OU") {
                    Cargar-NodoCompleto -node $child
                }
            }
        }
        
        # Propagar hacia abajo (primero cargar si es necesario)
        function Propagar-Hijos {
            param($node, $checked)
            
            # Si tiene placeholder, cargar primero
            if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq $script:PLACEHOLDER_TEXT) {
                Load-NodeChildren -node $node
            }
            
            foreach ($child in $node.Nodes) {
                # Saltar placeholders
                if ($child.Text -eq $script:PLACEHOLDER_TEXT) { continue }
                
                $child.Checked = $checked
                
                # Si es una OU, propagar recursivamente
                if ($child.Tag -and $child.Tag.Type -eq "OU") {
                    Propagar-Hijos -node $child -checked $checked
                }
            }
        }
        
        Propagar-Hijos -node $e.Node -checked $e.Node.Checked

        # Propagar hacia arriba: si todos los hermanos checked -> parent checked, si alguno unchecked -> parent unchecked
        function Actualizar-Padre {
            param($node)
            if ($node.Parent -ne $null) {
                $allChecked = $true
                foreach ($sib in $node.Parent.Nodes) { 
                    # Ignorar placeholders
                    if ($sib.Text -eq $script:PLACEHOLDER_TEXT) { continue }
                    if (-not $sib.Checked) { $allChecked = $false; break } 
                }
                $node.Parent.Checked = $allChecked
                Actualizar-Padre -node $node.Parent
            }
        }
        Actualizar-Padre -node $e.Node
    }
})

# ----------------------------------------
# OBTENER LISTADO DE EQUIPOS SELECCIONADOS (OPTIMIZADO - USA CACHE)
# ----------------------------------------
function Get-SelectedComputers {
    $selectedFqdns = @()
    # Recorrer nodos del tree para obtener nodos marcados tipo Computer o WSUSOnly
    foreach ($n in $tree.Nodes) {
        foreach ($child in $n.Nodes) {
            $stack = New-Object System.Collections.Stack
            $stack.Push($child)
            while ($stack.Count -gt 0) {
                $cur = $stack.Pop()
                if ($cur.Checked -and $null -ne $cur.Tag -and $cur.Tag.Type -in @("Computer","WSUSOnly")) {
                    if ($cur.Tag.FQDN) { $selectedFqdns += $cur.Tag.FQDN } else { $selectedFqdns += $cur.Text }
                }
                foreach ($sub in $cur.Nodes) { $stack.Push($sub) }
            }
        }
    }
    $selectedFqdns = $selectedFqdns | Sort-Object -Unique
    
    if ($selectedFqdns.Count -eq 0) { return @() }
    
    # Usar cache global WSUS
    $wsusDict = $script:wsusLookup
    
    # OPTIMIZACIÓN: Obtener todos los equipos AD de una sola vez con un filtro
    $computerNames = $selectedFqdns | ForEach-Object { $_.Split(".")[0] }
    $adCache = @{}
    
    # Consulta batch a AD (mucho más rápido que uno por uno)
    try {
        $filter = ($computerNames | ForEach-Object { "Name -eq '$_'" }) -join " -or "
        if ($filter) {
            $adComputers = Get-DirectoryComputers -SearchBase $script:ouBase -SearchScope Subtree -Filter $filter -Properties @("DNSHostName","OperatingSystem","OperatingSystemVersion","IPv4Address","Manufacturer","Model")
            foreach ($adc in $adComputers) {
                $adCache[$adc.Name.ToLower()] = $adc
            }
        }
    } catch {
        Mostrar-Mensaje "⚠️ Advertencia: No se pudo consultar AD en batch"
    }

    $results = @()
    foreach ($fqdn in $selectedFqdns) {
        $lower = $fqdn.ToLower()
        $shortName = $fqdn.Split(".")[0].ToLower()
        
        $adObj = if ($adCache.ContainsKey($shortName)) { $adCache[$shortName] } else { $null }
        $wsusObj = if ($wsusDict.ContainsKey($lower)) { $wsusDict[$lower] } else { $null }

        $results += [PSCustomObject]@{
            FQDN = $fqdn
            ADObject = $adObj
            WSUSTarget = $wsusObj
        }
    }
    return $results
}
# ----------------------------------------
# FUNCIÓN: Construir-TablaFinal (corrige columnas Versión, Marca, Modelo)
# ----------------------------------------
function Construir-TablaFinal {
    param([Parameter(Mandatory)] $computers)

    $resultados = @()
    $fechaLimite = (Get-Date).AddDays(-30)

    foreach ($c in $computers) {
        $fqdn = $c.FQDN
        $adObj = $c.ADObject
        $wsusObj = $c.WSUSTarget

        $nombre = if ($adObj) { $adObj.Name } else { $fqdn.Split(".")[0] }
        $ip = ""
        $version = ""
        $marca = ""
        $modelo = ""

        if ($wsusObj) {
            # Asegurar acceso a las propiedades WSUS (case-insensitive)
            try { $ip = $wsusObj.IpAddress -as [string] } catch {}
            try { $version = $wsusObj.OSDescription -as [string] } catch {}
            try { $marca = $wsusObj.Make -as [string] } catch {}
            try { $modelo = $wsusObj.Model -as [string] } catch {}
        }

        $estado = if ($wsusObj) {
            if ($wsusObj.LastSyncTime -lt $fechaLimite) { "Desactualizado" }
            else { "Correcto" }
        } else { "Sin conexión WSUS" }

        $obj = [PSCustomObject]@{
            Nombre                 = $nombre
            IP                     = $ip
            Version                = $version
            Marca                  = $marca
            Modelo                 = $modelo
            Error                  = if ($wsusObj) { $wsusObj.LastErrorCode } else { "N/A" }
            PorcentajeInstalacion  = if ($wsusObj) { [math]::Round($wsusObj.PercentComplete,2) } else { 0 }
            InformeUltimoEstado    = if ($wsusObj) { $wsusObj.LastReportedStatusTime } else { $null }
            UltimoContacto         = if ($wsusObj) { $wsusObj.LastSyncTime } else { $null }
            Estado                 = $estado
        }
        $resultados += $obj
    }
    return $resultados
}
# ----------------------------------------
# Mostrar tabla en RichTextBox (formato columnas y color)
# ----------------------------------------
function Mostrar-TablaColor {
    param($data)
    if (-not $data -or $data.Count -eq 0) { return }

    $txtSalida.Clear()

    # Anchos aproximados para fuente Consolas (ajústalos si lo ves necesario)
    $col1 = 28  # Nombre
    $col2 = 16  # IP
    $col3 = 32  # Versión (SO / OSDescription)
    $col4 = 16  # Marca
    $col5 = 22  # Modelo
    $col6 = 28  # Error
    $col7 = 14  # % Instalación
    $col8 = 26  # Informe último estado
    $col9 = 22  # Último contacto

    # Encabezado coherente con las propiedades que mostramos
    $header = ("Nombre".PadRight($col1) +
               "IP".PadRight($col2) +
               "Versión".PadRight($col3) +
               "Marca".PadRight($col4) +
               "Modelo".PadRight($col5) +
               "Error".PadRight($col6) +
               "% Inst.".PadRight($col7) +
               "Informe último estado".PadRight($col8) +
               "Último contacto".PadRight($col9))

    $txtSalida.AppendText($header + "`r`n")
    $txtSalida.AppendText(("-" * ($col1+$col2+$col3+$col4+$col5+$col6+$col7+$col8+$col9)) + "`r`n")

    foreach ($row in $data) {
        $nombre  = ($row.Nombre                 -as [string]).PadRight($col1)
        $ip      = ($row.IP                     -as [string]).PadRight($col2)
        $version = ($row.Version                -as [string]).PadRight($col3)
        $marca   = ($row.Marca                  -as [string]).PadRight($col4)
        $modelo  = ($row.Modelo                 -as [string]).PadRight($col5)
        $error   = ($row.Error                  -as [string]).PadRight($col6)
        $pct     = ($row.PorcentajeInstalacion  -as [string]).PadRight($col7)
        $informe = ($row.InformeUltimoEstado    -as [string]).PadRight($col8)
        $last    = ($row.UltimoContacto         -as [string]).PadRight($col9)

        $line = $nombre + $ip + $version + $marca + $modelo + $error + $pct + $informe + $last + "`r`n"

        $start = $txtSalida.TextLength
        $txtSalida.AppendText($line)

        # Colorear por estado (usa StatusColor si existe; por defecto, negro)
        switch ($row.StatusColor) {
            "Red"    { $txtSalida.Select($start, $line.Length); $txtSalida.SelectionColor = [System.Drawing.Color]::DarkRed }
            "Yellow" { $txtSalida.Select($start, $line.Length); $txtSalida.SelectionColor = [System.Drawing.Color]::DarkOrange }
            default  { $txtSalida.Select($start, $line.Length); $txtSalida.SelectionColor = [System.Drawing.Color]::Black }
        }

        $txtSalida.SelectionStart = $txtSalida.TextLength
        $txtSalida.SelectionColor = [System.Drawing.Color]::Black
    }
}

# Fin del BLOQUE 1
# ========================================
# SCRIPT WSUS + AD + SOLUCIONES AUTOMÁTICAS (MEJORADO - BLOQUE 2)
# Continuación del BLOQUE 1
# ========================================

# ----------------------------------------
# FUNCIONES PARA CALCULAR ESTADOS Y PORCENTAJES (WSUS)
# ----------------------------------------
# FUNCIÓN COMBINADA: Obtiene porcentaje y errores en una sola llamada (mucho más rápido)
function Get-WSUSStatusFast {
    param($wsusTarget)
    
    $result = @{
        Porcentaje = "N/A"
        Errores = @()
        TieneErrores = $false
    }
    
    if (-not $wsusTarget) { return $result }
    
    try {
        # UNA SOLA llamada a la API costosa
        $infos = $wsusTarget.GetUpdateInstallationInfoPerUpdate()
        if (-not $infos -or $infos.Count -eq 0) { 
            $result.Porcentaje = "0%"
            return $result 
        }
        
        $total = $infos.Count
        $installedOrNA = 0
        $failedItems = @()
        
        # Procesar todo en un solo loop
        foreach ($i in $infos) {
            $state = $i.UpdateInstallationState.ToString()
            if ($state -eq "Installed" -or $state -eq "NotApplicable") { 
                $installedOrNA++ 
            }
            elseif ($state -eq "Failed") {
                $failedItems += $i
            }
        }
        
        # Calcular porcentaje
        $pct = [math]::Round(($installedOrNA / $total) * 100, 0)
        $result.Porcentaje = "$pct`%"
        
        # Obtener nombres de errores (limitado a 5 para no ralentizar)
        $errorCount = 0
        foreach ($f in $failedItems) {
            if ($errorCount -ge 5) { 
                $result.Errores += "... y $($failedItems.Count - 5) más"
                break 
            }
            try {
                $upd = $script:wsus.GetUpdate($f.UpdateId)
                $result.Errores += $upd.Title
            } catch {
                $result.Errores += "UpdateId:$($f.UpdateId)"
            }
            $errorCount++
        }
        $result.TieneErrores = ($failedItems.Count -gt 0)
        
    } catch { }
    
    return $result
}

# Funciones de compatibilidad (por si se usan en otro lugar)
function Calcular-PorcentajeInstalacion {
    param($wsusTarget)
    return (Get-WSUSStatusFast $wsusTarget).Porcentaje
}

function Obtener-ErroresWSUS {
    param($wsusTarget)
    return (Get-WSUSStatusFast $wsusTarget).Errores
}

function Obtener-IP {
    param($adObj, $wsusObj)
    # Priorizar IPv4Address en AD si existe, si no, intentar con WSUS property
    if ($adObj -and $adObj.IPv4Address) { return $adObj.IPv4Address }
    if ($wsusObj -and $wsusObj.IpAddress) { return $wsusObj.IpAddress } # puede ser null
    # intentar resolución DNS del FQDN
    try {
        $fqdn = Get-ComputerFqdn -adObj $adObj -computerName $(if ($adObj) { $adObj.Name } else { $null }) -wsusObj $wsusObj
        if ($fqdn) {
            $a = [System.Net.Dns]::GetHostAddresses($fqdn) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1
            if ($a) { return $a.IPAddressToString }
        }
    } catch {}
    return ""
}

function Invoke-FilterErrores {
    param([switch]$Silent)

    if (-not $script:resultadosGlobal -or $script:resultadosGlobal.Count -eq 0) {
        if (-not $Silent) {
            Mostrar-Mensaje "❌ No hay datos listados. Primero usa 'Listar Equipos Seleccionados'."
        }
        return @()
    }

    $script:resultadosFiltrados = $script:resultadosGlobal | Where-Object {
        $_.StatusColor -eq "Red" -or
        $_.StatusColor -eq "Yellow" -or
        ($_.Error -and $_.Error -ne "" -and $_.Error -ne "N/A") -or
        $_.Estado -eq "No en WSUS" -or
        $_.Estado -eq "No notificado >30 días"
    }

    if ($script:resultadosFiltrados.Count -eq 0) {
        if (-not $Silent) {
            Mostrar-Mensaje "✅ No hay equipos con errores o problemas detectados."
            Actualizar-Estado "✅ Todos los equipos están correctos" ([System.Drawing.Color]::FromArgb(40,167,69))
        }
        return @()
    }

    Actualizar-Grid $script:resultadosFiltrados
    if (-not $Silent) {
        Actualizar-Estado "⚠️ Mostrando $($script:resultadosFiltrados.Count) equipos con problemas" ([System.Drawing.Color]::FromArgb(255,152,0))
        Mostrar-Mensaje "⚠️ Filtrado: $($script:resultadosFiltrados.Count) equipos con errores o incidencias de $($script:resultadosGlobal.Count) totales."
    }

    return @($script:resultadosFiltrados)
}

function Invoke-ListSelectedResults {
    param($SelectedItems)

    Mostrar-Mensaje "Listando equipos seleccionados (con progreso)..."

    $sel = if ($SelectedItems) { @($SelectedItems) } else { @(Get-SelectedComputers) }
    if (-not $sel -or $sel.Count -eq 0) {
        Mostrar-Mensaje "❌ No hay equipos seleccionados."
        return @()
    }

    $total = $sel.Count
    $i = 0
    $progressBar.Value = 0
    $script:resultadosGlobal = @()
    $fechaLimite = (Get-Date).AddDays(-30)

    foreach ($item in $sel) {
        $i++
        $pctProg = [math]::Max(1,[math]::Round(($i / $total) * 100,0))
        try { $progressBar.Value = $pctProg } catch {}

        $fqdn = $item.FQDN
        $adObj = $item.ADObject
        $wsusObj = $item.WSUSTarget

        $nombre = if ($adObj) { $adObj.Name } elseif ($wsusObj) { $wsusObj.FullDomainName.Split(".")[0] } else { $fqdn.Split(".")[0] }
        $ip = Obtener-IP -adObj $adObj -wsusObj $wsusObj

        $version = if     ($wsusObj -and $wsusObj.OSDescription)       { $wsusObj.OSDescription }
                  elseif ($adObj   -and $adObj.OperatingSystemVersion) { $adObj.OperatingSystemVersion }
                  elseif ($adObj   -and $adObj.OperatingSystem)        { $adObj.OperatingSystem }
                  else                                                 { "" }

        $marca   = if     ($wsusObj -and $wsusObj.Make)                { $wsusObj.Make }
                  elseif ($adObj   -and $adObj.Manufacturer)           { $adObj.Manufacturer }
                  else                                                 { "" }

        $modelo  = if     ($wsusObj -and $wsusObj.Model)               { $wsusObj.Model }
                  elseif ($adObj   -and $adObj.Model)                  { $adObj.Model }
                  else                                                 { "" }

        $wsusStatus = Get-WSUSStatusFast -wsusTarget $wsusObj
        $errores = $wsusStatus.Errores
        $porcentaje = $wsusStatus.Porcentaje

        $informe = ""
        $ultimoContacto = ""
        if ($wsusObj) {
            try { $informe = $wsusObj.LastReportedStatusTime } catch {}
            try { $ultimoContacto = $wsusObj.LastSyncTime } catch {}
        }

        $estado = "Ok"
        if ($wsusObj) {
            try { if ($wsusObj.LastSyncTime -lt $fechaLimite) { $estado = "No notificado >30 días" } } catch {}
        } else { $estado = "No en WSUS" }

        $statusColor = "Green"
        $sortPriority = 3
        if (-not $adObj -or -not $wsusObj) {
            $statusColor = "Red"
            $sortPriority = 1
        }
        elseif ($wsusStatus.TieneErrores -or $errores.Count -gt 0) {
            $statusColor = "Yellow"
            $sortPriority = 2
        }
        elseif ($estado -ne "Ok") {
            $statusColor = "Yellow"
            $sortPriority = 2
        }

        $script:resultadosGlobal += [PSCustomObject]@{
            Nombre = $nombre
            IP = $ip
            Version = $version
            Marca = $marca
            Modelo = $modelo
            Error = ($errores -join "; ")
            PorcentajeInstalacion = $porcentaje
            InformeUltimoEstado = ($informe -as [string])
            UltimoContacto = ($ultimoContacto -as [string])
            StatusColor = $statusColor
            SortPriority = $sortPriority
            FQDN = $fqdn
            ADObject = $adObj
            WSUSTarget = $wsusObj
            Estado = $estado
        }
    }

    $script:resultadosGlobal = $script:resultadosGlobal | Sort-Object SortPriority, Nombre
    try { $progressBar.Value = 0 } catch {}
    $script:resultadosFiltrados = $script:resultadosGlobal
    Actualizar-Grid $script:resultadosGlobal
    Mostrar-Mensaje "✅ Resultados listados: $($script:resultadosGlobal.Count) equipos."

    return @($script:resultadosGlobal)
}

function Export-ResultsToCsv {
    param(
        [Parameter(Mandatory)] [string]$Path,
        $Data = $script:resultadosGlobal
    )

    @($Data) |
        Select-Object Nombre, IP, Version, Marca, Modelo, Error, PorcentajeInstalacion, InformeUltimoEstado, UltimoContacto, Estado |
        Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Load-AllTreeNodes {
    param([System.Windows.Forms.TreeNodeCollection]$Nodes)

    foreach ($node in $Nodes) {
        if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq $script:PLACEHOLDER_TEXT) {
            Load-NodeChildren -node $node
        }
        if ($node.Nodes.Count -gt 0) {
            Load-AllTreeNodes -Nodes $node.Nodes
        }
    }
}

function Check-AllComputerNodes {
    param([System.Windows.Forms.TreeNodeCollection]$Nodes)

    foreach ($node in $Nodes) {
        if ($node.Tag -and $node.Tag.Type -in @("Computer","WSUSOnly")) {
            $node.Checked = $true
        }
        if ($node.Nodes.Count -gt 0) {
            Check-AllComputerNodes -Nodes $node.Nodes
        }
    }
}

function Invoke-SelfTest {
    if (-not $script:useMockMode) {
        throw "SelfTest requiere un perfil en modo simulación."
    }

    Mostrar-Mensaje "Iniciando SelfTest offline..."
    Build-TreeView
    Load-AllTreeNodes -Nodes $tree.Nodes
    Check-AllComputerNodes -Nodes $tree.Nodes

    $selected = @(Get-SelectedComputers)
    if ($selected.Count -eq 0) {
        throw "SelfTest: no se seleccionaron equipos desde el árbol."
    }

    $results = @(Invoke-ListSelectedResults -SelectedItems $selected)
    if ($results.Count -eq 0) {
        throw "SelfTest: el listado no devolvió resultados."
    }

    $issues = @(Invoke-FilterErrores -Silent)
    if ($issues.Count -eq 0) {
        throw "SelfTest: se esperaba al menos una incidencia en los datos mock."
    }

    $csvPath = Join-Path -Path $PSScriptRoot -ChildPath "wsus_auditor.selftest.csv"
    Export-ResultsToCsv -Path $csvPath -Data $results
    $csvRows = @(Import-Csv -Path $csvPath).Count
    if ($csvRows -ne $results.Count) {
        throw "SelfTest: el CSV exportado no coincide con el número de resultados."
    }

    $profilesTestPath = Join-Path -Path $PSScriptRoot -ChildPath "wsus_auditor.selftest.profiles.json"
    $profilesStore = New-EmptyProfilesStore
    $profilesStore = Upsert-EnvironmentProfile -store $profilesStore -profile ([PSCustomObject]@{
        Name = "SelfTest Profile"
        WsusServer = $script:environmentConfig.WsusServer
        UseSsl = $script:environmentConfig.UseSsl
        WsusPort = $script:environmentConfig.WsusPort
        DomainDnsRoot = $script:environmentConfig.DomainDnsRoot
        BaseOU = $script:environmentConfig.BaseOU
        TopLevelOUFilter = @($script:environmentConfig.TopLevelOUFilter)
        UseMockMode = $script:environmentConfig.UseMockMode
        MockDataPath = $script:environmentConfig.MockDataPath
    })
    $profilesStore.LastProfile = "SelfTest Profile"
    Save-EnvironmentProfiles -store $profilesStore -path $profilesTestPath
    $reloadedProfiles = Load-EnvironmentProfiles -path $profilesTestPath
    $reloadedProfile = Get-EnvironmentProfileByName -store $reloadedProfiles -name "SelfTest Profile"
    if (-not $reloadedProfile) {
        throw "SelfTest: no se pudo recargar el perfil JSON guardado."
    }
    Remove-Item -LiteralPath $profilesTestPath -Force -ErrorAction SilentlyContinue

    return [PSCustomObject]@{
        SelectedCount = $selected.Count
        ResultCount = $results.Count
        IssueCount = $issues.Count
        CsvPath = $csvPath
        CsvRows = $csvRows
    }
}

# BOTÓN 2: FILTRAR SOLO ERRORES (muestra equipos con problemas)
$btnErrores.Add_Click({
    [void](Invoke-FilterErrores)
})

# NOTA: El handler de $btnExportCSV está definido más abajo con confirmación y diálogo mejorado

# BOTÓN 4: APLICAR SOLUCIONES (usando PsExec para equipos cliente)
$btnSoluciones.Add_Click({
    # Verificar que hay filas seleccionadas en el grid
    $selectedRows = $grid.SelectedRows
    if ($selectedRows.Count -eq 0) {
        Mostrar-Mensaje "❌ Seleccione equipos en la tabla de resultados primero."
        [System.Windows.Forms.MessageBox]::Show(
            "Debe seleccionar uno o más equipos en la tabla de resultados antes de aplicar soluciones.`n`nUse Ctrl+Click o Shift+Click para seleccionar múltiples filas.",
            "Sin selección",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Obtener equipos seleccionados
    $equiposSeleccionados = @()
    foreach ($row in $selectedRows) {
        $nombre = $row.Cells["Nombre"].Value
        $ip = $row.Cells["IP"].Value
        # Buscar el objeto completo en resultadosGlobal
        $equipo = $script:resultadosGlobal | Where-Object { $_.Nombre -eq $nombre } | Select-Object -First 1
        if ($equipo) {
            $equiposSeleccionados += $equipo
        }
    }

    if ($equiposSeleccionados.Count -eq 0) {
        Mostrar-Mensaje "❌ No se pudieron obtener los equipos seleccionados."
        return
    }

    # Mostrar diálogo de opciones de solución
    $formOpciones = New-Object System.Windows.Forms.Form
    $formOpciones.Text = "Aplicar Soluciones - $($equiposSeleccionados.Count) equipo(s)"
    $formOpciones.Size = New-Object System.Drawing.Size(500,400)
    $formOpciones.StartPosition = "CenterParent"
    $formOpciones.FormBorderStyle = "FixedDialog"
    $formOpciones.MaximizeBox = $false

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Seleccione las acciones a ejecutar en los $($equiposSeleccionados.Count) equipo(s):"
    $lblInfo.Location = New-Object System.Drawing.Point(15,15)
    $lblInfo.Size = New-Object System.Drawing.Size(460,20)
    $formOpciones.Controls.Add($lblInfo)

    # Checkboxes de opciones
    $chkRestartWU = New-Object System.Windows.Forms.CheckBox
    $chkRestartWU.Text = "Reiniciar servicio Windows Update (wuauserv)"
    $chkRestartWU.Location = New-Object System.Drawing.Point(20,45)
    $chkRestartWU.Size = New-Object System.Drawing.Size(440,25)
    $chkRestartWU.Checked = $true
    $formOpciones.Controls.Add($chkRestartWU)

    $chkClearCache = New-Object System.Windows.Forms.CheckBox
    $chkClearCache.Text = "Limpiar caché de Windows Update (SoftwareDistribution)"
    $chkClearCache.Location = New-Object System.Drawing.Point(20,75)
    $chkClearCache.Size = New-Object System.Drawing.Size(440,25)
    $chkClearCache.Checked = $true
    $formOpciones.Controls.Add($chkClearCache)

    $chkDetectNow = New-Object System.Windows.Forms.CheckBox
    $chkDetectNow.Text = "Forzar detección de actualizaciones"
    $chkDetectNow.Location = New-Object System.Drawing.Point(20,105)
    $chkDetectNow.Size = New-Object System.Drawing.Size(440,25)
    $chkDetectNow.Checked = $true
    $formOpciones.Controls.Add($chkDetectNow)

    $chkReportNow = New-Object System.Windows.Forms.CheckBox
    $chkReportNow.Text = "Forzar reporte al servidor WSUS"
    $chkReportNow.Location = New-Object System.Drawing.Point(20,135)
    $chkReportNow.Size = New-Object System.Drawing.Size(440,25)
    $chkReportNow.Checked = $true
    $formOpciones.Controls.Add($chkReportNow)

    $chkRegisterDLLs = New-Object System.Windows.Forms.CheckBox
    $chkRegisterDLLs.Text = "Re-registrar DLLs de Windows Update"
    $chkRegisterDLLs.Location = New-Object System.Drawing.Point(20,165)
    $chkRegisterDLLs.Size = New-Object System.Drawing.Size(440,25)
    $chkRegisterDLLs.Checked = $false
    $formOpciones.Controls.Add($chkRegisterDLLs)

    $chkResetWU = New-Object System.Windows.Forms.CheckBox
    $chkResetWU.Text = "Reset completo de Windows Update (más agresivo)"
    $chkResetWU.Location = New-Object System.Drawing.Point(20,195)
    $chkResetWU.Size = New-Object System.Drawing.Size(440,25)
    $chkResetWU.Checked = $false
    $formOpciones.Controls.Add($chkResetWU)

    # Credenciales
    $lblCred = New-Object System.Windows.Forms.Label
    $lblCred.Text = "Credenciales (dominio\usuario):"
    $lblCred.Location = New-Object System.Drawing.Point(15,235)
    $lblCred.Size = New-Object System.Drawing.Size(200,20)
    $formOpciones.Controls.Add($lblCred)

    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(20,255)
    $txtUser.Size = New-Object System.Drawing.Size(200,23)
    $txtUser.Text = "$env:USERDOMAIN\$env:USERNAME"
    $formOpciones.Controls.Add($txtUser)

    $lblPass = New-Object System.Windows.Forms.Label
    $lblPass.Text = "La contraseña se solicitará de forma segura al ejecutar."
    $lblPass.Location = New-Object System.Drawing.Point(230,255)
    $lblPass.Size = New-Object System.Drawing.Size(250,35)
    $lblPass.ForeColor = [System.Drawing.Color]::DimGray
    $formOpciones.Controls.Add($lblPass)

    # Botones
    $btnEjecutar = New-Object System.Windows.Forms.Button
    $btnEjecutar.Text = "Ejecutar Soluciones"
    $btnEjecutar.Location = New-Object System.Drawing.Point(120,310)
    $btnEjecutar.Size = New-Object System.Drawing.Size(130,35)
    $btnEjecutar.BackColor = [System.Drawing.Color]::FromArgb(76,175,80)
    $btnEjecutar.ForeColor = [System.Drawing.Color]::White
    $btnEjecutar.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $formOpciones.Controls.Add($btnEjecutar)

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Location = New-Object System.Drawing.Point(260,310)
    $btnCancelar.Size = New-Object System.Drawing.Size(100,35)
    $btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $formOpciones.Controls.Add($btnCancelar)

    $formOpciones.AcceptButton = $btnEjecutar
    $formOpciones.CancelButton = $btnCancelar

    if ($formOpciones.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    # Validar usuario antes de abrir el diálogo seguro de credenciales.
    $usuario = $txtUser.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($usuario)) {
        Mostrar-Mensaje "❌ Debe indicar un usuario."
        return
    }

    # Verificar que al menos una opción está seleccionada
    if (-not ($chkRestartWU.Checked -or $chkClearCache.Checked -or $chkDetectNow.Checked -or $chkReportNow.Checked -or $chkRegisterDLLs.Checked -or $chkResetWU.Checked)) {
        Mostrar-Mensaje "❌ No se seleccionó ninguna acción."
        return
    }

    $credential = Get-Credential -UserName $usuario -Message "Credenciales para administrar los equipos seleccionados"
    if (-not $credential) {
        Mostrar-Mensaje "❌ No se proporcionaron credenciales."
        return
    }
    $usuario = $credential.UserName

    # Ventana de log de resultados
    $ventanaLog = New-Object System.Windows.Forms.Form
    $ventanaLog.Text = "Resultados - Aplicar Soluciones"
    $ventanaLog.Size = New-Object System.Drawing.Size(900,550)
    $ventanaLog.StartPosition = "CenterScreen"
    
    $txtLog = New-Object System.Windows.Forms.RichTextBox
    $txtLog.Dock = "Fill"
    $txtLog.ReadOnly = $true
    $txtLog.Font = New-Object System.Drawing.Font("Consolas",10)
    $txtLog.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
    $txtLog.ForeColor = [System.Drawing.Color]::LightGreen
    $ventanaLog.Controls.Add($txtLog)
    $ventanaLog.Show()

    $txtLog.AppendText("═══════════════════════════════════════════════════════════════`r`n")
    $txtLog.AppendText("  APLICAR SOLUCIONES WSUS - $($equiposSeleccionados.Count) equipo(s)`r`n")
    $txtLog.AppendText("  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n")
    $txtLog.AppendText("═══════════════════════════════════════════════════════════════`r`n`r`n")
    $txtLog.AppendText("Método: WMI (Win32_Process)`r`n")
    $txtLog.AppendText("Usuario: $usuario`r`n`r`n")
    $txtLog.AppendText("Acciones seleccionadas:`r`n")
    if ($chkRestartWU.Checked) { $txtLog.AppendText("  ✓ Reiniciar servicio Windows Update`r`n") }
    if ($chkClearCache.Checked) { $txtLog.AppendText("  ✓ Limpiar caché de Windows Update`r`n") }
    if ($chkDetectNow.Checked) { $txtLog.AppendText("  ✓ Forzar detección de actualizaciones`r`n") }
    if ($chkReportNow.Checked) { $txtLog.AppendText("  ✓ Forzar reporte a WSUS`r`n") }
    if ($chkRegisterDLLs.Checked) { $txtLog.AppendText("  ✓ Re-registrar DLLs`r`n") }
    if ($chkResetWU.Checked) { $txtLog.AppendText("  ✓ Reset completo de Windows Update`r`n") }
    $txtLog.AppendText("`r`n")

    $exitosos = 0
    $fallidos = 0

    foreach ($equipo in $equiposSeleccionados) {
        $nombre = $equipo.Nombre
        $ip = $equipo.IP
        $target = if ($ip -and $ip -ne "") { $ip } else { $nombre }

        $txtLog.AppendText("───────────────────────────────────────────────────────────────`r`n")
        $txtLog.AppendText("[$nombre] Conectando a $target...`r`n")
        $txtLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()

        # Primero verificar conectividad
        if (-not (Test-Connection -ComputerName $target -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            $txtLog.SelectionColor = [System.Drawing.Color]::Red
            $txtLog.AppendText("[$nombre] ❌ No responde a ping - saltando`r`n")
            $txtLog.SelectionColor = [System.Drawing.Color]::LightGreen
            $fallidos++
            continue
        }
        
        # El PSDrive usa la credencial sin exponerla en argumentos de procesos.
        $driveName = "WSUS_$([guid]::NewGuid().ToString('N').Substring(0,8))"
        $remoteDriveRoot = "${driveName}:"
        $cimSession = $null
        $txtLog.AppendText("[$nombre] Verificando acceso administrativo...`r`n")
        [System.Windows.Forms.Application]::DoEvents()

        $netUseResult = $null
        try {
            New-PSDrive -Name $driveName -PSProvider FileSystem -Root "\\$target\C$" -Credential $credential -Scope Local -ErrorAction Stop | Out-Null
            $netUseSuccess = $true
        } catch {
            $netUseSuccess = $false
            $netUseResult = $_.Exception.Message
        }
        
        if (-not $netUseSuccess) {
            $txtLog.SelectionColor = [System.Drawing.Color]::Red
            $txtLog.AppendText("[$nombre] ❌ No se puede acceder a ADMIN$ - Verificar:`r`n")
            $txtLog.AppendText("    - El servicio 'Server' (LanmanServer) está activo en el equipo remoto`r`n")
            $txtLog.AppendText("    - El firewall permite SMB (puerto 445)`r`n")
            $txtLog.AppendText("    - El usuario tiene permisos de administrador local`r`n")
            $txtLog.AppendText("    - Error: $netUseResult`r`n")
            $txtLog.SelectionColor = [System.Drawing.Color]::LightGreen
            Remove-PSDrive -Name $driveName -Force -ErrorAction SilentlyContinue
            $fallidos++
            continue
        }
        
        $txtLog.AppendText("[$nombre] ✓ Acceso a ADMIN$ verificado`r`n")

        try {
            # Crear script PowerShell con las acciones
            $scriptGuid = [guid]::NewGuid().ToString('N').Substring(0,8)
            $scriptName = "WSUS_Fix_$scriptGuid.ps1"
            $logName = "WSUS_Fix_$scriptGuid.log"
            $remoteScriptUNC = "$remoteDriveRoot\Temp\$scriptName"
            $remoteScriptLocal = "C:\Temp\$scriptName"
            $remoteLogUNC = "$remoteDriveRoot\Temp\$logName"
            $remoteLogLocal = "C:\Temp\$logName"
            
            # Crear contenido del script PS1 con logging detallado
            $ps1Content = @"
# Script de corrección WSUS - Generado automáticamente
# $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

`$logFile = "$remoteLogLocal"
`$ErrorActionPreference = 'Continue'

function Write-Log {
    param([string]`$mensaje, [string]`$tipo = "INFO")
    `$timestamp = Get-Date -Format 'HH:mm:ss'
    `$linea = "[`$timestamp] [`$tipo] `$mensaje"
    Add-Content -Path `$logFile -Value `$linea -Encoding UTF8
}

# Iniciar log
"═══════════════════════════════════════════════════════════" | Out-File `$logFile -Encoding UTF8
"  LOG DE EJECUCIÓN - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Add-Content `$logFile -Encoding UTF8
"  Equipo: `$env:COMPUTERNAME" | Add-Content `$logFile -Encoding UTF8
"═══════════════════════════════════════════════════════════" | Add-Content `$logFile -Encoding UTF8
"" | Add-Content `$logFile -Encoding UTF8

`$erroresTotal = 0
`$exitosTotal = 0

"@
            
            if ($chkRestartWU.Checked -or $chkClearCache.Checked -or $chkRegisterDLLs.Checked -or $chkResetWU.Checked) {
                $ps1Content += @"

# Detener servicios
Write-Log "Deteniendo servicio wuauserv (Windows Update)..." "ACCION"
try {
    `$svcWU = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
    if (`$svcWU) {
        `$estadoAntes = `$svcWU.Status
        Stop-Service -Name wuauserv -Force -ErrorAction Stop
        Write-Log "✓ Servicio wuauserv detenido (estado anterior: `$estadoAntes)" "OK"
        `$exitosTotal++
    } else {
        Write-Log "⚠ Servicio wuauserv no encontrado" "WARN"
    }
} catch {
    Write-Log "✗ Error deteniendo wuauserv: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

Write-Log "Deteniendo servicio BITS..." "ACCION"
try {
    `$svcBits = Get-Service -Name bits -ErrorAction SilentlyContinue
    if (`$svcBits) {
        `$estadoAntes = `$svcBits.Status
        Stop-Service -Name bits -Force -ErrorAction Stop
        Write-Log "✓ Servicio BITS detenido (estado anterior: `$estadoAntes)" "OK"
        `$exitosTotal++
    } else {
        Write-Log "⚠ Servicio BITS no encontrado" "WARN"
    }
} catch {
    Write-Log "✗ Error deteniendo BITS: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

"@
            }

            if ($chkClearCache.Checked) {
                $ps1Content += @"

# Limpiar cache de Windows Update
Write-Log "Limpiando cache de Windows Update..." "ACCION"
`$pathDownload = "C:\Windows\SoftwareDistribution\Download"
`$pathDataStore = "C:\Windows\SoftwareDistribution\DataStore"

try {
    if (Test-Path `$pathDownload) {
        `$itemsAntes = (Get-ChildItem `$pathDownload -Recurse -ErrorAction SilentlyContinue | Measure-Object).Count
        Remove-Item -Path "`$pathDownload\*" -Recurse -Force -ErrorAction Stop
        Write-Log "✓ Cache Download limpiado (`$itemsAntes elementos eliminados)" "OK"
        `$exitosTotal++
    } else {
        Write-Log "⚠ Carpeta Download no existe" "WARN"
    }
} catch {
    Write-Log "✗ Error limpiando Download: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

try {
    if (Test-Path `$pathDataStore) {
        `$itemsAntes = (Get-ChildItem `$pathDataStore -Recurse -ErrorAction SilentlyContinue | Measure-Object).Count
        Remove-Item -Path "`$pathDataStore\*" -Recurse -Force -ErrorAction Stop
        Write-Log "✓ Cache DataStore limpiado (`$itemsAntes elementos eliminados)" "OK"
        `$exitosTotal++
    } else {
        Write-Log "⚠ Carpeta DataStore no existe" "WARN"
    }
} catch {
    Write-Log "✗ Error limpiando DataStore: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

"@
            }

            if ($chkResetWU.Checked) {
                $ps1Content += @"

# Reset completo de Windows Update
Write-Log "Realizando reset completo de Windows Update..." "ACCION"

try {
    Stop-Service -Name cryptsvc -Force -ErrorAction Stop
    Write-Log "✓ Servicio cryptsvc detenido" "OK"
} catch {
    Write-Log "✗ Error deteniendo cryptsvc: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

try {
    if (Test-Path "C:\Windows\SoftwareDistribution") {
        `$newName = "SoftwareDistribution.old_`$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        Rename-Item -Path "C:\Windows\SoftwareDistribution" -NewName `$newName -Force -ErrorAction Stop
        Write-Log "✓ SoftwareDistribution renombrado a `$newName" "OK"
        `$exitosTotal++
    }
} catch {
    Write-Log "✗ Error renombrando SoftwareDistribution: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

try {
    if (Test-Path "C:\Windows\System32\catroot2") {
        `$newName = "catroot2.old_`$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        Rename-Item -Path "C:\Windows\System32\catroot2" -NewName `$newName -Force -ErrorAction Stop
        Write-Log "✓ catroot2 renombrado a `$newName" "OK"
        `$exitosTotal++
    }
} catch {
    Write-Log "✗ Error renombrando catroot2: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

try {
    Start-Service -Name cryptsvc -ErrorAction Stop
    Write-Log "✓ Servicio cryptsvc reiniciado" "OK"
} catch {
    Write-Log "✗ Error iniciando cryptsvc: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

"@
            }

            if ($chkRegisterDLLs.Checked) {
                $ps1Content += @"

# Re-registrar DLLs de Windows Update
Write-Log "Re-registrando DLLs de Windows Update..." "ACCION"
`$dlls = @("wuapi.dll","wuaueng.dll","wuaueng1.dll","wucltui.dll","wups.dll","wups2.dll","wuweb.dll","qmgr.dll","qmgrprxy.dll","wucltux.dll","muweb.dll","wuwebv.dll")
`$dllsOk = 0
`$dllsFail = 0

foreach (`$dll in `$dlls) {
    `$dllPath = "C:\Windows\System32\`$dll"
    if (Test-Path `$dllPath) {
        `$result = & regsvr32.exe /s `$dllPath 2>&1
        if (`$LASTEXITCODE -eq 0) {
            `$dllsOk++
        } else {
            Write-Log "  ⚠ Error registrando `$dll" "WARN"
            `$dllsFail++
        }
    } else {
        Write-Log "  ⚠ DLL no encontrada: `$dll" "WARN"
    }
}
Write-Log "✓ DLLs registradas: `$dllsOk OK, `$dllsFail fallidas" "OK"
if (`$dllsOk -gt 0) { `$exitosTotal++ }
if (`$dllsFail -gt 0) { `$erroresTotal++ }

"@
            }

            if ($chkRestartWU.Checked -or $chkClearCache.Checked -or $chkRegisterDLLs.Checked -or $chkResetWU.Checked) {
                $ps1Content += @"

# Iniciar servicios
Write-Log "Iniciando servicio BITS..." "ACCION"
try {
    Start-Service -Name bits -ErrorAction Stop
    Start-Sleep -Seconds 2
    `$svcBits = Get-Service -Name bits
    Write-Log "✓ Servicio BITS iniciado (estado: `$(`$svcBits.Status))" "OK"
    `$exitosTotal++
} catch {
    Write-Log "✗ Error iniciando BITS: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

Write-Log "Iniciando servicio wuauserv (Windows Update)..." "ACCION"
try {
    Start-Service -Name wuauserv -ErrorAction Stop
    Start-Sleep -Seconds 2
    `$svcWU = Get-Service -Name wuauserv
    Write-Log "✓ Servicio wuauserv iniciado (estado: `$(`$svcWU.Status))" "OK"
    `$exitosTotal++
} catch {
    Write-Log "✗ Error iniciando wuauserv: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

"@
            }

            if ($chkDetectNow.Checked) {
                $ps1Content += @"

# Forzar detección de actualizaciones
Write-Log "Forzando detección de actualizaciones..." "ACCION"
try {
    `$wuaucltPath = "C:\Windows\System32\wuauclt.exe"
    if (Test-Path `$wuaucltPath) {
        Start-Process -FilePath `$wuaucltPath -ArgumentList "/detectnow" -NoNewWindow -Wait -ErrorAction Stop
        Write-Log "✓ wuauclt /detectnow ejecutado" "OK"
        `$exitosTotal++
    } else {
        Write-Log "⚠ wuauclt.exe no encontrado (normal en Windows 10+)" "WARN"
    }
} catch {
    Write-Log "✗ Error ejecutando wuauclt: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

try {
    `$usoClientPath = "C:\Windows\System32\UsoClient.exe"
    if (Test-Path `$usoClientPath) {
        Start-Process -FilePath `$usoClientPath -ArgumentList "StartScan" -NoNewWindow -ErrorAction Stop
        Write-Log "✓ UsoClient StartScan iniciado" "OK"
        `$exitosTotal++
    } else {
        Write-Log "⚠ UsoClient.exe no encontrado" "WARN"
    }
} catch {
    Write-Log "✗ Error ejecutando UsoClient: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

"@
            }

            if ($chkReportNow.Checked) {
                $ps1Content += @"

# Forzar reporte a WSUS
Write-Log "Forzando reporte al servidor WSUS..." "ACCION"
try {
    `$wuaucltPath = "C:\Windows\System32\wuauclt.exe"
    if (Test-Path `$wuaucltPath) {
        Start-Process -FilePath `$wuaucltPath -ArgumentList "/reportnow" -NoNewWindow -Wait -ErrorAction Stop
        Write-Log "✓ wuauclt /reportnow ejecutado" "OK"
        `$exitosTotal++
    } else {
        Write-Log "⚠ wuauclt.exe no encontrado (normal en Windows 10+)" "WARN"
    }
} catch {
    Write-Log "✗ Error ejecutando wuauclt reportnow: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

try {
    `$usoClientPath = "C:\Windows\System32\UsoClient.exe"
    if (Test-Path `$usoClientPath) {
        Start-Process -FilePath `$usoClientPath -ArgumentList "RefreshSettings" -NoNewWindow -ErrorAction Stop
        Write-Log "✓ UsoClient RefreshSettings iniciado" "OK"
        `$exitosTotal++
    } else {
        Write-Log "⚠ UsoClient.exe no encontrado" "WARN"
    }
} catch {
    Write-Log "✗ Error ejecutando UsoClient RefreshSettings: `$(`$_.Exception.Message)" "ERROR"
    `$erroresTotal++
}

"@
            }

            $ps1Content += @"

# Resumen final
"" | Add-Content `$logFile -Encoding UTF8
"═══════════════════════════════════════════════════════════" | Add-Content `$logFile -Encoding UTF8
"  RESUMEN DE EJECUCIÓN" | Add-Content `$logFile -Encoding UTF8
"  Acciones exitosas: `$exitosTotal" | Add-Content `$logFile -Encoding UTF8
"  Errores: `$erroresTotal" | Add-Content `$logFile -Encoding UTF8
if (`$erroresTotal -eq 0) {
    "  ESTADO: ✓ COMPLETADO SIN ERRORES" | Add-Content `$logFile -Encoding UTF8
} else {
    "  ESTADO: ⚠ COMPLETADO CON ERRORES" | Add-Content `$logFile -Encoding UTF8
}
"═══════════════════════════════════════════════════════════" | Add-Content `$logFile -Encoding UTF8

"@
            
            # Asegurar que existe C:\Temp en el equipo remoto
            $remoteTempUNC = "$remoteDriveRoot\Temp"
            if (-not (Test-Path $remoteTempUNC)) {
                New-Item -Path $remoteTempUNC -ItemType Directory -Force | Out-Null
            }
            
            # Guardar script en equipo remoto
            $ps1Content | Out-File -FilePath $remoteScriptUNC -Encoding UTF8 -Force
            $txtLog.AppendText("[$nombre] Script copiado: $scriptName (log: $logName)`r`n")
            [System.Windows.Forms.Application]::DoEvents()
            
            # Ejecutar con WMI (Win32_Process) - más compatible que PsExec en entornos restrictivos
            $txtLog.AppendText("[$nombre] Ejecutando script remoto via WMI...`r`n")
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                # Comando a ejecutar: PowerShell con el script
                $comando = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$remoteScriptLocal`""

                $cimSession = New-CimSession -ComputerName $target -Credential $credential -SessionOption (New-CimSessionOption -Protocol Dcom) -ErrorAction Stop
                $scope = $cimSession

                if ($cimSession.State -eq "Opened") {
                    $txtLog.AppendText("[$nombre] Conectado via WMI`r`n")
                    
                    # Crear proceso remoto
                    $processResult = Invoke-CimMethod -CimSession $cimSession -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $comando } -ErrorAction Stop
                    $returnValue = $processResult.ReturnValue
                    $processId = $processResult.ProcessId
                    
                    if ($returnValue -eq 0) {
                        $txtLog.AppendText("[$nombre] Proceso iniciado (PID: $processId)`r`n")
                        
                        # Esperar a que termine el proceso (máximo 120 segundos)
                        $txtLog.AppendText("[$nombre] Esperando finalización...`r`n")
                        [System.Windows.Forms.Application]::DoEvents()
                        
                        $timeout = 120
                        $elapsed = 0
                        $processFinished = $false
                        
                        while ($elapsed -lt $timeout) {
                            Start-Sleep -Seconds 2
                            $elapsed += 2
                            
                            # Verificar si el proceso sigue corriendo
                            $query = "SELECT * FROM Win32_Process WHERE ProcessId = $processId"
                            $results = Get-CimInstance -CimSession $cimSession -Query $query -ErrorAction Stop
                            
                            if ($results.Count -eq 0) {
                                $processFinished = $true
                                break
                            }
                            
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                        
                        # Esperar un momento para que se escriba el log
                        Start-Sleep -Seconds 2
                        
                        # RECUPERAR Y MOSTRAR EL LOG DEL EQUIPO REMOTO
                        $txtLog.AppendText("[$nombre] Recuperando log de ejecución...`r`n")
                        [System.Windows.Forms.Application]::DoEvents()
                        
                        $logRecuperado = $false
                        $erroresEnLog = $false
                        
                        if (Test-Path $remoteLogUNC) {
                            try {
                                $contenidoLog = Get-Content -Path $remoteLogUNC -Encoding UTF8 -ErrorAction Stop
                                $txtLog.AppendText("`r`n")
                                $txtLog.SelectionColor = [System.Drawing.Color]::Cyan
                                $txtLog.AppendText("[$nombre] ─── LOG DE EJECUCIÓN REMOTA ───`r`n")
                                
                                foreach ($linea in $contenidoLog) {
                                    # Colorear según el tipo de mensaje
                                    if ($linea -match "\[ERROR\]" -or $linea -match "✗") {
                                        $txtLog.SelectionColor = [System.Drawing.Color]::Red
                                        $erroresEnLog = $true
                                    } elseif ($linea -match "\[WARN\]" -or $linea -match "⚠") {
                                        $txtLog.SelectionColor = [System.Drawing.Color]::Orange
                                    } elseif ($linea -match "\[OK\]" -or $linea -match "✓") {
                                        $txtLog.SelectionColor = [System.Drawing.Color]::LightGreen
                                    } elseif ($linea -match "═══") {
                                        $txtLog.SelectionColor = [System.Drawing.Color]::Cyan
                                    } else {
                                        $txtLog.SelectionColor = [System.Drawing.Color]::White
                                    }
                                    $txtLog.AppendText("  $linea`r`n")
                                }
                                
                                $txtLog.SelectionColor = [System.Drawing.Color]::Cyan
                                $txtLog.AppendText("[$nombre] ─── FIN DEL LOG ───`r`n")
                                $txtLog.SelectionColor = [System.Drawing.Color]::LightGreen
                                $logRecuperado = $true
                                
                                # Eliminar log remoto
                                Remove-Item $remoteLogUNC -Force -ErrorAction SilentlyContinue
                                
                            } catch {
                                $txtLog.SelectionColor = [System.Drawing.Color]::Orange
                                $txtLog.AppendText("[$nombre] ⚠️ No se pudo leer el log: $($_.Exception.Message)`r`n")
                                $txtLog.SelectionColor = [System.Drawing.Color]::LightGreen
                            }
                        } else {
                            $txtLog.SelectionColor = [System.Drawing.Color]::Orange
                            $txtLog.AppendText("[$nombre] ⚠️ No se encontró archivo de log en el equipo remoto`r`n")
                            $txtLog.SelectionColor = [System.Drawing.Color]::LightGreen
                        }
                        
                        $txtLog.AppendText("`r`n")
                        
                        if ($processFinished) {
                            if ($erroresEnLog) {
                                $txtLog.SelectionColor = [System.Drawing.Color]::Orange
                                $txtLog.AppendText("[$nombre] ⚠️ Proceso completado CON ERRORES (ver log arriba)`r`n")
                            } else {
                                $txtLog.SelectionColor = [System.Drawing.Color]::LightGreen
                                $txtLog.AppendText("[$nombre] ✅ Soluciones aplicadas correctamente`r`n")
                            }
                            $exitosos++
                        } else {
                            $txtLog.SelectionColor = [System.Drawing.Color]::Orange
                            $txtLog.AppendText("[$nombre] ⚠️ Proceso iniciado pero timeout esperando finalización`r`n")
                            $exitosos++  # Probablemente funcionó
                        }
                    } else {
                        $errorMsg = switch ($returnValue) {
                            2 { "Acceso denegado" }
                            3 { "Privilegios insuficientes" }
                            8 { "Error desconocido" }
                            9 { "Ruta no encontrada" }
                            21 { "Parámetro inválido" }
                            default { "Código de error: $returnValue" }
                        }
                        $txtLog.SelectionColor = [System.Drawing.Color]::Red
                        $txtLog.AppendText("[$nombre] ❌ Error WMI: $errorMsg`r`n")
                        $fallidos++
                    }
                } else {
                    throw "No se pudo conectar via WMI"
                }
                
            } catch {
                $txtLog.SelectionColor = [System.Drawing.Color]::Red
                $txtLog.AppendText("[$nombre] ❌ Error WMI: $($_.Exception.Message)`r`n")
                $txtLog.SelectionColor = [System.Drawing.Color]::LightGreen
                $fallidos++
            }
            
            $txtLog.SelectionColor = [System.Drawing.Color]::LightGreen
            
            # Limpiar referencias remotas y sesiones incluso si la operación falla.
            Start-Sleep -Seconds 2
            Remove-Item $remoteScriptUNC -Force -ErrorAction SilentlyContinue

        } catch {
            $txtLog.SelectionColor = [System.Drawing.Color]::Red
            $txtLog.AppendText("[$nombre] ❌ Error: $($_.Exception.Message)`r`n")
            $txtLog.SelectionColor = [System.Drawing.Color]::LightGreen
            $fallidos++
        }
        
        if ($cimSession) { Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue }
        Remove-PSDrive -Name $driveName -Force -ErrorAction SilentlyContinue

        $txtLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    $txtLog.AppendText("`r`n═══════════════════════════════════════════════════════════════`r`n")
    $txtLog.AppendText("  RESUMEN: $exitosos exitosos, $fallidos fallidos`r`n")
    $txtLog.AppendText("═══════════════════════════════════════════════════════════════`r`n")
    $txtLog.ScrollToCaret()

    Mostrar-Mensaje "✅ Proceso completado: $exitosos exitosos, $fallidos fallidos"
    $credential = $null
})

# ----------------------------------------
# Inicial: Mensaje y foco
# ----------------------------------------
Mostrar-Mensaje "Interfaz lista. Carga inicial completada."
$form.Activate()

# Context menu for grid rows
$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
$miExportRow = New-Object System.Windows.Forms.ToolStripMenuItem "Exportar fila seleccionada CSV"
$miShowDetails = New-Object System.Windows.Forms.ToolStripMenuItem "Ver detalles"
$ctxMenu.Items.AddRange(@($miExportRow,$miShowDetails))

# ----------------------------------------
# UTIL: Convertir resultados a DataTable para bind al DataGridView
# ----------------------------------------
function ConvertTo-DataTable {
    param([Parameter(Mandatory)] $data)

    $dt = New-Object System.Data.DataTable

    if ($data.Count -gt 0) {
        # Crear todas las columnas en orden fijo
        $columnNames = @("Nombre","IP","Version","Marca","Modelo","Error","PorcentajeInstalacion","InformeUltimoEstado","UltimoContacto","Estado")
        foreach ($col in $columnNames) {
            [void]$dt.Columns.Add($col)
        }

        foreach ($item in $data) {
            $row = $dt.NewRow()
            foreach ($col in $columnNames) {
                $row[$col] = if ($item.PSObject.Properties[$col]) { $item.$col } else { "" }
            }
            $dt.Rows.Add($row)
        }
    }
    return $dt
}

# ----------------------------------------
# FUNCIÓN: Actualizar Grid con resultados
# ----------------------------------------
function Actualizar-Grid {
    param($data)
    
    if (-not $data -or $data.Count -eq 0) {
        $grid.DataSource = $null
        Actualizar-Estado "ℹ️ No hay datos para mostrar" ([System.Drawing.Color]::Gray)
        return
    }
    
    # Convertir a DataTable
    $dt = New-Object System.Data.DataTable
    $cols = "Nombre","IP","Version","Marca","Modelo","Error","PorcentajeInstalacion","InformeUltimoEstado","UltimoContacto","Estado"
    foreach ($c in $cols) { [void]$dt.Columns.Add($c) }
    
    foreach ($obj in $data) {
        $row = $dt.NewRow()
        foreach ($c in $cols) {
            $prop = $obj.PSObject.Properties.Name | Where-Object { $_ -ieq $c }
            $row[$c] = if ($prop) { $obj.PSObject.Properties[$prop].Value } else { "" }
        }
        $dt.Rows.Add($row)
    }
    
    $grid.DataSource = $dt
    
    # Colorear filas según StatusColor
    foreach ($r in $grid.Rows) {
        $nombre = $r.Cells["Nombre"].Value
        $o = $data | Where-Object { $_.Nombre -eq $nombre } | Select-Object -First 1
        if ($o) {
            switch ($o.StatusColor) {
                "Red" { $r.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed }
                "Yellow" { $r.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkOrange }
                default { $r.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black }
            }
        }
    }
    
    Actualizar-Estado "✅ Mostrando $($data.Count) equipos" ([System.Drawing.Color]::FromArgb(40,167,69))
}

# ----------------------------------------
# BÚSQUEDA RÁPIDA EN GRID
# ----------------------------------------
$txtSearch.Add_TextChanged({
    $q = $txtSearch.Text.Trim().ToLower()
    if (-not $script:resultadosFiltrados -or $script:resultadosFiltrados.Count -eq 0) { return }
    
    if ($q -ne "") {
        $filtered = $script:resultadosFiltrados | Where-Object { 
            ($_.Nombre -as [string]).ToLower().Contains($q) -or 
            ($_.IP -as [string]).ToLower().Contains($q) -or 
            ($_.Marca -as [string]).ToLower().Contains($q) -or 
            ($_.Modelo -as [string]).ToLower().Contains($q) -or 
            ($_.Error -as [string]).ToLower().Contains($q) 
        }
        Actualizar-Grid $filtered
    } else {
        Actualizar-Grid $script:resultadosFiltrados
    }
})

# ----------------------------------------
# CONTEXT MENU HANDLERS (GRID)
# ----------------------------------------
$miExportRow.Add_Click({
    if (-not $grid.SelectedRows -or $grid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Seleccione una fila primero.",
            "Aviso",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    $rowData = @{}
    foreach ($c in $grid.Columns) {
        $rowData[$c.Name] = $grid.SelectedRows[0].Cells[$c.Index].Value
    }

    $ruta = Seleccionar-RutaCSV "Exportar fila seleccionada"
    if (-not $ruta) { return }

    try {
        $obj = New-Object PSObject
        foreach ($k in $rowData.Keys) {
            $obj | Add-Member -NotePropertyName $k -NotePropertyValue $rowData[$k]
        }

        $obj | Export-Csv -Path $ruta -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show(
            "✅ Fila exportada correctamente:`n$ruta",
            "Exportar CSV",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "❌ Error exportando fila: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# ----------------------------------------
# OPCIÓN CONTEXTUAL: Ver detalles de equipo
# ----------------------------------------
$miShowDetails.Add_Click({
    if (-not $grid.SelectedRows -or $grid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Seleccione una fila primero.",
            "Aviso",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    $nombre = $grid.SelectedRows[0].Cells["Nombre"].Value
    $obj = $script:resultadosFiltrados | Where-Object { $_.Nombre -eq $nombre } | Select-Object -First 1
    if (-not $obj) { return }

    $detalle = @"
Nombre: $($obj.Nombre)
IP: $($obj.IP)
Versión: $($obj.Version)
Marca: $($obj.Marca)
Modelo: $($obj.Modelo)
Error: $($obj.Error)
% Instalación: $($obj.PorcentajeInstalacion)
Informe Último Estado: $($obj.InformeUltimoEstado)
Último Contacto: $($obj.UltimoContacto)
Estado: $($obj.Estado)
"@

    [System.Windows.Forms.MessageBox]::Show(
        $detalle,
        "Detalles del equipo",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

# Permitir doble clic en fila para ver detalles
$grid.Add_CellDoubleClick({
    param($sender,$e)
    if ($e.RowIndex -lt 0) { return }
    $nombre = $grid.Rows[$e.RowIndex].Cells["Nombre"].Value
    $obj = $script:resultadosFiltrados | Where-Object { $_.Nombre -eq $nombre } | Select-Object -First 1
    if ($obj) {
        $detalle = @"
Nombre: $($obj.Nombre)
IP: $($obj.IP)
Versión: $($obj.Version)
Marca: $($obj.Marca)
Modelo: $($obj.Modelo)
Error: $($obj.Error)
% Instalación: $($obj.PorcentajeInstalacion)
Informe Último Estado: $($obj.InformeUltimoEstado)
Último Contacto: $($obj.UltimoContacto)
Estado: $($obj.Estado)
"@
        [System.Windows.Forms.MessageBox]::Show(
            $detalle,
            "Detalles del equipo",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# ----------------------------------------
# BOTÓN LISTAR EQUIPOS
# ----------------------------------------
$btnListar.Add_Click({
    try {
        [void](Invoke-ListSelectedResults)
    } catch {
        Mostrar-Mensaje "❌ Error al listar con progreso: $($_.Exception.Message)"
    }
})

# ----------------------------------------
# BOTÓN EXPORTAR A CSV (versión corregida)
# ----------------------------------------
$btnExportCSV.Add_Click({
    try {
        if (-not $script:resultadosGlobal -or $script:resultadosGlobal.Count -eq 0) {
            Mostrar-Mensaje "❌ No hay resultados para exportar."
            return
        }

        $resp = [System.Windows.Forms.MessageBox]::Show(
            "¿Desea exportar $($script:resultadosGlobal.Count) registros a CSV?",
            "Confirmar exportación",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) {
            Mostrar-Mensaje "Exportación cancelada por el usuario."
            return
        }

        # Diálogo para elegir la ruta de guardado
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV (*.csv)|*.csv"
        $saveDialog.Title = "Guardar informe de WSUS"
        $saveDialog.FileName = "Informe_WSUS_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $ruta = $saveDialog.FileName

            Export-ResultsToCsv -Path $ruta -Data $script:resultadosGlobal

            Mostrar-Mensaje "✅ Resultados exportados a: $ruta"
            [System.Windows.Forms.MessageBox]::Show(
                "Exportación completada correctamente:`n$ruta",
                "Exportar CSV",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    } catch {
        Mostrar-Mensaje "❌ Error exportando CSV: $($_.Exception.Message)"
    }
})

# Asociar menu contextual al grid
$grid.ContextMenuStrip = $ctxMenu

# ----------------------------------------
# EVENTO LOAD: Mostrar mensaje de conexión WSUS al iniciar
# ----------------------------------------
$form.Add_Load({
    # Mostrar estado de conexión WSUS en el log (sin bloquear)
    Mostrar-Mensaje "Ámbito AD configurado: $($script:ouBase)"
    if ($script:topLevelOUFilter.Count -gt 0) {
        Mostrar-Mensaje "Filtro de OUs de primer nivel: $($script:topLevelOUFilter -join ', ')"
    } else {
        Mostrar-Mensaje "Filtro de OUs de primer nivel: todas"
    }
    Mostrar-Mensaje $script:wsusConnectionMsg
    if (-not $script:wsusConnected) {
        Actualizar-Estado "⚠️ Sin conexión WSUS" ([System.Drawing.Color]::Red)
    } else {
        Actualizar-Estado "✅ Listo - Seleccione equipos y pulse 'Listar'" ([System.Drawing.Color]::Green)
    }
})

if ($SelfTest) {
    try {
        $selfTestResult = Invoke-SelfTest
        "SELFTEST_OK"
        "SelectedCount=$($selfTestResult.SelectedCount)"
        "ResultCount=$($selfTestResult.ResultCount)"
        "IssueCount=$($selfTestResult.IssueCount)"
        "CsvRows=$($selfTestResult.CsvRows)"
        "CsvPath=$($selfTestResult.CsvPath)"
    } catch {
        "SELFTEST_FAILED"
        "Error=$($_.Exception.Message)"
        exit 1
    }
    return
}

# ----------------------------------------
# MOSTRAR FORMULARIO
# ----------------------------------------
try {
    if (-not $form.Visible) {
        [void]$form.ShowDialog()
    }
} catch {}

# Fin del Script
