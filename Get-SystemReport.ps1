# Get-SystemReport.ps1
# Generates an HTML report with Bootstrap showing system specifications

$ErrorActionPreference = "SilentlyContinue"

# --- Gather system information ---

$cs = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$os = Get-CimInstance Win32_OperatingSystem
$cpu = Get-CimInstance Win32_Processor
$gpu = Get-CimInstance Win32_VideoController
$ram = Get-CimInstance Win32_PhysicalMemory
$disk = Get-CimInstance Win32_DiskDrive
$net = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
$mb = Get-CimInstance Win32_BaseBoard

$totalRAM = [math]::Round(($ram | Measure-Object Capacity -Sum).Sum / 1GB, 2)
$hostname = $cs.Name
$domain = $cs.Domain
$manufacturer = $cs.Manufacturer
$model = $cs.Model
$serial = $bios.SerialNumber
$biosVersion = $bios.SMBIOSBIOSVersion
$osName = $os.Caption
$osBuild = $os.BuildNumber
$osArch = $os.OSArchitecture
$installDate = $os.InstallDate.ToString("MM/dd/yyyy")
$cpuName = ($cpu | Select-Object -First 1).Name
$cpuCores = ($cpu | Select-Object -First 1).NumberOfCores
$cpuThreads = ($cpu | Select-Object -First 1).NumberOfLogicalProcessors
$mbManufacturer = $mb.Manufacturer
$mbProduct = $mb.Product
$mbSerial = $mb.SerialNumber
$reportDate = Get-Date -Format "MM/dd/yyyy HH:mm:ss"

# --- Build GPU rows ---
$gpuRows = ""
foreach ($g in $gpu) {
    $vram = if ($g.AdapterRAM -and $g.AdapterRAM -gt 0) {
        "$([math]::Round($g.AdapterRAM / 1GB, 2)) GB"
    } else { "N/A" }
    $gpuRows += @"
                        <tr>
                            <td>$($g.Name)</td>
                            <td>$vram</td>
                            <td>$($g.DriverVersion)</td>
                            <td>$($g.VideoModeDescription)</td>
                        </tr>
"@
}

# --- Build RAM rows ---
$ramRows = ""
$slotNum = 1
foreach ($stick in $ram) {
    $capGB = [math]::Round($stick.Capacity / 1GB, 2)
    $speed = if ($stick.Speed) { "$($stick.Speed) MHz" } else { "N/A" }
    $ramRows += @"
                        <tr>
                            <td>Slot $slotNum</td>
                            <td>$($stick.Manufacturer)</td>
                            <td>$capGB GB</td>
                            <td>$speed</td>
                            <td>$($stick.PartNumber)</td>
                        </tr>
"@
    $slotNum++
}

# --- Build Disk rows ---
$diskRows = ""
foreach ($d in $disk) {
    $sizeGB = [math]::Round($d.Size / 1GB, 2)
    $diskRows += @"
                        <tr>
                            <td>$($d.Model)</td>
                            <td>$sizeGB GB</td>
                            <td>$($d.InterfaceType)</td>
                            <td>$($d.SerialNumber)</td>
                        </tr>
"@
}

# --- Build Network rows ---
$netRows = ""
foreach ($n in $net) {
    $ips = ($n.IPAddress -join ", ")
    $mac = $n.MACAddress
    $desc = $n.Description
    $netRows += @"
                        <tr>
                            <td>$desc</td>
                            <td>$ips</td>
                            <td>$mac</td>
                        </tr>
"@
}

# --- Generate HTML ---
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>System Report - $hostname</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <style>
        body { background-color: #f0f2f5; }
        .report-header { background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%); color: white; padding: 2rem 0; }
        .card { border: none; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); margin-bottom: 1.5rem; }
        .card-header { border-radius: 12px 12px 0 0 !important; font-weight: 600; }
        .table { margin-bottom: 0; }
        .section-icon { font-size: 1.3rem; margin-right: 0.5rem; }
        .info-label { font-weight: 600; color: #495057; width: 35%; }
        @media print {
            body { background: white; }
            .report-header { background: #1a1a2e !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
    </style>
</head>
<body>
    <div class="report-header text-center">
        <div class="container">
            <h1 class="mb-1"><i class="bi bi-pc-display"></i> System Report</h1>
            <p class="mb-0 opacity-75">$hostname | Generated: $reportDate</p>
        </div>
    </div>

    <div class="container my-4">
        <div class="row">
            <!-- Computer -->
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header bg-primary text-white">
                        <i class="bi bi-laptop section-icon"></i>Computer
                    </div>
                    <div class="card-body p-0">
                        <table class="table table-striped">
                            <tbody>
                                <tr><td class="info-label">Manufacturer</td><td>$manufacturer</td></tr>
                                <tr><td class="info-label">Model</td><td>$model</td></tr>
                                <tr><td class="info-label">Serial Number</td><td>$serial</td></tr>
                                <tr><td class="info-label">Hostname</td><td>$hostname</td></tr>
                                <tr><td class="info-label">Domain</td><td>$domain</td></tr>
                                <tr><td class="info-label">BIOS</td><td>$biosVersion</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- Operating System -->
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header bg-success text-white">
                        <i class="bi bi-windows section-icon"></i>Operating System
                    </div>
                    <div class="card-body p-0">
                        <table class="table table-striped">
                            <tbody>
                                <tr><td class="info-label">OS</td><td>$osName</td></tr>
                                <tr><td class="info-label">Build</td><td>$osBuild</td></tr>
                                <tr><td class="info-label">Architecture</td><td>$osArch</td></tr>
                                <tr><td class="info-label">Install Date</td><td>$installDate</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <div class="row">
            <!-- Processor -->
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header bg-info text-white">
                        <i class="bi bi-cpu section-icon"></i>Processor
                    </div>
                    <div class="card-body p-0">
                        <table class="table table-striped">
                            <tbody>
                                <tr><td class="info-label">Processor</td><td>$cpuName</td></tr>
                                <tr><td class="info-label">Cores</td><td>$cpuCores</td></tr>
                                <tr><td class="info-label">Threads</td><td>$cpuThreads</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- Motherboard -->
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header bg-secondary text-white">
                        <i class="bi bi-motherboard section-icon"></i>Motherboard
                    </div>
                    <div class="card-body p-0">
                        <table class="table table-striped">
                            <tbody>
                                <tr><td class="info-label">Manufacturer</td><td>$mbManufacturer</td></tr>
                                <tr><td class="info-label">Product</td><td>$mbProduct</td></tr>
                                <tr><td class="info-label">Serial Number</td><td>$mbSerial</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <!-- Memory -->
        <div class="card">
            <div class="card-header bg-warning text-dark">
                <i class="bi bi-memory section-icon"></i>Memory (RAM) - Total: $totalRAM GB
            </div>
            <div class="card-body p-0">
                <table class="table table-striped">
                    <thead class="table-light">
                        <tr><th>Slot</th><th>Manufacturer</th><th>Capacity</th><th>Speed</th><th>Part Number</th></tr>
                    </thead>
                    <tbody>
                        $ramRows
                    </tbody>
                </table>
            </div>
        </div>

        <!-- GPU -->
        <div class="card">
            <div class="card-header bg-danger text-white">
                <i class="bi bi-gpu-card section-icon"></i>Graphics Card(s)
            </div>
            <div class="card-body p-0">
                <table class="table table-striped">
                    <thead class="table-light">
                        <tr><th>Name</th><th>VRAM</th><th>Driver</th><th>Resolution</th></tr>
                    </thead>
                    <tbody>
                        $gpuRows
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Storage -->
        <div class="card">
            <div class="card-header bg-dark text-white">
                <i class="bi bi-device-hdd section-icon"></i>Storage
            </div>
            <div class="card-body p-0">
                <table class="table table-striped">
                    <thead class="table-light">
                        <tr><th>Model</th><th>Capacity</th><th>Interface</th><th>Serial Number</th></tr>
                    </thead>
                    <tbody>
                        $diskRows
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Network -->
        <div class="card">
            <div class="card-header" style="background-color: #6f42c1; color: white;">
                <i class="bi bi-ethernet section-icon"></i>Network
            </div>
            <div class="card-body p-0">
                <table class="table table-striped">
                    <thead class="table-light">
                        <tr><th>Adapter</th><th>IP Address</th><th>MAC Address</th></tr>
                    </thead>
                    <tbody>
                        $netRows
                    </tbody>
                </table>
            </div>
        </div>

        <p class="text-center text-muted mt-3 mb-4">
            <small>Report automatically generated by PowerShell | $reportDate</small>
        </p>
    </div>
</body>
</html>
"@

# --- Save and open ---
$outputPath = Join-Path $PSScriptRoot "System-Report-$hostname.html"
$html | Out-File -FilePath $outputPath -Encoding UTF8
Write-Host "Report generated: $outputPath" -ForegroundColor Green
Start-Process $outputPath
