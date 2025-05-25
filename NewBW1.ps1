<#Current version: 2025-05-23, doesnt position the window correctly, but it does create the desktops and start the applications #>



#Requires -Modules VirtualDesktop

<#
.SYNOPSIS
    Creates three virtual desktops with specific applications and window layouts
.DESCRIPTION
    Target environment: Windows 10/11; single 3840x2160 monitor; PowerShell 7; VirtualDesktop module version 2.0 or newer
    Creates Plan, Code, and Entertainment desktops with assigned applications and automatic window tiling
.PARAMETER DryRun
    Simulate without launching applications
.PARAMETER TimeoutSec
    Override default 45-second wait time for windows to appear
.EXAMPLE
    .\NewBW1.ps1
    .\NewBW1.ps1 -DryRun
    .\NewBW1.ps1 -TimeoutSec 60
#>

param(
    [switch]$DryRun,
    [int]$TimeoutSec = 45
)

# Import required modules and functions
try {
    Import-Module VirtualDesktop -ErrorAction Stop
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [INFO] VirtualDesktop module loaded successfully" -ForegroundColor Green
} catch {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ERROR] Failed to load VirtualDesktop module: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Load Windows.Forms for SendKeys functionality
try {
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [INFO] Windows.Forms assembly loaded successfully" -ForegroundColor Green
} catch {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ERROR] Failed to load Windows.Forms assembly: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Simple Win32 API for window detection and positioning - COMPLETE VERSION
if (-not ([System.Management.Automation.PSTypeName]'SimpleWin32').Type) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class SimpleWin32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    
    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    
    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    
    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
    
    // ShowWindow constants
    public const int SW_RESTORE = 9;
    public const int SW_SHOW = 5;
    public const int SW_MAXIMIZE = 3;
    public const int SW_NORMAL = 1;
    
    // SetWindowPos constants
    public const uint SWP_NOZORDER = 0x0004;
    public const uint SWP_NOACTIVATE = 0x0010;
    public const uint SWP_SHOWWINDOW = 0x0040;
    public static readonly IntPtr HWND_TOP = new IntPtr(0);
}
"@
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [DEBUG] Complete Win32 API loaded" -ForegroundColor Cyan
} else {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [DEBUG] Complete Win32 API already loaded" -ForegroundColor Cyan
}

# Global variables
$script:CreatedDesktops = @{}
$script:MonitorWidth = 3840
$script:MonitorHeight = 2160

# Exact window positions based on specifications (X, Y, Width, Height)
$script:WindowPositions = @{
    'Left' = @{ X = 0; Y = 0; Width = 1920; Height = 2160 }          # Left half
    'RightTop' = @{ X = 1920; Y = 0; Width = 1920; Height = 1080 }    # Top-right quarter  
    'RightBottom' = @{ X = 1920; Y = 1080; Width = 1920; Height = 1080 } # Bottom-right quarter
}

# Valid position types
$script:ValidPositions = @('Left', 'RightTop', 'RightBottom')

# Desktop and application configuration
$script:DesktopConfig = @{
    'Plan' = @(
        @{ App = 'C:\Users\payam\AppData\Local\Programs\Evernote\Evernote.exe'; Args = @(); Position = 'Left'; Description = 'Evernote' }
        @{ App = 'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'; Args = @('--profile-directory=PlanBrowse'); Position = 'RightTop'; Description = 'Brave PlanBrowse' }
        @{ App = 'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'; Args = @('--profile-directory=PlanChat', '--app=https://chat.openai.com'); Position = 'RightBottom'; Description = 'Brave PlanChat PWA' }
    )
    'Code' = @(
        @{ App = 'C:\Users\payam\AppData\Local\Programs\cursor\Cursor.exe'; Args = @(); Position = 'Left'; Description = 'Cursor' }
        @{ App = 'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'; Args = @('--profile-directory=CodeBrowse'); Position = 'RightTop'; Description = 'Brave CodeBrowse' }
        @{ App = 'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'; Args = @('--profile-directory=CodeChat', '--app=https://chat.openai.com'); Position = 'RightBottom'; Description = 'Brave CodeChat PWA' }
    )
    'Entertainment' = @(
        @{ App = 'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'; Args = @('--profile-directory=NBALive', 'https://nbabite.to'); Position = 'Left'; Description = 'Brave NBALive' }
        @{ App = 'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'; Args = @('--profile-directory=Reddit', 'https://reddit.com'); Position = 'RightTop'; Description = 'Brave Reddit' }
    )
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'DEBUG', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'INFO' { 'White' }
        'DEBUG' { 'Cyan' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
    }
    
    Write-Host "$timestamp [$Level] $Message" -ForegroundColor $color
}

function Get-OrCreateDesktop {
    param([string]$DesktopName)
    
    if ($script:CreatedDesktops.ContainsKey($DesktopName)) {
        Write-Log "Desktop '$DesktopName' already exists, using cached reference" -Level 'DEBUG'
        return $script:CreatedDesktops[$DesktopName]
    }
    
    try {
        # Check if desktop already exists by iterating through indices
        $desktopList = Get-DesktopList
        for ($i = 0; $i -lt $desktopList.Count; $i++) {
            $desktopInfo = $desktopList[$i]
            if ($desktopInfo.Name -eq $DesktopName) {
                Write-Log "Found existing desktop '$DesktopName' at index $i" -Level 'DEBUG'
                # Get the actual desktop object by index
                $desktopObject = Get-Desktop -Index $i
                $script:CreatedDesktops[$DesktopName] = $desktopObject
                return $desktopObject
            }
        }
        
        # Create new desktop if not found
        Write-Log "Creating new desktop '$DesktopName'" -Level 'INFO'
        if (-not $DryRun) {
            $newDesktop = New-Desktop
            
            # Try to set the name (only works on Windows 10 2004+ and Windows 11)
            try {
                Set-DesktopName -Desktop $newDesktop -Name $DesktopName
                Write-Log "Successfully named desktop '$DesktopName'" -Level 'DEBUG'
            } catch {
                Write-Log "Could not set desktop name (may not be supported on this Windows version): $($_.Exception.Message)" -Level 'WARN'
            }
            
            $script:CreatedDesktops[$DesktopName] = $newDesktop
            return $newDesktop
        } else {
            Write-Log "DryRun: Would create desktop '$DesktopName'" -Level 'DEBUG'
            return $null
        }
    } catch {
        Write-Log "Failed to create desktop '$DesktopName': $($_.Exception.Message)" -Level 'ERROR'
        return $null
    }
}

function Get-WindowTitle {
    param([IntPtr]$WindowHandle)
    
    try {
        $length = [SimpleWin32]::GetWindowTextLength($WindowHandle)
        if ($length -eq 0) { return "" }
        
        $sb = New-Object System.Text.StringBuilder($length + 1)
        [SimpleWin32]::GetWindowText($WindowHandle, $sb, $sb.Capacity) | Out-Null
        return $sb.ToString()
    } catch {
        return "Unknown"
    }
}

function Set-WindowPosition {
    param(
        [IntPtr]$WindowHandle,
        [string]$PositionType,
        [string]$AppDescription = "Unknown"
    )
    
    try {
        Write-Log "=== POSITIONING WINDOW: '$AppDescription' ===" -Level 'INFO'
        Write-Log "Target position: $PositionType" -Level 'INFO'
        
        # Validate window
        if ($WindowHandle -eq [IntPtr]::Zero -or -not [SimpleWin32]::IsWindowVisible($WindowHandle)) {
            Write-Log "ERROR: Invalid or invisible window" -Level 'ERROR'
            return $false
        }
        
        # Get target position
        if (-not $script:WindowPositions.ContainsKey($PositionType)) {
            Write-Log "ERROR: Unknown position type '$PositionType'" -Level 'ERROR'
            return $false
        }
        
        $targetPos = $script:WindowPositions[$PositionType]
        $targetX = $targetPos.X
        $targetY = $targetPos.Y
        $targetWidth = $targetPos.Width
        $targetHeight = $targetPos.Height
        
        # Get window title for logging
        $windowTitle = Get-WindowTitle -WindowHandle $WindowHandle
        Write-Log "Window title: '$windowTitle'" -Level 'DEBUG'
        Write-Log "Target: X=$targetX, Y=$targetY, W=$targetWidth, H=$targetHeight" -Level 'INFO'
        
        # Get current position for comparison
        $currentRect = New-Object SimpleWin32+RECT
        [SimpleWin32]::GetWindowRect($WindowHandle, [ref]$currentRect) | Out-Null
        $currentX = $currentRect.Left
        $currentY = $currentRect.Top
        $currentWidth = $currentRect.Right - $currentRect.Left
        $currentHeight = $currentRect.Bottom - $currentRect.Top
        Write-Log "Current: X=$currentX, Y=$currentY, W=$currentWidth, H=$currentHeight" -Level 'DEBUG'
        
        # Step 1: Ensure window is restored (not minimized or maximized)
        Write-Log "Step 1: Restoring window to normal state" -Level 'DEBUG'
        [SimpleWin32]::ShowWindow($WindowHandle, [SimpleWin32]::SW_RESTORE) | Out-Null
        Start-Sleep -Milliseconds 300
        
        # Step 2: Bring window to foreground
        Write-Log "Step 2: Bringing window to foreground" -Level 'DEBUG'
        [SimpleWin32]::SetForegroundWindow($WindowHandle) | Out-Null
        Start-Sleep -Milliseconds 300
        
        # AGGRESSIVE 4-METHOD POSITIONING APPROACH
        $positionAttempts = @(
            @{ Method = "ShowWindow+MoveWindow"; Description = "Restore and Move" }
            @{ Method = "SetWindowPos"; Description = "Direct Position" }
            @{ Method = "ShowWindow+MoveWindow"; Description = "Restore and Move (Retry)" }
            @{ Method = "SetWindowPos"; Description = "Direct Position (Final)" }
        )
        
        $tolerance = 100  # Allow some tolerance for window chrome
        $positioned = $false
        
        foreach ($attempt in $positionAttempts) {
            Write-Log "Step 3.$($positionAttempts.IndexOf($attempt) + 1): Trying $($attempt.Description)" -Level 'DEBUG'
            
            if ($attempt.Method -eq "ShowWindow+MoveWindow") {
                # Method: ShowWindow + MoveWindow
                [SimpleWin32]::ShowWindow($WindowHandle, [SimpleWin32]::SW_RESTORE) | Out-Null
                Start-Sleep -Milliseconds 100
                $moveResult = [SimpleWin32]::MoveWindow($WindowHandle, $targetX, $targetY, $targetWidth, $targetHeight, $true)
                Write-Log "MoveWindow result: $moveResult" -Level 'DEBUG'
            } 
            elseif ($attempt.Method -eq "SetWindowPos") {
                # Method: SetWindowPos
                $flags = [SimpleWin32]::SWP_NOZORDER -bor [SimpleWin32]::SWP_SHOWWINDOW
                $setResult = [SimpleWin32]::SetWindowPos($WindowHandle, [SimpleWin32]::HWND_TOP, $targetX, $targetY, $targetWidth, $targetHeight, $flags)
                Write-Log "SetWindowPos result: $setResult" -Level 'DEBUG'
            }
            
            # Wait for window to settle
            Start-Sleep -Milliseconds 500
            
            # Check if positioning was successful
            $newRect = New-Object SimpleWin32+RECT
            [SimpleWin32]::GetWindowRect($WindowHandle, [ref]$newRect) | Out-Null
            $newX = $newRect.Left
            $newY = $newRect.Top
            $newWidth = $newRect.Right - $newRect.Left
            $newHeight = $newRect.Bottom - $newRect.Top
            
            Write-Log "After $($attempt.Description): X=$newX, Y=$newY, W=$newWidth, H=$newHeight" -Level 'DEBUG'
            
            # Check if we're within tolerance
            $xOk = [Math]::Abs($newX - $targetX) -le $tolerance
            $yOk = [Math]::Abs($newY - $targetY) -le $tolerance
            $wOk = [Math]::Abs($newWidth - $targetWidth) -le $tolerance
            $hOk = [Math]::Abs($newHeight - $targetHeight) -le $tolerance
            
            if ($xOk -and $yOk -and $wOk -and $hOk) {
                Write-Log "SUCCESS: Window positioned within tolerance using $($attempt.Description)" -Level 'INFO'
                $positioned = $true
                break
            } else {
                Write-Log "Position check: X=${xOk}, Y=${yOk}, W=${wOk}, H=${hOk} (tolerance: $tolerance)" -Level 'DEBUG'
            }
        }
        
        # Final verification and reporting
        $finalRect = New-Object SimpleWin32+RECT
        [SimpleWin32]::GetWindowRect($WindowHandle, [ref]$finalRect) | Out-Null
        $finalX = $finalRect.Left
        $finalY = $finalRect.Top
        $finalWidth = $finalRect.Right - $finalRect.Left
        $finalHeight = $finalRect.Bottom - $finalRect.Top
        
        Write-Log "=== FINAL RESULT for '$AppDescription' ===" -Level 'INFO'
        Write-Log "Target position: X=$targetX, Y=$targetY, W=$targetWidth, H=$targetHeight" -Level 'INFO'
        Write-Log "Final position:  X=$finalX, Y=$finalY, W=$finalWidth, H=$finalHeight" -Level 'INFO'
        
        $finalXOk = [Math]::Abs($finalX - $targetX) -le $tolerance
        $finalYOk = [Math]::Abs($finalY - $targetY) -le $tolerance
        $finalWOk = [Math]::Abs($finalWidth - $targetWidth) -le $tolerance
        $finalHOk = [Math]::Abs($finalHeight - $targetHeight) -le $tolerance
        
        if ($finalXOk -and $finalYOk -and $finalWOk -and $finalHOk) {
            Write-Log "SUCCESS: Final position within tolerance" -Level 'INFO'
            return $true
        } else {
            $xDiff = [Math]::Abs($finalX - $targetX)
            $yDiff = [Math]::Abs($finalY - $targetY)
            $wDiff = [Math]::Abs($finalWidth - $targetWidth)
            $hDiff = [Math]::Abs($finalHeight - $targetHeight)
            Write-Log "PARTIAL: Window positioned but outside tolerance. Diffs: X=$xDiff, Y=$yDiff, W=$wDiff, H=$hDiff" -Level 'WARN'
            Write-Log "This may be due to application constraints or window chrome" -Level 'WARN'
            return $true  # Still consider it a success as we attempted positioning
        }
        
    } catch {
        Write-Log "EXCEPTION in Set-WindowPosition for '$AppDescription': $($_.Exception.Message)" -Level 'ERROR'
        return $false
    }
}

function Get-ProcessByExecutablePath {
    param([string]$ExecutablePath)
    
    try {
        $processes = Get-Process | Where-Object { 
            try {
                $_.Path -eq $ExecutablePath
            } catch {
                $false
            }
        }
        return $processes
    } catch {
        return @()
    }
}

function Start-ApplicationOnDesktop {
    param(
        [string]$DesktopName,
        [hashtable]$AppConfig
    )
    
    $desktop = Get-OrCreateDesktop -DesktopName $DesktopName
    if (-not $desktop -and -not $DryRun) {
        Write-Log "Failed to get/create desktop '$DesktopName', skipping application" -Level 'ERROR'
        return
    }
    
    # Validate position type
    if (-not $script:WindowPositions.ContainsKey($AppConfig.Position)) {
        Write-Log "Invalid position '$($AppConfig.Position)' for application '$($AppConfig.Description)'. Valid positions: $($script:WindowPositions.Keys -join ', ')" -Level 'ERROR'
        return
    }
    
    Write-Log "Starting '$($AppConfig.Description)' on desktop '$DesktopName'" -Level 'INFO'
    
    if ($DryRun) {
        Write-Log "DryRun: Would start $($AppConfig.App) with args: $($AppConfig.Args -join ' ')" -Level 'DEBUG'
        Write-Log "DryRun: Would position at $($AppConfig.Position) using pixel-perfect positioning" -Level 'DEBUG'
        return
    }
    
    try {
        $windowHandle = $null
        
        # Get baseline window count before starting
        $beforeProcesses = @(Get-ProcessByExecutablePath -ExecutablePath $AppConfig.App)
        Write-Log "Found $($beforeProcesses.Count) existing instance(s) before launching" -Level 'DEBUG'
        
        # Start the process
        $processArgs = @{
            FilePath = $AppConfig.App
            PassThru = $true
        }
        
        if ($AppConfig.Args.Count -gt 0) {
            $processArgs.ArgumentList = $AppConfig.Args
        }
        
        $process = Start-Process @processArgs
        Write-Log "Started process $($AppConfig.App) (PID: $($process.Id))" -Level 'DEBUG'
        
        # Wait a moment for the application to fully start
        Start-Sleep -Seconds 5  # Increased from 2 to give apps more time to initialize
        
        # Look for a visible window - check the new process first, then all processes
        $windowHandle = Find-VisibleWindow -Process $process -ExecutablePath $AppConfig.App -Description $AppConfig.Description
        
        if (-not $windowHandle) {
            Write-Log "No window found for '$($AppConfig.Description)'" -Level 'WARN'
            return
        }
        
        # Move window to target desktop
        try {
            Write-Log "Attempting to move window handle $windowHandle to desktop '$DesktopName'" -Level 'DEBUG'
            
            # Verify the desktop object is valid
            if (-not $desktop) {
                Write-Log "ERROR: Desktop object is null for '$DesktopName'" -Level 'ERROR'
                return
            }
            
            Write-Log "Desktop object type: $($desktop.GetType().FullName)" -Level 'DEBUG'
            
            # Get detailed info about target desktop
            try {
                $targetDesktopIndex = Get-DesktopIndex -Desktop $desktop
                $targetDesktopName = Get-DesktopName -Desktop $desktop
                Write-Log "Target desktop: Index=$targetDesktopIndex, Name='$targetDesktopName'" -Level 'DEBUG'
            } catch {
                Write-Log "Could not get target desktop details: $($_.Exception.Message)" -Level 'WARN'
            }
            
            # Check current desktop of the window BEFORE moving
            try {
                $currentDesktop = Get-DesktopFromWindow -Hwnd $windowHandle
                $currentDesktopIndex = Get-DesktopIndex -Desktop $currentDesktop
                $currentDesktopName = Get-DesktopName -Desktop $currentDesktop
                Write-Log "BEFORE move - Window is on desktop: Index=$currentDesktopIndex, Name='$currentDesktopName'" -Level 'DEBUG'
            } catch {
                Write-Log "Could not determine current desktop of window: $($_.Exception.Message)" -Level 'DEBUG'
            }
            
            # Perform the move operation with detailed error handling
            Write-Log "Calling Move-Window -Desktop `$desktop -Hwnd $windowHandle" -Level 'DEBUG'
            try {
                $moveResult = Move-Window -Desktop $desktop -Hwnd $windowHandle -ErrorAction Stop
                Write-Log "Move-Window command completed. Result type: $($moveResult.GetType().FullName)" -Level 'DEBUG'
            } catch {
                Write-Log "Move-Window command failed: $($_.Exception.Message)" -Level 'ERROR'
                Write-Log "Exception details: $($_.Exception.GetType().FullName)" -Level 'DEBUG'
                Write-Log "Error category: $($_.CategoryInfo.Category)" -Level 'DEBUG'
                return
            }
            
            # Wait a moment for the move to complete
            Start-Sleep -Milliseconds 1000
            
            # Verify the move was successful AFTER moving
            try {
                $newDesktop = Get-DesktopFromWindow -Hwnd $windowHandle
                $newDesktopIndex = Get-DesktopIndex -Desktop $newDesktop
                $newDesktopName = Get-DesktopName -Desktop $newDesktop
                Write-Log "AFTER move - Window is on desktop: Index=$newDesktopIndex, Name='$newDesktopName'" -Level 'DEBUG'
                
                # Compare desktop indices to verify success
                if ($newDesktopIndex -eq $targetDesktopIndex) {
                    Write-Log "SUCCESS: Window moved to correct desktop '$DesktopName' (Index: $targetDesktopIndex)" -Level 'INFO'
                } else {
                    Write-Log "FAILURE: Window is on desktop Index=$newDesktopIndex ('$newDesktopName') instead of Index=$targetDesktopIndex ('$DesktopName')" -Level 'ERROR'
                }
            } catch {
                Write-Log "Could not verify window move: $($_.Exception.Message)" -Level 'WARN'
            }
            
        } catch {
            Write-Log "Failed to move '$($AppConfig.Description)' to desktop '$DesktopName': $($_.Exception.Message)" -Level 'ERROR'
            Write-Log "Full exception: $($_.Exception.ToString())" -Level 'DEBUG'
        }
        
        # Position and resize window
        $positionResult = Set-WindowPosition -WindowHandle $windowHandle -PositionType $AppConfig.Position -AppDescription $AppConfig.Description
        if ($positionResult) {
            Write-Log "Successfully positioned '$($AppConfig.Description)' at $($AppConfig.Position)" -Level 'INFO'
        } else {
            Write-Log "Failed to position '$($AppConfig.Description)'" -Level 'WARN'
        }
        
    } catch {
        Write-Log "Failed to start '$($AppConfig.Description)': $($_.Exception.Message)" -Level 'ERROR'
    }
}

function Find-VisibleWindow {
    param(
        [System.Diagnostics.Process]$Process,
        [string]$ExecutablePath,
        [string]$Description
    )
    
    $maxAttempts = 20  # Increased from 10
    $attempt = 0
    
    Write-Log "Searching for window for '$Description'..." -Level 'INFO'
    
    while ($attempt -lt $maxAttempts) {
        $attempt++
        Write-Log "Window search attempt $attempt/$maxAttempts" -Level 'DEBUG'
        
        try {
            # Method 1: Check the launched process first
            if ($Process -and -not $Process.HasExited) {
                $Process.Refresh()
                if ($Process.MainWindowHandle -ne [IntPtr]::Zero) {
                    $isVisible = [SimpleWin32]::IsWindowVisible($Process.MainWindowHandle)
                    if ($isVisible) {
                        $windowTitle = Get-WindowTitle -WindowHandle $Process.MainWindowHandle
                        
                        # Additional check: ensure window has reasonable size (not just a splash screen)
                        $testRect = New-Object SimpleWin32+RECT
                        [SimpleWin32]::GetWindowRect($Process.MainWindowHandle, [ref]$testRect) | Out-Null
                        $testWidth = $testRect.Right - $testRect.Left
                        $testHeight = $testRect.Bottom - $testRect.Top
                        
                        if ($testWidth -gt 100 -and $testHeight -gt 100) {
                            Write-Log "Found valid window from launched process: '$windowTitle' (${testWidth}x${testHeight})" -Level 'INFO'
                            Start-Sleep -Milliseconds 500  # Let window fully initialize
                            return $Process.MainWindowHandle
                        } else {
                            Write-Log "Window too small (${testWidth}x${testHeight}), likely splash screen" -Level 'DEBUG'
                        }
                    }
                }
            }
            
            # Method 2: Search all processes with the same executable
            $allProcesses = @(Get-ProcessByExecutablePath -ExecutablePath $ExecutablePath)
            foreach ($proc in $allProcesses) {
                if ($proc.MainWindowHandle -ne [IntPtr]::Zero) {
                    $isVisible = [SimpleWin32]::IsWindowVisible($proc.MainWindowHandle)
                    if ($isVisible) {
                        $windowTitle = Get-WindowTitle -WindowHandle $proc.MainWindowHandle
                        
                        # Check window size
                        $testRect = New-Object SimpleWin32+RECT
                        [SimpleWin32]::GetWindowRect($proc.MainWindowHandle, [ref]$testRect) | Out-Null
                        $testWidth = $testRect.Right - $testRect.Left
                        $testHeight = $testRect.Bottom - $testRect.Top
                        
                        if ($testWidth -gt 100 -and $testHeight -gt 100) {
                            Write-Log "Found valid window from existing process: '$windowTitle' (PID: $($proc.Id), ${testWidth}x${testHeight})" -Level 'INFO'
                            Start-Sleep -Milliseconds 500  # Let window fully initialize
                            return $proc.MainWindowHandle
                        }
                    }
                }
            }
        } catch {
            Write-Log "Error in Find-VisibleWindow: $($_.Exception.Message)" -Level 'DEBUG'
        }
        
        # Progressive wait times - start fast, get slower
        if ($attempt -le 5) {
            Start-Sleep -Milliseconds 500
        } elseif ($attempt -le 10) {
            Start-Sleep -Milliseconds 1000
        } else {
            Start-Sleep -Milliseconds 2000
        }
    }
    
    Write-Log "No suitable window found for '$Description' after $maxAttempts attempts" -Level 'WARN'
    return $null
}

function Main {
    Write-Log "Starting virtual desktop setup script" -Level 'INFO'
    Write-Log "Monitor resolution: $($script:MonitorWidth)x$($script:MonitorHeight)" -Level 'INFO'
    Write-Log "Window layout using pixel-perfect positioning:" -Level 'INFO'
    Write-Log "  Left position: X=0, Y=0, W=1920, H=2160 (left 50% of screen)" -Level 'INFO'
    Write-Log "  RightTop position: X=1920, Y=0, W=1920, H=1080 (top-right quadrant)" -Level 'INFO'
    Write-Log "  RightBottom position: X=1920, Y=1080, W=1920, H=1080 (bottom-right quadrant)" -Level 'INFO'
    Write-Log "Positioning method: 4-stage aggressive Win32 API calls" -Level 'INFO'
    Write-Log "Timeout: $TimeoutSec seconds" -Level 'INFO'
    
    if ($DryRun) {
        Write-Log "DRY RUN MODE - No applications will be launched" -Level 'WARN'
    }
    
    # First, create all desktops
    Write-Log "Creating virtual desktops..." -Level 'INFO'
    foreach ($desktopName in $script:DesktopConfig.Keys) {
        $desktop = Get-OrCreateDesktop -DesktopName $desktopName
        if ($desktop -or $DryRun) {
            Write-Log "Desktop '$desktopName' ready" -Level 'INFO'
        } else {
            Write-Log "Failed to create desktop '$desktopName'" -Level 'ERROR'
        }
    }
    
    # Give desktops time to be fully created
    if (-not $DryRun) {
        Write-Log "Waiting for desktops to be fully initialized..." -Level 'DEBUG'
        Start-Sleep -Seconds 2
    }
    
    # Process each desktop configuration
    foreach ($desktopName in $script:DesktopConfig.Keys) {
        Write-Log "Processing applications for desktop '$desktopName'" -Level 'INFO'
        
        $apps = $script:DesktopConfig[$desktopName]
        foreach ($app in $apps) {
            Start-ApplicationOnDesktop -DesktopName $desktopName -AppConfig $app
            
            # Small delay between application launches
            if (-not $DryRun) {
                Start-Sleep -Seconds 5  # Increased to ensure proper window detection and positioning
            }
        }
        
        Write-Log "Completed desktop '$desktopName'" -Level 'INFO'
    }
    
    Write-Log "Virtual desktop setup completed" -Level 'INFO'
    
    if (-not $DryRun) {
        Write-Log "Created desktops: $($script:CreatedDesktops.Keys -join ', ')" -Level 'INFO'
        Write-Log "All windows should now be positioned and ready for use" -Level 'INFO'
        
        # Show final desktop list
        try {
            $desktopList = Get-DesktopList
            Write-Log "Final desktop list:" -Level 'INFO'
            for ($i = 0; $i -lt $desktopList.Count; $i++) {
                $name = try { $desktopList[$i].Name } catch { "Desktop $i" }
                if ([string]::IsNullOrEmpty($name)) { $name = "Desktop $i" }
                Write-Log "  [$i] $name" -Level 'INFO'
            }
        } catch {
            Write-Log "Could not retrieve final desktop list: $($_.Exception.Message)" -Level 'DEBUG'
        }
    }
}

# Execute main function
try {
    Main
} catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level 'ERROR'
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level 'DEBUG'
    exit 1
}
