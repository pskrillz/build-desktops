param([switch]$DryRun)

# Fix for Unicode character display in console
$OutputEncoding = [System.Console]::OutputEncoding

Import-Module VirtualDesktop -DisableNameChecking

# Configuration for desktops and applications
$Layout = @(
    @{ Name='Sports'
        ; Exe='C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'
        ; Args='--new-window https://www.espn.com/nba/scoreboard'
        ; Proc='brave' },
    @{ Name='Files'
        ; Exe='C:\Windows\explorer.exe'
        ; Args='C:\Programs'
        ; Proc='explorer' }
    # Example for 'Coding' desktop (add if needed):
    # @{ Name='Coding'
    # ; Exe='notepad.exe' # Or your preferred code editor
    # ; Args='C:\Path\To\Your\code_snippet.txt' # Optional: path to a file to open
    # ; Proc='notepad' } # Or the process name of your editor
)

# Helper function for logging with timestamp and level
function WL($msg, $lvl='INFO'){
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formattedMsg = $msg.TrimStart()
    if ($formattedMsg.StartsWith("•")) {
        $formattedMsg = " " + $formattedMsg 
    } elseif (-not ($formattedMsg.StartsWith("===") -or $formattedMsg.StartsWith("Processing Desktop"))) {
        $formattedMsg = " • " + $formattedMsg 
    }
    Write-Host ("{0} [{1}]{2}" -f $timestamp, $lvl.ToUpper(), $formattedMsg)
}

WL "=== Workspace build started (DryRun=$DryRun) ==="

foreach ($slot in $Layout){

    WL "Processing Desktop '$($slot.Name)' …" -lvl 'INFO'
    $desk = $null # Initialize/reset $desk for the current slot

    # --- Phase 1: Get or Create Desktop ---
    $desk = Get-Desktop | Where-Object Name -eq $slot.Name -ErrorAction SilentlyContinue
    
    if ($desk) {
        WL "Desktop '$($slot.Name)' already exists (ID: $($desk.Id), Name: '$($desk.Name)')"
    } else {
        # Desktop does not exist, proceed with creation if not in DryRun
        if ($DryRun) {
            WL "Would create desktop '$($slot.Name)'."
            # For DryRun, we'll simulate a successful desktop prep for subsequent "would launch/move" logs
            # This is a placeholder object for DryRun logging consistency.
            $desk = [PSCustomObject]@{Id="[DryRun-NewID]"; Name=$slot.Name; IsCurrent=$false} 
        } else {
            WL "Desktop '$($slot.Name)' not found. Attempting to create and name..."
            $newlyCreatedDesktopObject = $null
            try {
                $newlyCreatedDesktopObject = New-Desktop -ErrorAction Stop
                if (-not $newlyCreatedDesktopObject) {
                    WL "FATAL: New-Desktop did not return a desktop object for '$($slot.Name)'." 'ERROR'
                    continue # Move to the next slot in $Layout
                }
                # Log the state of the object immediately after creation
                WL "Raw desktop object created. ID: '$($newlyCreatedDesktopObject.Id)', Name: '$($newlyCreatedDesktopObject.Name)', IsCurrent: '$($newlyCreatedDesktopObject.IsCurrent)'" 'DEBUG'

                Set-DesktopName -Desktop $newlyCreatedDesktopObject -Name $slot.Name -ErrorAction Stop
                WL "Set-DesktopName called for '$($slot.Name)' on desktop ID '$($newlyCreatedDesktopObject.Id)'."
                
                # PATIENCE: Give the system a significant pause to register the name and settle.
                $settleTimeSeconds = 4 # Increased delay
                WL "Pausing for $settleTimeSeconds seconds for desktop name to settle..."
                Start-Sleep -Seconds $settleTimeSeconds

                # We will now TRUST $newlyCreatedDesktopObject to be the correct, named desktop.
                # Let's check its properties again after the delay and naming.
                WL "Desktop object state after naming and delay. ID: '$($newlyCreatedDesktopObject.Id)', Name: '$($newlyCreatedDesktopObject.Name)', IsCurrent: '$($newlyCreatedDesktopObject.IsCurrent)'" 'DEBUG'
                
                # Even if $newlyCreatedDesktopObject.Name doesn't reflect $slot.Name immediately,
                # the object itself should be the handle to the correctly named desktop.
                $desk = $newlyCreatedDesktopObject

            } catch {
                WL "ERROR during desktop creation/naming for '$($slot.Name)': $($_.Exception.Message)" 'ERROR'
                if ($newlyCreatedDesktopObject) {
                    WL "Attempting to clean up potentially unnamed/misnamed desktop (ID: $($newlyCreatedDesktopObject.Id))" 'INFO'
                    Remove-Desktop -Desktop $newlyCreatedDesktopObject -ErrorAction SilentlyContinue
                }
                continue # Move to the next slot
            }
        } 
    } 

    # --- Phase 2: Launch Application ---
    # If DryRun, $desk might be a real existing one or the placeholder.
    # If not DryRun, $desk MUST be valid (either pre-existing or newly created and assumed good).
    if (-not $desk) {
        WL "Cannot proceed to launch/move for slot '$($slot.Name)' as a valid desktop object was not obtained." 'WARN'
        continue
    }

    # If in DryRun and $desk was a placeholder (ID starts with [DryRun-NewID]), log appropriately.
    if ($DryRun -and $desk.Id -like "[DryRun-NewID]*") {
         WL "Would launch: $($slot.Exe) $($slot.Args) for desktop '$($slot.Name)' (would be created)."
         WL "Would then hunt for process name: '$($slot.Proc)' and move to '$($slot.Name)'."
         continue
    }
    # If DryRun and $desk was an existing desktop.
    if ($DryRun -and $desk.Id -notlike "[DryRun-NewID]*") {
         WL "Would launch: $($slot.Exe) $($slot.Args) for desktop '$($desk.Name)' (existing)."
         WL "Would then hunt for process name: '$($slot.Proc)' and move to '$($desk.Name)'."
         continue
    }

    # --- Actual Launch for Non-DryRun ---
    WL "Launching: '$($slot.Exe)' $($slot.Args) for desktop '$($slot.Name)' (Actual Name on Object: '$($desk.Name)', ID: '$($desk.Id)')"
    $spParams = @{ FilePath = $slot.Exe; PassThru = $true }
    if (![string]::IsNullOrEmpty($slot.Args)) {
        $spParams.ArgumentList = $slot.Args
    } elseif ([string]::IsNullOrEmpty($slot.Args) -and $slot.PSBase.ContainsKey('Args')) {
        $spParams.ArgumentList = ''
    }

    $parentProcess = $null
    try {
        $parentProcess = Start-Process @spParams -ErrorAction Stop
        if ($parentProcess -and $parentProcess.Id) {
            WL "Successfully initiated process '$($slot.Exe)' (PID: $($parentProcess.Id))"
        } else {
            WL "Initiated process '$($slot.Exe)'. PID not captured by PassThru or process exited quickly." 'INFO'
        }
    } catch {
        WL "ERROR launching '$($slot.Exe)': $($_.Exception.Message)" 'ERROR'
        continue 
    }

    # --- Phase 3: Find Application Window ---
    $windowHuntTimeoutSeconds = 45 
    $huntDeadline = (Get-Date).AddSeconds($windowHuntTimeoutSeconds)
    $windowProcess = $null

    WL "Hunting for window of process name '$($slot.Proc)' for up to $windowHuntTimeoutSeconds seconds..."
    while (($null -eq $windowProcess) -and ((Get-Date) -lt $huntDeadline)){
        Start-Sleep -Milliseconds 500
        $foundProcesses = Get-Process -Name $slot.Proc -ErrorAction SilentlyContinue |
                         Where-Object {$_.MainWindowHandle -ne [System.IntPtr]::Zero -and !([string]::IsNullOrWhiteSpace($_.MainWindowTitle))}
        
        if ($foundProcesses) {
            $windowProcess = $foundProcesses | Select-Object -First 1
            WL "Found process '$($windowProcess.Name)' (PID: $($windowProcess.Id), Title: '$($windowProcess.MainWindowTitle)') with a main window."
        }
    }

    if (-not $windowProcess){
        WL "No window found for process name '$($slot.Proc)' after $windowHuntTimeoutSeconds seconds. Skipping move." 'WARN'
        continue 
    }
    
    # --- Phase 4: Move Window ---
    # We use $desk, which is the direct object reference we've been holding onto.
    WL "Attempting to move window (Handle: $($windowProcess.MainWindowHandle)) of '$($windowProcess.Name)' (PID: $($windowProcess.Id)) to target desktop '$($slot.Name)' (using object with ID '$($desk.Id)', Current Name on Object: '$($desk.Name)')..."
    try {
        $windowProcess.MainWindowHandle | Move-Window -Desktop $desk -ErrorAction Stop
        WL "Moved window for '$($windowProcess.Name)' to desktop '$($slot.Name)' successfully."
    } catch {
        WL "ERROR: Move-Window failed for '$($windowProcess.Name)' (PID: $($windowProcess.Id)) to desktop '$($slot.Name)'. Error: $($_.Exception.Message)" 'ERROR'
        WL "  Desktop object used: ID '$($desk.Id)', Name on Object '$($desk.Name)', IsCurrent '$($desk.IsCurrent)'" 'DEBUG'
        WL "  Window Handle was: $($windowProcess.MainWindowHandle)" 'DEBUG'
    }
}

WL "=== Workspace build finished ==="
