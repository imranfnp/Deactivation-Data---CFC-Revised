# ============================================================
# GrabInventoryStatus.ps1
# Automates Grab Merchant Portal login and extracts inventory
# status data for ALL active stores (skips CLOSED stores).
# Captures: Store Name, Item Name, Category, Status, Date, Time
# OUTPUT: Writes to Google Sheets (appends from last empty row)
# Also saves local Excel backup.
# OPTIMIZED FOR AZURE VM: Maximized window, large viewport.
# HANG PROTECTION: Per-store timeout, total runtime cap.
# Retry logic for skipped stores.
# Right-click > "Run with PowerShell"
# ============================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Configuration ---
$GrabLoginUrl = "https://weblogin.grab.com/merchant/saved-accounts?service_id=MEXUSERS&redirect=https%3A%2F%2Fmerchant.grab.com%2Fportal"
$GrabUsername = "automationtester"
$GrabPassword = "Operations@123"
$DebugPort    = 9222
$outputFolder = $PSScriptRoot
if (-not $outputFolder -or $outputFolder -eq "") {
    $outputFolder = [Environment]::GetFolderPath("Desktop")
}
# TARGET STORE: $null = scan ALL active stores (skip CLOSED)
$targetStoreName = $null

# GOOGLE SHEETS Configuration - Uses Apps Script Web App (same method as other scripts)
# You need to deploy a Google Apps Script for this sheet and paste the URL below
$GoogleAppsScriptUrl = "https://script.google.com/macros/s/AKfycbwKUB5adJeF7TY0s8D49sZ7ukgbD2dut-AZEZM_kEm8Frmoy29Gi4AFQ5KhdUYZWMFCuw/exec"
$EnableGoogleSheets = $true

# HANG TIMEOUT: Max seconds to wait for any single store before skipping
$StoreTimeoutSeconds = 120
# MAX TOTAL RUNTIME: If script exceeds this, terminate and save what we have
$MaxTotalRuntimeMinutes = 30

# --- Find Edge ---
$edgePath = @(
    "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edgePath) { [System.Windows.Forms.MessageBox]::Show("Microsoft Edge not found.","Error"); exit }

# --- UI Form ---
$form = New-Object System.Windows.Forms.Form -Property @{
    Text="Grab Inventory Status - All Stores"; Size="700,520"
    StartPosition="CenterScreen"; FormBorderStyle="FixedDialog"; MaximizeBox=$false
}
$lblTitle = New-Object System.Windows.Forms.Label -Property @{
    Text="Grab - Inventory Status (All 13 Active Stores)"; AutoSize=$true; Location="20,15"
    Font=New-Object System.Drawing.Font("Segoe UI",13,[System.Drawing.FontStyle]::Bold)
}
$form.Controls.Add($lblTitle)
$lblInfo = New-Object System.Windows.Forms.Label -Property @{
    Text="Click Start to login and extract inventory status from all active stores. CLOSED stores will be skipped."
    Size="650,20"; Location="20,50"; Font=New-Object System.Drawing.Font("Segoe UI",9)
}
$form.Controls.Add($lblInfo)
$btnStart = New-Object System.Windows.Forms.Button -Property @{
    Text="Start"; Size="150,40"; Location="20,80"
    Font=New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
}
$form.Controls.Add($btnStart)
$lblStatus = New-Object System.Windows.Forms.Label -Property @{
    Text="Status: Ready"; Size="640,20"; Location="20,130"
    Font=New-Object System.Drawing.Font("Segoe UI",9)
}
$form.Controls.Add($lblStatus)
$progressBar = New-Object System.Windows.Forms.ProgressBar -Property @{
    Location="20,155"; Size="645,22"; Maximum=100
}
$form.Controls.Add($progressBar)
$txtLog = New-Object System.Windows.Forms.TextBox -Property @{
    Location="20,185"; Size="645,295"; Multiline=$true
    ScrollBars="Vertical"; ReadOnly=$true
    Font=New-Object System.Drawing.Font("Consolas",8.5)
}
$form.Controls.Add($txtLog)

# --- Helper Functions ---
# Log file saved in the same folder as the script
$Global:logFile = Join-Path $outputFolder "GrabInventoryStatus_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log { param([string]$msg)
    $line = "[$((Get-Date).ToString('HH:mm:ss'))] $msg"
    $txtLog.AppendText("$line`r`n")
    $form.Refresh()
    Add-Content -Path $Global:logFile -Value $line -ErrorAction SilentlyContinue
}
$Global:cdpId = 0; $Global:ws = $null; $Global:debugPort = $DebugPort
$Global:scriptStartTime = $null

function Check-Timeout {
    # Returns $true if total runtime exceeded
    if ($Global:scriptStartTime -and ([DateTime]::Now - $Global:scriptStartTime).TotalMinutes -gt $MaxTotalRuntimeMinutes) {
        return $true
    }
    return $false
}

function Reconnect-CDP {
    try {
        if ($Global:ws) { try { $Global:ws.Dispose() } catch {}; $Global:ws = $null }
        Start-Sleep -Seconds 2
        $targets = Invoke-RestMethod "http://localhost:$Global:debugPort/json" -UseBasicParsing
        $page = $targets | Where-Object { $_.type -eq "page" } | Select-Object -First 1
        if ($page) {
            $Global:ws = New-Object System.Net.WebSockets.ClientWebSocket
            $Global:ws.ConnectAsync([Uri]$page.webSocketDebuggerUrl,[System.Threading.CancellationToken]::None).Wait()
            Send-CDP "Page.enable" | Out-Null
            Send-CDP "Runtime.enable" | Out-Null
            return $true
        }
    } catch {}
    return $false
}

function Send-CDP { param([string]$method, [hashtable]$params = @{})
    $Global:cdpId++
    $msg = @{ id=$Global:cdpId; method=$method; params=$params } | ConvertTo-Json -Depth 10 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($msg)
    try {
        $Global:ws.SendAsync(
            (New-Object System.ArraySegment[byte](,$bytes)),
            [System.Net.WebSockets.WebSocketMessageType]::Text, $true,
            [System.Threading.CancellationToken]::None
        ).Wait()
    } catch {
        Write-Log "    [WebSocket reconnecting...]"
        if (Reconnect-CDP) {
            $Global:cdpId++
            $msg = @{ id=$Global:cdpId; method=$method; params=$params } | ConvertTo-Json -Depth 10 -Compress
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($msg)
            try {
                $Global:ws.SendAsync(
                    (New-Object System.ArraySegment[byte](,$bytes)),
                    [System.Net.WebSockets.WebSocketMessageType]::Text, $true,
                    [System.Threading.CancellationToken]::None
                ).Wait()
            } catch { return $null }
        } else { return $null }
    }
    $buf = New-Object byte[] 2097152
    $deadline = [DateTime]::Now.AddSeconds(30)
    while ([DateTime]::Now -lt $deadline) {
        $res=""; $done=$false
        while (-not $done) {
            $seg = New-Object System.ArraySegment[byte](,$buf)
            try {
                $task = $Global:ws.ReceiveAsync($seg,[System.Threading.CancellationToken]::None)
                $task.Wait(30000)
                if ($task.IsCompleted) {
                    $res += [System.Text.Encoding]::UTF8.GetString($buf,0,$task.Result.Count)
                    $done = $task.Result.EndOfMessage
                } else { return $null }
            } catch { return $null }
        }
        try { $j=$res|ConvertFrom-Json; if($j.id -eq $Global:cdpId){return $j} } catch {}
    }
    return $null
}

function Run-JS { param([string]$js)
    $r = Send-CDP "Runtime.evaluate" @{ expression=$js; returnByValue=$true; awaitPromise=$true }
    if ($r -and $r.result -and $r.result.result) { return $r.result.result.value }
    return $null
}
function Wait-JS { param([string]$condition, [int]$seconds=30)
    $end=[DateTime]::Now.AddSeconds($seconds)
    while([DateTime]::Now -lt $end){ $v=Run-JS $condition; if($v -eq $true -or $v -eq "true" -or $v -eq "True"){return $true}; Start-Sleep -Milliseconds 500 }
    return $false
}
function Go-To { param([string]$url)
    Send-CDP "Page.navigate" @{url=$url} | Out-Null; Start-Sleep -Seconds 2
    for($i=0;$i -lt 20;$i++){if((Run-JS "document.readyState") -eq "complete"){break};Start-Sleep -Milliseconds 500}
    Start-Sleep -Seconds 1
}

# --- Force TLS 1.2 (required for Google APIs) ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Google Sheets Upload Function (via Apps Script Web App) ---
# Exact same pattern as the working CFC Deactivated Items script
function Send-ToGoogleSheets {
    param(
        [string]$WebAppUrl,
        [System.Collections.Generic.List[hashtable]]$ItemData
    )
    
    if (-not $WebAppUrl -or $WebAppUrl -eq "") {
        Write-Log "ERROR: Google Apps Script Web App URL not configured."
        return $false
    }
    
    # Build JSON payload
    $payload = @()
    foreach ($item in $ItemData) {
        $payload += @{
            StoreName = [string]$item.StoreName
            ItemName  = [string]$item.ItemName
            Category  = [string]$item.Category
            Status    = [string]$item.Status
            DateCheck = [string]$item.DateCheck
            TimeCheck = [string]$item.TimeCheck
        }
    }
    
    $jsonBody = $payload | ConvertTo-Json -Depth 5 -Compress
    Write-Log "Payload size: $($jsonBody.Length) bytes, Items: $($payload.Count)"
    
    # Force TLS 1.2
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    
    try {
        Write-Log "Uploading to Google Sheets..."
        
        $response = Invoke-RestMethod -Uri $WebAppUrl -Method POST `
            -Body $jsonBody -ContentType "application/json" `
            -UseBasicParsing -ErrorAction Stop
        
        Write-Log "Response: $($response | Out-String)"
        if ($response -match "success" -or ($response.status -eq "success")) {
            Write-Log "Google Sheets updated successfully! Rows: $($response.rows)"
            return $true
        }
        
        Write-Log "Unexpected response from Google Sheets."
        return $false
    } catch {
        Write-Log "ERROR uploading to Google Sheets: $($_.Exception.Message)"
        return $false
    }
}

# --- Main Logic ---
$btnStart.Add_Click({
    $btnStart.Enabled = $false; $txtLog.Clear(); $progressBar.Value = 0
    $edgeProc = $null; $excel = $null
    $Global:scriptStartTime = [DateTime]::Now

    try {
        # === STEP 1: Launch Edge ===
        Write-Log "Launching Microsoft Edge..."
        $lblStatus.Text = "Status: Starting Edge..."; $form.Refresh()
        $tempProfile = Join-Path $env:TEMP "GrabInventory_Edge"
        # Start every run with a FRESH, EMPTY profile.
        # A persistent profile keeps cookies from previous logins. When the saved session
        # token expires overnight, those stale cookies force Grab's "Welcome back"
        # saved-accounts page, which cannot be bypassed (both "Login as another user" and
        # the direct login URL redirect back to saved-accounts). This is what caused the
        # first-run-of-the-day LOGIN FAILED. A clean profile has no saved account, so Grab
        # shows the normal username + password login form directly.
        # First, make sure no leftover Edge is locking the profile folder.
        Get-Process -Name "msedge" -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Kill() } catch {} }
        Start-Sleep -Seconds 2
        if (Test-Path $tempProfile) {
            try {
                Remove-Item -Path $tempProfile -Recurse -Force -ErrorAction Stop
                Write-Log "Cleared previous Edge profile (fresh login)."
            } catch {
                Write-Log "WARNING: Could not fully clear profile: $($_.Exception.Message)"
            }
        }
        New-Item -ItemType Directory -Path $tempProfile -Force | Out-Null
        $edgeProc = Start-Process -FilePath $edgePath -ArgumentList "--headless=new --remote-debugging-port=$DebugPort --user-data-dir=`"$tempProfile`" --no-first-run --disable-extensions --window-size=1920,1080 --disable-gpu --disable-session-crashed-bubble --hide-crash-restore-bubble --disable-infobars `"$GrabLoginUrl`"" -PassThru
        Start-Sleep -Seconds 7
        Write-Log "Edge started (PID: $($edgeProc.Id))."
        $progressBar.Value = 5

        # === STEP 2: Connect via CDP ===
        Write-Log "Connecting to Edge..."
        $targets = $null
        for ($i=0; $i -lt 15; $i++) {
            try { $targets = Invoke-RestMethod "http://localhost:$DebugPort/json" -UseBasicParsing; if($targets){break} } catch {}
            Start-Sleep -Seconds 2
        }
        if (-not $targets) { throw "Cannot connect to Edge. Close ALL Edge windows and retry." }
        $page = $targets | Where-Object { $_.type -eq "page" } | Select-Object -First 1
        if (-not $page) { throw "No browser page found." }
        $Global:ws = New-Object System.Net.WebSockets.ClientWebSocket
        $Global:ws.ConnectAsync([Uri]$page.webSocketDebuggerUrl,[System.Threading.CancellationToken]::None).Wait()
        Send-CDP "Page.enable" | Out-Null; Send-CDP "Runtime.enable" | Out-Null
        Write-Log "Connected to browser."
        # Maximize browser window for Azure VM (prevent data loss from small viewport)
        Send-CDP "Browser.setWindowBounds" @{ windowId=1; bounds=@{ windowState="maximized" } } | Out-Null
        # Also set viewport size explicitly
        Send-CDP "Emulation.setDeviceMetricsOverride" @{ width=1920; height=1080; deviceScaleFactor=1; mobile=$false } | Out-Null
        Write-Log "Browser window maximized (1920x1080)."
        $progressBar.Value = 8

        # === STEP 3: Handle Login ===
        Write-Log "Waiting for login page..."
        $lblStatus.Text = "Status: Logging in..."; $form.Refresh()
        Start-Sleep -Seconds 8

        # Dismiss "Restore pages" popup if Edge crashed previously
        Run-JS @'
(function(){
    var els = document.querySelectorAll('button, span, div, a');
    for (var i=0; i<els.length; i++) {
        var el = els[i];
        if (el.offsetParent === null) continue;
        var t = (el.innerText||'').trim().toLowerCase();
        var aria = (el.getAttribute('aria-label')||'').toLowerCase();
        if (t === '\u00d7' || t === 'x' || t === '\u2715' || aria === 'close' || aria === 'dismiss' || t === 'close') {
            el.click();
            return 'dismissed';
        }
    }
    var body = document.querySelector('body');
    if (body) body.click();
    return 'clicked_body';
})();
'@
        Start-Sleep -Seconds 2

        $url = Run-JS "window.location.href"
        Write-Log "URL: $url"

        # Handle saved-accounts page
        if ($url -match "saved-accounts") {
            Write-Log "Saved accounts page - handling login..."
            
            # Dismiss any "Restore pages" popup first
            Start-Sleep -Seconds 2
            Run-JS @'
(function(){
    var els = document.querySelectorAll('button, span, div, a');
    for (var i=0; i<els.length; i++) {
        var el = els[i];
        if (el.offsetParent === null) continue;
        var t = (el.innerText||'').trim().toLowerCase();
        var aria = (el.getAttribute('aria-label')||'').toLowerCase();
        if (t === '\u00d7' || t === 'x' || t === '\u2715' || aria === 'close' || aria === 'dismiss') {
            el.click();
            return 'dismissed';
        }
    }
    return 'no_popup';
})();
'@
            Start-Sleep -Seconds 2
            
            # NOTE: We intentionally SKIP the saved-account "Continue" shortcut.
            # The account stays listed on the "Welcome back" page even after the underlying
            # login session has expired. Clicking "Continue" on an expired saved session
            # silently bounces back to saved-accounts, which caused intermittent LOGIN FAILED
            # errors (typically the first run of the day). Always force a clean username +
            # password login instead, which is reliable regardless of session age.
            #
            # Strategy 1: Click "Login as another user" to start a fresh login
            Write-Log "  Forcing fresh login (clicking 'Login as another user')..."
            Run-JS "window.scrollTo(0, document.body.scrollHeight);"
            Start-Sleep -Seconds 1
            $freshClick = Run-JS @'
(function(){
    var els = document.querySelectorAll('button, a, div, span, [role="button"]');
    for (var i=0; i<els.length; i++) {
        var el = els[i];
        if (el.offsetParent === null) continue;
        var t = (el.innerText||'').trim();
        if (t === 'Login as another user' || t.toLowerCase().indexOf('another user') > -1 || t.toLowerCase().indexOf('new user') > -1) {
            el.scrollIntoView({block:'center'});
            el.focus();
            var events = ['mousedown','mouseup','click'];
            for (var e=0; e<events.length; e++) {
                el.dispatchEvent(new MouseEvent(events[e], {bubbles:true, cancelable:true, view:window}));
            }
            return 'clicked_another';
        }
    }
    return 'not_found';
})();
'@
            Write-Log "  Fresh login click: $freshClick"
            Start-Sleep -Seconds 6

            # Strategy 2 (fallback): if still on saved-accounts, navigate directly to the login URL
            $url = Run-JS "window.location.href"
            if ($url -match "saved-accounts") {
                Write-Log "  Still on saved-accounts - navigating directly to login URL..."
                Go-To "https://weblogin.grab.com/merchant/login?service_id=MEXUSERS&redirect=https%3A%2F%2Fmerchant.grab.com%2Fportal"
                Start-Sleep -Seconds 5
            }

            # Wait for the username field (fresh login shows username step FIRST), or password, or dashboard
            Wait-JS "document.querySelector('#Username')!==null||document.querySelector('input[type=text]')!==null||document.querySelector('input[type=email]')!==null||document.querySelector('input[type=password]')!==null||(window.location.href.indexOf('merchant.grab.com')>-1&&window.location.href.indexOf('weblogin')===-1)" 25
        }

        # Check if already logged in
        $url = Run-JS "window.location.href"
        if ($url -match "merchant\.grab\.com" -and $url -notmatch "weblogin") {
            Write-Log "Already logged in!"
        } else {
            # Enter username. Wait for the username field to appear first (the fresh-login
            # flow shows the username step before the password step). Without this wait the
            # script could fall through to the password step too early and never log in.
            $hasUsernameField = Wait-JS "document.querySelector('#Username')!==null||document.querySelector('input[type=text]')!==null||document.querySelector('input[type=email]')!==null" 15
            if ($hasUsernameField -eq $true) {
                Write-Log "Entering username..."
                Start-Sleep -Seconds 2
                Run-JS @"
(function(){ var inp=document.querySelector('#Username')||document.querySelector('input[type=text]')||document.querySelector('input[type=email]'); if(!inp)return; inp.focus(); inp.value=''; var s=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set; s.call(inp,'$GrabUsername'); inp.dispatchEvent(new Event('input',{bubbles:true})); inp.dispatchEvent(new Event('change',{bubbles:true})); })();
"@
                Start-Sleep -Seconds 2
                Run-JS @'
(function(){ var b=document.querySelectorAll('button'); for(var i=0;i<b.length;i++){var t=(b[i].innerText||'').trim().toLowerCase(); if(t==='continue'||t==='next'){b[i].click();return;}} var s=document.querySelector('button[type=submit]'); if(s)s.click(); })();
'@
                Start-Sleep -Seconds 5
            }

            # Enter password
            $url = Run-JS "window.location.href"
            if (-not ($url -match "merchant\.grab\.com" -and $url -notmatch "weblogin")) {
                Write-Log "Waiting for password field..."
                Start-Sleep -Seconds 3
                $pwReady = Wait-JS "document.querySelector('input[type=password]')!==null" 25
                if ($pwReady) {
                    Write-Log "Entering password..."
                    Start-Sleep -Seconds 2
                    Run-JS @"
(function(){ var pw=document.querySelector('input[type=password]'); if(!pw)return; pw.focus(); pw.value=''; var s=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set; s.call(pw,'$GrabPassword'); pw.dispatchEvent(new Event('input',{bubbles:true})); pw.dispatchEvent(new Event('change',{bubbles:true})); })();
"@
                    Start-Sleep -Seconds 3
                    Write-Log "Clicking login button..."
                    Run-JS @'
(function(){ var b=document.querySelectorAll('button'); for(var i=0;i<b.length;i++){var t=(b[i].innerText||'').trim().toLowerCase(); if(t==='continue'||t==='log in'||t==='login'||t==='sign in'){b[i].click();return;}} var s=document.querySelector('button[type=submit]'); if(s)s.click(); })();
'@
                    Start-Sleep -Seconds 10
                } else {
                    Write-Log "Password field not found. Current URL: $(Run-JS 'window.location.href')"
                }
            }
        }
        $progressBar.Value = 13

        # === STEP 4: Verify Login ===
        Write-Log "Verifying login..."
        $loginSuccess = $false
        for ($i=0; $i -lt 20; $i++) {
            $url = Run-JS "window.location.href"
            if ($url -match "merchant\.grab\.com" -and $url -notmatch "weblogin") { $loginSuccess=$true; break }
            Start-Sleep -Seconds 2
        }
        if (-not $loginSuccess) { throw "Login failed. URL: $(Run-JS 'window.location.href')" }
        Write-Log "LOGIN SUCCESSFUL!"
        Start-Sleep -Seconds 3
        Run-JS @'
(function(){ var b=document.querySelectorAll('button,a,span'); for(var i=0;i<b.length;i++){var t=(b[i].innerText||'').trim().toLowerCase(); if(t==='close'||t==='skip'||t==='dismiss'||t==='got it'){if(b[i].offsetParent!==null){b[i].click();return;}}} })();
'@
        Start-Sleep -Seconds 2
        $progressBar.Value = 15

        # === STEP 5: Navigate to Inventory ===
        Write-Log "Navigating to Inventory page..."
        $lblStatus.Text = "Status: Going to Inventory..."; $form.Refresh()
        Go-To "https://merchant.grab.com/food/inventory"
        Start-Sleep -Seconds 3
        $invLoaded = Wait-JS "(document.body.innerText||'').indexOf('STORE NAME')>-1||(document.body.innerText||'').indexOf('ITEM NAME')>-1" 15
        if (-not $invLoaded) {
            # Try clicking sidebar
            Run-JS @'
(function(){ var els=document.querySelectorAll('a,div,span,li'); for(var i=0;i<els.length;i++){var t=(els[i].innerText||'').trim(); if(t==='Inventory'&&els[i].offsetParent!==null){els[i].click();return 'clicked';}} return 'nf'; })();
'@
            Start-Sleep -Seconds 5
        }
        Write-Log "Inventory page loaded."
        $progressBar.Value = 20

        # === STEP 6: Get ALL store names (robust scroll) ===
        Write-Log "Reading store list (aggressive scroll to load ALL stores)..."
        # Aggressive scroll: scroll multiple containers repeatedly
        $prevStoreCount = 0
        for ($scrollPass = 0; $scrollPass -lt 30; $scrollPass++) {
            Run-JS @'
(function(){
    // Scroll the window to bottom
    window.scrollTo(0, document.body.scrollHeight);
    // Also scroll any inner scrollable container
    var divs = document.querySelectorAll('div');
    for (var i=0; i<divs.length; i++) {
        var el = divs[i];
        if (el.scrollHeight > el.clientHeight + 20 && el.clientHeight > 80 && el.offsetParent !== null) {
            el.scrollTop = el.scrollHeight;
        }
    }
})();
'@
            Start-Sleep -Milliseconds 600
            $curStoreCount = Run-JS @'
(function(){
    var count = 0;
    var rows = document.querySelectorAll('tr');
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td');
        if (cells.length > 0 && cells[0].innerText.trim().length > 3) count++;
    }
    return count;
})();
'@
            if ($curStoreCount -ge 16) { break }  # We know there are 16 total stores
            if ($curStoreCount -eq $prevStoreCount -and $scrollPass -gt 5) { break }
            $prevStoreCount = $curStoreCount
        }
        Run-JS "window.scrollTo(0,0);"; Start-Sleep -Seconds 1

        # Collect all store names
        $storeNamesJson = Run-JS @'
(function(){
    var stores = [];
    var rows = document.querySelectorAll('tr');
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td');
        if (cells.length > 0) {
            var name = cells[0].innerText.trim();
            if (name.length > 3 && name !== 'STORE NAME' && stores.indexOf(name) === -1) {
                stores.push(name);
            }
        }
    }
    if (stores.length === 0) {
        var els = document.querySelectorAll('a, div, span, button');
        for (var i=0; i<els.length; i++) {
            var el = els[i];
            if (el.children.length > 2) continue;
            var t = (el.innerText||'').trim();
            if (t.length > 5 && t.length < 80 && el.offsetParent !== null &&
                (t.indexOf('FNP') > -1 || t.indexOf('Cakes') > -1 || t.indexOf('Flowers') > -1) &&
                stores.indexOf(t) === -1) {
                stores.push(t);
            }
        }
    }
    return JSON.stringify(stores);
})();
'@
        $storeNames = @()
        if ($storeNamesJson) { try { $storeNames = $storeNamesJson | ConvertFrom-Json } catch {} }

        # Filter: skip CLOSED stores
        $storesToProcess = @()
        $skippedStores = @()
        foreach ($s in $storeNames) {
            if ($s -match "CLOSED" -or $s -match "\[CLOSED\]") {
                $skippedStores += $s
            } else {
                if (-not $targetStoreName -or $s -eq $targetStoreName) {
                    $storesToProcess += $s
                }
            }
        }

        Write-Log "Found $($storeNames.Count) total stores."
        Write-Log "Active stores to process: $($storesToProcess.Count)"
        Write-Log "CLOSED stores skipped: $($skippedStores.Count)"
        foreach ($s in $skippedStores) { Write-Log "  [SKIPPED] $s" }
        Write-Log ""
        foreach ($s in $storesToProcess) { Write-Log "  [ACTIVE] $s" }
        if ($storesToProcess.Count -eq 0) { throw "No active stores found to process." }
        $progressBar.Value = 22

        # === STEP 7: Loop through each ACTIVE store ===
        $allItems = [System.Collections.Generic.List[hashtable]]::new()
        $skippedActiveStores = [System.Collections.Generic.List[string]]::new()
        $storeIdx = 0

        foreach ($storeName in $storesToProcess) {
            $storeIdx++
            $storeStartTime = [DateTime]::Now

            # Check total runtime timeout
            if (Check-Timeout) {
                Write-Log ""
                Write-Log "*** MAX RUNTIME ($MaxTotalRuntimeMinutes min) EXCEEDED. Stopping and saving data collected so far. ***"
                break
            }

            Write-Log ""
            Write-Log "======================================================"
            Write-Log "STORE $storeIdx/$($storesToProcess.Count): $storeName"
            Write-Log "======================================================"
            $lblStatus.Text = "Status: Store $storeIdx/$($storesToProcess.Count) - $storeName"
            $form.Refresh()
            $progressBar.Value = 22 + [int](($storeIdx / $storesToProcess.Count) * 65)

            try {
                # --- Find and click this store (with aggressive scroll) ---
                $clickResult = 'not_found'
                Run-JS "window.scrollTo(0,0);"
                Start-Sleep -Milliseconds 500

                for ($attempt = 0; $attempt -lt 25; $attempt++) {
                    $clickResult = Run-JS @"
(function(){
    var target = decodeURIComponent('$([System.Uri]::EscapeDataString($storeName))');
    var rows = document.querySelectorAll('tr');
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td');
        if (cells.length > 0 && cells[0].innerText.trim() === target) {
            rows[i].scrollIntoView({block:'center'});
            rows[i].click();
            return 'clicked_tr';
        }
    }
    var els = document.querySelectorAll('a, div, span, button');
    for (var i=0; i<els.length; i++) {
        var el = els[i];
        var t = (el.innerText||'').trim();
        if (t === target && el.offsetParent !== null && el.children.length <= 3) {
            el.scrollIntoView({block:'center'});
            el.click();
            return 'clicked_el';
        }
    }
    return 'not_found';
})();
"@
                    if ($clickResult -ne 'not_found') { break }
                    # Scroll down incrementally to find the store
                    Run-JS "window.scrollBy(0, 250);"
                    Start-Sleep -Milliseconds 400
                }

                if ($clickResult -eq 'not_found') {
                    Write-Log "  WARNING: Could not find store '$storeName' in list. SKIPPING. (Will retry later)"
                    $skippedActiveStores.Add($storeName)
                    continue
                }
                Write-Log "  Clicked store: $clickResult"
                Start-Sleep -Seconds 4

                # Dismiss popup
                Run-JS @'
(function(){ var b=document.querySelectorAll('button,a,span'); for(var i=0;i<b.length;i++){var t=(b[i].innerText||'').trim().toLowerCase(); if(t==='close'||t==='skip'||t==='dismiss'||t==='got it'){if(b[i].offsetParent!==null){b[i].click();return;}}} })();
'@
                Start-Sleep -Seconds 2

                # Wait for inventory table to load
                $tableLoaded = Wait-JS "(document.body.innerText||'').indexOf('ITEM NAME')>-1||(document.body.innerText||'').indexOf('CATEGORY')>-1||(document.body.innerText||'').indexOf('STATUS')>-1" 20
                if (-not $tableLoaded) {
                    Write-Log "  WARNING: Inventory table did not load for '$storeName'. SKIPPING. (Will retry later)"
                    $skippedActiveStores.Add($storeName)
                    # Navigate back
                    Go-To "https://merchant.grab.com/food/inventory"
                    Start-Sleep -Seconds 3
                    continue
                }
                Write-Log "  Store inventory page loaded."
                Start-Sleep -Seconds 2

                # --- Detect table column mapping ---
                Run-JS "window.scrollTo(0, 0);"
                Start-Sleep -Seconds 1

                $colMapJson = Run-JS @'
(function(){
    var headers = [];
    var thEls = document.querySelectorAll('th');
    if (thEls.length >= 3) {
        for (var i=0; i<thEls.length; i++) {
            headers.push(thEls[i].innerText.trim().toUpperCase());
        }
    }
    if (headers.length === 0) {
        var rows = document.querySelectorAll('tr');
        for (var i=0; i<rows.length; i++) {
            var txt = rows[i].innerText.toUpperCase();
            if (txt.indexOf('ITEM NAME') > -1 && txt.indexOf('STATUS') > -1) {
                var cells = rows[i].querySelectorAll('th, td');
                for (var c=0; c<cells.length; c++) {
                    headers.push(cells[c].innerText.trim().toUpperCase());
                }
                break;
            }
        }
    }
    var colMap = {nameIdx: -1, priceIdx: -1, categoryIdx: -1, sellingTimeIdx: -1, stockLevelIdx: -1, statusIdx: -1};
    for (var h=0; h<headers.length; h++) {
        var hdr = headers[h];
        if (hdr.indexOf('ITEM NAME') > -1 || hdr === 'ITEM' || hdr === 'NAME') colMap.nameIdx = h;
        else if (hdr.indexOf('PRICE') > -1) colMap.priceIdx = h;
        else if (hdr.indexOf('CATEGORY') > -1) colMap.categoryIdx = h;
        else if (hdr.indexOf('SELLING') > -1) colMap.sellingTimeIdx = h;
        else if (hdr.indexOf('STOCK') > -1) colMap.stockLevelIdx = h;
        else if (hdr.indexOf('STATUS') > -1) colMap.statusIdx = h;
    }
    colMap.headerCount = headers.length;
    return JSON.stringify(colMap);
})();
'@
                $colMap = $null
                if ($colMapJson) { try { $colMap = $colMapJson | ConvertFrom-Json } catch {} }
                $nameIdx = if ($colMap -and $colMap.nameIdx -ge 0) { $colMap.nameIdx } else { 0 }
                $catIdx2 = if ($colMap -and $colMap.categoryIdx -ge 0) { $colMap.categoryIdx } else { 2 }
                $statusIdx = if ($colMap -and $colMap.statusIdx -ge 0) { $colMap.statusIdx } else { 5 }
                Write-Log "  Columns: Name=$nameIdx, Category=$catIdx2, Status=$statusIdx"

                # --- AGGRESSIVE SCROLL & CAPTURE all items ---
                $itemsMap = @{}
                $prevItemCount = 0
                $noChangeCount = 0
                $scrollAttempts = 0
                $maxScrollAttempts = 100  # Very generous scroll limit

                # Reset all scrollable containers to top
                Run-JS @'
(function(){
    window.scrollTo(0, 0);
    var divs = document.querySelectorAll('div');
    for (var i=0; i<divs.length; i++) {
        var el = divs[i];
        if (el.scrollHeight > el.clientHeight + 50 && el.clientHeight > 150 && el.offsetParent !== null) {
            el.scrollTop = 0;
        }
    }
})();
'@
                Start-Sleep -Seconds 1

                for ($scrollStep = 0; $scrollStep -lt $maxScrollAttempts; $scrollStep++) {
                    # Check per-store timeout
                    if (([DateTime]::Now - $storeStartTime).TotalSeconds -gt $StoreTimeoutSeconds) {
                        Write-Log "  TIMEOUT: Store took > $StoreTimeoutSeconds seconds. Moving to next store."
                        break
                    }

                    # Collect all visible items
                    $visJson = Run-JS @"
(function(){
    var items = [];
    var nameIdx = $nameIdx;
    var catIdx = $catIdx2;
    var statusIdx = $statusIdx;
    
    // Store name patterns to filter out (store list rows appearing in DOM)
    var storePatterns = ['FNP Cakes', 'FNP Flowers', 'FNP Gifts', 'FNP Cakes &', '[CLOSED]'];
    var locationNames = ['singapore', 'malaysia', 'indonesia', 'thailand', 'vietnam', 'philippines'];
    
    function isStoreName(name) {
        var n = name.trim();
        for (var p = 0; p < storePatterns.length; p++) {
            if (n.indexOf(storePatterns[p]) === 0 || n.indexOf('[CLOSED]') > -1) return true;
        }
        // Also check pattern: "FNP ... - Location"
        if (/^FNP\s/.test(n) && n.indexOf(' - ') > -1) return true;
        return false;
    }
    
    function isLocationCategory(cat) {
        var c = cat.trim().toLowerCase();
        for (var l = 0; l < locationNames.length; l++) {
            if (c === locationNames[l]) return true;
        }
        return false;
    }
    
    var rows = document.querySelectorAll('tr');
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td');
        if (cells.length < 3) continue;
        
        // Get item name
        var nameCell = cells[nameIdx] ? cells[nameIdx].innerText.trim() : '';
        if (cells[nameIdx]) {
            var link = cells[nameIdx].querySelector('a');
            if (link) nameCell = link.innerText.trim();
        }
        if (!nameCell || nameCell.length < 2 || nameCell.length > 200) continue;
        if (nameCell.toUpperCase() === 'ITEM NAME' || nameCell.toUpperCase() === 'ITEM' || nameCell.toUpperCase() === 'NAME') continue;
        
        // FILTER: Skip store names that appear as items
        if (isStoreName(nameCell)) continue;
        
        // Get category
        var category = '';
        if (catIdx < cells.length && cells[catIdx]) {
            category = cells[catIdx].innerText.trim().split('\\n')[0].trim();
        }
        
        // FILTER: Skip rows where category is a location/country (store list rows)
        if (isLocationCategory(category)) continue;
        
        // Get status (custom dropdown - read visible text)
        var status = '';
        if (statusIdx < cells.length && cells[statusIdx]) {
            var statusCell = cells[statusIdx];
            var selectEl = statusCell.querySelector('select');
            if (selectEl && selectEl.selectedIndex >= 0) {
                status = selectEl.options[selectEl.selectedIndex].text.trim();
            }
            if (!status) {
                var cellText = statusCell.innerText.trim();
                cellText = cellText.replace(/[\u25BC\u25B6\u25B2\u25C0\u2304\u2303\uFE0F\u200B]/g, '').trim();
                var firstLine = cellText.split('\\n')[0].trim();
                if (firstLine && firstLine.length > 2 && firstLine.length < 50) {
                    status = firstLine;
                }
            }
        }
        
        if (nameCell) items.push({name: nameCell, category: category || 'Unknown', status: status || 'Unknown'});
    }
    return JSON.stringify(items);
})();
"@
                    if ($visJson) {
                        try {
                            $vis = $visJson | ConvertFrom-Json
                            foreach ($v in $vis) {
                                if (-not $itemsMap.ContainsKey($v.name)) {
                                    $itemsMap[$v.name] = $v
                                }
                            }
                        } catch {}
                    }

                    # Check progress
                    if ($itemsMap.Count -eq $prevItemCount) {
                        $noChangeCount++
                        # After 6 scroll attempts with no new items, we're done
                        if ($noChangeCount -ge 6) { break }
                    } else {
                        $noChangeCount = 0
                        $prevItemCount = $itemsMap.Count
                    }

                    # AGGRESSIVE SCROLL: try both inner containers and window
                    Run-JS @'
(function(){
    var scrolled = false;
    // Strategy 1: Scroll inner container holding the table
    var divs = document.querySelectorAll('div');
    for (var i=0; i<divs.length; i++) {
        var el = divs[i];
        if (el.scrollHeight > el.clientHeight + 50 && el.clientHeight > 150 && el.offsetParent !== null) {
            var text = el.innerText || '';
            if ((text.indexOf('ITEM NAME') > -1 || text.indexOf('CATEGORY') > -1 || text.indexOf('S$') > -1) &&
                text.indexOf('STORE NAME') === -1) {
                var before = el.scrollTop;
                el.scrollBy(0, 350);
                if (el.scrollTop !== before) { scrolled = true; }
                break;
            }
        }
    }
    // Strategy 2: Also scroll the window
    var wb = window.pageYOffset;
    window.scrollBy(0, 350);
    if (window.pageYOffset !== wb) scrolled = true;
    return scrolled ? 'MORE' : 'DONE';
})();
'@
                    Start-Sleep -Milliseconds 700
                }

                # Final collection pass
                Start-Sleep -Milliseconds 500
                $visJson = Run-JS @"
(function(){
    var items = [];
    var nameIdx = $nameIdx; var catIdx = $catIdx2; var statusIdx = $statusIdx;
    
    // Store name patterns to filter out
    var storePatterns = ['FNP Cakes', 'FNP Flowers', 'FNP Gifts', 'FNP Cakes &', '[CLOSED]'];
    var locationNames = ['singapore', 'malaysia', 'indonesia', 'thailand', 'vietnam', 'philippines'];
    function isStoreName(name) {
        var n = name.trim();
        for (var p = 0; p < storePatterns.length; p++) { if (n.indexOf(storePatterns[p]) === 0 || n.indexOf('[CLOSED]') > -1) return true; }
        if (/^FNP\s/.test(n) && n.indexOf(' - ') > -1) return true;
        return false;
    }
    function isLocationCategory(cat) {
        var c = cat.trim().toLowerCase();
        for (var l = 0; l < locationNames.length; l++) { if (c === locationNames[l]) return true; }
        return false;
    }
    
    var rows = document.querySelectorAll('tr');
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td');
        if (cells.length < 3) continue;
        var nameCell = cells[nameIdx] ? cells[nameIdx].innerText.trim() : '';
        if (cells[nameIdx]) { var link = cells[nameIdx].querySelector('a'); if (link) nameCell = link.innerText.trim(); }
        if (!nameCell || nameCell.length < 2 || nameCell.length > 200) continue;
        if (nameCell.toUpperCase() === 'ITEM NAME' || nameCell.toUpperCase() === 'ITEM') continue;
        if (isStoreName(nameCell)) continue;
        var category = (catIdx < cells.length && cells[catIdx]) ? cells[catIdx].innerText.trim().split('\\n')[0].trim() : 'Unknown';
        if (isLocationCategory(category)) continue;
        var status = '';
        if (statusIdx < cells.length && cells[statusIdx]) {
            var sc = cells[statusIdx]; var sel = sc.querySelector('select');
            if (sel && sel.selectedIndex >= 0) { status = sel.options[sel.selectedIndex].text.trim(); }
            if (!status) { var ct = sc.innerText.trim().replace(/[\u25BC\u25B6\u25B2\u25C0\u2304\u2303\uFE0F\u200B]/g,'').trim(); status = ct.split('\\n')[0].trim(); }
        }
        if (nameCell) items.push({name: nameCell, category: category || 'Unknown', status: status || 'Unknown'});
    }
    return JSON.stringify(items);
})();
"@
                if ($visJson) {
                    try {
                        $vis = $visJson | ConvertFrom-Json
                        foreach ($v in $vis) { if (-not $itemsMap.ContainsKey($v.name)) { $itemsMap[$v.name] = $v } }
                    } catch {}
                }

                # Add captured items to master list
                $sgtZone = [TimeZoneInfo]::FindSystemTimeZoneById("Singapore Standard Time")
                $sgtNow = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $sgtZone)
                $checkDate = $sgtNow.ToString("MM/dd/yyyy")
                $checkTime = $sgtNow.ToString("HH:mm:ss")
                $capturedItems = @($itemsMap.Values)

                if ($capturedItems.Count -eq 0) {
                    Write-Log "  No items found in this store."
                } else {
                    Write-Log "  Captured $($capturedItems.Count) items from '$storeName'."
                    foreach ($item in $capturedItems) {
                        $allItems.Add(@{
                            StoreName  = $storeName
                            ItemName   = [string]$item.name
                            Category   = [string]$item.category
                            Status     = [string]$item.status
                            DateCheck  = $checkDate
                            TimeCheck  = $checkTime
                        })
                    }
                }

                Write-Log "  Store complete. Running total: $($allItems.Count) items."

            } catch {
                Write-Log "  ERROR processing store '$storeName': $($_.Exception.Message)"
                Write-Log "  Attempting to recover and continue... (Will retry later)"
                $skippedActiveStores.Add($storeName)
                # Try to reconnect CDP if connection was lost
                Reconnect-CDP | Out-Null
            }

            # Navigate back to Inventory page for the next store
            if ($storeIdx -lt $storesToProcess.Count) {
                Write-Log "  Navigating back to store list..."
                Go-To "https://merchant.grab.com/food/inventory"
                Start-Sleep -Seconds 3

                # Wait for store list to reload
                $listReady = Wait-JS "(document.body.innerText||'').indexOf('STORE NAME')>-1" 15
                if (-not $listReady) {
                    Write-Log "  Store list did not reload. Trying again..."
                    Go-To "https://merchant.grab.com/food/inventory"
                    Start-Sleep -Seconds 5
                }

                # Aggressive scroll to ensure all stores are loaded
                for ($sp = 0; $sp -lt 15; $sp++) {
                    Run-JS @'
(function(){
    window.scrollTo(0, document.body.scrollHeight);
    var divs = document.querySelectorAll('div');
    for (var i=0; i<divs.length; i++) {
        var el = divs[i];
        if (el.scrollHeight > el.clientHeight + 20 && el.clientHeight > 80 && el.offsetParent !== null) {
            el.scrollTop = el.scrollHeight;
        }
    }
})();
'@
                    Start-Sleep -Milliseconds 400
                }
                Run-JS "window.scrollTo(0,0);"
                Start-Sleep -Seconds 1
            }
        } # end foreach store

        # === STEP 7B: RETRY any skipped stores (up to 2 more attempts) ===
        $maxRetries = 2
        $retryRound = 0
        while ($skippedActiveStores.Count -gt 0 -and $retryRound -lt $maxRetries) {
            $retryRound++
            if (Check-Timeout) {
                Write-Log ""
                Write-Log "*** MAX RUNTIME exceeded. Cannot retry skipped stores. ***"
                break
            }

            Write-Log ""
            Write-Log "########################################################"
            Write-Log "RETRY ROUND $retryRound/$maxRetries - $($skippedActiveStores.Count) store(s) to retry"
            Write-Log "########################################################"

            $retryList = @($skippedActiveStores)
            $skippedActiveStores.Clear()

            # Navigate to Inventory page
            Go-To "https://merchant.grab.com/food/inventory"
            Start-Sleep -Seconds 4
            Wait-JS "(document.body.innerText||'').indexOf('STORE NAME')>-1" 15

            # Scroll to load all stores
            for ($sp = 0; $sp -lt 20; $sp++) {
                Run-JS @'
(function(){
    window.scrollTo(0, document.body.scrollHeight);
    var divs = document.querySelectorAll('div');
    for (var i=0; i<divs.length; i++) {
        var el = divs[i];
        if (el.scrollHeight > el.clientHeight + 20 && el.clientHeight > 80 && el.offsetParent !== null) {
            el.scrollTop = el.scrollHeight;
        }
    }
})();
'@
                Start-Sleep -Milliseconds 500
            }
            Run-JS "window.scrollTo(0,0);"
            Start-Sleep -Seconds 1

            $retryIdx = 0
            foreach ($storeName in $retryList) {
                $retryIdx++
                $storeStartTime = [DateTime]::Now

                if (Check-Timeout) {
                    Write-Log "  MAX RUNTIME exceeded during retry. Stopping."
                    $skippedActiveStores.Add($storeName)
                    break
                }

                # Skip if we already have items for this store (from a timeout that captured partial data)
                $existingCount = ($allItems | Where-Object { $_.StoreName -eq $storeName }).Count
                if ($existingCount -gt 0) {
                    Write-Log "  Store '$storeName' already has $existingCount items. Skipping retry."
                    continue
                }

                Write-Log ""
                Write-Log "  RETRY $retryIdx/$($retryList.Count): $storeName"
                $lblStatus.Text = "Status: RETRY $retryRound - $storeName"
                $form.Refresh()

                try {
                    # Find and click the store
                    $clickResult = 'not_found'
                    Run-JS "window.scrollTo(0,0);"
                    Start-Sleep -Milliseconds 500

                    for ($attempt = 0; $attempt -lt 25; $attempt++) {
                        $clickResult = Run-JS @"
(function(){
    var target = decodeURIComponent('$([System.Uri]::EscapeDataString($storeName))');
    var rows = document.querySelectorAll('tr');
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td');
        if (cells.length > 0 && cells[0].innerText.trim() === target) {
            rows[i].scrollIntoView({block:'center'});
            rows[i].click();
            return 'clicked_tr';
        }
    }
    var els = document.querySelectorAll('a, div, span, button');
    for (var i=0; i<els.length; i++) {
        var el = els[i];
        var t = (el.innerText||'').trim();
        if (t === target && el.offsetParent !== null && el.children.length <= 3) {
            el.scrollIntoView({block:'center'});
            el.click();
            return 'clicked_el';
        }
    }
    return 'not_found';
})();
"@
                        if ($clickResult -ne 'not_found') { break }
                        Run-JS "window.scrollBy(0, 250);"
                        Start-Sleep -Milliseconds 400
                    }

                    if ($clickResult -eq 'not_found') {
                        Write-Log "    Still cannot find store. Skipping."
                        $skippedActiveStores.Add($storeName)
                        continue
                    }
                    Write-Log "    Clicked store: $clickResult"
                    Start-Sleep -Seconds 5

                    # Dismiss popup
                    Run-JS @'
(function(){ var b=document.querySelectorAll('button,a,span'); for(var i=0;i<b.length;i++){var t=(b[i].innerText||'').trim().toLowerCase(); if(t==='close'||t==='skip'||t==='dismiss'||t==='got it'){if(b[i].offsetParent!==null){b[i].click();return;}}} })();
'@
                    Start-Sleep -Seconds 2

                    # Wait for table (longer timeout on retry)
                    $tableLoaded = Wait-JS "(document.body.innerText||'').indexOf('ITEM NAME')>-1||(document.body.innerText||'').indexOf('CATEGORY')>-1||(document.body.innerText||'').indexOf('STATUS')>-1" 30
                    if (-not $tableLoaded) {
                        Write-Log "    Table still did not load on retry. Giving up on this store."
                        $skippedActiveStores.Add($storeName)
                        Go-To "https://merchant.grab.com/food/inventory"
                        Start-Sleep -Seconds 3
                        continue
                    }
                    Write-Log "    Table loaded on retry!"
                    Start-Sleep -Seconds 2

                    # Detect columns
                    Run-JS "window.scrollTo(0, 0);"
                    Start-Sleep -Seconds 1
                    $colMapJson = Run-JS @'
(function(){
    var headers = [];
    var thEls = document.querySelectorAll('th');
    if (thEls.length >= 3) { for (var i=0; i<thEls.length; i++) { headers.push(thEls[i].innerText.trim().toUpperCase()); } }
    if (headers.length === 0) { var rows = document.querySelectorAll('tr'); for (var i=0; i<rows.length; i++) { var txt = rows[i].innerText.toUpperCase(); if (txt.indexOf('ITEM NAME') > -1 && txt.indexOf('STATUS') > -1) { var cells = rows[i].querySelectorAll('th, td'); for (var c=0; c<cells.length; c++) { headers.push(cells[c].innerText.trim().toUpperCase()); } break; } } }
    var colMap = {nameIdx: -1, categoryIdx: -1, statusIdx: -1};
    for (var h=0; h<headers.length; h++) { var hdr = headers[h]; if (hdr.indexOf('ITEM NAME') > -1) colMap.nameIdx = h; else if (hdr.indexOf('CATEGORY') > -1) colMap.categoryIdx = h; else if (hdr.indexOf('STATUS') > -1) colMap.statusIdx = h; }
    return JSON.stringify(colMap);
})();
'@
                    $colMap = $null
                    if ($colMapJson) { try { $colMap = $colMapJson | ConvertFrom-Json } catch {} }
                    $nameIdx = if ($colMap -and $colMap.nameIdx -ge 0) { $colMap.nameIdx } else { 0 }
                    $catIdx2 = if ($colMap -and $colMap.categoryIdx -ge 0) { $colMap.categoryIdx } else { 2 }
                    $statusIdx = if ($colMap -and $colMap.statusIdx -ge 0) { $colMap.statusIdx } else { 5 }

                    # Scroll & capture items (same logic as main loop)
                    $itemsMap = @{}
                    $prevItemCount = 0; $noChangeCount = 0
                    Run-JS "window.scrollTo(0, 0);"
                    Start-Sleep -Seconds 1

                    for ($scrollStep = 0; $scrollStep -lt 100; $scrollStep++) {
                        if (([DateTime]::Now - $storeStartTime).TotalSeconds -gt $StoreTimeoutSeconds) {
                            Write-Log "    TIMEOUT on retry. Saving what was captured."
                            break
                        }
                        $visJson = Run-JS @"
(function(){
    var items = [];
    var nameIdx = $nameIdx; var catIdx = $catIdx2; var statusIdx = $statusIdx;
    
    // Store name patterns to filter out
    var storePatterns = ['FNP Cakes', 'FNP Flowers', 'FNP Gifts', 'FNP Cakes &', '[CLOSED]'];
    var locationNames = ['singapore', 'malaysia', 'indonesia', 'thailand', 'vietnam', 'philippines'];
    function isStoreName(name) {
        var n = name.trim();
        for (var p = 0; p < storePatterns.length; p++) { if (n.indexOf(storePatterns[p]) === 0 || n.indexOf('[CLOSED]') > -1) return true; }
        if (/^FNP\s/.test(n) && n.indexOf(' - ') > -1) return true;
        return false;
    }
    function isLocationCategory(cat) {
        var c = cat.trim().toLowerCase();
        for (var l = 0; l < locationNames.length; l++) { if (c === locationNames[l]) return true; }
        return false;
    }
    
    var rows = document.querySelectorAll('tr');
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td');
        if (cells.length < 3) continue;
        var nameCell = cells[nameIdx] ? cells[nameIdx].innerText.trim() : '';
        if (cells[nameIdx]) { var link = cells[nameIdx].querySelector('a'); if (link) nameCell = link.innerText.trim(); }
        if (!nameCell || nameCell.length < 2 || nameCell.length > 200) continue;
        if (nameCell.toUpperCase() === 'ITEM NAME' || nameCell.toUpperCase() === 'ITEM') continue;
        if (isStoreName(nameCell)) continue;
        var category = (catIdx < cells.length && cells[catIdx]) ? cells[catIdx].innerText.trim().split('\\n')[0].trim() : 'Unknown';
        if (isLocationCategory(category)) continue;
        var status = '';
        if (statusIdx < cells.length && cells[statusIdx]) {
            var sc = cells[statusIdx]; var sel = sc.querySelector('select');
            if (sel && sel.selectedIndex >= 0) { status = sel.options[sel.selectedIndex].text.trim(); }
            if (!status) { var ct = sc.innerText.trim().replace(/[\u25BC\u25B6\u25B2\u25C0\u2304\u2303\uFE0F\u200B]/g,'').trim(); status = ct.split('\\n')[0].trim(); }
        }
        if (nameCell) items.push({name: nameCell, category: category || 'Unknown', status: status || 'Unknown'});
    }
    return JSON.stringify(items);
})();
"@
                        if ($visJson) {
                            try {
                                $vis = $visJson | ConvertFrom-Json
                                foreach ($v in $vis) { if (-not $itemsMap.ContainsKey($v.name)) { $itemsMap[$v.name] = $v } }
                            } catch {}
                        }
                        if ($itemsMap.Count -eq $prevItemCount) { $noChangeCount++; if ($noChangeCount -ge 6) { break } } else { $noChangeCount = 0; $prevItemCount = $itemsMap.Count }
                        Run-JS @'
(function(){ var divs=document.querySelectorAll('div'); for(var i=0;i<divs.length;i++){var el=divs[i]; if(el.scrollHeight>el.clientHeight+50&&el.clientHeight>150&&el.offsetParent!==null){var t=el.innerText||''; if((t.indexOf('ITEM NAME')>-1||t.indexOf('CATEGORY')>-1||t.indexOf('S$')>-1)&&t.indexOf('STORE NAME')===-1){el.scrollBy(0,350);break;}}} window.scrollBy(0,350); })();
'@
                        Start-Sleep -Milliseconds 700
                    }

                    # Add items
                    $sgtZone = [TimeZoneInfo]::FindSystemTimeZoneById("Singapore Standard Time")
                    $sgtNow = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $sgtZone)
                    $checkDate = $sgtNow.ToString("MM/dd/yyyy")
                    $checkTime = $sgtNow.ToString("HH:mm:ss")
                    $capturedItems = @($itemsMap.Values)
                    if ($capturedItems.Count -gt 0) {
                        Write-Log "    RETRY SUCCESS: Captured $($capturedItems.Count) items from '$storeName'."
                        foreach ($item in $capturedItems) {
                            $allItems.Add(@{
                                StoreName = $storeName; ItemName = [string]$item.name
                                Category = [string]$item.category; Status = [string]$item.status
                                DateCheck = $checkDate; TimeCheck = $checkTime
                            })
                        }
                    } else {
                        Write-Log "    RETRY: Still no items found for '$storeName'."
                        $skippedActiveStores.Add($storeName)
                    }

                } catch {
                    Write-Log "    RETRY ERROR for '$storeName': $($_.Exception.Message)"
                    $skippedActiveStores.Add($storeName)
                    Reconnect-CDP | Out-Null
                }

                # Navigate back for next retry
                Go-To "https://merchant.grab.com/food/inventory"
                Start-Sleep -Seconds 3
                Wait-JS "(document.body.innerText||'').indexOf('STORE NAME')>-1" 10
                for ($sp = 0; $sp -lt 10; $sp++) {
                    Run-JS "window.scrollTo(0, document.body.scrollHeight);"
                    Start-Sleep -Milliseconds 400
                }
                Run-JS "window.scrollTo(0,0);"
                Start-Sleep -Seconds 1
            } # end retry foreach
        } # end retry while

        if ($skippedActiveStores.Count -gt 0) {
            Write-Log ""
            Write-Log "*** $($skippedActiveStores.Count) store(s) could not be captured after $maxRetries retries: $($skippedActiveStores -join ', ') ***"
        }

        # === STEP 8: Export to Google Sheets (via Apps Script Web App) + Local Excel backup ===
        Write-Log ""
        Write-Log "============================================"
        Write-Log "ALL STORES DONE!"
        Write-Log "Active stores processed: $storeIdx"
        Write-Log "Total items captured: $($allItems.Count)"
        Write-Log "Runtime: $(([DateTime]::Now - $Global:scriptStartTime).ToString('mm\:ss')) minutes"
        Write-Log "============================================"
        $progressBar.Value = 90

        if ($allItems.Count -eq 0) {
            Write-Log "No items found in any store."
            $lblStatus.Text = "Status: Complete - No items found"
        } else {
            # --- VALIDATION: Remove any store names that slipped through JS filters ---
            $beforeCount = $allItems.Count
            $itemsToRemove = [System.Collections.Generic.List[int]]::new()
            for ($vi = $allItems.Count - 1; $vi -ge 0; $vi--) {
                $itemName = [string]$allItems[$vi].ItemName
                $itemCat = [string]$allItems[$vi].Category
                # Filter: item name looks like a store name
                $isStore = $false
                if ($itemName -match "^FNP\s" -and $itemName -match " - ") { $isStore = $true }
                if ($itemName -match "^\[CLOSED\]") { $isStore = $true }
                # Filter: category is a country/location (store list artifact)
                if ($itemCat -match "^(Singapore|Malaysia|Indonesia|Thailand|Vietnam|Philippines)$") { $isStore = $true }
                if ($isStore) { $itemsToRemove.Add($vi) }
            }
            foreach ($idx in $itemsToRemove) { $allItems.RemoveAt($idx) }
            $removedCount = $beforeCount - $allItems.Count
            if ($removedCount -gt 0) {
                Write-Log "VALIDATION: Removed $removedCount store-name entries from item list."
            }
            Write-Log "Final clean item count: $($allItems.Count)"

            # --- PART A: Upload to Google Sheets via Apps Script Web App ---
            if ($EnableGoogleSheets) {
                Write-Log ""
                Write-Log "Uploading $($allItems.Count) items to Google Sheets..."
                $lblStatus.Text = "Status: Uploading to Google Sheets..."; $form.Refresh()
                try {
                    $uploadResult = Send-ToGoogleSheets -WebAppUrl $GoogleAppsScriptUrl -ItemData $allItems
                    if ($uploadResult) {
                        Write-Log "Google Sheets upload SUCCESSFUL!"
                    } else {
                        Write-Log "Google Sheets upload FAILED. Data saved to local Excel."
                    }
                } catch {
                    Write-Log "ERROR during Google Sheets upload: $($_.Exception.Message)"
                }
            } else {
                Write-Log "Google Sheets upload disabled."
            }

            # --- PART B: Save local CSV backup ---
            Write-Log "Saving local CSV backup..."
            $lblStatus.Text = "Status: Saving local CSV..."; $form.Refresh()
            
            $dateSuffix = (Get-Date).ToString("ddMMMyyyy_HHmmss")
            $outputPath = Join-Path $outputFolder "Grab_Inventory_Status_$dateSuffix.csv"

            try {
                $csvData = $allItems | ForEach-Object {
                    [PSCustomObject]@{
                        'Store Name'    = $_.StoreName
                        'Item Name'     = $_.ItemName
                        'Category'      = $_.Category
                        'Status'        = $_.Status
                        'Date of Check' = $_.DateCheck
                        'Time of Check' = $_.TimeCheck
                    }
                }
                $csvData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
                Write-Log "Local CSV saved: $outputPath"
            } catch {
                Write-Log "  Local CSV save failed: $($_.Exception.Message)"
            }

            $progressBar.Value = 100
            $lblStatus.Text = "Status: Done! $($allItems.Count) items from $storeIdx stores."
            Write-Log ""
            Write-Log "COMPLETE! $($allItems.Count) items uploaded to Google Sheets + local CSV."
        }

    }
    catch {
        $lblStatus.Text = "Status: Error"
        Write-Log "FATAL ERROR: $($_.Exception.Message)"

        # Even on error, try to save whatever data we have locally as CSV
        if ($allItems -and $allItems.Count -gt 0) {
            Write-Log "Attempting to save $($allItems.Count) items collected before error (local CSV)..."
            try {
                $dateSuffix = (Get-Date).ToString("ddMMMyyyy_HHmmss")
                $outputPath = Join-Path $outputFolder "Grab_Inventory_Status_PARTIAL_$dateSuffix.csv"
                $allItems | ForEach-Object {
                    [PSCustomObject]@{
                        'Store Name'=$_.StoreName; 'Item Name'=$_.ItemName
                        'Category'=$_.Category; 'Status'=$_.Status
                        'Date of Check'=$_.DateCheck; 'Time of Check'=$_.TimeCheck
                    }
                } | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
                Write-Log "PARTIAL data saved locally: $outputPath"
            } catch {
                Write-Log "Could not save partial data: $($_.Exception.Message)"
            }
        }

        Write-Log "Error details logged above."
    }
    finally {
        if ($Global:ws -and $Global:ws.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
            try { $Global:ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure,"done",[System.Threading.CancellationToken]::None).Wait(5000) } catch {}
        }
        if ($Global:ws) { try { $Global:ws.Dispose() } catch {}; $Global:ws = $null }
        if ($edgeProc -and -not $edgeProc.HasExited) { try { $edgeProc.Kill() } catch {} }
        if ($excel) { try { $excel.Quit() } catch {}; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $btnStart.Enabled = $true

        # Auto-close the form after completion so script exits cleanly for Task Scheduler
        Write-Log "Script finished. Closing in 5 seconds..."
        Start-Sleep -Seconds 5
        $form.Close()
    }
})

# --- Auto-Start for Task Scheduler ---
# When run from Task Scheduler, auto-click Start after form loads
$script:autoStartTimer = New-Object System.Windows.Forms.Timer
$script:autoStartTimer.Interval = 3000  # 3 seconds delay to let form render
$script:autoStartTimer.Add_Tick({
    $script:autoStartTimer.Stop()
    $script:autoStartTimer.Dispose()
    if ($btnStart.Enabled) {
        $btnStart.PerformClick()
    }
})
$form.Add_Shown({
    $script:autoStartTimer.Start()
})

# --- Show Form ---
[void]$form.ShowDialog()
