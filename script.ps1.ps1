# ============================================================
# GrabDeactivatedItems_Edge.ps1
# Checks Grab Merchant stores for deactivated/out-of-stock items.
# NO INSTALLATION REQUIRED - uses Edge + PowerShell + Excel only.
# Right-click > "Run with PowerShell"
# ============================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$outputFolder = $PSScriptRoot
$GrabLoginUrl = "https://merchant.grab.com/portal"
$GrabUsername = "automationtester"
$GrabPassword = "Operations@123"
$DebugPort    = 9222

$edgePath = @(
    "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edgePath) { [System.Windows.Forms.MessageBox]::Show("Edge not found.","Error"); exit }

# --- UI ---
$form = New-Object System.Windows.Forms.Form -Property @{ Text="Grab Deactivated Items Checker"; Size="620,460"; StartPosition="CenterScreen"; FormBorderStyle="FixedDialog"; MaximizeBox=$false }
$lblTitle = New-Object System.Windows.Forms.Label -Property @{ Text="Grab - Deactivated Items Checker"; AutoSize=$true; Location="20,15"; Font=New-Object System.Drawing.Font("Segoe UI",13,[System.Drawing.FontStyle]::Bold) }
$form.Controls.Add($lblTitle)
$btnRun = New-Object System.Windows.Forms.Button -Property @{ Text="Run Check"; Size="150,40"; Location="20,55"; Font=New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold) }
$form.Controls.Add($btnRun)
$lblStatus = New-Object System.Windows.Forms.Label -Property @{ Text="Status: Ready"; Size="570,20"; Location="20,105" }
$form.Controls.Add($lblStatus)
$progressBar = New-Object System.Windows.Forms.ProgressBar -Property @{ Location="20,130"; Size="565,22"; Maximum=100 }
$form.Controls.Add($progressBar)
$txtLog = New-Object System.Windows.Forms.TextBox -Property @{ Location="20,160"; Size="565,260"; Multiline=$true; ScrollBars="Vertical"; ReadOnly=$true; Font=New-Object System.Drawing.Font("Consolas",8.5) }
$form.Controls.Add($txtLog)

# --- Helpers ---
function Write-Log { param([string]$msg); $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] $msg`r`n"); $form.Refresh() }
$Global:cdpId = 0; $Global:ws = $null
function Send-CDP { param([string]$method, [hashtable]$params = @{})
    $Global:cdpId++
    $msg = @{ id=$Global:cdpId; method=$method; params=$params } | ConvertTo-Json -Depth 10 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($msg)
    $Global:ws.SendAsync((New-Object System.ArraySegment[byte](,$bytes)),[System.Net.WebSockets.WebSocketMessageType]::Text,$true,[System.Threading.CancellationToken]::None).Wait()
    $buf = New-Object byte[] 2097152
    $deadline = [DateTime]::Now.AddSeconds(30)
    while ([DateTime]::Now -lt $deadline) {
        $res=""; $done=$false
        while (-not $done) { $seg=New-Object System.ArraySegment[byte](,$buf); $task=$Global:ws.ReceiveAsync($seg,[System.Threading.CancellationToken]::None); $task.Wait(30000); if($task.IsCompleted){$res+=[System.Text.Encoding]::UTF8.GetString($buf,0,$task.Result.Count);$done=$task.Result.EndOfMessage}else{return $null} }
        try { $j=$res|ConvertFrom-Json; if($j.id -eq $Global:cdpId){return $j} } catch {}
    }
    return $null
}
function Run-JS { param([string]$js)
    $r = Send-CDP "Runtime.evaluate" @{ expression=$js; returnByValue=$true; awaitPromise=$true }
    if ($r -and $r.result -and $r.result.result) { return $r.result.result.value }; return $null
}
function Wait-JS { param([string]$condition, [int]$seconds=30)
    $end=[DateTime]::Now.AddSeconds($seconds)
    while([DateTime]::Now -lt $end){ $v=Run-JS $condition; if($v -eq $true -or $v -eq "true" -or $v -eq "True"){return $true}; Start-Sleep -Seconds 1 }
    return $false
}
function Go-To { param([string]$url)
    Send-CDP "Page.navigate" @{url=$url} | Out-Null; Start-Sleep -Seconds 3
    for($i=0;$i -lt 30;$i++){if((Run-JS "document.readyState") -eq "complete"){break};Start-Sleep -Seconds 1}; Start-Sleep -Seconds 2
}

# --- Main Logic ---
$btnRun.Add_Click({
    $btnRun.Enabled = $false; $txtLog.Clear(); $progressBar.Value = 0
    $edgeProc = $null; $excel = $null

    try {
        # === STEP 1: Launch Edge ===
        Write-Log "Launching Edge..."
        $lblStatus.Text = "Status: Starting Edge..."; $form.Refresh()
        $tempProfile = Join-Path $env:TEMP "GrabChecker_Edge"
        if (-not (Test-Path $tempProfile)) { New-Item -ItemType Directory -Path $tempProfile -Force | Out-Null }
        $edgeProc = Start-Process -FilePath $edgePath -ArgumentList "--remote-debugging-port=$DebugPort --user-data-dir=`"$tempProfile`" --no-first-run --disable-extensions `"$GrabLoginUrl`"" -PassThru
        Start-Sleep -Seconds 7
        Write-Log "Edge started."
        $progressBar.Value = 5

        # === STEP 2: Connect CDP ===
        Write-Log "Connecting..."
        $targets = $null
        for ($i=0; $i -lt 15; $i++) { try { $targets = Invoke-RestMethod "http://localhost:$DebugPort/json" -UseBasicParsing; if($targets){break} } catch {}; Start-Sleep -Seconds 2 }
        if (-not $targets) { throw "Cannot connect to Edge. Close ALL Edge windows and retry." }
        $page = $targets | Where-Object { $_.type -eq "page" } | Select-Object -First 1
        $Global:ws = New-Object System.Net.WebSockets.ClientWebSocket
        $Global:ws.ConnectAsync([Uri]$page.webSocketDebuggerUrl,[System.Threading.CancellationToken]::None).Wait()
        Send-CDP "Page.enable" | Out-Null; Send-CDP "Runtime.enable" | Out-Null
        Write-Log "Connected."
        $progressBar.Value = 8

        # === STEP 3: LOGIN ===
        Write-Log "Waiting for page..."
        $lblStatus.Text = "Status: Logging in..."; $form.Refresh()
        Start-Sleep -Seconds 8
        $url = Run-JS "window.location.href"
        Write-Log "URL: $url"

        # Handle saved-accounts page
        if ($url -match "saved-accounts") {
            Write-Log "Saved accounts page - clicking 'Login as another user'..."
            Run-JS "window.scrollTo(0, document.body.scrollHeight);"
            Start-Sleep -Seconds 2
            Run-JS @'
(function(){ var els=document.querySelectorAll('button,a,div,span'); for(var i=0;i<els.length;i++){var t=(els[i].innerText||'').trim(); if(t==='Login as another user'||t.toLowerCase().includes('another user')){els[i].click();return;}} })();
'@
            Start-Sleep -Seconds 5
            Wait-JS "document.querySelector('#Username') !== null || document.querySelector('input') !== null" 20
        }

        $url = Run-JS "window.location.href"
        if ($url -match "merchant\.grab\.com" -and $url -notmatch "weblogin") {
            Write-Log "Already logged in!"
        } else {
            # Enter username
            Write-Log "Entering username..."
            Start-Sleep -Seconds 2
            Run-JS @'
(function(){ var inp=document.querySelector('#Username')||document.querySelector('input[type=text]')||document.querySelector('input[type=email]'); if(!inp)return; inp.focus();inp.select();document.execCommand('delete'); var s=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set; s.call(inp,'automationtester'); inp.dispatchEvent(new Event('input',{bubbles:true})); inp.dispatchEvent(new Event('change',{bubbles:true})); })();
'@
            Start-Sleep -Seconds 2
            Run-JS @'
(function(){ var b=document.querySelectorAll('button'); for(var i=0;i<b.length;i++){var t=(b[i].innerText||'').trim().toLowerCase(); if(t==='continue'||t==='next'){b[i].click();return;}} var s=document.querySelector('button[type=submit]'); if(s)s.click(); })();
'@
            Start-Sleep -Seconds 5

            # Enter password if needed
            $url = Run-JS "window.location.href"
            if (-not ($url -match "merchant\.grab\.com" -and $url -notmatch "weblogin")) {
                $pwReady = Wait-JS 'document.querySelector("input[type=password]") !== null' 15
                if ($pwReady) {
                    Write-Log "Entering password..."
                    Run-JS @'
(function(){ var pw=document.querySelector('input[type=password]'); if(!pw)return; pw.focus();pw.select();document.execCommand('delete'); var s=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set; s.call(pw,'Operations@123'); pw.dispatchEvent(new Event('input',{bubbles:true})); pw.dispatchEvent(new Event('change',{bubbles:true})); })();
'@
                    Start-Sleep -Seconds 2
                    Run-JS @'
(function(){ var b=document.querySelectorAll('button'); for(var i=0;i<b.length;i++){var t=(b[i].innerText||'').trim().toLowerCase(); if(t==='continue'||t==='log in'){b[i].click();return;}} var s=document.querySelector('button[type=submit]'); if(s)s.click(); })();
'@
                    Start-Sleep -Seconds 8
                }
            }
        }

        # Verify login
        for ($i=0; $i -lt 15; $i++) { $url=Run-JS "window.location.href"; if($url -match "merchant\.grab\.com" -and $url -notmatch "weblogin"){break}; Start-Sleep -Seconds 2 }
        if (-not ($url -match "merchant\.grab\.com" -and $url -notmatch "weblogin")) { throw "Login failed. URL: $url" }
        Write-Log "Login OK!"
        $progressBar.Value = 15

        # Dismiss welcome popup
        Start-Sleep -Seconds 3
        Run-JS @'
(function(){ var b=document.querySelectorAll('button,a,span'); for(var i=0;i<b.length;i++){var t=(b[i].innerText||'').trim().toLowerCase(); if(t==='close'||t==='skip'||t==='dismiss'){if(b[i].offsetParent!==null){b[i].click();return;}}} })();
'@
        Start-Sleep -Seconds 2

        # === STEP 4: Navigate to MENU section ===
        Write-Log "Going to Menu..."
        $lblStatus.Text = "Status: Going to Menu..."; $form.Refresh()
        Go-To "https://merchant.grab.com/food/menu"
        Wait-JS "(document.body.innerText||'').indexOf('STORE NAME') > -1" 15
        Write-Log "Menu page loaded."
        $progressBar.Value = 20

        # === STEP 5: Get all store names from the table ===
        Write-Log "Reading all stores from table..."
        # Scroll the store table to load all rows
        for ($ts = 0; $ts -lt 10; $ts++) {
            Run-JS "window.scrollBy(0, 300);"
            Start-Sleep -Milliseconds 600
        }
        Run-JS "window.scrollTo(0, 0);"
        Start-Sleep -Seconds 1

        $storeNamesJson = Run-JS @'
(function(){
    var stores = [];
    var rows = document.querySelectorAll('tr.dui-table-row, [class*="dui-table-row"]');
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td, [class*="dui-table-cell"]');
        if (cells.length > 0) {
            var name = cells[0].innerText.trim();
            if (name.length > 3 && name !== 'STORE NAME') {
                stores.push(name);
            }
        }
    }
    return JSON.stringify(stores);
})();
'@
        $storeNames = @()
        if ($storeNamesJson) { try { $storeNames = $storeNamesJson | ConvertFrom-Json } catch {} }
        Write-Log "Found $($storeNames.Count) stores:"
        foreach ($s in $storeNames) { Write-Log "  - $s" }
        if ($storeNames.Count -eq 0) { throw "No stores found in table." }

        # === STEP 6: Loop through each store ===
        $allItems = [System.Collections.Generic.List[hashtable]]::new()
        $storeIdx = 0

        foreach ($storeObj in $storeNames) {
            $storeIdx++
            $targetStore = $storeObj
            # Count how many times this name appeared so far (for duplicate handling)
            $occurrence = 0
            for ($k = 0; $k -lt $storeIdx; $k++) {
                if ($storeNames[$k] -eq $targetStore) { $occurrence++ }
            }
            Write-Log ""
            Write-Log "=============================="
            Write-Log "=== Store $storeIdx/$($storeNames.Count): $targetStore ==="
            Write-Log "=============================="
            $lblStatus.Text = "Status: Store $storeIdx/$($storeNames.Count) - $targetStore"; $form.Refresh()
            $progressBar.Value = 20 + [int](($storeIdx / $storeNames.Count) * 65)

            # Click on this store by name (use occurrence index for duplicates)
            $clickStore = Run-JS @"
(function(){
    var target = '$($targetStore -replace "'","\\'")';
    var occurrence = $occurrence;
    var count = 0;
    var rows = document.querySelectorAll('tr.dui-table-row, [class*=dui-table-row]');
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td, [class*="dui-table-cell"]');
        var rowName = cells.length > 0 ? cells[0].innerText.trim() : (rows[i].innerText||'').trim().split('\n')[0].trim();
        if (rowName === target) {
            count++;
            if (count === occurrence) {
                rows[i].scrollIntoView({block:'center'});
                rows[i].click();
                return 'clicked';
            }
        }
    }
    // Fallback: indexOf match
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td, [class*="dui-table-cell"]');
        var rowName = cells.length > 0 ? cells[0].innerText.trim() : '';
        if (rowName === target) {
            rows[i].scrollIntoView({block:'center'});
            rows[i].click();
            return 'clicked_fallback';
        }
    }
    return 'not_found';
})();
"@
            Write-Log "  Click: $clickStore"
            if ($clickStore -eq 'not_found') {
                # Scroll page down and retry finding the store
                for ($retry = 0; $retry -lt 15; $retry++) {
                    Run-JS "window.scrollBy(0, 300);"
                    Start-Sleep -Milliseconds 500
                    $clickStore = Run-JS @"
(function(){
    var target = '$($targetStore -replace "'","\\'")';
    var occurrence = $occurrence;
    var count = 0;
    var rows = document.querySelectorAll('tr.dui-table-row, [class*=dui-table-row]');
    for (var i=0; i<rows.length; i++) {
        var cells = rows[i].querySelectorAll('td, [class*="dui-table-cell"]');
        var rowName = cells.length > 0 ? cells[0].innerText.trim() : '';
        if (rowName === target) {
            count++;
            if (count === occurrence) {
                rows[i].scrollIntoView({block:'center'});
                rows[i].click();
                return 'clicked';
            }
        }
    }
    return 'not_found';
})();
"@
                    if ($clickStore -ne 'not_found') { break }
                }
                if ($clickStore -eq 'not_found') {
                    Write-Log "  SKIPPED - could not find store in table."
                    continue
                }
            }
            Start-Sleep -Seconds 3

            # Dismiss popup
            Run-JS @'
(function(){ var b=document.querySelectorAll('button,a,span'); for(var i=0;i<b.length;i++){var t=(b[i].innerText||'').trim().toLowerCase(); if(t==='close'||t==='skip'){if(b[i].offsetParent!==null){b[i].click();return;}}} })();
'@
            Start-Sleep -Seconds 1

        # === STEP 7: Get all categories from the left column ===
        Write-Log "  Reading categories..."
        Start-Sleep -Seconds 2
        # Scroll categories column and collect category names at each position
        $catList = [System.Collections.Generic.List[string]]::new()
        
        # Scroll page down to load all categories (Add Ons may be below viewport)
        Run-JS "window.scrollTo(0, 0);"
        Start-Sleep -Milliseconds 500
        
        for ($catPass = 0; $catPass -lt 5; $catPass++) {
            $visCatJson = Run-JS @'
(function(){
    var cats = [];
    var rows = document.querySelectorAll('[class*="dui-row-middle"]');
    for (var i=0; i<rows.length; i++) {
        var cols = rows[i].querySelectorAll('[class*="dui-col"]');
        if (cols.length > 0) {
            var nameEl = cols[0].querySelector('div');
            if (nameEl) {
                var name = (nameEl.textContent||'').trim().replace(/\d+$/, '').trim();
                if (name.length > 1 && name.length < 50 && cats.indexOf(name) === -1 &&
                    !name.match(/^[\d]+$/) && !name.match(/S?\$/) &&
                    !name.match(/^(CATEGORIES|ITEMS|OVERVIEW|OPTION|BULK|Add new|All |Search)/i)) {
                    cats.push(name);
                }
            }
        }
    }
    return JSON.stringify(cats);
})();
'@
            if ($visCatJson) {
                try {
                    $visCats = $visCatJson | ConvertFrom-Json
                    foreach ($vc in $visCats) { if (-not $catList.Contains($vc)) { $catList.Add($vc) } }
                } catch {}
            }
            # Scroll page down to reveal more categories below viewport
            Run-JS "window.scrollBy(0, 300);"
            Start-Sleep -Milliseconds 600
        }
        Run-JS "window.scrollTo(0, 0);"
        Start-Sleep -Milliseconds 500
        
        $categories = $catList.ToArray()
        if ($categories.Count -eq 0) {
            # Fallback: just get all short text items on the page that could be categories
            Write-Log "  Trying fallback category detection..."
            $catJson = Run-JS @'
(function(){
    var cats = [];
    var seen = {};
    var els = document.querySelectorAll('div, span, a');
    for (var i=0; i<els.length; i++) {
        var el = els[i];
        if (el.children.length > 1) continue;
        var t = (el.textContent||'').trim();
        if (t.length > 2 && t.length < 50 && !seen[t] && el.offsetParent !== null &&
            !t.match(/^\d+$/) && !t.match(/^S?\$/) &&
            !t.match(/^(CATEGORIES|ITEMS|OVERVIEW|OPTION GROUPS|BULK UPDATE|Add new|All Items|All availability|Search|Menu|Home|Orders|Insights|Feedback|Finance|Marketing|Inventory|Employees|Help|Get the App)/i)) {
            // Check if nearby sibling has a number (item count) - typical for categories
            var next = el.nextElementSibling;
            var prev = el.previousElementSibling;
            var hasCount = false;
            if (next && (next.textContent||'').trim().match(/^\d+$/)) hasCount = true;
            if (prev && (prev.textContent||'').trim().match(/^\d+$/)) hasCount = true;
            if ((el.parentElement.textContent||'').match(/\d+/) && t.match(/[a-zA-Z]/)) hasCount = true;
            if (hasCount) { seen[t] = true; cats.push(t); }
        }
    }
    return JSON.stringify(cats);
})();
'@
            if ($catJson) { try { $categories = $catJson | ConvertFrom-Json } catch {} }
        }
        Write-Log "  Categories found ($($categories.Count)): $($categories -join ', ')"
        $progressBar.Value = 30

        # === STEP 8: For each category, click it, scroll items, find deactivated ===
        $storeName = $targetStore
        $catIdx = 0

        foreach ($cat in $categories) {
            $catIdx++
            Write-Log ""
            Write-Log "--- Category $catIdx/$($categories.Count): $cat ---"
            $lblStatus.Text = "Status: $storeName > $cat"; $form.Refresh()

            # Click this category in the left panel
            $escapedCat = $cat -replace "'","\\'"
            Run-JS @"
(function(){
    var els = document.querySelectorAll('div, span, li, a');
    for (var i=0; i<els.length; i++) {
        var t = (els[i].innerText||'').trim().split('\n')[0].replace(/\s+\d+$/, '').trim();
        if (t === '$escapedCat' && els[i].offsetParent !== null) {
            els[i].click();
            return 'clicked';
        }
    }
    return 'nf';
})();
"@
            Start-Sleep -Seconds 4

            # Scroll page back to top so items panel is properly positioned
            Run-JS "window.scrollTo(0, 0);"
            Start-Sleep -Seconds 2
            
            # Wait for items to load in the new category
            Wait-JS "(document.body.innerText||'').includes('S$')" 8

            # Now scroll the ITEMS column to find all deactivated items
            $deactivatedMap = @{}

            # Reset items container to top and do initial scan
            Run-JS @'
(function(){
    var allDivs = document.querySelectorAll('div');
    for (var i=0; i<allDivs.length; i++) {
        var el = allDivs[i];
        if (el.scrollHeight > el.clientHeight + 50 && el.clientHeight > 200) {
            var text = el.innerText || '';
            if (text.includes('S$') && !text.includes('Orders') && !text.includes('Inventory')) {
                el.scrollTop = 0;
                return 'reset_items_container';
            }
        }
    }
    return 'not_found';
})();
'@
            Start-Sleep -Seconds 1

            # PowerShell-controlled scroll loop on the items container
            for ($scrollStep = 0; $scrollStep -lt 50; $scrollStep++) {
                # Collect visible deactivated items
                $visJson = Run-JS @'
(function(){
    var items=[];
    var kw=['Hidden','Out of stock indefinitely','Out of stock today'];
    var tags=document.querySelectorAll('[class*="tag"], span, div');
    for(var i=0;i<tags.length;i++){
        var tt=(tags[i].innerText||'').trim();
        var ms='';
        for(var s=0;s<kw.length;s++){if(tt===kw[s]){ms=tt;break;}}
        if(!ms)continue;
        // Find item name from parent row
        var row=tags[i].parentElement;
        for(var p=0;p<8;p++){if(!row)break; var rt=(row.innerText||''); if(rt.split('\n').length>=2&&(rt.includes('S$')||rt.includes('$'))){break;} row=row.parentElement;}
        if(!row)continue;
        var lines=(row.innerText||'').split('\n').map(function(l){return l.trim()}).filter(function(l){return l.length>0});
        var nm='';
        for(var li=0;li<lines.length;li++){
            var ln=lines[li];
            if(ln===ms||ln==='Hidden'||ln.match(/^Out of stock/)||ln.match(/^S?\$\d/)||ln.length<=1)continue;
            if(ln.length>2&&ln.length<150){nm=ln;break;}
        }
        if(nm&&nm!==ms)items.push({name:nm,status:ms});
    }
    return JSON.stringify(items);
})();
'@
                if ($visJson) {
                    try {
                        $vis = $visJson | ConvertFrom-Json
                        foreach ($v in $vis) {
                            $key = "$($v.name)|$($v.status)"
                            if (-not $deactivatedMap.ContainsKey($key)) { $deactivatedMap[$key] = $v }
                        }
                    } catch {}
                }

                # Scroll the items container down
                $scrollResult = Run-JS @'
(function(){
    var scrolled = false;
    // Try to find and scroll the items container
    var allDivs = document.querySelectorAll('div');
    for (var i=0; i<allDivs.length; i++) {
        var el = allDivs[i];
        if (el.scrollHeight > el.clientHeight + 50 && el.clientHeight > 200) {
            var text = el.innerText || '';
            if (text.includes('S$') && !text.includes('Orders') && !text.includes('Inventory')) {
                var before = el.scrollTop;
                el.scrollBy(0, 500);
                if (el.scrollTop !== before) { scrolled = true; break; }
            }
        }
    }
    // Fallback: scroll the window
    if (!scrolled) {
        var wb = window.pageYOffset;
        window.scrollBy(0, 400);
        if (window.pageYOffset !== wb) scrolled = true;
    }
    return scrolled ? 'MORE' : 'DONE';
})();
'@
                Start-Sleep -Milliseconds 1200
                if ($scrollResult -eq 'DONE') {
                    # Scroll one more time to be absolutely sure we're at the bottom
                    Start-Sleep -Milliseconds 500
                    Run-JS @'
(function(){ var allDivs=document.querySelectorAll('div'); for(var i=0;i<allDivs.length;i++){var el=allDivs[i]; if(el.scrollHeight>el.clientHeight+50&&el.clientHeight>200){var text=el.innerText||''; if(text.includes('S$')&&!text.includes('Orders')&&!text.includes('Inventory')){el.scrollBy(0,500);}}} })();
'@
                    Start-Sleep -Milliseconds 800
                    # Do 3 extra collection passes to catch the last items
                    for ($extra = 0; $extra -lt 3; $extra++) {
                        Start-Sleep -Milliseconds 500
                        $visJson = Run-JS @'
(function(){
    var items=[];
    var kw=['Hidden','Out of stock indefinitely','Out of stock today'];
    var tags=document.querySelectorAll('[class*="tag"], span, div');
    for(var i=0;i<tags.length;i++){
        var tt=(tags[i].innerText||'').trim();
        var ms='';
        for(var s=0;s<kw.length;s++){if(tt===kw[s]){ms=tt;break;}}
        if(!ms)continue;
        var row=tags[i].parentElement;
        for(var p=0;p<8;p++){if(!row)break; var rt=(row.innerText||''); if(rt.split('\n').length>=2&&(rt.includes('S$')||rt.includes('$'))){break;} row=row.parentElement;}
        if(!row)continue;
        var lines=(row.innerText||'').split('\n').map(function(l){return l.trim()}).filter(function(l){return l.length>0});
        var nm='';
        for(var li=0;li<lines.length;li++){
            var ln=lines[li];
            if(ln===ms||ln==='Hidden'||ln.match(/^Out of stock/)||ln.match(/^S?\$\d/)||ln.length<=1)continue;
            if(ln.length>2&&ln.length<150){nm=ln;break;}
        }
        if(nm&&nm!==ms)items.push({name:nm,status:ms});
    }
    return JSON.stringify(items);
})();
'@
                        if ($visJson) { try { $vis=$visJson|ConvertFrom-Json; foreach($v in $vis){$key="$($v.name)|$($v.status)"; if(-not $deactivatedMap.ContainsKey($key)){$deactivatedMap[$key]=$v}} } catch {} }
                    }
                    break
                }
            }

            $deactivated = @($deactivatedMap.Values)
            if ($deactivated.Count -eq 0) {
                Write-Log "  No deactivated items in this category."
                continue
            }
            Write-Log "  Found $($deactivated.Count) deactivated item(s)"

            # For each deactivated item: click it, read SKU, close panel
            foreach ($dItem in $deactivated) {
                $itemName = $dItem.name
                $status = $dItem.status
                $escapedItem = $itemName -replace "'","\\'" -replace '"','\"'

                # Close any open edit panel first
                Run-JS @'
(function(){ var x=document.querySelector('[class*="close"]'); if(x&&x.offsetParent){x.click();} var b=document.querySelectorAll('button,span'); for(var i=0;i<b.length;i++){var t=(b[i].innerText||'').trim(); if((t==='X'||t==='\u00d7')&&b[i].offsetParent){b[i].click();return;}} })();
'@
                Start-Sleep -Seconds 1

                # Scroll the items container to find this item
                $clickRes = Run-JS @"
(function(){
    var target = '$escapedItem';
    
    // First try to find and click directly
    var els = document.querySelectorAll('div, span, p, a');
    for (var i=0; i<els.length; i++) {
        var t = (els[i].textContent||'').trim();
        if (t === target && els[i].offsetParent !== null && els[i].children.length === 0) {
            els[i].scrollIntoView({block:'center'});
            els[i].click();
            return 'CLICKED';
        }
    }
    for (var i=0; i<els.length; i++) {
        var t = (els[i].innerText||'').trim().split('\n')[0].trim();
        if (t === target && els[i].offsetParent !== null) {
            els[i].scrollIntoView({block:'center'});
            els[i].click();
            return 'CLICKED2';
        }
    }
    
    // If not found, scroll items container to top and try again
    var allDivs = document.querySelectorAll('div');
    var container = null;
    for (var i=0; i<allDivs.length; i++) {
        if (allDivs[i].scrollHeight > allDivs[i].clientHeight + 50 && allDivs[i].clientHeight > 200) {
            var text = allDivs[i].innerText || '';
            if (text.includes('S`$') && !text.includes('Orders') && !text.includes('Inventory')) {
                container = allDivs[i];
                break;
            }
        }
    }
    if (container) {
        container.scrollTop = 0;
    }
    return 'NF_NEED_SCROLL';
})();
"@
                Write-Log "    Click '$itemName': $clickRes"
                
                # If item not found, scroll through container to find it
                if ($clickRes -eq 'NF_NEED_SCROLL') {
                    Start-Sleep -Seconds 1
                    for ($findIt = 0; $findIt -lt 20; $findIt++) {
                        $tryClick = Run-JS @"
(function(){
    var target = '$escapedItem';
    var els = document.querySelectorAll('div, span, p, a');
    for (var i=0; i<els.length; i++) {
        var t = (els[i].textContent||'').trim();
        if (t === target && els[i].offsetParent !== null && els[i].children.length === 0) {
            els[i].scrollIntoView({block:'center'});
            els[i].click();
            return 'CLICKED';
        }
    }
    // Scroll container down
    var allDivs = document.querySelectorAll('div');
    for (var i=0; i<allDivs.length; i++) {
        if (allDivs[i].scrollHeight > allDivs[i].clientHeight + 50 && allDivs[i].clientHeight > 200) {
            var text = allDivs[i].innerText || '';
            if (text.includes('S`$') && !text.includes('Orders')) {
                allDivs[i].scrollBy(0, 300);
                return 'SCROLLED';
            }
        }
    }
    return 'CANT_SCROLL';
})();
"@
                        if ($tryClick -match 'CLICKED') { $clickRes = $tryClick; break }
                        if ($tryClick -eq 'CANT_SCROLL') { break }
                        Start-Sleep -Milliseconds 500
                    }
                    Write-Log "    Retry result: $clickRes"
                }
                
                Start-Sleep -Seconds 2

                # Wait for EDIT ITEM panel with #itemCode
                Wait-JS "document.querySelector('#itemCode') !== null" 5
                Start-Sleep -Seconds 1

                # Read SKU from input#itemCode
                $sku = Run-JS @'
(function(){
    var inp = document.querySelector('#itemCode');
    if (inp && inp.value) return inp.value.trim();
    return '';
})();
'@
                $checkTime = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss")
                Write-Log "    SKU: $sku | Status: $status"

                # Close the EDIT ITEM panel (click the X button)
                Run-JS @'
(function(){
    // The EDIT ITEM panel has X at top-left
    var btns = document.querySelectorAll('button, span, div');
    for (var i=0; i<btns.length; i++) {
        var t = (btns[i].innerText||'').trim();
        if ((t === '\u00d7' || t === 'X' || t === '\u2715') && btns[i].offsetParent !== null) {
            btns[i].click(); return 'closed';
        }
    }
    // Try aria-label
    var x = document.querySelector('[aria-label="Close"]');
    if (x) { x.click(); return 'closed2'; }
    return 'no_close';
})();
'@
                Start-Sleep -Seconds 1

                $allItems.Add(@{
                    StoreName = $storeName
                    Category  = $cat
                    ItemName  = $itemName
                    SKU       = if ($sku) { [string]$sku } else { "" }
                    Status    = $status
                    CheckTime = $checkTime
                })
            } # end foreach dItem
        } # end foreach cat

            Write-Log "  Store $targetStore done. Items so far: $($allItems.Count)"

            # Go back to Menu page for next store
            Go-To "https://merchant.grab.com/food/menu"
            Start-Sleep -Seconds 2
            Wait-JS "(document.body.innerText||'').indexOf('STORE NAME') > -1" 10
            # Scroll page down to load all store rows, then back to top
            for ($ts = 0; $ts -lt 10; $ts++) {
                Run-JS "window.scrollBy(0, 300);"
                Start-Sleep -Milliseconds 400
            }
            Run-JS "window.scrollTo(0, 0);"
            Start-Sleep -Seconds 1
        } # end foreach store

        Write-Log ""
        Write-Log "=== ALL STORES DONE: $($allItems.Count) deactivated items found ==="
        $progressBar.Value = 88

        # === STEP 9: Export to Excel ===
        Write-Log "Creating Excel..."
        $lblStatus.Text = "Status: Saving Excel..."; $form.Refresh()
        $outputPath = Join-Path $outputFolder "CFC Grab Deactivated Items_$((Get-Date).ToString('ddMMMyyyy'))T$((Get-Date).ToString('HHmmss')).xlsx"

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
        $ws.Name = "Deactivated Items"

        @("Store Name","Category","Item Name","SKU Code","Status","Time of Check") | ForEach-Object -Begin{$c=1} { $ws.Cells.Item(1,$c)=$_; $c++ }
        $ws.Rows.Item(1).Font.Bold = $true

        $row = 2
        foreach ($item in $allItems) {
            $ws.Cells.Item($row,1) = [string]$item.StoreName
            $ws.Cells.Item($row,2) = [string]$item.Category
            $ws.Cells.Item($row,3) = [string]$item.ItemName
            $ws.Cells.Item($row,4) = [string]$item.SKU
            $ws.Cells.Item($row,5) = [string]$item.Status
            $ws.Cells.Item($row,6) = [string]$item.CheckTime
            $row++
        }
        $ws.UsedRange.Columns.AutoFit() | Out-Null
        $wb.SaveAs($outputPath, 51); $wb.Close($false); $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel); $excel=$null

        $progressBar.Value = 100
        $lblStatus.Text = "Status: Done! $($allItems.Count) items found."
        Write-Log "Excel saved: $outputPath"
        [System.Windows.Forms.MessageBox]::Show("Done!`nItems: $($allItems.Count)`nFile: $outputPath","Success")
    }
    catch {
        $lblStatus.Text = "Status: Error"
        Write-Log "ERROR: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error:`n$($_.Exception.Message)","Error")
    }
    finally {
        if ($Global:ws -and $Global:ws.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
            try{$Global:ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure,"done",[System.Threading.CancellationToken]::None).Wait(5000)}catch{}
        }
        if ($Global:ws) { try{$Global:ws.Dispose()}catch{}; $Global:ws=$null }
        if ($edgeProc -and -not $edgeProc.HasExited) { try{$edgeProc.Kill()}catch{} }
        if ($excel) { try{$excel.Quit()}catch{}; try{[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}catch{} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $btnRun.Enabled = $true
    }
})
[void]$form.ShowDialog()
