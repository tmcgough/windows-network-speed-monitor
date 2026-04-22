param(
    [int]$RefreshSeconds = 1
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# Script-scope state is used so the timer tick, button click handlers, and
# grid callbacks can all read/write the same live monitor data.
$script:hostCache = @{}
$script:processCache = @{}
$script:connectionHistory = @{}
$script:displayUnit = "Mbps"
$script:stats = @{
    InMin        = $null
    InMax        = 0.0
    OutMin       = $null
    OutMax       = 0.0
    CombinedMin  = $null
    CombinedMax  = 0.0
}
$script:gridSortState = @{
    Adapter = @{
        Column    = $null
        Direction = [System.ComponentModel.ListSortDirection]::Ascending
    }
    Connections = @{
        Column    = $null
        Direction = [System.ComponentModel.ListSortDirection]::Ascending
    }
}
$script:lastDisplayState = $null
$script:lastAdapterRows = @()
# Ignore adapters that either are not real network paths for user traffic or
# tend to add noise to the dashboard.
$script:adapterPatternsToIgnore = @(
    '^Loopback',
    'isatap',
    'Teredo',
    '6to4',
    'Pseudo-Interface',
    'Bluetooth'
)

function Convert-ToMbps {
    param([double]$BytesPerSecond)

    return [math]::Round(($BytesPerSecond * 8) / 1MB, 2)
}

function Convert-ToMBps {
    param([double]$BytesPerSecond)

    return [math]::Round($BytesPerSecond / 1MB, 2)
}

function Get-SpeedDisplayValue {
    param([double]$BytesPerSecond)

    # The monitor always samples in bytes/second internally, then converts only
    # for display so unit toggling does not affect stored stats.
    if ($script:displayUnit -ceq "MBps") {
        return Convert-ToMBps -BytesPerSecond $BytesPerSecond
    }

    return Convert-ToMbps -BytesPerSecond $BytesPerSecond
}

function Format-SpeedText {
    param([double]$BytesPerSecond)

    $value = Get-SpeedDisplayValue -BytesPerSecond $BytesPerSecond
    return ("{0:N2} {1}" -f $value, $script:displayUnit)
}

function Format-BytesText {
    param([double]$Bytes)

    if ($Bytes -ge 1TB) { return ("{0:N2} TB" -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    return ("{0:N0} B" -f $Bytes)
}

function Get-ActiveInterfaceRows {
    $rows = @()

    try {
        $stats = Get-NetAdapterStatistics -ErrorAction Stop
    }
    catch {
        return @()
    }

    foreach ($stat in $stats) {
        if ($stat.Name -match ($script:adapterPatternsToIgnore -join '|')) {
            continue
        }

        # Get-NetAdapterStatistics can return entries that are not currently up.
        # Pair the byte counters with the adapter status so the dashboard only
        # shows adapters that are actively available.
        $linkSpeedText = ""
        try {
            $adapter = Get-NetAdapter -Name $stat.Name -ErrorAction Stop
            if ($adapter.Status -ne 'Up') {
                continue
            }

            $linkSpeedText = $adapter.LinkSpeed
        }
        catch {
            continue
        }

        $rows += [pscustomobject]@{
            Name          = $stat.Name
            SentBytes     = [double]$stat.SentBytes
            ReceivedBytes = [double]$stat.ReceivedBytes
            LinkSpeed     = $linkSpeedText
        }
    }

    return $rows
}

function Get-ResolvedHost {
    param([string]$IpAddress)

    if ([string]::IsNullOrWhiteSpace($IpAddress)) {
        return ""
    }

    if ($script:hostCache.ContainsKey($IpAddress)) {
        return $script:hostCache[$IpAddress]
    }

    # Reverse DNS can be slow or unavailable. Cache both hits and misses so the
    # connections grid does not repeatedly block on the same remote address.
    $result = ""
    try {
        $entry = [System.Net.Dns]::GetHostEntry($IpAddress)
        if ($entry.HostName) {
            $result = $entry.HostName
        }
    }
    catch {
        $result = ""
    }

    $script:hostCache[$IpAddress] = $result
    return $result
}

function Get-ProcessNameCached {
    param([uint32]$ProcessId)

    if ($ProcessId -le 0) {
        return ""
    }

    if ($script:processCache.ContainsKey($ProcessId)) {
        return $script:processCache[$ProcessId]
    }

    # Process lookup is repeated across refreshes, so cache the friendly name
    # by PID until the app closes.
    $name = ""
    try {
        $name = (Get-Process -Id $ProcessId -ErrorAction Stop).ProcessName
    }
    catch {
        $name = ""
    }

    $script:processCache[$ProcessId] = $name
    return $name
}

function Get-ConnectionKey {
    param($Connection)

    return "{0}|{1}|{2}|{3}|{4}" -f `
        $Connection.OwningProcess,
        $Connection.LocalAddress,
        $Connection.LocalPort,
        $Connection.RemoteAddress,
        $Connection.RemotePort
}

function Update-ConnectionHistory {
    param($Connections, [datetime]$Timestamp)

    $activeKeys = @{}

    foreach ($connection in $Connections) {
        $key = Get-ConnectionKey -Connection $connection
        $activeKeys[$key] = $true

        if (-not $script:connectionHistory.ContainsKey($key)) {
            $script:connectionHistory[$key] = @{
                FirstSeen = $Timestamp
                LastSeen  = $Timestamp
            }
        }
        else {
            $script:connectionHistory[$key].LastSeen = $Timestamp
        }
    }

    # Drop stale entries after a while so the in-memory table reflects current
    # app usage instead of growing forever during long sessions.
    $staleKeys = @()
    foreach ($key in $script:connectionHistory.Keys) {
        if ($activeKeys.ContainsKey($key)) {
            continue
        }

        if (($Timestamp - $script:connectionHistory[$key].LastSeen).TotalMinutes -gt 30) {
            $staleKeys += $key
        }
    }

    foreach ($key in $staleKeys) {
        $script:connectionHistory.Remove($key)
    }
}

function New-Label {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [float]$Size = 10,
        [bool]$Bold = $false
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, $Height)
    $fontStyle = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $label.Font = New-Object System.Drawing.Font("Segoe UI", $Size, $fontStyle)
    $label.Text = $Text
    return $label
}

function Apply-GridSort {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [hashtable]$State
    )

    if (-not $State.Column) {
        return
    }

    if (-not $Grid.Columns.Contains($State.Column)) {
        return
    }

    $sortColumnName = $State.Column
    if ($Grid.Name -eq "adapterGrid") {
        # Speed columns are displayed as formatted text, so the adapter table
        # keeps hidden numeric columns for accurate sorting.
        switch ($State.Column) {
            "Incoming" { $sortColumnName = "IncomingValue" }
            "Outgoing" { $sortColumnName = "OutgoingValue" }
        }
    }

    $column = $Grid.Columns[$sortColumnName]
    $Grid.Sort($column, $State.Direction)
    $glyphColumn = $Grid.Columns[$State.Column]
    $glyphColumn.HeaderCell.SortGlyphDirection = if ($State.Direction -eq [System.ComponentModel.ListSortDirection]::Ascending) {
        [System.Windows.Forms.SortOrder]::Ascending
    }
    else {
        [System.Windows.Forms.SortOrder]::Descending
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Network Speed Monitor"
$form.Size = New-Object System.Drawing.Size(1180, 760)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$title = New-Label -Text "Live Network Monitor" -X 20 -Y 15 -Width 300 -Height 30 -Size 16 -Bold $true
$subtitle = New-Label -Text "Shows current, min, max, and total traffic across active adapters plus live TCP connections." -X 20 -Y 45 -Width 820 -Height 22

$inLabel = New-Label -Text "Incoming" -X 20 -Y 90 -Width 200 -Height 20 -Bold $true
$outLabel = New-Label -Text "Outgoing" -X 310 -Y 90 -Width 200 -Height 20 -Bold $true
$combinedLabel = New-Label -Text "Combined" -X 600 -Y 90 -Width 200 -Height 20 -Bold $true
$statusLabel = New-Label -Text "Sampling..." -X 900 -Y 20 -Width 240 -Height 20
$noteLabel = New-Label -Text "Hostname is best effort from reverse DNS. Exact site/page URLs are not available from normal Windows socket data." -X 20 -Y 670 -Width 1120 -Height 40
$unitLabel = New-Label -Text "Display Unit" -X 900 -Y 45 -Width 90 -Height 22

$unitSelector = New-Object System.Windows.Forms.ComboBox
$unitSelector.Location = New-Object System.Drawing.Point(990, 41)
$unitSelector.Size = New-Object System.Drawing.Size(150, 28)
$unitSelector.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$unitSelector.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$unitSelector.Items.Add("Mbps")
[void]$unitSelector.Items.Add("MBps")
$unitSelector.SelectedItem = "Mbps"

$inCurrent = New-Label -Text "Current: -" -X 20 -Y 120 -Width 240 -Height 28 -Size 12
$inMin = New-Label -Text "Min: -" -X 20 -Y 150 -Width 240 -Height 24
$inMax = New-Label -Text "Max: -" -X 20 -Y 176 -Width 240 -Height 24
$inTotal = New-Label -Text "Total Received: -" -X 20 -Y 202 -Width 240 -Height 24

$outCurrent = New-Label -Text "Current: -" -X 310 -Y 120 -Width 240 -Height 28 -Size 12
$outMin = New-Label -Text "Min: -" -X 310 -Y 150 -Width 240 -Height 24
$outMax = New-Label -Text "Max: -" -X 310 -Y 176 -Width 240 -Height 24
$outTotal = New-Label -Text "Total Sent: -" -X 310 -Y 202 -Width 240 -Height 24

$combinedCurrent = New-Label -Text "Current: -" -X 600 -Y 120 -Width 240 -Height 28 -Size 12
$combinedMin = New-Label -Text "Min: -" -X 600 -Y 150 -Width 240 -Height 24
$combinedMax = New-Label -Text "Max: -" -X 600 -Y 176 -Width 240 -Height 24
$activeAdapters = New-Label -Text "Adapters: -" -X 600 -Y 202 -Width 520 -Height 24

$adapterTitle = New-Label -Text "Active Adapters" -X 20 -Y 250 -Width 220 -Height 20 -Bold $true
$connectionsTitle = New-Label -Text "Established Connections" -X 20 -Y 440 -Width 260 -Height 20 -Bold $true

$adapterGrid = New-Object System.Windows.Forms.DataGridView
$adapterGrid.Location = New-Object System.Drawing.Point(20, 275)
$adapterGrid.Size = New-Object System.Drawing.Size(1120, 145)
$adapterGrid.ReadOnly = $true
$adapterGrid.AllowUserToAddRows = $false
$adapterGrid.AllowUserToDeleteRows = $false
$adapterGrid.RowHeadersVisible = $false
$adapterGrid.AutoSizeColumnsMode = "Fill"
$adapterGrid.SelectionMode = "FullRowSelect"
$adapterGrid.BackgroundColor = [System.Drawing.Color]::White
$adapterGrid.Name = "adapterGrid"
$adapterGrid.DataSource = [System.Data.DataTable]::new()

$connectionsGrid = New-Object System.Windows.Forms.DataGridView
$connectionsGrid.Location = New-Object System.Drawing.Point(20, 465)
$connectionsGrid.Size = New-Object System.Drawing.Size(1120, 190)
$connectionsGrid.ReadOnly = $true
$connectionsGrid.AllowUserToAddRows = $false
$connectionsGrid.AllowUserToDeleteRows = $false
$connectionsGrid.RowHeadersVisible = $false
$connectionsGrid.AutoSizeColumnsMode = "Fill"
$connectionsGrid.SelectionMode = "FullRowSelect"
$connectionsGrid.BackgroundColor = [System.Drawing.Color]::White
$connectionsGrid.Name = "connectionsGrid"
$connectionsGrid.DataSource = [System.Data.DataTable]::new()

foreach ($control in @(
    $title, $subtitle, $inLabel, $outLabel, $combinedLabel, $statusLabel, $noteLabel,
    $unitLabel, $unitSelector,
    $inCurrent, $inMin, $inMax, $inTotal, $outCurrent, $outMin, $outMax, $outTotal,
    $combinedCurrent, $combinedMin, $combinedMax, $activeAdapters,
    $adapterTitle, $connectionsTitle, $adapterGrid, $connectionsGrid
)) {
    [void]$form.Controls.Add($control)
}

$script:previousSample = $null

function Update-AdapterGrid {
    param($Rows, [double]$IntervalSeconds)

    # Rebuild a fresh table each refresh. This keeps the grid simple, and the
    # saved sort state is reapplied immediately after rebinding.
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add("Adapter", [string])
    [void]$table.Columns.Add("Incoming", [string])
    [void]$table.Columns.Add("Outgoing", [string])
    [void]$table.Columns.Add("Link Speed", [string])
    [void]$table.Columns.Add("IncomingValue", [double])
    [void]$table.Columns.Add("OutgoingValue", [double])

    foreach ($row in $Rows) {
        $incoming = "-"
        $outgoing = "-"
        $incomingValue = 0.0
        $outgoingValue = 0.0

        if ($script:previousSample -and $script:previousSample.ContainsKey($row.Name)) {
            # Adapter statistics are cumulative counters. Current rate is the
            # delta between the latest and previous sample divided by the refresh interval.
            $previous = $script:previousSample[$row.Name]
            $incomingRate = [math]::Max(0.0, ($row.ReceivedBytes - $previous.ReceivedBytes) / $IntervalSeconds)
            $outgoingRate = [math]::Max(0.0, ($row.SentBytes - $previous.SentBytes) / $IntervalSeconds)
            $incomingValue = Get-SpeedDisplayValue -BytesPerSecond $incomingRate
            $outgoingValue = Get-SpeedDisplayValue -BytesPerSecond $outgoingRate
            $incoming = Format-SpeedText -BytesPerSecond $incomingRate
            $outgoing = Format-SpeedText -BytesPerSecond $outgoingRate
        }

        [void]$table.Rows.Add($row.Name, $incoming, $outgoing, $row.LinkSpeed, $incomingValue, $outgoingValue)
    }

    $adapterGrid.DataSource = $table
    # Hidden numeric columns back the visible formatted text columns so sorting
    # stays numeric in both Mbps and MBps modes.
    $adapterGrid.Columns["IncomingValue"].Visible = $false
    $adapterGrid.Columns["OutgoingValue"].Visible = $false
    foreach ($column in $adapterGrid.Columns) {
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
    }
    Apply-GridSort -Grid $adapterGrid -State $script:gridSortState.Adapter
}

function Update-ConnectionsGrid {
    param([datetime]$Timestamp)

    # The connections table is separate from the speed sampling logic: it shows
    # the latest established TCP connections, not traffic volume per connection.
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add("Process", [string])
    [void]$table.Columns.Add("PID", [int])
    [void]$table.Columns.Add("Local", [string])
    [void]$table.Columns.Add("Remote IP", [string])
    [void]$table.Columns.Add("Remote Host", [string])
    [void]$table.Columns.Add("Remote Port", [int])
    [void]$table.Columns.Add("First Seen", [datetime])
    [void]$table.Columns.Add("Last Seen", [datetime])

    try {
        $connections = Get-NetTCPConnection -State Established -ErrorAction Stop |
            Sort-Object -Property OwningProcess, RemoteAddress, RemotePort -Unique
    }
    catch {
        $connections = @()
    }

    Update-ConnectionHistory -Connections $connections -Timestamp $Timestamp

    # Limit the table size so refresh stays responsive even on busy systems.
    $shown = 0
    foreach ($connection in $connections) {
        if ($connection.RemoteAddress -in @("127.0.0.1", "::1", "0.0.0.0", "::")) {
            continue
        }

        $processName = Get-ProcessNameCached -ProcessId $connection.OwningProcess
        $remoteHost = Get-ResolvedHost -IpAddress $connection.RemoteAddress
        $localAddress = "{0}:{1}" -f $connection.LocalAddress, $connection.LocalPort
        $historyKey = Get-ConnectionKey -Connection $connection
        $history = $script:connectionHistory[$historyKey]

        [void]$table.Rows.Add(
            $processName,
            [int]$connection.OwningProcess,
            $localAddress,
            $connection.RemoteAddress,
            $remoteHost,
            [int]$connection.RemotePort,
            [datetime]$history.FirstSeen,
            [datetime]$history.LastSeen
        )

        $shown++
        if ($shown -ge 50) {
            break
        }
    }

    $connectionsGrid.DataSource = $table
    $connectionsGrid.Columns["First Seen"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss"
    $connectionsGrid.Columns["Last Seen"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss"
    foreach ($column in $connectionsGrid.Columns) {
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
    }
    Apply-GridSort -Grid $connectionsGrid -State $script:gridSortState.Connections
}

function Update-DisplayLabels {
    if (-not $script:lastDisplayState) {
        return
    }

    # All labels are derived from the most recent computed sample so a unit
    # toggle can redraw text instantly without waiting for the next timer tick.
    $state = $script:lastDisplayState
    $unitLabel.Text = "Display Unit"
    if ($unitSelector.SelectedItem -ne $script:displayUnit) {
        $unitSelector.SelectedItem = $script:displayUnit
    }

    $inCurrent.Text = "Current: $(Format-SpeedText -BytesPerSecond $state.ReceivedDelta)"
    $inMin.Text = "Min: $(Format-SpeedText -BytesPerSecond $script:stats.InMin)"
    $inMax.Text = "Max: $(Format-SpeedText -BytesPerSecond $script:stats.InMax)"
    $inTotal.Text = "Total Received: $(Format-BytesText -Bytes $state.TotalReceived)"

    $outCurrent.Text = "Current: $(Format-SpeedText -BytesPerSecond $state.SentDelta)"
    $outMin.Text = "Min: $(Format-SpeedText -BytesPerSecond $script:stats.OutMin)"
    $outMax.Text = "Max: $(Format-SpeedText -BytesPerSecond $script:stats.OutMax)"
    $outTotal.Text = "Total Sent: $(Format-BytesText -Bytes $state.TotalSent)"

    $combinedCurrent.Text = "Current: $(Format-SpeedText -BytesPerSecond $state.CombinedDelta)"
    $combinedMin.Text = "Min: $(Format-SpeedText -BytesPerSecond $script:stats.CombinedMin)"
    $combinedMax.Text = "Max: $(Format-SpeedText -BytesPerSecond $script:stats.CombinedMax)"
    $activeAdapters.Text = "Adapters: $($state.AdapterNames -join ', ')"
    $statusLabel.Text = "Last update: $($state.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))"
}

function Refresh-Monitor {
    $rows = Get-ActiveInterfaceRows
    $now = Get-Date

    if (-not $rows -or $rows.Count -eq 0) {
        $statusLabel.Text = "No active adapters found"
        return
    }

    $currentMap = @{}
    foreach ($row in $rows) {
        $currentMap[$row.Name] = $row
    }

    if (-not $script:previousSample) {
        # The first refresh establishes the baseline counters. A speed value is
        # only meaningful once we have two samples to compare.
        $script:previousSample = $currentMap
        $statusLabel.Text = "Collecting baseline..."
        Update-AdapterGrid -Rows $rows -IntervalSeconds $RefreshSeconds
        Update-ConnectionsGrid -Timestamp $now
        return
    }

    $totalReceived = 0.0
    $totalSent = 0.0
    $receivedDelta = 0.0
    $sentDelta = 0.0

    foreach ($row in $rows) {
        $totalReceived += $row.ReceivedBytes
        $totalSent += $row.SentBytes

        if ($script:previousSample.ContainsKey($row.Name)) {
            $previous = $script:previousSample[$row.Name]
            # Clamp negative deltas to zero in case an adapter resets its
            # counters or briefly disappears between samples.
            $receivedDelta += [math]::Max(0.0, ($row.ReceivedBytes - $previous.ReceivedBytes) / $RefreshSeconds)
            $sentDelta += [math]::Max(0.0, ($row.SentBytes - $previous.SentBytes) / $RefreshSeconds)
        }
    }

    $combinedDelta = $receivedDelta + $sentDelta

    if ($null -eq $script:stats.InMin -or $receivedDelta -lt $script:stats.InMin) {
        $script:stats.InMin = $receivedDelta
    }
    if ($receivedDelta -gt $script:stats.InMax) {
        $script:stats.InMax = $receivedDelta
    }

    if ($null -eq $script:stats.OutMin -or $sentDelta -lt $script:stats.OutMin) {
        $script:stats.OutMin = $sentDelta
    }
    if ($sentDelta -gt $script:stats.OutMax) {
        $script:stats.OutMax = $sentDelta
    }

    if ($null -eq $script:stats.CombinedMin -or $combinedDelta -lt $script:stats.CombinedMin) {
        $script:stats.CombinedMin = $combinedDelta
    }
    if ($combinedDelta -gt $script:stats.CombinedMax) {
        $script:stats.CombinedMax = $combinedDelta
    }

    $script:lastDisplayState = @{
        ReceivedDelta = $receivedDelta
        SentDelta     = $sentDelta
        CombinedDelta = $combinedDelta
        TotalReceived = $totalReceived
        TotalSent     = $totalSent
        AdapterNames  = @($rows.Name)
        Timestamp     = $now
    }
    $script:lastAdapterRows = @($rows)

    Update-DisplayLabels

    Update-AdapterGrid -Rows $rows -IntervalSeconds $RefreshSeconds
    Update-ConnectionsGrid -Timestamp $now

    # Store the full current sample so the next timer tick can compute rates
    # from cumulative adapter byte counters.
    $script:previousSample = $currentMap
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [math]::Max(500, $RefreshSeconds * 1000)
$timer.Add_Tick({
    Refresh-Monitor
})

$adapterGrid.Add_ColumnHeaderMouseClick({
    param($sender, $eventArgs)

    # Programmatic sorting lets the app keep user-selected sort order across
    # data refreshes instead of losing it whenever the grid is rebound.
    $clickedColumn = $adapterGrid.Columns[$eventArgs.ColumnIndex]
    $newDirection = [System.ComponentModel.ListSortDirection]::Ascending
    if ($script:gridSortState.Adapter.Column -eq $clickedColumn.Name -and $script:gridSortState.Adapter.Direction -eq [System.ComponentModel.ListSortDirection]::Ascending) {
        $newDirection = [System.ComponentModel.ListSortDirection]::Descending
    }

    foreach ($column in $adapterGrid.Columns) {
        $column.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
    }

    $script:gridSortState.Adapter.Column = $clickedColumn.Name
    $script:gridSortState.Adapter.Direction = $newDirection
    Apply-GridSort -Grid $adapterGrid -State $script:gridSortState.Adapter
})

$connectionsGrid.Add_ColumnHeaderMouseClick({
    param($sender, $eventArgs)

    $clickedColumn = $connectionsGrid.Columns[$eventArgs.ColumnIndex]
    $newDirection = [System.ComponentModel.ListSortDirection]::Ascending
    if ($script:gridSortState.Connections.Column -eq $clickedColumn.Name -and $script:gridSortState.Connections.Direction -eq [System.ComponentModel.ListSortDirection]::Ascending) {
        $newDirection = [System.ComponentModel.ListSortDirection]::Descending
    }

    foreach ($column in $connectionsGrid.Columns) {
        $column.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
    }

    $script:gridSortState.Connections.Column = $clickedColumn.Name
    $script:gridSortState.Connections.Direction = $newDirection
    Apply-GridSort -Grid $connectionsGrid -State $script:gridSortState.Connections
})

$unitSelector.Add_SelectedIndexChanged({
    if (-not $unitSelector.SelectedItem) {
        return
    }

    # Unit changes only affect formatting. The underlying sampled values remain
    # in bytes/second so min/max and sorting stay internally consistent.
    $script:displayUnit = [string]$unitSelector.SelectedItem
    Update-DisplayLabels
    if ($script:lastDisplayState -and $script:lastAdapterRows.Count -gt 0) {
        Update-AdapterGrid -Rows $script:lastAdapterRows -IntervalSeconds $RefreshSeconds
    }
})

$form.Add_Shown({
    Refresh-Monitor
    $timer.Start()
})

$form.Add_FormClosing({
    $timer.Stop()
})

[void]$form.ShowDialog()
