try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace('MAPI')
    $folder = $namespace.GetDefaultFolder(9)
    $items = $folder.Items
    $items.Sort('[Start]')
    $items.IncludeRecurrences = $true

    $startDate = '11/18/2025'
    $endDate = '11/30/2025 23:59:59'
    $filter = "[Start] >= '$startDate' AND [End] <= '$endDate'"

    Write-Host "Filter: $filter"

    $restricted = $items.Restrict($filter)
    Write-Host "Found items: $($restricted.Count)"

    foreach ($item in $restricted) {
        if ($item.Class -eq 26) {
            Write-Host "Event: $($item.Subject) - Start: $($item.Start) - End: $($item.End)"
        }
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)"
    Write-Host $_.Exception.StackTrace
}
