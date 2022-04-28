function Install-TryModule ($moduleName) {
    try {
        $status = Get-Module -ListAvailable -Name $moduleName
        if (!$status) {
            Write-Host "Installing $moduleName" -ForegroundColor Yellow
            Install-Module -Name $moduleName -Scope CurrentUser -Force
        }
    } catch {
        # If anything goes wrong, just install the damn thing anyway
        Write-Host "Installing module $moduleName" -ForegroundColor Yellow
        Install-Module -Name $moduleName -Scope CurrentUser -Force
    }

}