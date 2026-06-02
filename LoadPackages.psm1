Function Load-RelatedAssemblies {
    param (
        [string] $directory = '..\',
        [string] $assemblyPath
    )
    
    $assembly = [System.Reflection.Assembly]::LoadFrom($assemblyPath)
    foreach ($referencedAssembly in $assembly.GetReferencedAssemblies()) {
        try {
            [System.Reflection.Assembly]::Load($referencedAssembly)
        }
        catch {
            #Assembly not found in GAC or current directory, attempt to load from bin folder
            Write-Host "Failed to load assembly: $($referencedAssembly.FullName)" -ForegroundColor Red
            $foundPaths = Get-ChildItem "$directory\bin\$($referencedAssembly.Name).*\lib\net462\$($referencedAssembly.Name).dll"
            $firstFoundPath = $foundPaths | Select-Object -First 1
            if ($firstFoundPath) {
                Write-Host "Attempting to load from: $($firstFoundPath.FullName)" -ForegroundColor Yellow
                [System.Reflection.Assembly]::LoadFrom($firstFoundPath.FullName)
                #Add-Type -Path $firstFoundPath.FullName
            }
            else {
                Write-Host "Could not find assembly in bin folder: $($referencedAssembly.Name).dll" -ForegroundColor Red
            }
        }
    }
}


Function Load-Packages {
    param ([string] $directory = '.\')
    
    try {
        if(-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Mail")) {
            Install-Module -Name "Microsoft.Graph.Mail" -Scope AllUsers -Force -Confirm:$false
            Import-Module -Name Microsoft.Graph.Mail
        }
        if(-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Users.Actions")) {
            Install-Module -Name "Microsoft.Graph.Users.Actions" -Scope AllUsers -Force -Confirm:$false
            Import-Module -Name Microsoft.Graph.Users.Actions
        }
        if(-not (Get-Module -ListAvailable -Name "MSAL.PS")) {
            Install-Module -Name "MSAL.PS" -Scope AllUsers -Force -Confirm:$false
            Import-Module -Name MSAL.PS
        }
        if(-not (Get-Module -ListAvailable -Name "Logging")) {
            Install-Module -Name "Logging" -Scope AllUsers -Force -Confirm:$false
            Import-Module -Name Logging
        }
        if($null -eq (Get-Module -Name "Microsoft.Xrm.Tooling.CrmConnector.PowerShell")) {
            Install-Module -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell -Force -Confirm:$false
        }
        # Install nuget if not present
        if (!(Get-Command 'nuget.exe' -ErrorAction SilentlyContinue) -and !(get-command '.\nuget.exe' -ErrorAction SilentlyContinue)) {
            Invoke-WebRequest -Uri 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -OutFile 'nuget.exe'
        }
        # Install the package
        & nuget.exe install Microsoft.CrmSdk.CoreAssemblies -OutputDirectory "$directory\bin"
        
        # Load Microsoft.Xrm.Sdk.dll
        $assemblyPath = Get-ChildItem -Path "$directory\bin\Microsoft.CrmSdk.CoreAssemblies.*\lib\net462\Microsoft.Xrm.Sdk.dll" | Select-Object -First 1
        Load-RelatedAssemblies -directory $directory -assemblyPath $assemblyPath.FullName
        Add-Type -Path $assemblyPath.FullName
        
        # Load Microsoft.Crm.Sdk.Proxy.dll
        $assemblyPath = Get-ChildItem -Path "$directory\bin\Microsoft.CrmSdk.CoreAssemblies.*\lib\net462\Microsoft.Crm.Sdk.Proxy.dll" | Select-Object -First 1
        Load-RelatedAssemblies -directory $directory -assemblyPath $assemblyPath.FullName
        Add-Type -Path $assemblyPath.FullName
    }
    catch [System.Reflection.ReflectionTypeLoadException] {
        foreach ($loaderException in $_.Exception.LoaderExceptions)
        {
            Write-Host $loaderException.Message -ForegroundColor Red
        }
        Write-Host "Message: $($_.Exception.Message)"
        Write-Host "StackTrace: $($_.Exception.StackTrace)"
        Write-Host "LoaderExceptions: $($_.Exception.LoaderExceptions)"
    }
}

Export-ModuleMember -Function Load-Packages
