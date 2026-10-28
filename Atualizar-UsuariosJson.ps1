# ==============================================================================
# EXPORTADOR DO ACTIVE DIRECTORY PARA O USERS.JSON DO SUPLEMENTO
# Executar este script no servidor do AD ou estação de administração
# ==============================================================================

$outputFile = Join-Path $PSScriptRoot "users.json"
Write-Host "Consultando usuarios ativos no Active Directory..." -ForegroundColor Cyan

Import-Module ActiveDirectory -ErrorAction Stop

# Buscar todos os usuários ativos com e-mail corporativo
$adUsers = Get-ADUser -Filter "Enabled -eq '$true' -and Mail -like '*@*'" `
    -Properties Title, Department, StreetAddress, City, TelephoneNumber, Mobile, Pager, Mail, UserPrincipalName

Write-Host "Total de usuarios encontrados: $($adUsers.Count)" -ForegroundColor Cyan

$catalog = [ordered]@{}

foreach ($u in $adUsers) {
    # Determina o e-mail chave (em minúsculas)
    $email = if ($u.Mail) { $u.Mail.ToLower().Trim() } else { $u.UserPrincipalName.ToLower().Trim() }
    
    $catalog[$email] = [ordered]@{
        name       = $u.Name
        title      = if ($u.Title) { $u.Title.Trim() } else { "" }
        department = if ($u.Department) { $u.Department.Trim() } else { "" }
        address    = if ($u.StreetAddress) { $u.StreetAddress.Trim() } else { "" }
        city       = if ($u.City) { $u.City.Trim() } else { "" }
        phone      = if ($u.TelephoneNumber) { $u.TelephoneNumber.Trim() } else { "" }
        pager      = if ($u.Pager) { $u.Pager.Trim() } else { "" }
        mobile     = if ($u.Mobile) { $u.Mobile.Trim() } else { "" }
    }
}

$catalog | ConvertTo-Json -Depth 3 | Out-File -FilePath $outputFile -Encoding utf8
Write-Host "Catalogo exportado com sucesso em: $outputFile" -ForegroundColor Green
