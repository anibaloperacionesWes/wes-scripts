# Script para crear usuario en WES API
# Uso: .\crear_usuario.ps1 -NombreEmpresa "Lo valledor" -Nombre "Jose Luis" -Apellido "Otarola" -Email "joseluisricardo@hotmail.com"
# Uso con ID: .\crear_usuario.ps1 -NombreEmpresa "Lo valledor" -CompanyId "000002" -Nombre "Jose Luis" -Apellido "Otarola" -Email "joseluisricardo@hotmail.com"

param(
    [Parameter(Mandatory=$true)]
    [string]$NombreEmpresa,
    
    [Parameter(Mandatory=$false)]
    [string]$CompanyId,
    
    [Parameter(Mandatory=$true)]
    [string]$Nombre,
    
    [Parameter(Mandatory=$true)]
    [string]$Apellido,
    
    [Parameter(Mandatory=$true)]
    [string]$Email
)

$baseUrl = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
$headers = @{
    "Accept" = "application/json"
    "Content-Type" = "application/json"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  CREACION DE USUARIO EN WES API" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Nombre Empresa: $NombreEmpresa" -ForegroundColor White
if ($CompanyId) {
    Write-Host "  Compania ID: $CompanyId" -ForegroundColor White
}
Write-Host "  Nombre: $Nombre $Apellido" -ForegroundColor White
Write-Host "  Email: $Email" -ForegroundColor White
Write-Host ""

try {
    $foundCompanyId = $null
    $companyResponse = $null
    
    # Si no se proporciona CompanyId, buscar por nombre
    if (-not $CompanyId) {
        Write-Host "[0/3] Buscando empresa por nombre: '$NombreEmpresa'..." -ForegroundColor Cyan
        Write-Host "  Buscando desde 000001 hasta 000100..." -ForegroundColor Gray
        
        for ($i = 1; $i -le 100; $i++) {
            $testCompanyId = $i.ToString("000000")
            
            try {
                $testResponse = Invoke-RestMethod -Uri "$baseUrl/companies/$testCompanyId" `
                    -Method Get `
                    -Headers $headers `
                    -ErrorAction Stop
                
                # Comparar nombre (case-insensitive)
                if ($testResponse.name -and $testResponse.name.Trim() -eq $NombreEmpresa.Trim()) {
                    $foundCompanyId = $testCompanyId
                    $companyResponse = $testResponse
                    Write-Host "  [OK] Empresa encontrada: ID $foundCompanyId" -ForegroundColor Green
                    break
                }
            } catch {
                # Empresa no existe o error 404, continuar buscando
                continue
            }
        }
        
        if (-not $foundCompanyId) {
            Write-Host ""
            Write-Host "  [ERROR] No se encontro la empresa '$NombreEmpresa' en el rango 000001-000100" -ForegroundColor Red
            Write-Host "  Por favor, proporciona el CompanyId manualmente" -ForegroundColor Yellow
            exit 1
        }
        
        $CompanyId = $foundCompanyId
        Write-Host ""
    } else {
        # Si se proporciona CompanyId, validar que existe y coincide con el nombre
        Write-Host "[0/3] Validando empresa..." -ForegroundColor Cyan
        
        $companyResponse = Invoke-RestMethod -Uri "$baseUrl/companies/$CompanyId" `
            -Method Get `
            -Headers $headers
        
        # Verificar que el nombre coincide
        if ($companyResponse.name.Trim() -ne $NombreEmpresa.Trim()) {
            Write-Host "  [ADVERTENCIA] El nombre de la empresa no coincide" -ForegroundColor Yellow
            Write-Host "    Esperado: '$NombreEmpresa'" -ForegroundColor Yellow
            Write-Host "    Encontrado: '$($companyResponse.name)'" -ForegroundColor Yellow
            Write-Host "    Continuando con el ID proporcionado..." -ForegroundColor Yellow
        } else {
            Write-Host "  [OK] Empresa validada correctamente" -ForegroundColor Green
        }
        Write-Host ""
    }
    
    Write-Host "[1/3] Obteniendo nodos de la compania $CompanyId ($($companyResponse.name))..." -ForegroundColor Cyan
    
    Write-Host "  Compania: $($companyResponse.name)" -ForegroundColor Green
    Write-Host "  Nodos encontrados: $($companyResponse.nodes.Count)" -ForegroundColor Green
    
    $allowedNodes = @()
    foreach ($node in $companyResponse.nodes) {
        $allowedNodes += $node.nodeId
        Write-Host "    - $($node.nodeId): $($node.name)" -ForegroundColor Yellow
    }
    
    if ($allowedNodes.Count -eq 0) {
        Write-Host "  [ADVERTENCIA] La compania no tiene nodos" -ForegroundColor Yellow
    }
    
    Write-Host ""
    
    Write-Host "[2/3] Creando usuario..." -ForegroundColor Cyan
    
    $userData = @{
        username = $Email
        name = $Nombre
        lastName = $Apellido
        companyId = $CompanyId
        allowedNodes = $allowedNodes
    }
    
    $jsonBody = $userData | ConvertTo-Json -Depth 10 -Compress
    
    $response = Invoke-WebRequest -Uri "$baseUrl/configuration/users" `
        -Method Post `
        -Headers @{
            "Accept" = "*/*"
            "Content-Type" = "application/json"
        } `
        -Body $jsonBody `
        -UseBasicParsing
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  USUARIO CREADO EXITOSAMENTE" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  Email: $Email" -ForegroundColor Cyan
    Write-Host "  Nombre: $Nombre $Apellido" -ForegroundColor Cyan
    Write-Host "  Compania: $($companyResponse.name) (ID: $CompanyId)" -ForegroundColor Cyan
    Write-Host "  Nodos asignados: $($allowedNodes.Count)" -ForegroundColor Cyan
    Write-Host "  Respuesta: $($response.Content)" -ForegroundColor Green
    Write-Host ""
    
} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "  ERROR AL CREAR USUARIO" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "  Mensaje: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.Exception.Response) {
        $statusCode = $_.Exception.Response.StatusCode.value__
        Write-Host "  Status Code: $statusCode" -ForegroundColor Red
        
        try {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $responseBody = $reader.ReadToEnd()
            if ($responseBody) {
                Write-Host "  Detalles: $responseBody" -ForegroundColor Red
            }
        } catch {
        }
    }
    
    Write-Host ""
    exit 1
}
