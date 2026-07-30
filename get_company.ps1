# Comando curl para obtener la compañía por ID (PowerShell)
# ID: 000002
# Endpoint: retrieveCompanyUsingGET_1

$headers = @{
    "Accept" = "application/json"
    "Content-Type" = "application/json"
}

# Si requiere autenticación, descomenta y agrega:
# $headers["Authorization"] = "Bearer TU_TOKEN_AQUI"

Invoke-RestMethod -Uri "http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/000002" `
    -Method Get `
    -Headers $headers

# Alternativa usando curl.exe directamente:
# curl.exe -X GET "http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/000002" `
#   -H "Accept: application/json" `
#   -H "Content-Type: application/json" `
#   -v


