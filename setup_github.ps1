<#
.SYNOPSIS
    Script automatizado para inicializar y publicar el repositorio en GitHub.
.DESCRIPTION
    1. Verifica e inicializa el repositorio local Git en la rama 'main'.
    2. Comprueba que .gitignore proteja .env y archivos pesados.
    3. Agrega los archivos de código y configuración al staging.
    4. Solicita o configura la URL remota de GitHub.
    5. Realiza el commit inicial y el push hacia GitHub.
#>

param(
    [string]$RemoteUrl = "",
    [string]$CommitMessage = "feat: pipeline de auditoria contable automatizada Numeral 22.3 c"
)

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   CONFIGURACIÓN Y PUBLICACIÓN DEL PROYECTO EN GITHUB       " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# 1. Comprobar instalación de Git
try {
    $gitVer = git --version
    Write-Host "[OK] Git detectado: $gitVer" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Git no está instalado o no se encuentra en el PATH." -ForegroundColor Red
    exit 1
}

# 2. Inicializar repositorio si no existe
if (-not (Test-Path ".git")) {
    Write-Host "[INFO] Inicializando repositorio Git local..." -ForegroundColor Yellow
    git init
    git branch -M main
    Write-Host "[OK] Repositorio Git inicializado en rama 'main'." -ForegroundColor Green
} else {
    Write-Host "[OK] Repositorio Git ya inicializado." -ForegroundColor Green
}

# 3. Validar que .gitignore esté activo y proteja .env
Write-Host "[INFO] Validando exclusión de archivos confidenciales (.env)..." -ForegroundColor Yellow
$checkEnv = git check-ignore .env
if ($checkEnv) {
    Write-Host "[SEGURIDAD OK] Archivo .env está correctamente ignorado por .gitignore." -ForegroundColor Green
} else {
    Write-Host "[ALERTA DE SEGURIDAD] .env NO está ignorado. Por favor revise .gitignore antes de continuar." -ForegroundColor Red
    exit 1
}

# 4. Agregar archivos permitidos
Write-Host "[INFO] Añadiendo archivos al área de preparación (git add)..." -ForegroundColor Yellow
git add .

# Mostrar estado resumido
Write-Host "`nArchivos que serán incluidos en el commit:" -ForegroundColor Cyan
git status -s

# 5. Realizar commit
Write-Host "`n[INFO] Creando commit..." -ForegroundColor Yellow
git commit -m $CommitMessage

# 6. Configurar o verificar control remoto (remote origin)
$existingRemote = git remote get-url origin 2>$null
if (-not $existingRemote) {
    if ([string]::IsNullOrWhiteSpace($RemoteUrl)) {
        Write-Host "`n------------------------------------------------------------" -ForegroundColor Yellow
        Write-Host "Crea un repositorio vacío en tu cuenta de GitHub (sin README):"
        Write-Host "Ejemplo: https://github.com/TU_USUARIO/Auditoria-Aeropuerto-Barranquilla.git"
        Write-Host "------------------------------------------------------------" -ForegroundColor Yellow
        $RemoteUrl = Read-Host "Ingresa la URL HTTPS o SSH de tu repositorio de GitHub"
    }

    if (-not [string]::IsNullOrWhiteSpace($RemoteUrl)) {
        git remote add origin $RemoteUrl.Trim()
        Write-Host "[OK] Remoto 'origin' vinculado a: $RemoteUrl" -ForegroundColor Green
    } else {
        Write-Host "[AVISO] No se proporcionó URL remota. El commit local se guardó pero no se subió a GitHub." -ForegroundColor Yellow
        exit 0
    }
} else {
    Write-Host "[INFO] Remoto 'origin' ya configurado: $existingRemote" -ForegroundColor Green
}

# 7. Push hacia GitHub
Write-Host "`n[INFO] Subiendo código a GitHub (git push -u origin main)..." -ForegroundColor Yellow
git push -u origin main

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n============================================================" -ForegroundColor Green
    Write-Host " [ÉXITO] ¡Código publicado exitosamente en GitHub!          " -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
} else {
    Write-Host "`n[NOTA] Si GitHub te solicitó autenticación, asegúrate de iniciar sesión con tu cuenta o Personal Access Token (PAT)." -ForegroundColor Yellow
}
