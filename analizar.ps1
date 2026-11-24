#requires -Version 5.1
<#
.SYNOPSIS
    Resume las sesiones del Monitor de batería y energía en un informe HTML.
.DESCRIPTION
    Arriba va el resumen de cada sesión: cuánto grabó, si se cerró bien, qué juegos y aplicaciones
    se usaron y lo raro que encontró. Debajo va el detalle completo. Solo lee los registros:
    no necesita permisos de administrador ni cambia nada del equipo.
.EXAMPLE
    .\analizar.ps1                         # sesiones de las últimas 24 horas
.EXAMPLE
    .\analizar.ps1 -Horas 72               # sesiones de los últimos 3 días
.EXAMPLE
    .\analizar.ps1 -Todas                  # todas las sesiones
.EXAMPLE
    .\analizar.ps1 -Sesion 2026-11-18_20-13-55
#>
param(
    [int]$Horas = 24,
    [switch]$Todas,
    [string]$Sesion,
    [switch]$NoAbrir,
    [string]$Registros = (Join-Path $env:USERPROFILE 'MonitorBateria\registros')
)

$ErrorActionPreference = 'Stop'
$inv = [Globalization.CultureInfo]::InvariantCulture
$es = New-Object Globalization.CultureInfo 'es-ES'
$estilo = [Globalization.NumberStyles]::Float
$fmtHora = 'yyyy-MM-dd HH:mm:ss.fff'
$NaN = [double]::NaN
$cronometro = [Diagnostics.Stopwatch]::StartNew()

# Programas que suelen llegar empaquetados con otros instaladores
$patronNoDeseados = '(?i)coccoc|utweb|utorrent|browser_assistant|opera_autoupdate|onelaunch|wavebrowser|webcompanion|pcappstore'

# Códigos de salida que indican que el programa falló (no un cierre normal)
$codigosFallo = @{
    '0xC0000005' = 'acceso a memoria no válido'; '0xC0000409' = 'desbordamiento de búfer'; '0xC0000374' = 'memoria dañada'
    '0xC00000FD' = 'desbordamiento de pila'; '0xC0000142' = 'no pudo iniciarse'; '0xC000001D' = 'instrucción no válida'
    '0xC0000017' = 'sin memoria'; '0xE0434352' = 'error de .NET'; '0xC06D007E' = 'falta un módulo'
}

# ------------------------------------------------------------------ utilidades

# Separa una línea CSV respetando las comillas
function Dividir([string]$linea) {
    if ($linea.IndexOf('"') -lt 0) { return ,$linea.Split(',') }
    $campos = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $entre = $false
    for ($i = 0; $i -lt $linea.Length; $i++) {
        $ch = $linea[$i]
        if ($entre) {
            if ($ch -eq '"') {
                if ($i + 1 -lt $linea.Length -and $linea[$i + 1] -eq '"') { [void]$sb.Append('"'); $i++ } else { $entre = $false }
            }
            else { [void]$sb.Append($ch) }
        }
        elseif ($ch -eq '"') { $entre = $true }
        elseif ($ch -eq ',') { $campos.Add($sb.ToString()); [void]$sb.Clear() }
        else { [void]$sb.Append($ch) }
    }
    $campos.Add($sb.ToString())
    return ,$campos.ToArray()
}

# Abre un CSV del monitor aunque el monitor lo esté escribiendo en ese momento
function Abrir-Csv([string]$ruta) {
    if (-not (Test-Path -LiteralPath $ruta)) { return $null }
    $fs = New-Object IO.FileStream($ruta, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $lector = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8)
    $cabecera = $lector.ReadLine()
    if ($cabecera -eq $null) { $lector.Dispose(); return $null }
    $col = @{}
    $nombres = Dividir $cabecera
    for ($i = 0; $i -lt $nombres.Length; $i++) { $col[$nombres[$i].Trim([char]0xFEFF)] = $i }
    return [pscustomobject]@{ Lector = $lector; Col = $col; Columnas = $nombres.Length }
}

# CSV pequeño entero en memoria
function Leer-Tabla([string]$ruta) {
    $tabla = [pscustomobject]@{ Col = @{}; Filas = New-Object 'System.Collections.Generic.List[string[]]' }
    $csv = Abrir-Csv $ruta
    if ($csv -eq $null) { return $tabla }
    $tabla.Col = $csv.Col
    try {
        while (($linea = $csv.Lector.ReadLine()) -ne $null) {
            if ($linea.Length -eq 0) { continue }
            $campos = Dividir $linea
            if ($campos.Length -ge $csv.Columnas) { $tabla.Filas.Add($campos) }
        }
    }
    finally { $csv.Lector.Dispose() }
    return $tabla
}

function Num([string]$s) { $v = 0.0; if ([double]::TryParse($s, $estilo, $inv, [ref]$v)) { return $v } return $NaN }
function Hx($texto) { return [Net.WebUtility]::HtmlEncode([string]$texto) }
function Fmt([double]$v, [int]$dec = 0, [string]$unidad = '') {
    if ([double]::IsNaN($v) -or [double]::IsInfinity($v)) { return 'n/d' }
    return $v.ToString("N$dec", $es) + $unidad
}
function Dur([double]$seg) {
    if ([double]::IsNaN($seg)) { return 'n/d' }
    $t = [TimeSpan]::FromSeconds([math]::Round([math]::Max(0, $seg)))
    if ($t.TotalHours -ge 1) { return '{0} h {1:00} min' -f [int][math]::Floor($t.TotalHours), $t.Minutes }
    if ($t.TotalMinutes -ge 1) { return '{0} min {1:00} s' -f $t.Minutes, $t.Seconds }
    return '{0} s' -f $t.Seconds
}
function Corta([datetime]$t) { return $t.ToString('dd/MM HH:mm:ss', $inv) }

# Tabla HTML a partir de objetos; las columnas en $numericas se alinean a la derecha
function Html-Tabla($filas, [string[]]$columnas, [string[]]$numericas = @()) {
    $lista = @($filas)
    if ($lista.Count -eq 0) { return '<p class="tenue">Sin datos.</p>' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="desliza"><table><thead><tr>')
    foreach ($c in $columnas) { [void]$sb.Append('<th>' + (Hx $c) + '</th>') }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($f in $lista) {
        [void]$sb.Append('<tr>')
        foreach ($c in $columnas) {
            $clase = ''
            if ($numericas -contains $c) { $clase = ' class="n"' }
            [void]$sb.Append("<td$clase>" + (Hx $f.$c) + '</td>')
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table></div>')
    return $sb.ToString()
}

# Firma digital de un ejecutable (se recuerda para no comprobar dos veces el mismo)
$firmas = @{}
function Firma([string]$ruta) {
    if ($firmas.ContainsKey($ruta)) { return $firmas[$ruta] }
    $res = 'no se pudo comprobar'
    if (-not (Test-Path -LiteralPath $ruta)) { $res = 'el archivo ya no existe' }
    else {
        try {
            $s = Get-AuthenticodeSignature -LiteralPath $ruta
            if ($s.Status -eq 'Valid') { $res = 'firmado por ' + ((($s.SignerCertificate.Subject -split ', ')[0]) -replace '^CN=', '' -replace '"', '') }
            else { $res = 'SIN FIRMA VÁLIDA (' + $s.Status + ')' }
        }
        catch { }
    }
    $firmas[$ruta] = $res
    return $res
}

# ------------------------------------------------------------------ programas instalados

# Carpeta en minúsculas y sin barra final; null si es demasiado corta para identificar nada
function Normalizar([string]$d) {
    if (-not $d) { return $null }
    $d = $d.Trim().Trim('"').TrimEnd('\').ToLowerInvariant()
    if ($d.Length -lt 4) { return $null }
    return $d
}

# Carpeta del ejecutable que aparece en DisplayIcon o UninstallString ("C:\x\y.exe",0 o C:\x\y.exe /args)
function Carpeta-De([string]$s) {
    if (-not $s) { return $null }
    $s = $s.Trim()
    if ($s.StartsWith('"')) { $fin = $s.IndexOf('"', 1); if ($fin -gt 0) { $s = $s.Substring(1, $fin - 1) } }
    else { $i = $s.ToLowerInvariant().IndexOf('.exe'); if ($i -gt 0) { $s = $s.Substring(0, $i + 4) } }
    $s = $s -replace ',\s*-?\d+$', ''
    if ($s -notmatch '^[a-zA-Z]:\\' -or $s -match '(?i)\\(msiexec|rundll32)\.exe$') { return $null }
    try { return Normalizar ([IO.Path]::GetDirectoryName($s)) } catch { return $null }
}

$perfil = Normalizar $env:USERPROFILE
$raices = @(
    @{ Clave = 'pf'; Ruta = 'c:\program files\common files\' }, @{ Clave = 'pf'; Ruta = 'c:\program files (x86)\common files\' },
    @{ Clave = 'pf'; Ruta = 'c:\program files\' }, @{ Clave = 'pf'; Ruta = 'c:\program files (x86)\' },
    @{ Clave = 'lp'; Ruta = $perfil + '\appdata\local\programs\' }, @{ Clave = 'l'; Ruta = $perfil + '\appdata\local\' },
    @{ Clave = 'r'; Ruta = $perfil + '\appdata\roaming\' }, @{ Clave = 'pd'; Ruta = 'c:\programdata\' }
)
# Carpetas demasiado generales: que un programa las declare no dice nada de lo que hay dentro
$genericas = @('c:\program files', 'c:\program files (x86)', 'c:\program files\common files', 'c:\program files (x86)\common files',
    'c:\programdata', 'c:\windows', 'c:\windows\system32', 'c:\windows\syswow64', $perfil, ($perfil + '\appdata\local'),
    ($perfil + '\appdata\roaming'), ($perfil + '\appdata\local\programs'))

# "Fabricante" de una carpeta: la primera carpeta bajo Archivos de programa, AppData, etc. (p. ej. pf:coccoc)
function Clave-Fabricante([string]$dir) {
    if (-not $dir) { return $null }
    foreach ($r in $raices) {
        if ($dir.StartsWith($r.Ruta)) {
            $resto = $dir.Substring($r.Ruta.Length)
            if ($resto) { return $r.Clave + ':' + ($resto -split '\\')[0] }
        }
    }
    return $null
}

$instalados = $null
function Cargar-Instalados {
    $lista = New-Object System.Collections.Generic.List[object]
    $claves = @('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*')
    foreach ($k in $claves) {
        foreach ($e in @(Get-ItemProperty -Path $k -ErrorAction SilentlyContinue)) {
            if (-not $e.DisplayName) { continue }
            $candidatas = @((Normalizar $e.InstallLocation), (Carpeta-De $e.DisplayIcon), (Carpeta-De $e.UninstallString))
            # Los juegos de Steam, Epic, etc. se desinstalan con el lanzador ("steam.exe steam://uninstall/..."):
            # la carpeta del lanzador no es la del juego, así que solo vale su InstallLocation
            if ([string]$e.UninstallString -match '://') { $candidatas = @((Normalizar $e.InstallLocation)) }
            $dirs = $candidatas |
                Where-Object { $_ -and $genericas -notcontains $_ -and $_ -notlike 'c:\windows\installer*' -and $_ -notlike 'c:\programdata\package cache*' } | Select-Object -Unique
            $lista.Add([pscustomobject]@{
                Nombre = [string]$e.DisplayName; Editor = [string]$e.Publisher; Version = [string]$e.DisplayVersion
                Desinstalar = [string]$e.UninstallString; Oculto = ($e.SystemComponent -eq 1); Dirs = @($dirs)
                Fabricantes = @($dirs | ForEach-Object { Clave-Fabricante $_ } | Where-Object { $_ } | Select-Object -Unique)
            })
        }
    }
    return ,$lista
}

# A qué programa instalado pertenece un ejecutable y cómo quitarlo
$cacheProgramas = @{}
function Identificar([string]$ruta) {
    if (-not $ruta) { return $null }
    if ($cacheProgramas.ContainsKey($ruta)) { return $cacheProgramas[$ruta] }
    if ($script:instalados -eq $null) { $script:instalados = Cargar-Instalados }
    $dir = Normalizar ([IO.Path]::GetDirectoryName($ruta))
    $res = [pscustomobject]@{ Programa = ''; Editor = ''; Version = ''; Como = ''; Comando = ''; Carpeta = [IO.Path]::GetDirectoryName($ruta); Tipo = 'suelto'; Certeza = '' }
    if ($dir -like 'c:\windows\system32\driverstore\filerepository\*') {
        $inf = ($dir.Substring('c:\windows\system32\driverstore\filerepository\'.Length) -split '\\')[0] -replace '_amd64_[0-9a-f]+$', '' -replace '_x86_[0-9a-f]+$', ''
        $res.Programa = 'Controlador (driver): ' + $inf; $res.Tipo = 'controlador'
        $res.Como = 'Va con un controlador: se quita desde el Administrador de dispositivos o con el desinstalador del fabricante.'
    }
    elseif ($dir -like 'c:\windows*') { $res.Programa = 'Windows'; $res.Tipo = 'windows'; $res.Como = 'Es parte de Windows: no se desinstala.' }
    elseif ($dir -like 'c:\program files\windowsapps\*') {
        $paquete = (($dir.Substring('c:\program files\windowsapps\'.Length) -split '\\')[0] -split '_')[0]
        $res.Programa = 'App de Microsoft Store: ' + $paquete; $res.Tipo = 'store'
        $res.Como = 'Configuración > Aplicaciones > Aplicaciones instaladas > busca la app > Desinstalar.'
        $res.Comando = 'Get-AppxPackage ' + $paquete + ' | Remove-AppxPackage'
    }
    elseif ($ruta -match '(?i)^[a-z]:\\(xboxgames|xbox)\\([^\\]+)') {
        $res.Programa = 'Juego de Xbox: ' + $matches[2]; $res.Tipo = 'store'
        $res.Como = 'Desde la app Xbox (Mi biblioteca > ... > Desinstalar) o Configuración > Aplicaciones > Aplicaciones instaladas.'
    }
    else {
        # 1) El programa cuya carpeta contiene al ejecutable (la más concreta gana)
        $mejor = $null; $largo = 0
        foreach ($p in $script:instalados) {
            foreach ($d in $p.Dirs) {
                if (($dir -eq $d -or $dir.StartsWith($d + '\')) -and ($d.Length -gt $largo -or ($d.Length -eq $largo -and $mejor.Oculto -and -not $p.Oculto))) { $mejor = $p; $largo = $d.Length }
            }
        }
        # 2) Si no, el del mismo fabricante (p. ej. el actualizador de CocCoc está en otra carpeta que el
        #    navegador). Primero con la misma raíz (Archivos de programa, AppData...) y luego con cualquiera.
        if ($mejor -eq $null) {
            $clave = Clave-Fabricante $dir
            if ($clave) {
                $fabricante = $clave.Substring($clave.IndexOf(':') + 1)
                $candidatos = @($script:instalados | Where-Object { $_.Fabricantes -contains $clave })
                if ($candidatos.Count -eq 0) { $candidatos = @($script:instalados | Where-Object { @($_.Fabricantes | Where-Object { $_.EndsWith(':' + $fabricante) }).Count -gt 0 }) }
                if ($candidatos.Count -gt 0) {
                    # Mejor el que se llame como la carpeta que sigue a la del fabricante (p. ej. "NVIDIA app")
                    $siguiente = ''
                    $i = $dir.IndexOf('\' + $fabricante + '\')
                    if ($i -ge 0) { $siguiente = ($dir.Substring($i + $fabricante.Length + 2) -split '\\')[0] }
                    $candidatos = @($candidatos | Sort-Object { if ($siguiente -and $_.Nombre.ToLowerInvariant().Contains($siguiente)) { 0 } else { 1 } }, Oculto)
                    $mejor = $candidatos[0]
                    $otros = @($candidatos | Select-Object -ExpandProperty Nombre -Unique | Where-Object { $_ -ne $mejor.Nombre })
                    $res.Certeza = 'deducido por la carpeta del fabricante'
                    if ($otros.Count -gt 0) { $res.Certeza += '; también podría ser: ' + (@($otros | Select-Object -First 3) -join ', ') }
                }
            }
        }
        if ($mejor -ne $null) {
            $res.Programa = $mejor.Nombre; $res.Editor = $mejor.Editor; $res.Version = $mejor.Version; $res.Comando = $mejor.Desinstalar; $res.Tipo = 'instalado'
            if ($mejor.Dirs.Count -gt 0) { $res.Carpeta = $mejor.Dirs[0] }
            if ($mejor.Desinstalar -match 'steam://uninstall') { $res.Como = 'Desde Steam: Biblioteca > clic derecho en el juego > Administrar > Desinstalar.' }
            elseif ($mejor.Oculto) { $res.Como = 'Es un componente oculto de otro programa: se quita al desinstalar el programa principal.' }
            else { $res.Como = 'Configuración > Aplicaciones > Aplicaciones instaladas > busca «' + $mejor.Nombre + '» > Desinstalar.' }
        }
        else {
            $res.Programa = '(no figura como programa instalado)'
            $res.Como = 'No tiene desinstalador registrado: puede ser un programa portátil o haberlo dejado otro. Para quitarlo, ciérralo y borra su carpeta, solo si sabes qué es.'
            if ($ruta -match '(?i)\\AppData\\Local\\Temp\\') { $res.Como = 'Se ejecutó desde la carpeta temporal: revísalo con el antivirus antes de nada.' }
        }
    }
    $cacheProgramas[$ruta] = $res
    return $res
}

function Texto-Programa($id) {
    if ($id -eq $null) { return '' }
    $t = $id.Programa
    if ($id.Editor) { $t += ', de ' + $id.Editor }
    if ($id.Certeza) { $t += ' (' + $id.Certeza + ')' }
    return $t
}

# Eventos de Windows relacionados con energía, hardware y bloqueos en un intervalo
function Eventos-Windows([datetime]$desde, [datetime]$hasta) {
    $lista = New-Object System.Collections.Generic.List[object]
    $filtros = @(
        @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Power'; Id = 41 },
        @{ LogName = 'System'; ProviderName = 'EventLog'; Id = 6008 },
        @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-WHEA-Logger' },
        @{ LogName = 'System'; ProviderName = 'Display' },
        @{ LogName = 'System'; ProviderName = 'nvlddmkm' },
        @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Processor-Power'; Id = 37 },
        @{ LogName = 'Application'; ProviderName = 'Application Error'; Id = 1000 },
        @{ LogName = 'Application'; ProviderName = 'Windows Error Reporting'; Id = 1001 }
    )
    foreach ($fh in $filtros) {
        $fh.StartTime = $desde
        $fh.EndTime = $hasta
        try {
            foreach ($e in Get-WinEvent -FilterHashtable $fh -ErrorAction Stop) {
                $d = Describir-Evento $e $desde
                if ($d -ne $null) { $lista.Add($d) }
            }
        }
        catch { }
    }
    return ,$lista
}

function Describir-Evento($e, [datetime]$inicioSesion) {
    $msg = [string]$e.Message
    $primera = (($msg -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1)
    $nivel = 'atencion'
    $texto = $primera
    switch ($e.ProviderName) {
        'Microsoft-Windows-Kernel-Power' { $nivel = 'grave'; $texto = 'Apagado brusco: el equipo perdió la energía sin apagarse (Kernel-Power 41)' }
        'EventLog' { $texto = 'Windows registró que el apagado anterior fue inesperado (6008)' }
        'Microsoft-Windows-WHEA-Logger' { $nivel = 'grave'; $texto = 'Error de hardware (WHEA): ' + $primera }
        'Display' { $texto = 'El controlador de la GPU dejó de responder y se recuperó: ' + $primera }
        'nvlddmkm' { $texto = 'Error del controlador NVIDIA: ' + $primera }
        'Microsoft-Windows-Kernel-Processor-Power' { $texto = 'El firmware limitó la velocidad del procesador: ' + $primera }
        'Application Error' {
            $app = ''; $cod = ''
            try { $app = [string]$e.Properties[0].Value; $cod = [string]$e.Properties[6].Value } catch { }
            $texto = 'Se cerró por un error: ' + $app + $(if ($cod) { ' (código ' + $cod + ')' } else { '' })
        }
        'Windows Error Reporting' {
            $nombre = ''
            if ($msg -match '(Nombre de evento|Event Name):\s*(\S+)') { $nombre = $matches[2] }
            if ($nombre -ne 'LiveKernelEvent' -and $nombre -notlike 'AppHang*') { return $null }   # los APPCRASH ya salen como Application Error
            $p1 = ''
            if ($msg -match 'P1:\s*(\S+)') { $p1 = $matches[1] }
            if ($nombre -like 'AppHang*') {
                $texto = 'Un programa dejó de responder: ' + $p1
            }
            else {
                $que = 'bloqueo de hardware o del controlador (código ' + $p1 + ')'
                if ($p1 -eq '141' -or $p1 -eq '117') { $que = 'la GPU dejó de responder y se reinició (código ' + $p1 + ')' }
                $texto = 'Informe de ' + $que
                if ($msg -match '-(\d{8})-\d{4}\.dmp') {
                    $fechaDmp = [datetime]::ParseExact($matches[1], 'yyyyMMdd', $inv)
                    if ($fechaDmp -lt $inicioSesion.Date) {
                        $nivel = 'info'
                        $texto = 'Windows reintentó enviar un informe antiguo, del ' + $fechaDmp.ToString('dd/MM/yyyy') + ': ' + $que + '. No es un fallo nuevo.'
                    }
                }
            }
        }
    }
    return [pscustomobject]@{ T = $e.TimeCreated; Hora = (Corta $e.TimeCreated); Origen = ($e.ProviderName -replace '^Microsoft-Windows-', ''); Id = $e.Id; Nivel = $nivel; Texto = $texto }
}

# ------------------------------------------------------------------ análisis de una sesión

function Analizar-Sesion($ses) {
    $r = $ses.Ruta
    $hallazgos = New-Object System.Collections.Generic.List[object]
    $kpis = New-Object System.Collections.Generic.List[object]
    $detalle = New-Object System.Text.StringBuilder
    function Hallazgo([string]$nivel, [string]$texto) { $hallazgos.Add([pscustomobject]@{ Nivel = $nivel; Texto = $texto; I = $hallazgos.Count }) }
    function Kpi([string]$etiqueta, [string]$valor, [string]$nota, [string]$nivel) { $kpis.Add([pscustomobject]@{ Etiqueta = $etiqueta; Valor = $valor; Nota = $nota; Nivel = $nivel }) }
    function Seccion([string]$titulo, [string]$html) { [void]$detalle.Append('<details><summary>' + (Hx $titulo) + '</summary><div class="cuerpo">' + $html + '</div></details>') }

    # --- info.txt y FIN_DE_SESION.txt
    $info = @{}
    $rutaInfo = Join-Path $r 'info.txt'
    if (Test-Path -LiteralPath $rutaInfo) {
        foreach ($l in [IO.File]::ReadAllLines($rutaInfo, [Text.Encoding]::UTF8)) {
            if ($l -match '^([^:=]{2,40}):\s*(.*)$') { if (-not $info.ContainsKey($matches[1])) { $info[$matches[1]] = $matches[2] } }
        }
    }
    $esAdmin = ([string]$info['Permisos']) -like 'administrador*'
    $intervalo = 1; $intProc = 3
    if ([string]$info['Muestreo'] -match 'cada (\d+) s; procesos y GPU por proceso cada (\d+) s') { $intervalo = [int]$matches[1]; $intProc = [int]$matches[2] }
    $mhzBase = 2592.0
    if ([string]$info['CPU'] -match 'base (\d+) MHz') { $mhzBase = [double]$matches[1] }

    $rutaFin = Join-Path $r 'FIN_DE_SESION.txt'
    $finLineas = @()
    if (Test-Path -LiteralPath $rutaFin) { $finLineas = [IO.File]::ReadAllLines($rutaFin, [Text.Encoding]::UTF8) }
    $abrupto = $false
    $nivelCierre = 'ok'
    if ($finLineas.Count -gt 0 -and $finLineas[0] -like 'Cierre limpio*') {
        $cierre = 'Cerrado correctamente (' + ($finLineas[0] -replace '^Cierre limpio del monitor:\s*', '') + ')'
    }
    elseif ($finLineas.Count -gt 0) {
        $clasif = (($finLineas | Where-Object { $_ -like 'Clasificación:*' } | Select-Object -First 1) -replace '^Clasificación:\s*', '')
        $cierre = 'No se cerró bien: ' + $clasif
        $abrupto = $clasif -like 'APAGADO BRUSCO*'
        $nivelCierre = 'atencion'
        if ($abrupto) { $nivelCierre = 'grave' }
    }
    elseif ($ses.EnCurso) { $cierre = 'En curso: el monitor sigue grabando'; $nivelCierre = 'info' }
    else { $cierre = 'Terminó sin cierre limpio y aún no se ha revisado: abre el monitor y vuelve a generar el informe'; $nivelCierre = 'atencion' }

    # --- eventos.csv
    $ev = Leer-Tabla (Join-Path $r 'eventos.csv')
    $eH = $ev.Col['fecha_hora']; $eTipo = $ev.Col['tipo']; $eDet = $ev.Col['detalle']
    if ($eH -eq $null) { $ev.Filas.Clear() }

    # --- sistema.csv: una pasada leyendo línea a línea
    $cs = Abrir-Csv (Join-Path $r 'sistema.csv')
    if ($cs -eq $null) {
        Hallazgo 'grave' 'La sesión no tiene sistema.csv: el monitor no llegó a grabar.'
        return [pscustomobject]@{ Nombre = $ses.Nombre; Inicio = $ses.Inicio; Fin = $ses.Fin; Duracion = 0; Cierre = $cierre; NivelCierre = $nivelCierre; Abrupto = $abrupto; Hallazgos = $hallazgos; Kpis = $kpis; Resumen = ''; Detalle = '' }
    }
    $c = $cs.Col
    function Ix([string]$nombre) { if ($c.ContainsKey($nombre)) { return [int]$c[$nombre] } return -1 }
    $iT = Ix 'fecha_hora'; $iAc = Ix 'cargador'; $iCarg = Ix 'cargando'; $iPct = Ix 'bateria_pct_windows'; $iCap = Ix 'capacidad_mwh'
    $iFcc = Ix 'capacidad_completa_mwh'; $iVolt = Ix 'voltaje_mv'; $iW = Ix 'cpu_mas_gpu_w'; $iCpuW = Ix 'cpu_paquete_w'; $iMhz = Ix 'cpu_mhz'
    $iGT = Ix 'gpu_temp_c'; $iGW = Ix 'gpu_w'; $iRaz = Ix 'gpu_razones'; $iRam = Ix 'ram_libre_mb'; $iCom = Ix 'ram_comprometida_pct'
    $iJ = Ix 'juego'; $iFps = Ix 'juego_fps'; $iAcpi = Ix 'temp_acpi_c'; $iOk = Ix 'bateria_ok'

    $lTicks = New-Object 'System.Collections.Generic.List[long]'
    $lW = New-Object 'System.Collections.Generic.List[double]'
    $lAc = New-Object 'System.Collections.Generic.List[bool]'
    $cola = New-Object 'System.Collections.Generic.Queue[string[]]'
    $huecos = New-Object System.Collections.Generic.List[object]
    $cambiosFcc = New-Object System.Collections.Generic.List[string]
    $juegos = New-Object System.Collections.Generic.List[object]
    $ventanas = [ordered]@{}
    $razones = @{}
    $tick20 = [TimeSpan]::FromMinutes(20).Ticks
    $x = 0.0
    $t0 = $null; $tPrev = $null; $tUlt = $null
    $segHuecos = 0.0; $segAc = 0; $segBat = 0; $segCarga = 0; $segBatMal = 0; $segSinCargar = 0
    $pctIni = -1; $pctFin = -1; $pctMin = 101; $pctMax = -1; $pctSinCargarMin = 101
    $capIni = $NaN; $capFin = $NaN; $capPrev = $NaN; $perdidoAc = 0.0; $perdidoBat = 0.0; $ganado = 0.0
    $fccIni = $NaN; $fccPrev = $NaN
    $voltMin = [double]::MaxValue; $voltMax = 0.0
    $sumW = 0.0; $nW = 0; $maxW = 0.0; $maxCpuW = 0.0; $maxMhz = 0.0
    $maxGT = 0.0; $segGT86 = 0; $maxAcpi = 0.0
    $ramMin = [double]::MaxValue; $ramMinT = $null; $segRam500 = 0; $comMax = 0.0
    $jAct = ''; $jIni = $null; $jFin = $null; $jSumW = 0.0; $jNW = 0; $jMaxGT = 0.0; $jSumFps = 0.0; $jNFps = 0

    try {
        while (($linea = $cs.Lector.ReadLine()) -ne $null) {
            if ($linea.Length -lt 30) { continue }
            if ($linea.IndexOf('"') -ge 0) { $f = Dividir $linea } else { $f = $linea.Split(',') }
            if ($f.Length -lt $cs.Columnas) { continue }   # línea cortada por un apagado brusco
            $tm = [datetime]::MinValue
            if (-not [datetime]::TryParseExact($f[$iT], $fmtHora, $inv, [Globalization.DateTimeStyles]::None, [ref]$tm)) { continue }
            if ($t0 -eq $null) { $t0 = $tm; $tPrev = $tm }
            $dt = ($tm - $tPrev).TotalSeconds
            if ($dt -gt 3 * $intervalo + 1) {
                $huecos.Add([pscustomobject]@{ Desde = (Corta $tPrev); Hasta = (Corta $tm); Duración = (Dur $dt) })
                $segHuecos += $dt - $intervalo
            }
            $tPrev = $tm; $tUlt = $tm
            $lTicks.Add($tm.Ticks)
            $ac = $f[$iAc] -eq '1'
            $lAc.Add($ac)
            if ($ac) { $segAc += $intervalo } else { $segBat += $intervalo }
            $carg = $f[$iCarg] -eq '1'
            if ($carg) { $segCarga += $intervalo }
            if ($iOk -ge 0 -and $f[$iOk] -eq '0') { $segBatMal += $intervalo }

            $pct = -1
            if ([int]::TryParse($f[$iPct], [ref]$pct)) {
                if ($pctIni -lt 0) { $pctIni = $pct }
                $pctFin = $pct
                if ($pct -lt $pctMin) { $pctMin = $pct }
                if ($pct -gt $pctMax) { $pctMax = $pct }
                if ($ac -and -not $carg -and $pct -lt 95) { $segSinCargar += $intervalo; if ($pct -lt $pctSinCargarMin) { $pctSinCargarMin = $pct } }
            }
            $cap = $NaN
            if ([double]::TryParse($f[$iCap], $estilo, $inv, [ref]$x)) { $cap = $x }
            if (-not [double]::IsNaN($cap)) {
                if ([double]::IsNaN($capIni)) { $capIni = $cap }
                if (-not [double]::IsNaN($capPrev)) {
                    $d = $cap - $capPrev
                    if ($d -lt 0) { if ($ac) { $perdidoAc -= $d } else { $perdidoBat -= $d } }
                    elseif ($d -gt 0) { $ganado += $d }
                }
                $capPrev = $cap; $capFin = $cap
            }
            if ([double]::TryParse($f[$iFcc], $estilo, $inv, [ref]$x)) {
                if ([double]::IsNaN($fccIni)) { $fccIni = $x }
                if (-not [double]::IsNaN($fccPrev) -and $x -ne $fccPrev) { $cambiosFcc.Add((Corta $tm) + ': ' + (Fmt $fccPrev) + ' → ' + (Fmt $x) + ' mWh') }
                $fccPrev = $x
            }
            if ([double]::TryParse($f[$iVolt], $estilo, $inv, [ref]$x)) { if ($x -lt $voltMin) { $voltMin = $x }; if ($x -gt $voltMax) { $voltMax = $x } }
            $w = $NaN
            if ([double]::TryParse($f[$iW], $estilo, $inv, [ref]$x)) { $w = $x; $sumW += $x; $nW++; if ($x -gt $maxW) { $maxW = $x } }
            $lW.Add($w)
            if ([double]::TryParse($f[$iCpuW], $estilo, $inv, [ref]$x) -and $x -gt $maxCpuW) { $maxCpuW = $x }
            if ([double]::TryParse($f[$iMhz], $estilo, $inv, [ref]$x) -and $x -gt $maxMhz) { $maxMhz = $x }
            $gt = $NaN
            if ([double]::TryParse($f[$iGT], $estilo, $inv, [ref]$x)) { $gt = $x; if ($x -gt $maxGT) { $maxGT = $x }; if ($x -ge 86) { $segGT86 += $intervalo } }
            if ([double]::TryParse($f[$iAcpi], $estilo, $inv, [ref]$x) -and $x -gt $maxAcpi) { $maxAcpi = $x }
            $rz = $f[$iRaz]
            if ($rz -ne '' -and $rz -ne 'reposo') { $razones[$rz] = $intervalo + [int]$razones[$rz] }
            if ([double]::TryParse($f[$iRam], $estilo, $inv, [ref]$x)) {
                if ($x -lt $ramMin) { $ramMin = $x; $ramMinT = $tm }
                if ($x -lt 500) { $segRam500 += $intervalo }
            }
            if ([double]::TryParse($f[$iCom], $estilo, $inv, [ref]$x) -and $x -gt $comMax) { $comMax = $x }
            $fps = $NaN
            if ($iFps -ge 0 -and [double]::TryParse($f[$iFps], $estilo, $inv, [ref]$x)) { $fps = $x }

            # Tramos de juego
            $j = $f[$iJ]
            if ($j -ne $jAct) {
                if ($jAct -ne '') { $juegos.Add([pscustomobject]@{ Juego = $jAct; Desde = $jIni; Hasta = $jFin; SumW = $jSumW; NW = $jNW; MaxGT = $jMaxGT; SumFps = $jSumFps; NFps = $jNFps }) }
                $jAct = $j; $jIni = $tm; $jSumW = 0.0; $jNW = 0; $jMaxGT = 0.0; $jSumFps = 0.0; $jNFps = 0
            }
            $jFin = $tm
            if ($jAct -ne '') {
                if (-not [double]::IsNaN($w)) { $jSumW += $w; $jNW++ }
                if ($gt -gt $jMaxGT) { $jMaxGT = $gt }
                if (-not [double]::IsNaN($fps)) { $jSumFps += $fps; $jNFps++ }
            }

            # Ventanas de 20 minutos (clave de texto: con un número el diccionario lo tomaría por una posición)
            $vn = [long][math]::Floor($tm.Ticks / $tick20)
            $vk = $vn.ToString()
            $v = $ventanas[$vk]
            if ($v -eq $null) {
                $v = @{ Inicio = (New-Object DateTime ($vn * $tick20)); CapIni = $cap; CapFin = $cap; SumW = 0.0; NW = 0; MaxW = 0.0; MaxGT = 0.0; SumFps = 0.0; NFps = 0; Juegos = @{}; SumMa = 0.0; NMa = 0 }
                $ventanas[$vk] = $v
            }
            if (-not [double]::IsNaN($cap)) { if ([double]::IsNaN($v.CapIni)) { $v.CapIni = $cap }; $v.CapFin = $cap }
            if (-not [double]::IsNaN($w)) { $v.SumW = $v.SumW + $w; $v.NW = $v.NW + 1; if ($w -gt $v.MaxW) { $v.MaxW = $w } }
            if ($gt -gt $v.MaxGT) { $v.MaxGT = $gt }
            if (-not [double]::IsNaN($fps)) { $v.SumFps = $v.SumFps + $fps; $v.NFps = $v.NFps + 1 }
            if ($j -ne '') { $v.Juegos[$j] = 1 + [int]$v.Juegos[$j] }

            $cola.Enqueue($f)
            if ($cola.Count -gt 30) { [void]$cola.Dequeue() }
        }
    }
    finally { $cs.Lector.Dispose() }
    if ($jAct -ne '') { $juegos.Add([pscustomobject]@{ Juego = $jAct; Desde = $jIni; Hasta = $jFin; SumW = $jSumW; NW = $jNW; MaxGT = $jMaxGT; SumFps = $jSumFps; NFps = $jNFps }) }

    $n = $lTicks.Count
    if ($n -lt 2) {
        Hallazgo 'grave' 'La sesión casi no tiene muestras.'
        return [pscustomobject]@{ Nombre = $ses.Nombre; Inicio = $ses.Inicio; Fin = $ses.Fin; Duracion = 0; Cierre = $cierre; NivelCierre = $nivelCierre; Abrupto = $abrupto; Hallazgos = $hallazgos; Kpis = $kpis; Resumen = ''; Detalle = '' }
    }
    $durSeg = ($tUlt - $t0).TotalSeconds
    $esperadas = [math]::Floor($durSeg / $intervalo) + 1
    $cobertura = [math]::Min(100, 100.0 * $n / $esperadas)

    # --- datos de MSI (solo sesiones con administrador): corriente y temperatura reales de la batería
    $hayMsi = $false
    $snap = New-Object System.Collections.Generic.List[object]
    $crudo = [ordered]@{}
    $significado = @{
        'MSI_Master_Battery|1' = 'Capacidad de diseño (mAh)'; 'MSI_Master_Battery|2' = 'Voltaje de diseño (mV)'
        'MSI_Master_Battery|6' = 'Carga (%)'; 'MSI_Master_Battery|7' = 'Capacidad completa según el chip (mAh)'
        'MSI_Master_Battery|8' = 'Corriente (mA; negativa = se descarga)'; 'MSI_Master_Battery|9' = 'Carga restante (mAh)'
        'MSI_Master_Battery|10' = 'Voltaje (mV)'; 'MSI_Master_Battery|11' = 'Temperatura de la batería (décimas de kelvin)'
        'MSI_CPU|1' = 'Temperatura de la CPU según el EC (°C)'; 'MSI_CPU|2' = 'Ventilador de la CPU (%)'
        'MSI_VGA|1' = 'Temperatura de la GPU según el EC (°C)'; 'MSI_VGA|2' = 'Ventilador de la GPU (%)'
    }
    $claveCampo = @{
        'MSI_Master_Battery|1' = 'Diseno'; 'MSI_Master_Battery|6' = 'Pct'; 'MSI_Master_Battery|7' = 'Fcc'; 'MSI_Master_Battery|8' = 'Ma'
        'MSI_Master_Battery|9' = 'Rem'; 'MSI_Master_Battery|10' = 'Mv'; 'MSI_Master_Battery|11' = 'Temp'
        'MSI_CPU|1' = 'CpuC'; 'MSI_CPU|2' = 'CpuFan'; 'MSI_VGA|1' = 'GpuC'; 'MSI_VGA|2' = 'GpuFan'
    }
    $cm = Abrir-Csv (Join-Path $r 'msi_wmi.csv')
    if ($cm -ne $null) {
        $act = @{}
        foreach ($k in $claveCampo.Values) { $act[$k] = $NaN }
        $mH = $cm.Col['fecha_hora']; $mC = $cm.Col['clase']; $mI = $cm.Col['instancia']; $mP = $cm.Col['propiedad']; $mV = $cm.Col['valor']
        $hPrev = $null
        try {
            while (($linea = $cm.Lector.ReadLine()) -ne $null) {
                $f = Dividir $linea
                if ($f.Length -lt $cm.Columnas) { continue }
                $h = $f[$mH]
                if ($hPrev -ne $null -and $h -ne $hPrev) {
                    $o = [pscustomobject]@{ T = [datetime]::ParseExact($hPrev, $fmtHora, $inv).Ticks; Diseno = $act.Diseno; Pct = $act.Pct; Fcc = $act.Fcc; Ma = $act.Ma; Rem = $act.Rem; Mv = $act.Mv; Temp = $act.Temp; CpuC = $act.CpuC; CpuFan = $act.CpuFan; GpuC = $act.GpuC; GpuFan = $act.GpuFan }
                    $snap.Add($o)
                }
                $hPrev = $h
                if ($f[$mP] -eq '(error)') { continue }
                $inst = $f[$mI]
                $clave = $f[$mC] + '|' + $inst.Substring($inst.LastIndexOf('_') + 1)
                $val = Num $f[$mV]
                $cr = $crudo[$clave]
                if ($cr -eq $null) { $cr = @{ Primero = $f[$mV]; Ultimo = $f[$mV]; Min = $val; Max = $val; Cambios = 0 }; $crudo[$clave] = $cr }
                $cr.Ultimo = $f[$mV]; $cr.Cambios = $cr.Cambios + 1
                if ($val -lt $cr.Min) { $cr.Min = $val }
                if ($val -gt $cr.Max) { $cr.Max = $val }
                if ($claveCampo.ContainsKey($clave)) { $act[$claveCampo[$clave]] = $val }
            }
        }
        finally { $cm.Lector.Dispose() }
        if ($hPrev -ne $null) { $snap.Add([pscustomobject]@{ T = [datetime]::ParseExact($hPrev, $fmtHora, $inv).Ticks; Diseno = $act.Diseno; Pct = $act.Pct; Fcc = $act.Fcc; Ma = $act.Ma; Rem = $act.Rem; Mv = $act.Mv; Temp = $act.Temp; CpuC = $act.CpuC; CpuFan = $act.CpuFan; GpuC = $act.GpuC; GpuFan = $act.GpuFan }) }
        $hayMsi = $snap.Count -gt 2 -and -not [double]::IsNaN($snap[$snap.Count - 1].Ma)
    }

    # Corriente real de la batería frente al consumo de CPU+GPU de los 10 s anteriores
    $bandas = [ordered]@{}
    foreach ($b in '< 40 W', '40–50 W', '50–60 W', '60–70 W', '70–80 W', '80–90 W', '≥ 90 W') { $bandas[$b] = @{ N = 0; SumMa = 0.0; MinMa = 0.0; NDesc = 0 } }
    $mahAc = 0.0; $mahBat = 0.0; $mahCarga = 0.0; $segDescAc = 0.0; $segCargaMsi = 0.0
    $tempMin = $NaN; $tempMax = $NaN; $maPeor = 0.0; $cpuCMax = $NaN; $gpuCMax = $NaN; $cpuFanMax = $NaN; $gpuFanMax = $NaN
    $umbral = ''
    if ($hayMsi) {
        $jx = 0
        $diez = [TimeSpan]::FromSeconds(10).Ticks
        for ($q = 0; $q -lt $snap.Count; $q++) {
            $s = $snap[$q]
            $T = $s.T
            while ($jx -lt $n -and $lTicks[$jx] -le $T) { $jx++ }
            $suma = 0.0; $cnt = 0; $b = $jx - 1
            while ($b -ge 0 -and $lTicks[$b] -gt $T - $diez) { if (-not [double]::IsNaN($lW[$b])) { $suma += $lW[$b]; $cnt++ }; $b-- }
            $acAhora = $jx -gt 0 -and $lAc[$jx - 1]
            $dtS = 10.0
            if ($q + 1 -lt $snap.Count) { $dtS = [math]::Min(30.0, ($snap[$q + 1].T - $T) / 1e7) }
            $ma = $s.Ma
            if (-not [double]::IsNaN($ma)) {
                if ($ma -lt 0) {
                    if ($acAhora) { $mahAc += -$ma * $dtS / 3600; $segDescAc += $dtS } else { $mahBat += -$ma * $dtS / 3600 }
                    if ($ma -lt $maPeor) { $maPeor = $ma }
                }
                elseif ($ma -gt 0) { $mahCarga += $ma * $dtS / 3600; $segCargaMsi += $dtS }
                if ($cnt -gt 0 -and $acAhora) {
                    $wm = $suma / $cnt
                    $nb = '≥ 90 W'
                    if ($wm -lt 40) { $nb = '< 40 W' } elseif ($wm -lt 50) { $nb = '40–50 W' } elseif ($wm -lt 60) { $nb = '50–60 W' } elseif ($wm -lt 70) { $nb = '60–70 W' } elseif ($wm -lt 80) { $nb = '70–80 W' } elseif ($wm -lt 90) { $nb = '80–90 W' }
                    $bd = $bandas[$nb]
                    $bd.N = $bd.N + 1; $bd.SumMa = $bd.SumMa + $ma
                    if ($ma -lt $bd.MinMa) { $bd.MinMa = $ma }
                    if ($ma -lt 0) { $bd.NDesc = $bd.NDesc + 1 }
                }
                $vk = ([long][math]::Floor($T / $tick20)).ToString()
                if ($ventanas.Contains($vk)) { $vv = $ventanas[$vk]; $vv.SumMa = $vv.SumMa + $ma; $vv.NMa = $vv.NMa + 1 }
            }
            $tc = $s.Temp / 10.0 - 273.15
            if (-not [double]::IsNaN($tc) -and $tc -gt -50) {
                if ([double]::IsNaN($tempMin) -or $tc -lt $tempMin) { $tempMin = $tc }
                if ([double]::IsNaN($tempMax) -or $tc -gt $tempMax) { $tempMax = $tc }
            }
            if (-not [double]::IsNaN($s.CpuC) -and ([double]::IsNaN($cpuCMax) -or $s.CpuC -gt $cpuCMax)) { $cpuCMax = $s.CpuC }
            if (-not [double]::IsNaN($s.GpuC) -and ([double]::IsNaN($gpuCMax) -or $s.GpuC -gt $gpuCMax)) { $gpuCMax = $s.GpuC }
            if (-not [double]::IsNaN($s.CpuFan) -and ([double]::IsNaN($cpuFanMax) -or $s.CpuFan -gt $cpuFanMax)) { $cpuFanMax = $s.CpuFan }
            if (-not [double]::IsNaN($s.GpuFan) -and ([double]::IsNaN($gpuFanMax) -or $s.GpuFan -gt $gpuFanMax)) { $gpuFanMax = $s.GpuFan }
        }
        foreach ($nb in $bandas.Keys) { $bd = $bandas[$nb]; if ($umbral -eq '' -and $bd.N -ge 6 -and $bd.NDesc / $bd.N -ge 0.5) { $umbral = $nb } }
    }

    # --- fps.csv
    $fpsPorProceso = @{}
    $cf = Abrir-Csv (Join-Path $r 'fps.csv')
    if ($cf -ne $null) {
        $fN = $cf.Col['nombre']; $fF = $cf.Col['fps']; $fM = $cf.Col['frametime_max_ms']
        try {
            while (($linea = $cf.Lector.ReadLine()) -ne $null) {
                if ($linea.IndexOf('"') -ge 0) { $f = Dividir $linea } else { $f = $linea.Split(',') }
                if ($f.Length -lt $cf.Columnas) { continue }
                $nom = $f[$fN]
                $o = $fpsPorProceso[$nom]
                if ($o -eq $null) { $o = @{ Fps = New-Object 'System.Collections.Generic.List[double]'; Lentos = 0; Peor = 0.0 }; $fpsPorProceso[$nom] = $o }
                if ([double]::TryParse($f[$fF], $estilo, $inv, [ref]$x)) { $o.Fps.Add($x) }
                if ([double]::TryParse($f[$fM], $estilo, $inv, [ref]$x)) { if ($x -gt 50) { $o.Lentos = $o.Lentos + 1 }; if ($x -gt $o.Peor) { $o.Peor = $x } }
            }
        }
        finally { $cf.Lector.Dispose() }
    }
    $tablaFps = foreach ($nom in $fpsPorProceso.Keys) {
        $o = $fpsPorProceso[$nom]
        if ($o.Fps.Count -eq 0) { continue }
        $orden = $o.Fps.ToArray(); [Array]::Sort($orden)
        $media = ($orden | Measure-Object -Average).Average
        [pscustomobject]@{ Proceso = $nom; 'Tiempo dibujando' = (Dur ($o.Fps.Count * $intervalo)); 'FPS medio' = (Fmt $media 0); 'FPS en el 5 % peor' = (Fmt $orden[[int][math]::Floor($orden.Length * 0.05)] 0); 'Segundos con tirones (>50 ms)' = $o.Lentos; 'Peor cuadro' = (Fmt $o.Peor 0 ' ms'); _N = $o.Fps.Count; _Media = $media }
    }
    $tablaFps = @($tablaFps | Sort-Object _N -Descending)

    # --- ubicación de cada programa (las sesiones antiguas solo la tienen para los procesos que se abrieron)
    $rutasPorNombre = @{}
    foreach ($archivo in 'procesos_al_inicio.csv', 'rutas_procesos.csv', 'procesos_inicio_fin.csv', 'procesos_nuevos.csv') {
        $tr = Leer-Tabla (Join-Path $r $archivo)
        if (-not $tr.Col.ContainsKey('nombre') -or -not $tr.Col.ContainsKey('ruta')) { continue }
        $cN = $tr.Col['nombre']; $cR = $tr.Col['ruta']
        foreach ($f in $tr.Filas) {
            $nom = $f[$cN]; $ru = $f[$cR]
            if (-not $nom -or -not $ru) { continue }
            $lr = $rutasPorNombre[$nom]
            if ($lr -eq $null) { $lr = New-Object System.Collections.Generic.List[string]; $rutasPorNombre[$nom] = $lr }
            if (-not $lr.Contains($ru)) { $lr.Add($ru) }
        }
    }
    $hayRutas = Test-Path -LiteralPath (Join-Path $r 'rutas_procesos.csv')
    function Ubicacion([string]$nombre) {
        $lr = $rutasPorNombre[$nombre]
        if ($lr -eq $null -or $lr.Count -eq 0) {
            if (-not $hayRutas) { return '(no consta: sesión grabada con una versión anterior del monitor)' }
            return '(Windows no dio la ubicación: proceso protegido; con administrador se obtienen casi todas)'
        }
        if ($lr.Count -le 2) { return ($lr -join ' · ') }
        return ($lr[0] + ' · ' + $lr[1] + ' (y ' + ($lr.Count - 2) + ' más)')
    }
    function Programa-De([string]$nombre) {
        $lr = $rutasPorNombre[$nombre]
        if ($lr -eq $null -or $lr.Count -eq 0) { return '' }
        return Texto-Programa (Identificar $lr[0])
    }

    # --- procesos.csv: aplicaciones y procesos en segundo plano
    $grupos = @{}; $fondo = @{}
    $nTot = 0; $sumTotCpu = 0.0; $maxTotCpu = 0.0; $maxTotPriv = 0.0
    $enRamMin = New-Object System.Collections.Generic.List[object]
    $desdeRam = $null; $hastaRam = $null
    if ($ramMinT -ne $null) { $desdeRam = $ramMinT.AddSeconds(-$intProc).ToString($fmtHora, $inv); $hastaRam = $ramMinT.AddSeconds($intProc).ToString($fmtHora, $inv) }
    $cp = Abrir-Csv (Join-Path $r 'procesos.csv')
    if ($cp -ne $null) {
        $pH = $cp.Col['fecha_hora']; $pTipo = $cp.Col['tipo']; $pNom = $cp.Col['nombre']; $pInst = $cp.Col['instancias']; $pCpu = $cp.Col['cpu_pct']; $pPriv = $cp.Col['privada_mb']; $pRam = $cp.Col['ram_mb']
        try {
            while (($linea = $cp.Lector.ReadLine()) -ne $null) {
                if ($linea.IndexOf('"') -ge 0) { $f = Dividir $linea } else { $f = $linea.Split(',') }
                if ($f.Length -lt $cp.Columnas) { continue }
                $tipo = $f[$pTipo]
                $cpu = 0.0; [void][double]::TryParse($f[$pCpu], $estilo, $inv, [ref]$cpu)
                $priv = 0.0; [void][double]::TryParse($f[$pPriv], $estilo, $inv, [ref]$priv)
                if ($tipo -eq 'total') {
                    $nTot++; $sumTotCpu += $cpu
                    if ($cpu -gt $maxTotCpu) { $maxTotCpu = $cpu }
                    if ($priv -gt $maxTotPriv) { $maxTotPriv = $priv }
                    continue
                }
                $nom = $f[$pNom]
                if ($tipo -eq 'grupo') {
                    $g = $grupos[$nom]
                    if ($g -eq $null) { $g = @{ N = 0; SumCpu = 0.0; MaxCpu = 0.0; MaxPriv = 0.0; MaxInst = 0 }; $grupos[$nom] = $g }
                    $g.N = $g.N + 1; $g.SumCpu = $g.SumCpu + $cpu
                    if ($cpu -gt $g.MaxCpu) { $g.MaxCpu = $cpu }
                    if ($priv -gt $g.MaxPriv) { $g.MaxPriv = $priv }
                    $inst = 0; [void][int]::TryParse($f[$pInst], [ref]$inst)
                    if ($inst -gt $g.MaxInst) { $g.MaxInst = $inst }
                    if ($desdeRam -ne $null -and [string]::CompareOrdinal($f[$pH], $desdeRam) -ge 0 -and [string]::CompareOrdinal($f[$pH], $hastaRam) -le 0) {
                        $enRamMin.Add([pscustomobject]@{ Aplicación = $nom; Instancias = $inst; 'Memoria privada' = (Fmt $priv 0 ' MB'); 'En RAM' = ((Fmt (Num $f[$pRam]) 0) + ' MB'); _Priv = $priv })
                    }
                }
                elseif ($tipo -eq 'top') {
                    $g = $fondo[$nom]
                    if ($g -eq $null) { $g = @{ N = 0; SumCpu = 0.0; MaxCpu = 0.0; MaxPriv = 0.0 }; $fondo[$nom] = $g }
                    $g.N = $g.N + 1; $g.SumCpu = $g.SumCpu + $cpu
                    if ($cpu -gt $g.MaxCpu) { $g.MaxCpu = $cpu }
                    if ($priv -gt $g.MaxPriv) { $g.MaxPriv = $priv }
                }
            }
        }
        finally { $cp.Lector.Dispose() }
    }
    $nTotDiv = [math]::Max(1, $nTot)
    $tablaApps = @(foreach ($nom in $grupos.Keys) {
        $g = $grupos[$nom]
        [pscustomobject]@{ Aplicación = $nom; 'Tiempo abierta' = (Dur ($g.N * $intProc)); 'CPU media' = (Fmt ($g.SumCpu / $nTotDiv) 1 ' %'); 'CPU máxima' = (Fmt $g.MaxCpu 0 ' %'); 'Memoria máx.' = (Fmt $g.MaxPriv 0 ' MB'); 'Procesos' = $g.MaxInst; Ubicación = (Ubicacion $nom); 'Programa instalado' = (Programa-De $nom); _Cpu = $g.SumCpu }
    }) | Sort-Object _Cpu -Descending
    $tablaFondo = @(foreach ($nom in $fondo.Keys) {
        if ($grupos.ContainsKey($nom) -or $nom -eq 'System' -or $nom -eq 'Idle') { continue }
        $g = $fondo[$nom]
        [pscustomobject]@{ Proceso = $nom; 'Veces entre los que más CPU usaban' = $g.N; 'CPU media' = (Fmt ($g.SumCpu / $nTotDiv) 2 ' %'); 'CPU máxima' = (Fmt $g.MaxCpu 1 ' %'); 'Memoria máx.' = (Fmt $g.MaxPriv 0 ' MB'); Ubicación = (Ubicacion $nom); 'Programa instalado' = (Programa-De $nom); _Cpu = $g.SumCpu }
    }) | Sort-Object _Cpu -Descending

    # --- procesos que se abrieron y cerraron
    $rutaIF = Join-Path $r 'procesos_inicio_fin.csv'
    $antiguo = -not (Test-Path -LiteralPath $rutaIF)
    if ($antiguo) { $rutaIF = Join-Path $r 'procesos_nuevos.csv' }
    $pf = Leer-Tabla $rutaIF
    $lanz = @{}
    $iEvt = -1; $iVida = -1; $iCod = -1
    if ($pf.Col.ContainsKey('evento')) { $iEvt = $pf.Col['evento']; $iVida = $pf.Col['vida_s']; $iCod = $pf.Col['codigo_salida'] }
    $iNom = $pf.Col['nombre']; $iPad = $pf.Col['padre']; $iRuta = $pf.Col['ruta']; $iHora = $pf.Col['fecha_hora']
    $fallos = New-Object System.Collections.Generic.List[object]
    $totalInicios = 0; $totalFugaces = 0
    $iPid = $pf.Col['pid']
    $clavePorPid = @{}
    foreach ($f in $pf.Filas) {
        # Programas con el mismo nombre pero distinta ruta (p. ej. dos updater.exe) van por separado;
        # el cierre no trae ruta, así que se asocia por pid con su inicio
        $nom = $f[$iNom]
        $evt = 'inicio'
        if ($iEvt -ge 0) { $evt = $f[$iEvt] }
        if ($evt -eq 'inicio') { $clave = $nom + '|' + $f[$iRuta]; $clavePorPid[$f[$iPid]] = $clave }
        else { $clave = $clavePorPid[$f[$iPid]]; if (-not $clave -or -not $clave.StartsWith($nom + '|')) { $clave = $nom + '|' } }
        $o = $lanz[$clave]
        if ($o -eq $null) { $o = @{ Nombre = $nom; Inicios = 0; Fugaces = 0; Fallos = 0; Padres = @{}; Ruta = ''; Primera = $f[$iHora] }; $lanz[$clave] = $o }
        if ($evt -eq 'inicio') {
            $o.Inicios = $o.Inicios + 1; $totalInicios++
            $pad = $f[$iPad]; if (-not $pad) { $pad = '(desconocido)' }
            $o.Padres[$pad] = 1 + [int]$o.Padres[$pad]
            if (-not $o.Ruta -and $f[$iRuta]) { $o.Ruta = $f[$iRuta] }
        }
        else {
            $vida = Num $f[$iVida]
            if (-not [double]::IsNaN($vida) -and $vida -lt 2) { $o.Fugaces = $o.Fugaces + 1; $totalFugaces++ }
            $cod = $f[$iCod]
            if ($codigosFallo.ContainsKey($cod)) {
                $o.Fallos = $o.Fallos + 1
                $fallos.Add([pscustomobject]@{ Hora = $f[$iHora].Substring(11, 8); Proceso = $nom; Código = $cod; Significado = $codigosFallo[$cod]; Ubicación = (Ubicacion $nom) })
            }
        }
    }
    $tablaLanz = @(foreach ($clave in $lanz.Keys) {
        $o = $lanz[$clave]
        if ($o.Inicios -eq 0) { continue }
        $padres = ($o.Padres.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 2 | ForEach-Object { $_.Key }) -join ', '
        $firma = ''
        # Las apps de Microsoft Store van firmadas como paquete, no archivo por archivo
        if ($o.Ruta -match '^(?i)C:\\Program Files\\WindowsApps\\') { $firma = 'app de Microsoft Store' }
        elseif ($o.Ruta -and $o.Ruta -notmatch '^(?i)C:\\Windows\\') { $firma = Firma $o.Ruta }
        $id = Identificar $o.Ruta
        $como = ''
        if ($id -ne $null) { $como = $id.Como }
        $tipo = ''
        if ($id -ne $null) { $tipo = $id.Tipo }
        [pscustomobject]@{ Programa = $o.Nombre; Veces = $o.Inicios; 'Duraron < 2 s' = $o.Fugaces; 'Lo abrió' = $padres; Ubicación = $o.Ruta; 'Programa instalado' = (Texto-Programa $id); Firma = $firma; 'Cómo quitarlo' = $como; _N = $o.Inicios; _Tipo = $tipo }
    }) | Sort-Object _N -Descending
    # Sin firma y sin un programa instalado detrás, o desde carpetas donde no suele haber programas
    $raros = @($tablaLanz | Where-Object { ($_.Firma -like 'SIN FIRMA*' -and $_._Tipo -ne 'instalado' -and $_._Tipo -ne 'store') -or $_.Ubicación -match '(?i)\\AppData\\Local\\Temp\\|\\Downloads\\|\\Users\\Public\\' })
    $sinFirmaConocidos = @($tablaLanz | Where-Object { $_.Firma -like 'SIN FIRMA*' -and ($_._Tipo -eq 'instalado' -or $_._Tipo -eq 'store') })
    $noDeseados = @($tablaLanz | Where-Object { $_.Programa -match $patronNoDeseados -or $_.Ubicación -match $patronNoDeseados -or $_.'Programa instalado' -match $patronNoDeseados })

    # Programas instalados a los que pertenece todo lo que se vio ejecutarse (sin contar Windows)
    $porPrograma = @{}
    foreach ($nom in $rutasPorNombre.Keys) {
        foreach ($ru in $rutasPorNombre[$nom]) {
            $id = Identificar $ru
            if ($id -eq $null -or $id.Tipo -eq 'windows') { continue }
            $clave = Texto-Programa $id
            if ($id.Tipo -eq 'suelto') { $clave += ' ' + $id.Carpeta }
            $g = $porPrograma[$clave]
            if ($g -eq $null) { $g = @{ Id = $id; Ejecutables = New-Object System.Collections.Generic.List[string] }; $porPrograma[$clave] = $g }
            if (-not $g.Ejecutables.Contains($nom)) { $g.Ejecutables.Add($nom) }
        }
    }
    $ordenTipo = @{ suelto = 0; instalado = 1; store = 2; controlador = 3 }
    $tablaInstalados = @(foreach ($clave in $porPrograma.Keys) {
        $g = $porPrograma[$clave]; $id = $g.Id
        $tipo = @{ suelto = 'Sin desinstalador'; instalado = 'Programa instalado'; store = 'App de la Store'; controlador = 'Controlador' }[$id.Tipo]
        [pscustomobject]@{ 'Programa instalado' = (Texto-Programa $id); Versión = $id.Version; Tipo = $tipo; Ejecutables = ($g.Ejecutables -join ', '); Carpeta = $id.Carpeta; 'Cómo quitarlo' = $id.Como; 'Comando de desinstalación' = $id.Comando; _O = $ordenTipo[$id.Tipo]; _N = $id.Programa }
    }) | Sort-Object _O, _N

    # --- eventos de Windows durante la sesión (y el arranque siguiente, si se cortó la luz)
    $evWin = Eventos-Windows $t0 $tUlt.AddHours(6)
    $evWin = @($evWin | Where-Object { $_.T -le $tUlt.AddMinutes(2) -or $_.Id -eq 41 -or $_.Id -eq 6008 } | Sort-Object T)

    # --- cronología resumida de eventos.csv
    $crono = New-Object System.Collections.Generic.List[object]
    $consumos = New-Object System.Collections.Generic.List[double]
    $ramMonitor = 0.0
    $estadoFps = ''; $estadoProcesos = ''
    $gpuEp = $null; $batEp = $null
    foreach ($f in $ev.Filas) {
        $tipo = $f[$eTipo]; $det = $f[$eDet]
        $t = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($f[$eH], $fmtHora, $inv, [Globalization.DateTimeStyles]::None, [ref]$t)) { continue }
        if ($tipo -eq 'CONSUMO_MONITOR') {
            if ($det -match 'CPU ([\d\.]+) %, RAM (\d+) MB') { $consumos.Add([double]::Parse($matches[1], $inv)); $ramMonitor = [math]::Max($ramMonitor, [double]$matches[2]) }
        }
        elseif ($tipo -eq 'FPS') { $estadoFps = $det -replace '^[^:]*:\s*', '' }
        elseif ($tipo -eq 'PROCESOS') { $estadoProcesos = $det -replace '^[^:]*:\s*', '' }
        elseif ($tipo -eq 'GPU') {
            # Las limitaciones de la GPU se agrupan en episodios para no llenar la cronología
            if ($det -like 'Limitación por hardware:*') {
                $razon = $det -replace '^Limitación por hardware:\s*', ''
                if ($gpuEp -ne $null -and $gpuEp.Razon -eq $razon -and ($t - $gpuEp.Ultima).TotalSeconds -le 180) { $gpuEp.Veces = $gpuEp.Veces + 1; $gpuEp.Abierta = $t }
                else {
                    if ($gpuEp -ne $null) { $crono.Add([pscustomobject]@{ T = $gpuEp.Desde; Hora = (Corta $gpuEp.Desde); Tipo = 'GPU'; Detalle = ('Limitada por ' + $gpuEp.Razon + ' ' + $gpuEp.Veces + ' veces hasta las ' + $gpuEp.Ultima.ToString('HH:mm:ss') + ' (' + (Dur $gpuEp.Seg) + ' en total)') }) }
                    $gpuEp = [pscustomobject]@{ Razon = $razon; Desde = $t; Ultima = $t; Abierta = $t; Veces = 1; Seg = 0.0 }
                }
            }
            elseif ($gpuEp -ne $null) { $gpuEp.Seg = $gpuEp.Seg + ($t - $gpuEp.Abierta).TotalSeconds; $gpuEp.Ultima = $t }
        }
        elseif ($tipo -eq 'BATERIA' -and $det -match '^(\d+) % -> (\d+) %\s*(.*)$') {
            # Los cambios de porcentaje seguidos en el mismo sentido se agrupan en un tramo
            $de = [int]$matches[1]; $a = [int]$matches[2]; $suf = $matches[3]
            if ($batEp -ne $null -and $batEp.Suf -eq $suf -and $batEp.A -eq $de -and [math]::Sign($a - $de) -eq $batEp.Signo) { $batEp.A = $a; $batEp.Hasta = $t; $batEp.Pasos = $batEp.Pasos + 1 }
            else {
                if ($batEp -ne $null) { $crono.Add([pscustomobject]@{ T = $batEp.Desde; Hora = (Corta $batEp.Desde); Tipo = 'BATERÍA'; Detalle = ($batEp.De.ToString() + ' % → ' + $batEp.A + ' % ' + $batEp.Suf + $(if ($batEp.Pasos -gt 1) { ', hasta las ' + $batEp.Hasta.ToString('HH:mm:ss') } else { '' })) }) }
                $batEp = [pscustomobject]@{ De = $de; A = $a; Suf = $suf; Desde = $t; Hasta = $t; Pasos = 1; Signo = [math]::Sign($a - $de) }
            }
        }
        else { $crono.Add([pscustomobject]@{ T = $t; Hora = (Corta $t); Tipo = $tipo; Detalle = $det }) }
    }
    if ($gpuEp -ne $null) { $crono.Add([pscustomobject]@{ T = $gpuEp.Desde; Hora = (Corta $gpuEp.Desde); Tipo = 'GPU'; Detalle = ('Limitada por ' + $gpuEp.Razon + ' ' + $gpuEp.Veces + ' veces hasta las ' + $gpuEp.Ultima.ToString('HH:mm:ss') + ' (' + (Dur $gpuEp.Seg) + ' en total)') }) }
    if ($batEp -ne $null) { $crono.Add([pscustomobject]@{ T = $batEp.Desde; Hora = (Corta $batEp.Desde); Tipo = 'BATERÍA'; Detalle = ($batEp.De.ToString() + ' % → ' + $batEp.A + ' % ' + $batEp.Suf + $(if ($batEp.Pasos -gt 1) { ', hasta las ' + $batEp.Hasta.ToString('HH:mm:ss') } else { '' })) }) }
    $crono = @($crono | Sort-Object T)

    # ================================================================== hallazgos

    Hallazgo $nivelCierre ('Final de la sesión: ' + $cierre + '.')
    if ($huecos.Count -gt 0) { Hallazgo 'atencion' ('Hubo ' + $huecos.Count + ' huecos sin muestras (' + (Dur $segHuecos) + ' en total): el equipo pudo suspenderse o el monitor quedarse congelado.') }
    else { Hallazgo 'ok' ('El monitor grabó sin huecos: ' + (Fmt $n) + ' muestras en ' + (Dur $durSeg) + ' (una cada ' + $intervalo + ' s).') }
    if (-not $esAdmin) { Hallazgo 'info' 'Sesión sin permisos de administrador: faltan los FPS, la corriente y la temperatura reales de la batería (datos de MSI), y pueden faltar procesos muy cortos.' }

    foreach ($e in @($evWin | Where-Object { $_.Nivel -eq 'grave' })) { Hallazgo 'grave' ($e.Hora + ' · ' + $e.Texto) }
    if ($segBatMal -gt 0) { Hallazgo 'grave' ('La batería dejó de responder durante ' + (Dur $segBatMal) + '.') }

    if ($hayMsi -and $mahAc -gt 20) {
        $wh = $mahAc * 12.0 / 1000
        $txt = 'Estando enchufada, la batería aportó ' + (Fmt $mahAc 0 ' mAh') + ' (≈ ' + (Fmt $wh 1 ' Wh') + ') en ' + (Dur $segDescAc) + ' en total; el peor momento fue ' + (Fmt $maPeor 0 ' mA') + '.'
        if ($umbral -like '<*') { $txt += ' Se descarga incluso con menos de 40 W de CPU+GPU: el cargador apenas da para el equipo en reposo.' }
        elseif ($umbral -match '(\d+)') { $txt += ' Empieza a descargarse cuando la CPU y la GPU juntas pasan de unos ' + $matches[1] + ' W (tramo ' + $umbral + '): por encima de eso el cargador, o el límite que le pone el equipo, no da abasto.' }
        Hallazgo 'atencion' $txt
    }
    elseif (-not $hayMsi -and $perdidoAc -gt 200) {
        Hallazgo 'atencion' ('Estando enchufada, la batería perdió ' + (Fmt $perdidoAc 0 ' mWh') + ' según Windows. Sin administrador no se ve la corriente real; ábrelo como administrador para saber a partir de qué consumo pasa.')
    }
    elseif ($hayMsi) { Hallazgo 'ok' 'Con el cargador conectado, la batería no tuvo que aportar energía.' }

    if ($segSinCargar -gt 600) {
        Hallazgo 'atencion' ('Estuvo ' + (Dur $segSinCargar) + ' enchufada sin cargar por debajo del 95 % (llegó a ' + $pctSinCargarMin + ' %). Puede ser el modo de carga de MSI Center o que el equipo espera a bajar de cierto nivel para volver a cargar.')
    }
    if ($cambiosFcc.Count -gt 0) { Hallazgo 'atencion' ('Windows cambió ' + $cambiosFcc.Count + ' veces la capacidad completa de la batería: el medidor se está recalibrando (suele pasar tras un apagado brusco).') }
    $saludWin = $NaN
    if ($info['Batería'] -match 'diseño (\d+) mWh') { $saludWin = 100.0 * $fccPrev / [double]$matches[1] }
    $saludChip = $NaN
    if ($hayMsi) {
        $ult = $snap[$snap.Count - 1]
        if ($ult.Diseno -gt 0 -and $ult.Fcc -gt 0) { $saludChip = 100.0 * $ult.Fcc / $ult.Diseno }
        $t1 = 'Según el chip de la batería: capacidad completa ' + (Fmt $ult.Fcc 0 ' mAh') + ' de ' + (Fmt $ult.Diseno 0 ' mAh') + ' de diseño (' + (Fmt $saludChip 1 ' %') + ').'
        if (-not [double]::IsNaN($saludWin) -and [math]::Abs($saludWin - $saludChip) -gt 15) { $t1 += ' Windows indica ' + (Fmt $saludWin 1 ' %') + ': ese dato lo calcula el controlador del equipo y se reinicia tras cada apagado brusco, así que el del chip es más fiable.' }
        $nivelSalud = 'info'
        if ($saludChip -lt 60) { $nivelSalud = 'atencion' }
        if ($saludChip -lt 40) { $nivelSalud = 'grave' }
        Hallazgo $nivelSalud $t1
        if (-not [double]::IsNaN($tempMax)) {
            if ($tempMax -gt 50) { Hallazgo 'grave' ('La batería llegó a ' + (Fmt $tempMax 1 ' °C') + ': está demasiado caliente.') }
            elseif ($tempMax -gt 45) { Hallazgo 'atencion' ('La batería llegó a ' + (Fmt $tempMax 1 ' °C') + ': algo caliente.') }
            else { Hallazgo 'ok' ('Temperatura de la batería normal: entre ' + (Fmt $tempMin 1) + ' y ' + (Fmt $tempMax 1 ' °C') + '.') }
        }
    }
    if ($voltMin -lt 10800) { Hallazgo 'atencion' ('El voltaje de la batería bajó a ' + (Fmt ($voltMin / 1000) 2 ' V') + ': muy bajo para una batería de 3 celdas.') }
    if ($segBat -gt 60) { Hallazgo 'info' ('Funcionó con batería (sin cargador) durante ' + (Dur $segBat) + '.') }

    $segFreno = 0; $segFuente = 0; $segTermica = 0
    foreach ($k in $razones.Keys) {
        if ($k -like '*fuente*') { $segFuente += $razones[$k] }
        if ($k -like '*freno de hardware*' -or $k -like '*temperatura (hardware)*') { $segFreno += $razones[$k] }
        if ($k -like '*temperatura*') { $segTermica += $razones[$k] }
    }
    if ($segFuente -gt 0) { Hallazgo 'grave' ('La GPU se frenó ' + (Dur $segFuente) + ' por una señal de la fuente de alimentación: el equipo detectó que no le llegaba suficiente energía.') }
    if ($segGT86 -gt 300) { Hallazgo 'atencion' ('La GPU pasó ' + (Dur $segGT86) + ' a 86 °C o más (máximo ' + (Fmt $maxGT 1 ' °C') + ') y se limitó por temperatura ' + (Dur $segTermica) + $(if (-not [double]::IsNaN($gpuFanMax)) { ', con su ventilador al ' + (Fmt $gpuFanMax 0 ' %') } else { '' }) + '. La refrigeración va justa: limpieza del polvo, pasta térmica o limitar los FPS.') }
    if ($segFreno -gt 0) { Hallazgo 'atencion' ('La GPU activó su freno de hardware durante ' + (Dur $segFreno) + ' (temperatura o potencia al límite).') }
    if ($maxMhz -gt $mhzBase * 1.05) { Hallazgo 'info' ('El Turbo del CPU estuvo activo: llegó a ' + (Fmt $maxMhz 0 ' MHz') + '.') }
    else { Hallazgo 'info' ('El Turbo del CPU estuvo desactivado: no pasó de ' + (Fmt $maxMhz 0 ' MHz') + '.') }
    # Una fila por aplicación (la ventana de tiempo alrededor del mínimo puede tener varias lecturas)
    $enRamMin = @($enRamMin | Group-Object Aplicación | ForEach-Object { $_.Group | Sort-Object _Priv -Descending | Select-Object -First 1 })
    if ($ramMinT -ne $null -and $ramMin -lt 300) {
        $quien = (@($enRamMin | Sort-Object _Priv -Descending | Select-Object -First 3 | ForEach-Object { $_.Aplicación + ' (' + $_.'Memoria privada' + ')' })) -join ', '
        Hallazgo 'atencion' ('La memoria RAM libre bajó a ' + (Fmt $ramMin 0 ' MB') + ' a las ' + $ramMinT.ToString('HH:mm:ss') + ' y estuvo por debajo de 500 MB durante ' + (Dur $segRam500) + '.' + $(if ($quien) { ' En ese momento lo que más ocupaba: ' + $quien + '.' } else { '' }))
    }
    foreach ($e in @($evWin | Where-Object { $_.Nivel -eq 'atencion' } | Group-Object Texto)) { Hallazgo 'atencion' ($e.Group[0].Hora + ' · ' + $e.Name + $(if ($e.Count -gt 1) { ' (' + $e.Count + ' veces)' } else { '' })) }
    foreach ($e in @($evWin | Where-Object { $_.Nivel -eq 'info' } | Group-Object Texto)) { Hallazgo 'info' ($e.Name + $(if ($e.Count -gt 1) { ' (' + $e.Count + ' veces)' } else { '' })) }
    if ($fallos.Count -gt 0) { Hallazgo 'atencion' ('Se cerraron por un error ' + $fallos.Count + ' procesos: ' + ((@($fallos | Group-Object Proceso | ForEach-Object { $_.Name + ' (' + $_.Count + ')' })) -join ', ') + '.') }
    foreach ($o in $noDeseados) {
        Hallazgo 'atencion' ('Programa no deseado activo: ' + $o.Programa + ' se ejecutó ' + $o.Veces + ' veces. Ubicación: ' + $o.Ubicación + '. Pertenece a: ' +
            $o.'Programa instalado' + '. Si no lo instalaste a propósito, desinstálalo así: ' + $o.'Cómo quitarlo')
    }
    foreach ($o in $raros) {
        Hallazgo 'atencion' ('Proceso a revisar: ' + $o.Programa + ', ' + $o.Firma + '. Ubicación: ' + $o.Ubicación + '. Pertenece a: ' + $o.'Programa instalado' + '. ' + $o.'Cómo quitarlo')
    }
    if ($sinFirmaConocidos.Count -gt 0) {
        Hallazgo 'info' ('Sin firma digital, pero pertenecen a programas instalados: ' + ((@($sinFirmaConocidos | ForEach-Object { $_.Programa + ' (' + ($_.'Programa instalado' -replace ',.*$', '') + ')' })) -join ', ') + '. Es habitual en juegos.')
    }
    $deTerceros = @($tablaInstalados | Where-Object { $_.Tipo -ne 'Controlador' })
    if ($deTerceros.Count -gt 0) {
        Hallazgo 'info' ('Además de Windows, se ejecutaron programas de ' + $deTerceros.Count + ' aplicaciones instaladas. En el detalle, «Programas instalados detrás de lo que se ejecutó», tienes la ubicación de cada uno y cómo desinstalarlo.')
    }
    if ($raros.Count -eq 0 -and $tablaLanz.Count -gt 0) { Hallazgo 'ok' ('Los ' + $totalInicios + ' procesos que se abrieron tienen firma válida o son de Windows; ' + $totalFugaces + ' duraron menos de 2 segundos.') }
    $juegosRapidos = @($tablaFps | Where-Object { $_._Media -gt 90 -and $_.Proceso -match '(?i)rocketleague|aoe|reliccardinal|frostpunk' })
    if ($juegosRapidos.Count -gt 0 -and ($mahAc -gt 20 -or $segGT86 -gt 300)) {
        Hallazgo 'info' ('Sugerencia: ' + (($juegosRapidos | ForEach-Object { $_.Proceso + ' iba a ' + $_.'FPS medio' + ' FPS' }) -join ' y ') + '. Limitar los FPS (a 60 o 90) baja el consumo y la temperatura, y así la batería no tiene que ayudar al cargador.')
    }

    $orden = @{ grave = 0; atencion = 1; info = 2; ok = 3 }
    $hallazgos = @($hallazgos | Sort-Object { $orden[$_.Nivel] }, I)

    # ================================================================== indicadores

    $nivelBat = ''
    if ($pctFin -lt $pctIni) { $nivelBat = 'atencion' }
    Kpi 'Grabado' (Dur $durSeg) ((Fmt $n) + ' muestras · ' + (Fmt $cobertura 1 ' %') + ' de cobertura') $(if ($huecos.Count -gt 0) { 'atencion' } else { 'ok' })
    Kpi 'Batería' ("$pctIni % → $pctFin %") ('mín. ' + $pctMin + ' % · cargando ' + (Dur $segCarga)) $nivelBat
    if ($hayMsi) {
        Kpi 'Aportó enchufada' (Fmt $mahAc 0 ' mAh') ('≈ ' + (Fmt ($mahAc * 12 / 1000) 1 ' Wh') + ' · peor ' + (Fmt $maPeor 0 ' mA')) $(if ($mahAc -gt 20) { 'atencion' } else { 'ok' })
        Kpi 'Se descarga desde' $(if ($umbral) { $umbral } else { 'no se descargó' }) 'CPU+GPU, media de 10 s' $(if ($umbral) { 'atencion' } else { 'ok' })
        Kpi 'Salud (chip)' (Fmt $saludChip 1 ' %') ('Windows dice ' + (Fmt $saludWin 1 ' %')) ''
        Kpi 'Temp. batería' ((Fmt $tempMin 1) + '–' + (Fmt $tempMax 1 ' °C')) 'según el chip' $(if ($tempMax -gt 45) { 'atencion' } else { 'ok' })
    }
    else {
        Kpi 'Perdida enchufada' (Fmt $perdidoAc 0 ' mWh') 'según Windows (sin datos de MSI)' $(if ($perdidoAc -gt 200) { 'atencion' } else { '' })
    }
    Kpi 'CPU + GPU' ((Fmt ($sumW / [math]::Max(1, $nW)) 0) + ' / ' + (Fmt $maxW 0 ' W')) 'media / máxima' ''
    Kpi 'GPU' (Fmt $maxGT 0 ' °C') ((Dur $segGT86) + ' a ≥ 86 °C') $(if ($segGT86 -gt 300) { 'atencion' } else { '' })
    Kpi 'RAM libre mínima' (Fmt $ramMin 0 ' MB') ('memoria comprometida máx. ' + (Fmt $comMax 0 ' %')) $(if ($ramMin -lt 300) { 'atencion' } else { '' })

    # ================================================================== resumen (HTML)

    $sbR = New-Object System.Text.StringBuilder
    [void]$sbR.Append('<div class="kpis">')
    foreach ($k in $kpis) { [void]$sbR.Append('<div class="kpi ' + $k.Nivel + '"><div class="e">' + (Hx $k.Etiqueta) + '</div><div class="v">' + (Hx $k.Valor) + '</div><div class="nota">' + (Hx $k.Nota) + '</div></div>') }
    [void]$sbR.Append('</div><h3>Lo que encontró</h3><ul class="hallazgos">')
    $etiquetas = @{ grave = 'Grave'; atencion = 'Atención'; info = 'Info'; ok = 'Bien' }
    foreach ($h in $hallazgos) { [void]$sbR.Append('<li class="' + $h.Nivel + '"><span class="chip">' + $etiquetas[$h.Nivel] + '</span>' + (Hx $h.Texto) + '</li>') }
    [void]$sbR.Append('</ul>')

    $consumoTxt = 'sin datos'
    if ($consumos.Count -gt 0) { $consumoTxt = 'CPU ' + (Fmt (($consumos | Measure-Object -Average).Average) 2 ' %') + ' de media (máx. ' + (Fmt (($consumos | Measure-Object -Maximum).Maximum) 2 ' %') + '), RAM hasta ' + (Fmt $ramMonitor 0 ' MB') }
    $estado = @(
        [pscustomobject]@{ Dato = 'Inicio'; Valor = $t0.ToString('dddd dd/MM/yyyy HH:mm:ss', $es) },
        [pscustomobject]@{ Dato = 'Última muestra'; Valor = $tUlt.ToString('dddd dd/MM/yyyy HH:mm:ss', $es) },
        [pscustomobject]@{ Dato = 'Duración'; Valor = (Dur $durSeg) },
        [pscustomobject]@{ Dato = 'Muestras'; Valor = ((Fmt $n) + ' de ' + (Fmt $esperadas) + ' esperadas (' + (Fmt $cobertura 1 ' %') + '), ' + $huecos.Count + ' huecos') },
        [pscustomobject]@{ Dato = 'Cómo terminó'; Valor = $cierre },
        [pscustomobject]@{ Dato = 'Permisos'; Valor = $(if ($esAdmin) { 'Administrador' } else { 'Usuario normal' }) },
        [pscustomobject]@{ Dato = 'FPS'; Valor = $(if ($estadoFps) { $estadoFps } elseif ($esAdmin) { 'sin datos' } else { 'requiere administrador' }) },
        [pscustomobject]@{ Dato = 'Procesos'; Valor = $(if ($estadoProcesos) { $estadoProcesos } else { 'revisados cada ' + $intervalo + ' s' }) },
        [pscustomobject]@{ Dato = 'Datos de MSI'; Valor = $(if ($hayMsi) { $crudo.Count.ToString() + ' valores de ' + (@($crudo.Keys | ForEach-Object { ($_ -split '\|')[0] } | Select-Object -Unique)).Count + ' clases' } elseif ($esAdmin) { 'sin datos' } else { 'requiere administrador' }) },
        [pscustomobject]@{ Dato = 'Consumo del monitor'; Valor = $consumoTxt }
    )
    [void]$sbR.Append('<div class="dos"><div><h3>Estado del monitor</h3>' + (Html-Tabla $estado @('Dato', 'Valor')) + '</div>')

    $tablaJuegos = @(foreach ($jj in $juegos) {
        $fpsTxt = 'n/d'
        if ($jj.NFps -gt 0) { $fpsTxt = Fmt ($jj.SumFps / $jj.NFps) 0 }
        [pscustomobject]@{ Juego = $jj.Juego; Desde = $jj.Desde.ToString('HH:mm'); Hasta = $jj.Hasta.ToString('HH:mm'); Duración = (Dur ($jj.Hasta - $jj.Desde).TotalSeconds); 'FPS medio' = $fpsTxt; 'CPU+GPU medio' = (Fmt ($jj.SumW / [math]::Max(1, $jj.NW)) 0 ' W'); 'GPU máx.' = (Fmt $jj.MaxGT 0 ' °C') }
    })
    [void]$sbR.Append('<div><h3>Juegos</h3>' + (Html-Tabla $tablaJuegos @('Juego', 'Desde', 'Hasta', 'Duración', 'FPS medio', 'CPU+GPU medio', 'GPU máx.') @('FPS medio', 'CPU+GPU medio', 'GPU máx.')) + '</div></div>')
    [void]$sbR.Append('<h3>Aplicaciones abiertas</h3>' + (Html-Tabla @($tablaApps | Select-Object -First 10) @('Aplicación', 'Tiempo abierta', 'CPU media', 'Memoria máx.', 'Ubicación') @('CPU media', 'Memoria máx.')))
    [void]$sbR.Append('<h3>En segundo plano (los que más CPU usaron)</h3>' + (Html-Tabla @($tablaFondo | Select-Object -First 10) @('Proceso', 'CPU media', 'CPU máxima', 'Memoria máx.', 'Ubicación') @('CPU media', 'CPU máxima', 'Memoria máx.')))

    # ================================================================== detalle (HTML)

    Seccion 'Cronología' (Html-Tabla $crono @('Hora', 'Tipo', 'Detalle'))

    $tablaVent = @(foreach ($vk in $ventanas.Keys) {
        $v = $ventanas[$vk]
        $jv = ''
        if ($v.Juegos.Count -gt 0) { $jv = ($v.Juegos.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key }
        $cambio = $v.CapFin - $v.CapIni
        [pscustomobject]@{ Desde = $v.Inicio.ToString('dd/MM HH:mm'); Juego = $jv; 'Cambio de batería' = ((Fmt $cambio 0) + ' mWh'); 'Corriente media' = $(if ($v.NMa -gt 0) { Fmt ($v.SumMa / $v.NMa) 0 ' mA' } else { 'n/d' }); 'CPU+GPU medio' = (Fmt ($v.SumW / [math]::Max(1, $v.NW)) 0 ' W'); 'CPU+GPU máx.' = (Fmt $v.MaxW 0 ' W'); 'GPU máx.' = (Fmt $v.MaxGT 0 ' °C'); 'FPS medio' = $(if ($v.NFps -gt 0) { Fmt ($v.SumFps / $v.NFps) 0 } else { '' }) }
    })
    $htmlBat = '<p>Batería según Windows: ' + "$pctIni % → $pctFin %" + ' (' + (Fmt $capIni 0) + ' → ' + (Fmt $capFin 0 ' mWh') + '), capacidad completa ' + (Fmt $fccPrev 0 ' mWh') + '. Cargando: ' + (Dur $segCarga) + '. Con cargador: ' + (Dur $segAc) + '; sin cargador: ' + (Dur $segBat) + '. Perdido estando enchufada: ' + (Fmt $perdidoAc 0 ' mWh') + '; cargado: ' + (Fmt $ganado 0 ' mWh') + '. Voltaje: ' + (Fmt ($voltMin / 1000) 2) + '–' + (Fmt ($voltMax / 1000) 2 ' V') + '.</p>'
    $htmlBat += '<h4>Cada 20 minutos</h4>' + (Html-Tabla $tablaVent @('Desde', 'Juego', 'Cambio de batería', 'Corriente media', 'CPU+GPU medio', 'CPU+GPU máx.', 'GPU máx.', 'FPS medio') @('Cambio de batería', 'Corriente media', 'CPU+GPU medio', 'CPU+GPU máx.', 'GPU máx.', 'FPS medio'))
    if ($hayMsi) {
        $tablaBandas = @(foreach ($nb in $bandas.Keys) {
            $bd = $bandas[$nb]
            if ($bd.N -eq 0) { continue }
            [pscustomobject]@{ 'CPU+GPU (10 s)' = $nb; Muestras = $bd.N; 'Corriente media' = (Fmt ($bd.SumMa / $bd.N) 0 ' mA'); 'Peor' = (Fmt $bd.MinMa 0 ' mA'); 'Tiempo descargando' = (Fmt (100.0 * $bd.NDesc / $bd.N) 0 ' %'); 'Aporte aprox.' = (Fmt (-12.0 * $bd.SumMa / $bd.N / 1000) 2 ' W') }
        })
        $htmlBat += '<h4>Corriente real de la batería según el consumo (solo con cargador)</h4><p class="tenue">Corriente negativa: la batería ayuda al cargador. Cada fila agrupa las lecturas de MSI (cada ~10 s) según la media de CPU+GPU de los 10 segundos anteriores.</p>' + (Html-Tabla $tablaBandas @('CPU+GPU (10 s)', 'Muestras', 'Corriente media', 'Peor', 'Tiempo descargando', 'Aporte aprox.') @('Muestras', 'Corriente media', 'Peor', 'Tiempo descargando', 'Aporte aprox.'))
        $ult = $snap[$snap.Count - 1]; $pri = $snap[0]
        $htmlBat += '<p>Según el chip de la batería: carga ' + (Fmt $pri.Pct 0) + ' % → ' + (Fmt $ult.Pct 0 ' %') + ', restante ' + (Fmt $pri.Rem 0) + ' → ' + (Fmt $ult.Rem 0 ' mAh') + ', capacidad completa ' + (Fmt $ult.Fcc 0 ' mAh') + ' de ' + (Fmt $ult.Diseno 0 ' mAh') + '. Salió de la batería: ' + (Fmt $mahAc 0 ' mAh') + ' con cargador y ' + (Fmt $mahBat 0 ' mAh') + ' sin él; entró al cargar: ' + (Fmt $mahCarga 0 ' mAh') + '.</p>'
    }
    if ($cambiosFcc.Count -gt 0) { $htmlBat += '<h4>Cambios de la capacidad completa según Windows</h4><ul>' + (($cambiosFcc | ForEach-Object { '<li>' + (Hx $_) + '</li>' }) -join '') + '</ul>' }
    if ($huecos.Count -gt 0) { $htmlBat += '<h4>Huecos sin muestras</h4>' + (Html-Tabla $huecos @('Desde', 'Hasta', 'Duración')) }
    Seccion 'Batería y cargador' $htmlBat

    $tablaRaz = @(foreach ($k in $razones.Keys) { [pscustomobject]@{ Motivo = $k; Tiempo = (Dur $razones[$k]); _S = $razones[$k] } }) | Sort-Object _S -Descending
    $htmlTemp = '<p>GPU: máximo ' + (Fmt $maxGT 1 ' °C') + ', ' + (Dur $segGT86) + ' a 86 °C o más. Zona térmica ACPI: máximo ' + (Fmt $maxAcpi 1 ' °C') + '. CPU: hasta ' + (Fmt $maxMhz 0 ' MHz') + ' y ' + (Fmt $maxCpuW 1 ' W') + '.'
    if ($hayMsi) { $htmlTemp += ' Según el controlador del equipo: CPU hasta ' + (Fmt $cpuCMax 0 ' °C') + ' (ventilador hasta ' + (Fmt $cpuFanMax 0 ' %') + '), GPU hasta ' + (Fmt $gpuCMax 0 ' °C') + ' (ventilador hasta ' + (Fmt $gpuFanMax 0 ' %') + ').' }
    $htmlTemp += '</p><h4>Por qué se limitó la GPU</h4>' + (Html-Tabla $tablaRaz @('Motivo', 'Tiempo') @('Tiempo'))
    Seccion 'Temperaturas y limitaciones' $htmlTemp

    Seccion 'FPS' ('<p class="tenue">Cada proceso que dibuja en pantalla. Que Chrome o Discord dibujen mientras juegas le quita GPU al juego.</p>' + (Html-Tabla $tablaFps @('Proceso', 'Tiempo dibujando', 'FPS medio', 'FPS en el 5 % peor', 'Segundos con tirones (>50 ms)', 'Peor cuadro') @('FPS medio', 'FPS en el 5 % peor', 'Segundos con tirones (>50 ms)', 'Peor cuadro')))

    $htmlMem = '<p>RAM libre mínima: ' + (Fmt $ramMin 0 ' MB') + $(if ($ramMinT -ne $null) { ' a las ' + $ramMinT.ToString('HH:mm:ss') } else { '' }) + '; por debajo de 500 MB durante ' + (Dur $segRam500) + '. Memoria comprometida máxima: ' + (Fmt $comMax 1 ' %') + '. Memoria privada de todos los procesos: hasta ' + (Fmt $maxTotPriv 0 ' MB') + '.</p>'
    if ($enRamMin.Count -gt 0) { $htmlMem += '<h4>Lo que más memoria ocupaba en ese momento</h4>' + (Html-Tabla @($enRamMin | Sort-Object _Priv -Descending | Select-Object -First 10) @('Aplicación', 'Instancias', 'Memoria privada', 'En RAM') @('Instancias', 'Memoria privada', 'En RAM')) }
    Seccion 'Memoria' $htmlMem

    Seccion 'Programas instalados detrás de lo que se ejecutó' ('<p class="tenue">A qué programa pertenece cada ejecutable que se vio funcionar (sin contar los de Windows) y cómo desinstalarlo. «Probable» quiere decir que se dedujo por la carpeta del fabricante. Antes de desinstalar algo, asegúrate de que no lo necesitas.</p>' +
        (Html-Tabla $tablaInstalados @('Programa instalado', 'Versión', 'Tipo', 'Ejecutables', 'Carpeta', 'Cómo quitarlo', 'Comando de desinstalación')))

    $htmlProc = '<h4>Aplicaciones (todas las vigiladas)</h4>' + (Html-Tabla $tablaApps @('Aplicación', 'Tiempo abierta', 'CPU media', 'CPU máxima', 'Memoria máx.', 'Procesos', 'Ubicación', 'Programa instalado') @('CPU media', 'CPU máxima', 'Memoria máx.', 'Procesos'))
    $htmlProc += '<h4>Procesos en segundo plano que usaron CPU</h4><p class="tenue">Solo se registran los que estaban entre los que más CPU usaban en cada momento.</p>' + (Html-Tabla $tablaFondo @('Proceso', 'Veces entre los que más CPU usaban', 'CPU media', 'CPU máxima', 'Memoria máx.', 'Ubicación', 'Programa instalado') @('Veces entre los que más CPU usaban', 'CPU media', 'CPU máxima', 'Memoria máx.'))
    $htmlProc += '<h4>Programas que se abrieron (' + $totalInicios + ', ' + $totalFugaces + ' duraron menos de 2 s)</h4>' + (Html-Tabla $tablaLanz @('Programa', 'Veces', 'Duraron < 2 s', 'Lo abrió', 'Ubicación', 'Programa instalado', 'Firma', 'Cómo quitarlo') @('Veces', 'Duraron < 2 s'))
    if ($fallos.Count -gt 0) { $htmlProc += '<h4>Procesos que se cerraron por un error</h4>' + (Html-Tabla $fallos @('Hora', 'Proceso', 'Código', 'Significado', 'Ubicación')) }
    if ($antiguo) { $htmlProc += '<p class="tenue">Sesión de una versión anterior del monitor: solo hay inicios de procesos, sin cierres.</p>' }
    Seccion 'Aplicaciones y procesos en segundo plano' $htmlProc

    Seccion 'Eventos de Windows' ('<p class="tenue">Apagados bruscos, errores de hardware, bloqueos de la GPU, programas que fallaron y límites del firmware durante la sesión (y en el arranque siguiente).</p>' + (Html-Tabla $evWin @('Hora', 'Origen', 'Id', 'Texto')))

    $colsFinal = @('fecha_hora', 'cargador', 'cargando', 'bateria_pct_windows', 'capacidad_mwh', 'voltaje_mv', 'cpu_mas_gpu_w', 'gpu_temp_c', 'gpu_razones', 'ram_libre_mb', 'juego', 'juego_fps')
    $finFilas = @(foreach ($f in $cola) {
        $o = [ordered]@{}
        foreach ($nc in $colsFinal) { if ($c.ContainsKey($nc)) { $o[$nc] = $f[$c[$nc]] } }
        [pscustomobject]$o
    })
    $titFinal = 'Últimos 30 segundos grabados'
    if ($abrupto) { $titFinal = 'Últimos 30 segundos antes del apagado brusco' }
    Seccion $titFinal (Html-Tabla $finFilas @($colsFinal | Where-Object { $c.ContainsKey($_) }))

    if ($crudo.Count -gt 0) {
        $tablaCrudo = @(foreach ($k in $crudo.Keys) {
            $cr = $crudo[$k]
            $partes = $k -split '\|'
            [pscustomobject]@{ Clase = $partes[0]; Índice = $partes[1]; Significado = [string]$significado[$k]; Primero = $cr.Primero; Mínimo = (Fmt $cr.Min 0); Máximo = (Fmt $cr.Max 0); Último = $cr.Ultimo; Cambios = $cr.Cambios }
        })
        Seccion 'Datos de MSI (todos los valores)' ('<p class="tenue">MSI no documenta estos valores: el significado está deducido comparándolos con los datos de Windows.</p>' + (Html-Tabla $tablaCrudo @('Clase', 'Índice', 'Significado', 'Primero', 'Mínimo', 'Máximo', 'Último', 'Cambios') @('Mínimo', 'Máximo', 'Cambios')))
    }

    return [pscustomobject]@{
        Nombre = $ses.Nombre; Inicio = $t0; Fin = $tUlt; Duracion = $durSeg; Cierre = $cierre; NivelCierre = $nivelCierre; Abrupto = $abrupto
        Hallazgos = $hallazgos; Kpis = $kpis; Resumen = $sbR.ToString(); Detalle = $detalle.ToString()
        PctIni = $pctIni; PctFin = $pctFin; Juegos = @($juegos | ForEach-Object { $_.Juego } | Select-Object -Unique)
    }
}

# ------------------------------------------------------------------ sesiones a analizar

if (-not (Test-Path -LiteralPath $Registros)) { Write-Host "No hay registros en $Registros. Abre primero el monitor."; exit 1 }
$activa = ''
$rutaActiva = Join-Path $Registros 'sesion_activa.txt'
if (Test-Path -LiteralPath $rutaActiva) { $activa = ([IO.File]::ReadAllText($rutaActiva)).Trim() }
$monitorAbierto = @(Get-Process -Name MonitorBateria -ErrorAction SilentlyContinue).Count -gt 0
$todasSes = @(Get-ChildItem -LiteralPath $Registros -Directory | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}$' } | Sort-Object Name | ForEach-Object {
    $sis = Join-Path $_.FullName 'sistema.csv'
    $fin = $_.LastWriteTime
    if (Test-Path -LiteralPath $sis) { $fin = (Get-Item -LiteralPath $sis).LastWriteTime }
    [pscustomobject]@{ Nombre = $_.Name; Ruta = $_.FullName; Inicio = [datetime]::ParseExact($_.Name, 'yyyy-MM-dd_HH-mm-ss', $inv); Fin = $fin; EnCurso = ($_.Name -eq $activa -and $monitorAbierto) }
})
if ($Sesion) { $elegidas = @($todasSes | Where-Object { $_.Nombre -eq $Sesion }) ; $periodo = 'sesión ' + $Sesion }
elseif ($Todas) { $elegidas = $todasSes; $periodo = 'todas las sesiones' }
else { $limite = (Get-Date).AddHours(-$Horas); $elegidas = @($todasSes | Where-Object { $_.Fin -ge $limite }); $periodo = 'últimas ' + $Horas + ' horas' }

Write-Host ("Sesiones encontradas: {0}; a analizar ({1}): {2}" -f $todasSes.Count, $periodo, $elegidas.Count)
$resultados = New-Object System.Collections.Generic.List[object]
foreach ($s in $elegidas) {
    Write-Host ("Analizando {0}..." -f $s.Nombre)
    try { $resultados.Add((Analizar-Sesion $s)) }
    catch { Write-Host ("  No se pudo analizar: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
}

# ------------------------------------------------------------------ informe HTML

$css = @'
:root { color-scheme: light dark; --fondo:#f6f6f3; --tarjeta:#ffffff; --texto:#1d2025; --tenue:#62666d; --linea:#e2e2dc;
  --ok:#166534; --ok-f:#dcfce7; --at:#92400e; --at-f:#fef3c7; --gr:#991b1b; --gr-f:#fee2e2; --in:#1e40af; --in-f:#dbeafe; }
@media (prefers-color-scheme: dark) { :root { --fondo:#15171b; --tarjeta:#1e2127; --texto:#e7e7e4; --tenue:#a2a6ad; --linea:#30343b;
  --ok:#86efac; --ok-f:#173322; --at:#fcd34d; --at-f:#3a2d0e; --gr:#fca5a5; --gr-f:#3d1818; --in:#93c5fd; --in-f:#172a47; } }
* { box-sizing: border-box; }
body { margin: 0; background: var(--fondo); color: var(--texto); font: 15px/1.5 "Segoe UI", system-ui, sans-serif; }
main { max-width: 1180px; margin: 0 auto; padding: 24px 16px 48px; }
h1 { font-size: 26px; margin: 0 0 4px; } h2 { font-size: 20px; margin: 0 0 4px; } h3 { font-size: 16px; margin: 20px 0 8px; } h4 { font-size: 15px; margin: 18px 0 6px; }
.tenue { color: var(--tenue); }
.tarjeta { background: var(--tarjeta); border: 1px solid var(--linea); border-radius: 12px; padding: 18px 20px; margin: 18px 0; }
.cabecera { display: flex; flex-wrap: wrap; gap: 8px 16px; align-items: baseline; justify-content: space-between; }
.kpis { display: grid; grid-template-columns: repeat(auto-fill, minmax(170px, 1fr)); gap: 10px; margin-top: 14px; }
.kpi { border: 1px solid var(--linea); border-radius: 10px; padding: 10px 12px; }
.kpi .e { color: var(--tenue); font-size: 13px; } .kpi .v { font-size: 21px; font-weight: 600; font-variant-numeric: tabular-nums; } .kpi .nota { color: var(--tenue); font-size: 12.5px; }
.kpi.atencion { border-color: var(--at); } .kpi.grave { border-color: var(--gr); } .kpi.ok .v { color: var(--ok); } .kpi.atencion .v { color: var(--at); } .kpi.grave .v { color: var(--gr); }
ul.hallazgos { list-style: none; padding: 0; margin: 0; }
ul.hallazgos li { padding: 8px 12px; border-radius: 8px; margin: 6px 0; }
li.grave { background: var(--gr-f); } li.atencion { background: var(--at-f); } li.info { background: var(--in-f); } li.ok { background: var(--ok-f); }
.chip { display: inline-block; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .03em; margin-right: 8px; min-width: 64px; }
li.grave .chip { color: var(--gr); } li.atencion .chip { color: var(--at); } li.info .chip { color: var(--in); } li.ok .chip { color: var(--ok); }
.estado { font-weight: 600; padding: 3px 10px; border-radius: 999px; font-size: 13px; }
.estado.ok { background: var(--ok-f); color: var(--ok); } .estado.atencion { background: var(--at-f); color: var(--at); } .estado.grave { background: var(--gr-f); color: var(--gr); } .estado.info { background: var(--in-f); color: var(--in); }
.dos { display: grid; grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); gap: 0 24px; }
.desliza { overflow-x: auto; }
table { border-collapse: collapse; width: 100%; font-size: 13.5px; }
th, td { text-align: left; padding: 6px 8px; border-bottom: 1px solid var(--linea); vertical-align: top; }
th { color: var(--tenue); font-weight: 600; white-space: nowrap; } td.n { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
details { border-top: 1px solid var(--linea); } details:first-of-type { border-top: none; }
summary { cursor: pointer; font-weight: 600; padding: 10px 0; }
.cuerpo { padding-bottom: 12px; }
@media (max-width: 560px) { .dos { grid-template-columns: 1fr; } .kpis { grid-template-columns: repeat(2, 1fr); } }
'@

$html = New-Object System.Text.StringBuilder
[void]$html.Append('<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Informe del monitor de batería</title><style>' + $css + '</style></head><body><main>')
[void]$html.Append('<h1>Informe del monitor de batería</h1><p class="tenue">Generado el ' + (Hx (Get-Date).ToString('dddd dd/MM/yyyy HH:mm', $es)) + ' · ' + (Hx $periodo) + ' · ' + $resultados.Count + ' sesiones · ' + (Hx $Registros) + '</p>')

if ($resultados.Count -eq 0) {
    [void]$html.Append('<div class="tarjeta"><p>No hay sesiones en este periodo. Abre el monitor para grabar, o genera el informe con <code>-Todas</code> o <code>-Horas 72</code>.</p></div>')
}
else {
    # Resumen del periodo
    $totalSeg = ($resultados | Measure-Object Duracion -Sum).Sum
    $bruscos = @($resultados | Where-Object Abrupto).Count
    $graves = @($resultados | ForEach-Object { $_.Hallazgos } | Where-Object { $_.Nivel -eq 'grave' }).Count
    $atenciones = @($resultados | ForEach-Object { $_.Hallazgos } | Where-Object { $_.Nivel -eq 'atencion' }).Count
    [void]$html.Append('<div class="tarjeta"><h2>Resumen del periodo</h2><div class="kpis">')
    [void]$html.Append('<div class="kpi"><div class="e">Sesiones</div><div class="v">' + $resultados.Count + '</div><div class="nota">' + (Hx (Dur $totalSeg)) + ' grabadas</div></div>')
    [void]$html.Append('<div class="kpi ' + $(if ($bruscos -gt 0) { 'grave' } else { 'ok' }) + '"><div class="e">Apagados bruscos</div><div class="v">' + $bruscos + '</div><div class="nota">con el monitor abierto</div></div>')
    [void]$html.Append('<div class="kpi ' + $(if ($graves -gt 0) { 'grave' } else { 'ok' }) + '"><div class="e">Hallazgos graves</div><div class="v">' + $graves + '</div><div class="nota">' + $atenciones + ' para revisar</div></div>')
    [void]$html.Append('</div><h3>Sesiones</h3>')
    $filasSes = @(foreach ($res in $resultados) { [pscustomobject]@{ Sesión = $res.Inicio.ToString('dd/MM HH:mm') + ' → ' + $res.Fin.ToString('dd/MM HH:mm'); Duración = (Dur $res.Duracion); Batería = "$($res.PctIni) % → $($res.PctFin) %"; Juegos = ($res.Juegos -join ', '); Final = $res.Cierre } })
    [void]$html.Append((Html-Tabla $filasSes @('Sesión', 'Duración', 'Batería', 'Juegos', 'Final')) + '</div>')

    foreach ($res in ($resultados | Sort-Object Inicio -Descending)) {
        [void]$html.Append('<div class="tarjeta"><div class="cabecera"><h2>Sesión del ' + (Hx $res.Inicio.ToString('dd/MM HH:mm')) + ' al ' + (Hx $res.Fin.ToString('dd/MM HH:mm')) + ' · ' + (Hx (Dur $res.Duracion)) + '</h2><span class="estado ' + $res.NivelCierre + '">' + (Hx $res.Cierre) + '</span></div>')
        [void]$html.Append('<p class="tenue">Carpeta: ' + (Hx $res.Nombre) + '</p>' + $res.Resumen + '<h3>Detalle</h3>' + $res.Detalle + '</div>')
    }
}
[void]$html.Append('<p class="tenue">Informe generado en ' + (Fmt $cronometro.Elapsed.TotalSeconds 1 ' s') + ' por analizar.ps1. Para otro periodo: analizar.ps1 -Horas 72, -Todas o -Sesion AAAA-MM-DD_HH-MM-SS.</p></main></body></html>')

$dirInformes = Join-Path (Split-Path $Registros -Parent) 'informes'
[void](New-Item -ItemType Directory -Force -Path $dirInformes)
$salida = Join-Path $dirInformes ('informe_' + (Get-Date).ToString('yyyy-MM-dd_HH-mm', $inv) + '.html')
[IO.File]::WriteAllText($salida, $html.ToString(), (New-Object Text.UTF8Encoding $true))
Write-Host ("Informe guardado en: {0} ({1:N1} s)" -f $salida, $cronometro.Elapsed.TotalSeconds)
if (-not $NoAbrir) { Start-Process $salida }
