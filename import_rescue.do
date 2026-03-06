*******************************************************
* IMPORT RESCUE (Stata 19) — robusto para Excel
* Ruta fija: D:\LOGAN\TIE\RESUTADOS.xlsx
* Salidas de diagnóstico: import_log.txt
*******************************************************

version 19.0
clear all
set more off
set linesize 255

*------------------------------------------------------
* 0) RUTA / LOG
*------------------------------------------------------
local base "D:\LOGAN\TIE"
local xlsx "RESUTADOS.xlsx"     // <- forzado a tu archivo
local out  "`base'\salidas"

capture noisily mkdir "`out'"
cd "`out'"

log close _all
log using "import_log.txt", replace text

di as txt "PWD: " c(pwd)
di as txt "Listando contenido de `base':"
shell dir /b "D:\LOGAN\TIE"

* Verifica existencia
capture confirm file "`base'\`xlsx'"
if _rc {
    di as err "No existe el archivo: `base'\`xlsx'"
    di as txt "Confirma la ruta y el nombre exacto (extensión incluida)."
    log close
    exit 601
}

local full "`base'\`xlsx'"
di as res "Intentando importar: `full'"

*------------------------------------------------------
* 1) DESCRIBE HOJAS (visual)
*------------------------------------------------------
di as txt _n "==> Hojas disponibles (ver nombres exactos abajo):"
capture noisily import excel using "`full'", describe

*------------------------------------------------------
* 2) INTENTOS DE IMPORTACIÓN (A: xlsx nativo)
*    - Probar varias etiquetas de hoja comunes
*    - Probar sheet(1)
*    - Probar allstring si falla
*------------------------------------------------------
local tried 0
local ok    0

foreach s in "Hoja1" "Hoja 1" "Sheet1" "Sheet 1" {
    di as txt _n ">>> Intento: sheet(`s') firstrow"
    capture noisily import excel using "`full'", sheet("`s'") firstrow clear
    if !_rc {
        local ok 1
        local tried 1
        di as res "Éxito con sheet(`s') firstrow."
        continue, break
    }
}

if !`ok' {
    di as txt _n ">>> Intento: sheet(1) firstrow"
    capture noisily import excel using "`full'", sheet(1) firstrow clear
    if !_rc {
        local ok 1
        local tried 1
        di as res "Éxito con sheet(1) firstrow."
    }
}

if !`ok' {
    di as txt _n ">>> Reintentando con ALLSTRING (tipos a string) y firstrow"
    foreach s in "Hoja1" "Hoja 1" "Sheet1" "Sheet 1" {
        di as txt ">>> sheet(`s') allstring firstrow"
        capture noisily import excel using "`full'", sheet("`s'") allstring firstrow clear
        if !_rc {
            local ok 1
            local tried 1
            di as res "Éxito con sheet(`s') allstring firstrow."
            continue, break
        }
    }
    if !`ok' {
        di as txt ">>> sheet(1) allstring firstrow"
        capture noisily import excel using "`full'", sheet(1) allstring firstrow clear
        if !_rc {
            local ok 1
            local tried 1
            di as res "Éxito con sheet(1) allstring firstrow."
        }
    }
}

if `ok' {
    di as res _n "==> IMPORTACIÓN OK"
    di as txt "Obs: " _N
    di as txt "Variables:"
    describe
    list in 1/5, abbrev(24)
    log close
    exit
}

*------------------------------------------------------
* 3) PLAN B: ODBC (driver de Microsoft Excel)
*   Requiere el driver 'Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)'
*   Si no lo tienes, salta a la sección 4 (CSV).
*------------------------------------------------------
di as err _n "Aún falla import excel. Probando ODBC..."
capture noisily odbc load, exec("SELECT * FROM [Hoja1$]") ///
    connectionstring("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=`full';ReadOnly=1;") clear
if !_rc {
    di as res "ODBC OK con Hoja1$"
    describe
    list in 1/5, abbrev(24)
    log close
    exit
}

capture noisily odbc load, exec("SELECT * FROM [Sheet1$]") ///
    connectionstring("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=`full';ReadOnly=1;") clear
if !_rc {
    di as res "ODBC OK con Sheet1$"
    describe
    list in 1/5, abbrev(24)
    log close
    exit
}

di as err "ODBC también falló. Puede no estar instalado el driver o el libro estar protegido/corrupto."

*------------------------------------------------------
* 4) PLAN C: CSV (camino garantizado)
*   Abre el archivo en Excel y haz: Archivo > Guardar como > CSV (UTF-8)
*   Nómbralo: RESUTADOS.csv en D:\LOGAN\TIE\
*   Luego descomenta y ejecuta el bloque siguiente.
*------------------------------------------------------

/***
di as txt _n ">>> Intentando importar CSV como último recurso..."
capture noisily import delimited using "D:\LOGAN\TIE\RESUTADOS.csv", ///
    varnames(1) encoding("UTF-8") bindquote(strict) clear
if _rc {
    di as err "Falló import delimited. Verifica que exportaste CSV UTF-8 y que no esté abierto."
    log close
    exit 603
}
di as res "CSV importado correctamente."
describe
list in 1/5, abbrev(24)
log close
exit
***/

log close
di as err _n "No fue posible importar el Excel con ningún método automático."
di as txt "Revisa import_log.txt en: `c(pwd)' para ver el detalle."
exit 604
