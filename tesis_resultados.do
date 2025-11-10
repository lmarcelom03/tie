*******************************************************
* RESULTADOS – TESIS (Stata .do)  [Versión final]
* Integra análisis A/B/C/D, shifts y prueba de impacto diferencial
*******************************************************

version 19.0
clear all
set more off
set linesize 255
macro drop _all

*******************************************************
* 0) Rutas y configuración
*******************************************************
local base "D:\LOGAN\TIE"
local xlsname "RESUTADOS.xlsx"
local preferred_sheet "Hoja1"
local logname "tesis_resultados_log.txt"

capture noisily cd "`base'"
if _rc {
    di as err "No puedo cambiar a la carpeta base: `base'. Verifica la ruta."
    exit 610
}

log close _all
log using "`logname'", replace text

di as txt "Trabajando en: " c(pwd)

di as txt "Contenido de la carpeta base:"
capture noisily shell dir /b
if _rc di as txt "(No se pudo listar el contenido con shell dir.)"

local xlspath ""
capture confirm file "`xlsname'"
if !_rc local xlspath "`xlsname'"

if "`xlspath'"=="" {
    di as err "No se encontró el archivo `xlsname' en `base'."
    log close
    exit 601
}

di as res "Archivo detectado: `xlspath'"

di as txt _n "==> Hojas disponibles en el libro:"
capture noisily import excel using "`xlspath'", describe
if _rc {
    di as err "Stata no pudo leer la estructura del Excel. ¿Está abierto o protegido?"
    log close
    exit 602
}

local imported 0
local sheet_used ""

foreach s in "`preferred_sheet'" "Hoja 1" "Sheet1" "Sheet 1" {
    di as txt _n "Intentando importar sheet(`s') firstrow..."
    capture noisily import excel using "`xlspath'", sheet("`s'") firstrow clear
    if !_rc {
        local imported 1
        local sheet_used "`s'"
        continue, break
    }
}

if !`imported' {
    di as txt _n "Intentando importar sheet(1) firstrow..."
    capture noisily import excel using "`xlspath'", sheet(1) firstrow clear
    if !_rc {
        local imported 1
        local sheet_used "sheet(1)"
    }
}

if !`imported' {
    di as err "No fue posible importar el Excel ni por nombre ni por índice."
    di as txt "Cierra el archivo en Excel y confirma que no esté protegido."
    log close
    exit 603
}

di as res _n "Importación exitosa desde `sheet_used'."

di as txt "Observaciones cargadas: " _N

di as txt "Variables importadas:"
describe

compress

*******************************************************
* 1) Preparar datos
*******************************************************
* 1.1) Mantener solo filas de participantes (id presente)
capture confirm variable participant_id_in_session
if _rc!=0 {
    di as err "No se encuentra 'participant_id_in_session' en la hoja. Revisa nombres."
    describe
    log close
    exit 459
}
drop if missing(participant_id_in_session)

* 1.2) Reconstruir etiqueta de grupo (Control/Tratamiento)
capture drop group_label
capture confirm variable Unnamed__0
if _rc==0 {
    gen str80 group_label = Unnamed__0
}
else {
    gen str80 group_label = ""
    ds, has(type string)
    local stringvars `r(varlist)'
    foreach v of local stringvars {
        quietly count if strpos(lower(`v'), "grupo") | strpos(lower(`v'), "trat") | strpos(lower(`v'), "control")
        if r(N)>0 {
            replace group_label = `v' if group_label==""
        }
    }
}

gen long __obs = _n
sort __obs
replace group_label = group_label[_n-1] if missing(group_label) & _n>1

gen byte treat = .
replace treat = 1 if strpos(lower(group_label),"trat")>0
replace treat = 0 if strpos(lower(group_label),"control")>0
label define lb_treat 0 "Control" 1 "Tratamiento", replace
label values treat lb_treat

drop __obs

count if inlist(treat,0,1)
if r(N)==0 {
    di as err "No se pudo identificar el grupo de tratamiento/control. Revisa la columna de grupos."
    log close
    exit 611
}

* Controles
capture rename TESIS_TOTAL_C_1_player_edad edad
capture confirm variable edad
if _rc!=0 {
    gen double edad = .
    di as txt "Aviso: no se encontró la variable de edad; se crea como missing."
}

capture rename TESIS_TOTAL_C_1_player_Sexo sexo_str
capture confirm variable sexo_str
if _rc==0 {
    capture drop mujer
    gen byte mujer = (lower(sexo_str)=="female" | lower(sexo_str)=="mujer")
}
else {
    capture drop mujer
    gen byte mujer = .
    di as txt "Aviso: no se encontró la variable de sexo; 'mujer' queda en missing."
}
label define lb_sex 0 "Hombre" 1 "Mujer", replace
label values mujer lb_sex

* Encuesta (autorreporte)
capture rename TESIS_TOTAL_C_1_player_Preg_Optimismo preg_optimismo
capture confirm variable preg_optimismo
if _rc!=0 {
    gen double preg_optimismo = .
}

capture rename TESIS_TOTAL_C_1_player_Preg_Confianza preg_confianza
capture confirm variable preg_confianza
if _rc!=0 {
    gen double preg_confianza = .
}

* Desempeño (proxy de P)
capture egen gk_ok_total = rowtotal(TESIS_TOTAL_C_1_player_gk_1_ok ///
                                    TESIS_TOTAL_C_1_player_gk_2_ok ///
                                    TESIS_TOTAL_C_1_player_gk_3_ok ///
                                    TESIS_TOTAL_C_1_player_gk_4_ok)

*******************************************************
* 2) Promedios por sección A/B/C/D (p_blue) y SHIFTS
*******************************************************
* Detectar columnas por patrón (robusto a mayúsculas)
ds *p_blue*_*_A*, has(type numeric)
local Avars `r(varlist)'
if "`Avars'"=="" {
    di as err "No se encontraron columnas para la sección A (*p_blue*_*_A*)."
    log close
    exit 620
}

ds *p_blue*_*_B*, has(type numeric)
local Bvars `r(varlist)'

ds *p_blue*_*_C*, has(type numeric)
local Cvars `r(varlist)'

ds *p_blue*_*_D*, has(type numeric)
local Dvars `r(varlist)'
if "`Dvars'"=="" {
    di as err "No se encontraron columnas para la sección D (*p_blue*_*_D*)."
    log close
    exit 621
}

* Promedios por sesión
capture drop pA pB pC pD
egen pA = rowmean(`Avars')
if "`Bvars'"!="" egen pB = rowmean(`Bvars')
if "`Cvars'"!="" egen pC = rowmean(`Cvars')
egen pD = rowmean(`Dvars')

* Shifts (distancias respecto a A)
capture drop shift_B shift_C shift_D
if "`Bvars'"!="" gen shift_B = pB - pA if !missing(pB, pA)
if "`Cvars'"!="" gen shift_C = pC - pA if !missing(pC, pA)
gen shift_D = pD - pA if !missing(pD, pA)

capture confirm variable shift_B
if _rc!=0 gen shift_B = .
capture confirm variable shift_C
if _rc!=0 gen shift_C = .

* Variables “pre” y “post” estilo anterior (para anexos)
capture drop optim_pre optim_post d_optim
egen optim_pre  = rowmean(`Avars')
egen optim_post = rowmean(`Dvars')
gen d_optim = optim_post - optim_pre

* Mantener muestra válida para análisis principal (requiere A y D)
drop if missing(pA, pD, treat)

count
if r(N)==0 {
    di as err "La muestra quedó vacía tras filtrar por pA, pD y treat. Revisa los datos."
    log close
    exit 622
}

*******************************************************
* 3) Descriptivos y verificaciones
*******************************************************
local sumvars "pA pD shift_D"
if "`Bvars'"!="" local sumvars "`sumvars' pB shift_B"
if "`Cvars'"!="" local sumvars "`sumvars' pC shift_C"
local sumvars "`sumvars' preg_optimismo preg_confianza edad mujer"

di as txt "=== Descriptivos base ==="
summ `sumvars'

if "`Bvars'"!="" {
    di as txt "=== Balance pre-tratamiento: shift_B por grupo ==="
    ttest shift_B, by(treat)
}
if "`Cvars'"!="" {
    di as txt "=== Balance pre-tratamiento: shift_C por grupo ==="
    ttest shift_C, by(treat)
}

*******************************************************
* 4) Gráficos de distribuciones (A vs D por grupo)
*******************************************************
twoway (kdensity pA if treat==0) (kdensity pA if treat==1), ///
       legend(order(1 "Control" 2 "Tratamiento")) ///
       title("Distribución p(A) por grupo") name(gA, replace)

twoway (kdensity pD if treat==0) (kdensity pD if treat==1), ///
       legend(order(1 "Control" 2 "Tratamiento")) ///
       title("Distribución p(D) por grupo (post)") name(gD, replace)

* Barras de medias de shifts por grupo
preserve
local collapsevars "shift_D"
local idx = 1
local legend_items "`idx' \"Combinado (D−A)\""
if "`Bvars'"!="" {
    local collapsevars "`collapsevars' shift_B"
    local ++idx
    local legend_items "`legend_items' `idx' \"Optimismo (B−A)\""
}
if "`Cvars'"!="" {
    local collapsevars "`collapsevars' shift_C"
    local ++idx
    local legend_items "`legend_items' `idx' \"Exceso (C−A)\""
}
collapse (mean) `collapsevars', by(treat)

graph bar `collapsevars', over(treat) ///
    title("Shifts promedio por grupo") ///
    legend(order(`legend_items'))
restore

*******************************************************
* 5) Modelos principales
*******************************************************
* Asegurar estout disponible (eststo/esttab)
capture which eststo
if _rc {
    di as txt "Instalando estout (requerido para eststo/esttab)..."
    ssc install estout, replace
}

eststo clear

* 5.1) ANCOVA del shift combinado (D−A) controlando niveles pre (B−A y C−A)
reg shift_D i.treat c.shift_B c.shift_C c.edad i.mujer, vce(robust)
eststo m_ancova_comb

* 5.2) Impacto diferencial del feedback: ¿afecta por igual optimismo y exceso?
reg shift_D c.shift_B##i.treat c.shift_C##i.treat c.edad i.mujer, vce(robust)
eststo m_diffimpact

di as txt "=== Prueba conjunta: igual corrección en ambos sesgos ==="
test 1.treat#c.shift_B = 1.treat#c.shift_C

* 5.3) Chequeo estilo DiD simple (A→D) como en versión previa
reg optim_post i.treat c.optim_pre c.edad i.mujer, vce(robust)
eststo m_ancova_level
reg d_optim i.treat c.edad i.mujer, vce(robust)
eststo m_did_level

*******************************************************
* 6) Robustez opcional
*******************************************************
* 6.1) Añadir desempeño (proxy de habilidad) al marco Heger & Papageorge
capture confirm variable gk_ok_total
if _rc==0 {
    reg shift_D c.shift_B##i.treat c.shift_C##i.treat c.gk_ok_total c.edad i.mujer, vce(robust)
    eststo m_diffimpact_perf
}

* 6.2) Autorreporte (encuesta) como covariables
capture confirm variable preg_optimismo
capture confirm variable preg_confianza
if _rc==0 {
    reg shift_D c.shift_B##i.treat c.shift_C##i.treat c.preg_optimismo c.preg_confianza c.edad i.mujer, vce(robust)
    eststo m_diffimpact_enc
}

*******************************************************
* 7) Exportar tablas
*******************************************************
esttab m_ancova_comb m_diffimpact m_ancova_level m_did_level ///
       using "resultados_modelos_clave.rtf", replace ///
       title("Efectos de retroalimentación y prueba de impacto diferencial") ///
       b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
       label mtitles("ANCOVA D−A" "Diferencial (int.)" "ANCOVA niveles" "DiD niveles")

capture confirm estimation m_diffimpact_perf
if _rc==0 {
    esttab m_diffimpact_perf using "robustez_perf.rtf", replace ///
        title("Robustez con desempeño (P)") b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01)
}

capture confirm estimation m_diffimpact_enc
if _rc==0 {
    esttab m_diffimpact_enc using "robustez_encuesta.rtf", replace ///
        title("Robustez con encuesta (autorreporte)") b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01)
}

* Guardar base analítica
save "resultados_analiticos_shifts.dta", replace

log close

di as txt _n "Proceso completado. Revisa el log en: " c(pwd) "\`logname'"

di as txt "Notas:"
di as txt "- shift_B ≈ ε_D (optimismo), shift_C ≈ ε_P (exceso), shift_D ≈ ε_D + ε_P."
di as txt "- En m_diffimpact, compara 1.treat#c.shift_B vs 1.treat#c.shift_C para la prueba clave."
