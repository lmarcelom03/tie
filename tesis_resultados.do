
version 17.0
clear all
set more off
macro drop _all

* Rutas

local xls "Combinación_Resultados.xlsx"
local sheet "Hoja1"

* Importar

import excel using "`xls'", sheet("`sheet'") firstrow clear

* Mantener solo filas de participantes (id presente)
capture confirm variable participant_id_in_session
if _rc!=0 {
    di as err "No se encuentra 'participant_id_in_session' en la hoja. Revisa nombres."
    describe
    exit 459
}
drop if missing(participant_id_in_session)

* Reconstruir etiqueta de grupo (Control/Tratamiento)
capture confirm variable Unnamed__0
if _rc==0 {
    gen str80 group_label = Unnamed__0
}
else {
    gen str80 group_label = ""
    ds, has(type string)
    foreach v of varlist `r(varlist)' {
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
label define lb_treat 0 "Control" 1 "Tratamiento"
label values treat lb_treat

* Controles
capture rename TESIS_TOTAL_C_1_player_edad edad
capture rename TESIS_TOTAL_C_1_player_Sexo sexo_str
gen byte mujer = (lower(sexo_str)=="female")
label define lb_sex 0 "Hombre" 1 "Mujer"
label values mujer lb_sex

* Encuesta 
capture rename TESIS_TOTAL_C_1_player_Preg_Optimismo preg_optimismo
capture rename TESIS_TOTAL_C_1_player_Preg_Confianza preg_confianza

* Desempeño (proxy de P)
capture egen gk_ok_total = rowtotal(TESIS_TOTAL_C_1_player_gk_1_ok ///
                                    TESIS_TOTAL_C_1_player_gk_2_ok ///
                                    TESIS_TOTAL_C_1_player_gk_3_ok ///
                                    TESIS_TOTAL_C_1_player_gk_4_ok)

* Promedios por sección A/B/C/D (p_blue)

* Detectar columnas por patrón (robusto a mayúsculas)
ds *p_blue*_*_A*, has(type numeric)
local Avars `r(varlist)'
ds *p_blue*_*_B*, has(type numeric)
local Bvars `r(varlist)'
ds *p_blue*_*_C*, has(type numeric)
local Cvars `r(varlist)'
ds *p_blue*_*_D*, has(type numeric)
local Dvars `r(varlist)'

* Promedios por sesión
capture drop pA pB pC pD
if "`Avars'"!="" {
    egen pA = rowmean(`Avars')
}
if "`Bvars'"!="" {
    egen pB = rowmean(`Bvars')
}
if "`Cvars'"!="" {
    egen pC = rowmean(`Cvars')
}
if "`Dvars'"!="" {
    egen pD = rowmean(`Dvars')
}

* Shifts (distancias respecto a A)
capture drop shift_B shift_C shift_D
gen shift_B = pB - pA if !missing(pB, pA)
gen shift_C = pC - pA if !missing(pC, pA)
gen shift_D = pD - pA if !missing(pD, pA)

* Variables “pre” y “post” estilo anterior (EN CASO DE SER NECESARIO EN ANEXOS)
capture drop optim_pre optim_post d_optim
if "`Avars'"!="" {
    egen optim_pre  = rowmean(`Avars')
}
if "`Dvars'"!="" {
    egen optim_post = rowmean(`Dvars')
}
gen d_optim = optim_post - optim_pre

* Mantener muestra válida para análisis principal (requiere A y D)
drop if missing(pA, pD, treat)

* Descriptivos

di as txt "=== Descriptivos base ==="
summ pA pB pC pD shift_B shift_C shift_D preg_optimismo preg_confianza edad mujer

di as txt "=== Balance pre-tratamiento (shifts B y C no deben diferir por grupo) ==="
ttest shift_B, by(treat)
ttest shift_C, by(treat)


* Gráficos de distribuciones (A vs D por grupo)

twoway (kdensity pA if treat==0) (kdensity pA if treat==1), ///
       legend(order(1 "Control" 2 "Tratamiento")) ///
       title("Distribución p(A) por grupo") name(gA, replace)

twoway (kdensity pD if treat==0) (kdensity pD if treat==1), ///
       legend(order(1 "Control" 2 "Tratamiento")) ///
       title("Distribución p(D) por grupo (post)") name(gD, replace)

* Barras de medias de shifts por grupo
preserve
collapse (mean) shift_B shift_C shift_D, by(treat)
graph bar shift_B shift_C shift_D, over(treat) ///
    title("Shifts promedio por grupo") ///
    legend(order(1 "Optimismo (B−A)" 2 "Exceso (C−A)" 3 "Combinado (D−A)"))
restore

* Modelos principales

eststo clear

* ANCOVA del shift combinado (D−A) controlando niveles pre (B−A y C−A)
reg shift_D i.treat c.shift_B c.shift_C c.edad i.mujer, vce(robust)
eststo m_ancova_comb

* Impacto diferencial del feedback: ¿afecta por igual optimismo y exceso?
*     Interacciones: cambio en la pendiente de D−A respecto a B−A y C−A en Tratamiento.
reg shift_D c.shift_B##i.treat c.shift_C##i.treat c.edad i.mujer, vce(robust)
eststo m_diffimpact

* Prueba conjunta: igual corrección en ambos sesgos
test 1.treat#c.shift_B = 1.treat#c.shift_C

* Chequeo estilo DiD simple (A→D) como en versión previa
reg optim_post i.treat c.optim_pre c.edad i.mujer, vce(robust)
eststo m_ancova_level
reg d_optim i.treat c.edad i.mujer, vce(robust)
eststo m_did_level


* Robustez (VERIFICAR) OPCIONAL

* Añadir desempeño (proxy de habilidad) al marco Heger & Papageorge
*      Si gk_ok_total está disponible, controla por P (performance)
capture confirm variable gk_ok_total
if _rc==0 {
    reg shift_D c.shift_B##i.treat c.shift_C##i.treat c.gk_ok_total c.edad i.mujer, vce(robust)
    eststo m_diffimpact_perf
}

* Encuesta como covariables
capture confirm variable preg_optimismo
capture confirm variable preg_confianza
if _rc==0 {
    reg shift_D c.shift_B##i.treat c.shift_C##i.treat c.preg_optimismo c.preg_confianza c.edad i.mujer, vce(robust)
    eststo m_diffimpact_enc
}

* Exportar

capture which esttab
if _rc ssc install estout, replace

esttab m_ancova_comb m_diffimpact m_ancova_level m_did_level ///
       using "resultados_modelos_clave.rtf", replace ///
       title("Efectos de retroalimentación y prueba de impacto diferencial") ///
       b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
       label mtitles("ANCOVA D−A" "Diferencial (int.)" "ANCOVA niveles" "DiD niveles")

capture confirm estimation m_diffimpact_perf
if _rc==0 esttab m_diffimpact_perf using "robustez_perf.rtf", replace ///
    title("Robustez con desempeño (P)") b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01)

capture confirm estimation m_diffimpact_enc
if _rc==0 esttab m_diffimpact_enc using "robustez_encuesta.rtf", replace ///
    title("Robustez con encuesta (autorreporte)") b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01)

* Guardar base
save "resultados_analiticos_shifts.dta", replace


* Notas
* - shift_B ≈ ε_D (optimismo), shift_C ≈ ε_P (exceso), shift_D ≈ ε_D + ε_P.
* - Prueba clave: en m_diffimpact, comparar 1.treat#c.shift_B vs 1.treat#c.shift_C.
*   Si son iguales (test no rechaza), el feedback afecta “por igual”.
*   Si difieren, identifica cuál sesgo se corrige más con el tratamiento.

