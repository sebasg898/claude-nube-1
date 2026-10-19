# Control de asistencia — QALEN

Página que toma el reporte del reloj facial (`.xls`) y entrega un Excel listo para trabajar:
el reporte de la máquina, debajo la entrada y salida de cada día, **celdas amarillas** donde
falta un dato, y un **Resumen de la bolsa de horas** que se recalcula solo con fórmulas.

## Para Yasmín (uso)

1. Abre `Control-de-asistencia-QALEN.html` (doble clic; funciona sin internet).
2. **Procesar mes** → arrastra el `.xls` del reloj. Elige "Calcular hasta el día" (por defecto: último día hábil del mes).
3. Revisa en pantalla (cada persona se expande día por día) y pulsa **Descargar Excel**.
4. En el Excel, hoja **Asistencia**: completa lo amarillo (hora o motivo). **Resumen** y **Semanal** se actualizan solos.

Nada sale de la computadora. La configuración (personas y reglas) se guarda en el navegador, con respaldo JSON.

## Excel generado (4 hojas)

| Hoja | Qué tiene |
|---|---|
| Leeme | Cómo usarlo, motivos, horas extras y colores. |
| Asistencia | Por persona: Reloj · Entrada (verde) · Salida (naranja) · Motivo · Horas trabajadas · Extra 25% · Extra 35% · Dom./Feriado · En contra · Compensadas (+2 filas ocultas: extras a pagar después de descontar compensadas). Columnas = días, con líneas por semana. |
| Semanal | Por semana (lunes a sábado): trabajadas, extra 25%, extra 35%, dom./feriado, en contra, compensadas. |
| Resumen | Faltas · horas trabajadas · base del mes · **extras 25% / 35%** · compensadas · **a pagar 25% / 35% / dom.-feriado 100%** · valor hora (editable) · **TOTAL A PAGAR (S/)** · nos debe · horas en contra · datos por completar. |

- La página decide **entrada y salida** (o las deja vacías si hay duda); todo lo demás son fórmulas simples (`SI`, `MIN`, `MAX`, `SUMA`, `CONTAR.SI`, `CONTAR.SI.CONJUNTO`, `SUMAR.SI`, `SUMAPRODUCTO`). El horario de cada persona va en columnas ocultas de Asistencia; el valor hora está a la vista en el Resumen.
- Hojas protegidas **sin contraseña**. `fullCalcOnLoad` activado. Validaciones: hora fuera de formato o salida antes que la entrada (aviso con confirmación); motivo solo de la lista.

## Reglas

- Entrada 08:30; salida L-V 18:00; sábado 12:00 (por persona se puede cambiar en la página).
- **Horas extras** (D.S. 007-2002-TR, art. 10): cuentan si sale 20 min o más después de su hora (18:20 / 12:20), desde las 18:00 / 12:00.
  Cada día, las **2 primeras horas al 25%** y **desde la 3.ª al 35%**. Llegar temprano no suma.
- **Domingo / feriado** trabajado: columna aparte, al **100%**. No se le descuentan compensadas.
- **Compensadas** (motivo *Compensado*): se descuentan de las extras de los **primeros días del período**, y dentro de cada día **primero las del 35%**.
  Si las compensadas superan las extras, la diferencia aparece en **"Nos debe"**.
- **Horas en contra** (tardanza después de las 08:30, salida antes de hora, *Estudio*, *Falta*): van **aparte**; no se descuentan de las extras.
- **Motivos** (lista cerrada): *Vacaciones* / *Salud* = día completo trabajado; *Compensado*; *Estudio* y *Falta* = en contra (*Falta* también se cuenta).
  Jornada de un día completo = horas semanales ÷ 6 (8 h; sábado máx. 3:30).
- **TOTAL A PAGAR** = (extras 25% netas × 1.25 + extras 35% netas × 1.35 + dom./feriado × 2) × valor hora, redondeado a céntimos.
  Valor hora = sueldo mensual ÷ 240. Los recargos son configurables (mínimos de ley; confirmar con el contador).
- **Horas base del mes** = días laborables (L-S, sin feriados) × horas semanales ÷ 6.
- Personas con "Descuenta" = No (casos especiales): no se les resta nada.
- Marcadas a menos de 3 min se consideran duplicadas.

### Clasificación de marcadas

Entrada = primera marcada antes de las 11:00. Salida = última marcada desde las 16:30 (L-V; o 90 min antes de su salida) o
después de las 11:00 (sábado). Lo que no encaja **no se adivina**: queda en amarillo (sin salida, sin entrada, una sola
marcada, día sin marcadas), con la explicación como comentario en la celda de la fila Reloj.

## Diferencias conocidas

**Con el reloj:** las tardanzas de las hojas de cálculo del propio reloj usan su propia tolerancia; aquí solo se usa
"Reporte de Asistencia" y la tardanza se mide desde 08:30 exacto.

**Con la hoja manual de setiembre** (pruebas privadas en `/samples/tests`): entradas/salidas coinciden en todos los días
normales (185); solo difieren 3 días por errores de tipeo en la hoja manual, y todo día completado a mano aparece en amarillo.

## Desarrollo

```bash
npm install
npm run dev        # desarrollo
npm test           # pruebas (las de coherencia con Excel requieren LibreOffice Calc: apt install libreoffice-calc)
npm run entrega    # genera entrega/Control-de-asistencia-QALEN.html (un solo archivo autocontenido)
```

Código: `src/lib/parser.ts` (lectura del reloj), `rules.ts` (clasificación y cálculo de referencia),
`model.ts` (mes), `excel.ts` (Excel con fórmulas que replican `calcularDia`), `src/ui/*` (pantallas).
La prueba `tests/excel-sintetico.test.ts` (datos inventados) recalcula el Excel con LibreOffice y verifica que cada número coincida
con la página (incluido el reparto de compensadas y el monto en soles), y que completar horas o elegir motivos cambie Semanal y Resumen.

**Privacidad:** `/samples` y todo `.xls/.xlsx` están en `.gitignore`. No subir datos de personas.
