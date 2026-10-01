# Control de asistencia — QALEN

Página que toma el reporte del reloj facial (`.xls`) y entrega un Excel listo para trabajar:
el reporte de la máquina, debajo la entrada y salida de cada día, **celdas amarillas** donde
falta un dato, y un **Resumen de la bolsa de horas** que se recalcula solo con fórmulas.

## Para Yasmín (uso)

1. Abre `Control-de-asistencia-QALEN.html` (doble clic; funciona sin internet).
2. **Procesar mes** → arrastra el `.xls` del reloj. Elige "Calcular hasta el día" (por defecto: último día hábil del mes).
3. Revisa en pantalla (cada persona se expande día por día) y pulsa **Descargar Excel**.
4. En el Excel: hoja **Por revisar** → "Ir ›" → completa lo amarillo en **Asistencia**. **Resumen** y **Semanal** se actualizan solos.
5. Al cerrar el mes: **Personas y reglas → Leer saldos de un Excel terminado** para pasar los saldos al mes siguiente.

Nada sale de la computadora. La configuración se guarda en el navegador (con respaldo JSON).

## Excel generado

| Hoja | Qué tiene |
|---|---|
| Leeme | Cómo usarlo, colores y reglas. |
| Por revisar | Lista de casos con enlace a la celda y estado que cambia solo (Falta completar / Listo). |
| Asistencia | Por persona: Reloj · Entrada · Salida · Motivo · Horas trabajadas · A favor · En contra · Estado. Columnas = días (como la hoja de Yasmín). |
| Semanal | Trabajadas / a favor / en contra / neto por semana (lunes a sábado), con `SUMAR.SI`. |
| Resumen | Bolsa de horas: saldo anterior + a favor − en contra − castigo − pagadas = saldo final (le debemos / nos debe). |
| Configuracion | Reglas, personas (horarios), motivos y feriados. Lo único que las fórmulas consultan como parámetros. |
| Reloj original | Copia fiel del reporte del reloj. |

- La página decide **entrada y salida** (o las deja vacías si hay duda); todo lo demás son fórmulas simples (`SI`, `MIN`, `MAX`, `BUSCARV`, `CONTAR.SI`, `SUMAR.SI`), sin funciones dinámicas.
- Hojas protegidas **sin contraseña** (aviso, no traba). `fullCalcOnLoad` activado.
- Validaciones con confirmación: hora fuera de formato, salida antes que la entrada (típico de escribir 6:14 en vez de 18:14).
- Colores: amarillo = falta dato · azul = hora escrita a mano (no está en el reloj) · rojo = revisar horas · verde = con motivo · gris = fórmula.

## Reglas (configurables)

- Entrada 08:30; salida L-V 18:00; sábado 12:00. Por persona se puede cambiar (p. ej. medio tiempo).
- **Extras**: solo si sale 20 min o más después de su hora (18:20 / 12:20); se cuentan desde su hora (18:00 / 12:00).
- **Tardanza**: resta de la bolsa solo si es de **60 min o más** (criterio observado en la hoja de Yasmín; con 0 resta toda tardanza).
- **Salida antes de hora**: resta.
- **Llegar temprano**: suma solo a quien tiene configurado "cuenta si llega antes de" (p. ej. 07:45).
- **Permiso / Estudia** sin marcadas: resta la jornada (horas semanales ÷ 6, máx. el horario del día). **Falta**: no toca la bolsa; se cuenta aparte.
- **Domingo / feriado** trabajado: todo el tiempo suma a favor (factor configurable).
- Personas con "Descuenta" = No (casos especiales): no se les resta nada.
- Marcadas a menos de 3 min se consideran duplicadas.

### Clasificación de marcadas

Entrada = primera marcada antes de las 11:00. Salida = última marcada desde las 16:30 (L-V; o 90 min antes de su salida) o
después de las 11:00 (sábado). Almuerzo (informativo) = marcadas entre 12:00 y 15:30. Lo que no encaja **no se adivina**:
queda en amarillo (sin salida, sin entrada, una sola marcada, día sin marcadas). Para personas de reparto se sugiere la
hora de salida, pero siempre se pregunta. Horas antes de 06:30 / después de 22:00 y más de 3 h de extras se listan para mirar.

## Diferencias conocidas

**Con el reloj:** las tardanzas del "Reporte Estadístico" del reloj usan su propia tolerancia; aquí no se usan esas hojas
(solo "Reporte de Asistencia") y la tardanza se mide desde 08:30 exacto. No tienen por qué coincidir.

**Con la hoja SETIEMBRE de Yasmín** (validado con pruebas que viven en `/samples/tests`, fuera del repo, porque usan datos reales):
- Entradas/salidas: coinciden en todos los días normales (185). Solo difieren 3 días, por errores de tipeo en la hoja manual.
- Todo día que ella completó a mano aparece en "Por revisar".
- "En contra" (compensadas) coincide en todas las personas. "A favor" coincide exacto en 3 personas; en el resto la diferencia
  se explica por salidas de 18:21–18:29 de la última semana que no se contaron a mano, los tipeos y un caso de llegada temprana en sábado.

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
con la página, y que editar una celda cambie Semanal y Resumen.

**Privacidad:** `/samples` y todo `.xls/.xlsx` están en `.gitignore`. No subir datos de personas.
