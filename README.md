# K2 Captura

Captura en el celular las horas (REG/OT/DT/PD/ML), el equipo por orden y el
Equipment & Material Log del libro **TS-CEM-BRO-K2** (.xlsb). El celular nunca
toca el Excel: exporta una captura y la PC la escribe con Excel mismo, así el
.xlsb conserva todas sus fórmulas, formato y macros.

## Flujo
1. PC: suelta el `.xlsb` sobre el icono **K2 Captura** del escritorio → sale
   `<libro>.base.json` junto al Excel.
2. Pásalo al celular y cárgalo en la app.
3. Captura (funciona sin señal; todo se guarda en el celular).
4. **Enviar → Exportar captura**, pásala a la PC y suéltala sobre el mismo icono.
   La PC hace respaldo en `respaldos K2/`, avisa si alguna celda cambió en la PC
   desde la base, nunca escribe sobre celdas con fórmula, guarda y genera una base
   nueva. Carga esa base en el celular y lo enviado se limpia solo.

## Archivos
- `index.html`, `js/`, `styles.css`, `sw.js`: la PWA (sin datos en el código).
- `pc/k2_sync.py`: puente PC (Python + pywin32, usa Excel por COM).
