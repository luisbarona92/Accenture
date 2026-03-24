# CLAUDE.md — Canal Integrador · MasOrange

Este fichero da instrucciones y permisos a Claude Code para trabajar de forma autónoma en este proyecto.

## Permisos pre-autorizados (no preguntar)

- **Leer y editar** `Seguimiento_dashboard.html` y cualquier fichero HTML/CSS/JS de esta carpeta
- **Leer** el fichero PPT de seguimiento (`*.pptx`) para extraer datos actualizados
- **Copiar** el HTML actualizado al clon git local (`C:/Users/luis.barona.arroyo/tmpAccenture/Canal_Integrador/`)
- **Hacer git add, commit y push** al repo `github.com/luisbarona92/Accenture` (rama `main`)
- **Ejecutar** scripts Node.js o bash para parsear el PPTX y extraer datos
- **Instalar paquetes npm** temporales en `/tmp` para parseo (jszip, pptx2json, etc.)

## Workflow estándar de actualización (PMO nuevo)

Cuando el usuario diga "actualiza el dashboard" o suba un nuevo PPT:

1. Leer el PPTX con Node.js + jszip para extraer texto de slides relevantes
2. Comparar con el HTML actual y detectar cambios en:
   - Contadores del funnel (H1, H2, H3, H4)
   - Grupos de integradores (verde/posible/hold/churn) y sus conteos
   - Observaciones por integrador en la tabla
   - KPIs de ventas (oportunidades, leads, ofertas, ventas)
   - Tareas en curso (equipo, fecha, estado)
   - Fecha y número de sesión PMO
3. Editar el HTML con los cambios
4. Copiar a `/tmp/Accenture/Canal_Integrador/Seguimiento_dashboard.html`
5. `git -C /tmp/Accenture pull origin main`
6. `git -C /tmp/Accenture add Canal_Integrador/Seguimiento_dashboard.html`
7. `git -C /tmp/Accenture commit -m "PMO #XX - [resumen de cambios]"`
8. `git -C /tmp/Accenture push origin main`

## Rutas clave

- **HTML local**: `C:/Users/luis.barona.arroyo/OneDrive - Accenture/Documents/MAS ORANGE/Canal_Integrador/Seguimiento_dashboard.html`
- **PPT fuente**: `C:/Users/luis.barona.arroyo/OneDrive - Accenture/Documents/MAS ORANGE/` (buscar el último `*.pptx`)
- **Repo git local**: `C:/Users/luis.barona.arroyo/tmpAccenture/` (clon de `github.com/luisbarona92/Accenture`)
- **GitHub Pages**: `https://luisbarona92.github.io/Accenture/Canal_Integrador/Seguimiento_dashboard.html`

## Notas de diseño

- No modificar el CSS ni la estructura JS salvo bugs
- Colores: naranja (#FF6600) para MasOrange, verde (#059669) para activados/firmados
- Grupos integradores: verde = leads compartidos alta participación, posible = moderada, hold = gris, churn = rojo
- El sticky del thead usa `top:60px` (altura del topbar)
- `overflow:clip` en `.card` para no romper sticky
- Contadores animados con `data-t` attribute en elementos `.cnt`
