# Sistema de Gestión en Google Sheets
### Inventario · Compras · Ventas · Costos · Facturación

Solución completa para Google Sheets impulsada por **Google Apps Script**. Un único
script construye automáticamente todo el libro (hojas, fórmulas, validaciones,
tablero de control, gráficos y facturación con exportación a PDF), reemplazando los
procesos manuales por un control integral y automático.

---

## 0. ⭐ Forma más rápida: abre el archivo Excel ya construido

En `generator/Sistema_Gestion_Inventario.xlsx` tienes el **libro completo ya armado**
(todas las pestañas, fórmulas, validaciones, formato, KPIs y plantilla de factura),
con **datos de ejemplo** para que veas todo funcionando.

**Para usarlo como Google Sheet:**
1. Entra a [Google Drive](https://drive.google.com) → **Nuevo → Subir archivo** y sube
   `Sistema_Gestion_Inventario.xlsx`.
2. Haz doble clic en el archivo → **Abrir con → Hojas de cálculo de Google**.
3. (Recomendado) **Archivo → Guardar como Hojas de cálculo de Google** para tener una
   copia nativa. ¡Listo! Las fórmulas de inventario, costos, ventas y la factura ya funcionan.
4. Borra los datos de ejemplo (filas con productos `P001…`, compras, ventas, etc.) cuando
   quieras empezar con tu información real, y personaliza la pestaña **Config**.

> El `.xlsx` usa fórmulas estándar (VLOOKUP, SUMIF, SUMIFS, IF, EOMONTH…) que funcionan
> igual en Excel y en Google Sheets, **sin necesidad de programar nada**.

**¿Quieres además los botones automáticos** (número de factura automático, generar PDF,
enviar por correo, compartir por WhatsApp y gráficos que se recrean solos)? Esas
automatizaciones no caben dentro de un archivo estático: añádelas pegando una sola vez el
script de la carpeta `apps-script/` siguiendo la sección 2. **Es opcional**: el control de
inventario, costos, ventas y la factura funcionan sin el script.

---

## 1. ¿Qué incluye?

| Módulo | Hoja | Función |
|--------|------|---------|
| **Dashboard Ejecutivo** | `Dashboard` | KPIs automáticos del mes y gráficos de ventas/utilidades. |
| **Catálogo de Productos** | `Catálogo` | Maestro de productos (cada artículo se registra una sola vez). |
| **Compras** | `Compras` | Suma stock a Casa y recalcula el costo promedio. |
| **Inventario por Ubicación** | `Inventario` | Reporte consolidado: Casa, Mercado Libre y Total. |
| **Traslados** | `Traslados` | Movimientos Casa ⇄ Mercado Libre. |
| **Ventas** | `Ventas` | Descuenta inventario y calcula la utilidad con el costo promedio vigente. |
| **Ajustes de Inventario** | `Ajustes` | Daños, pérdidas y correcciones sin tocar compras/ventas. |
| **Clientes** | `Clientes` | Base de datos de clientes para la facturación. |
| **Facturación** | `Factura` | Plantilla profesional + numeración, totales, PDF, correo y WhatsApp. |
| **Configuración** | `Config` | Datos de la empresa, IVA, prefijo de factura, listas, etc. |

---

## 2. Instalación de las automatizaciones (opcional)

> Solo si quieres los botones (PDF, número de factura automático, correo, WhatsApp y
> gráficos auto-recreados). El sistema de inventario y la factura ya funcionan sin esto.

1. Crea una hoja de cálculo nueva en **Google Sheets** (https://sheets.new).
2. Menú **Extensiones → Apps Script**.
3. En el editor de Apps Script:
   - Borra el contenido del archivo `Código.gs` que viene por defecto.
   - Crea un archivo por cada `.gs` de la carpeta `apps-script/` (botón **+ → Secuencia de comandos**)
     y pega el contenido. Los archivos son:
     - `Codigo.gs`
     - `Setup.gs`
     - `Dashboard.gs`
     - `Facturacion.gs`
   - Activa el manifiesto: **Configuración del proyecto ⚙ → marca "Mostrar archivo de manifiesto `appsscript.json`"**
     y reemplaza su contenido por el de `apps-script/appsscript.json`.
4. Guarda todo (**Ctrl/Cmd + S**).
5. Vuelve a la hoja de cálculo y **recárgala** (F5). Aparecerá el menú **📊 Gestión**.
6. Ejecuta **📊 Gestión → 🚀 Inicializar / Reconstruir sistema**.
   - La primera vez Google pedirá **autorizar permisos**: acepta con tu cuenta.
     (Como el código no está verificado, pulsa *Configuración avanzada → Ir a (proyecto)*).
7. ¡Listo! Se construyen todas las hojas. Abre **Config** y personaliza los datos de tu empresa.

> Los permisos solicitados sirven para: editar el libro, crear el menú, generar el PDF
> en tu Drive y enviar el correo de la factura.

---

## 3. ¿Cómo funciona la automatización?

Toda la lógica de inventario y costos vive en **fórmulas**, por lo que se recalcula sola y
es transparente y auditable:

- **Costo promedio ponderado** (por producto):
  `Σ(costo total de compras) ÷ Σ(cantidades compradas)` — basado en **todas** las compras.
- **Stock por ubicación**, calculado con `ARRAYFORMULA` + `SUMIF` sobre claves
  `Código|Ubicación`:
  - **Casa** = compras + traslados que entran − traslados que salen − ventas desde Casa + ajustes en Casa.
  - **Mercado Libre** = traslados que entran − traslados que salen − ventas desde ML + ajustes en ML.
- **Stock Total** = Casa + Mercado Libre. **Valor inventario** = Stock Total × Costo Promedio.
- **Semáforo de stock**: `Agotado` (≤ 0), `Stock bajo` (≤ stock mínimo) o `Disponible`.
- **Utilidad de la venta** = Valor recibido − (Cantidad × Costo promedio vigente).

Las columnas con fondo azul claro son **automáticas**: no se escriben a mano.

---

## 4. Uso diario

> Regla de oro: solo se escribe en las celdas **blancas/amarillas**. Las **azul claro** son fórmulas.

### Catálogo
Registra cada producto **una sola vez**: `Código`, `Nombre`, `Categoría`, `Estado` y
`Stock Mínimo` (umbral para la alerta de stock bajo). El resto se calcula solo.

### Compras
`Fecha`, `Código`, `Cantidad`, `Costo Unitario`. Cada compra **suma stock a Casa** y
**recalcula el costo promedio**.

**Control de pagos y tarjetas (columnas a la derecha de Compras):**
- **Costos Fijos**: se suman solos a cada pedido (envío + bolsa/etiqueta, definidos en `Config`).
- **Total Pedido** = Costo Total + Costos Fijos.
- **Medio de Pago**: elige la tarjeta o medio (lista que configuras en `Config`).
- **¿A crédito?**: `Sí` / `No`. Si es `No`, el pedido queda como "Pagado (contado)".
- **Fecha Cierre** y **Fecha Límite Pago**: se calculan solas. Con cierre el **30** y pago el
  **16** del mes siguiente, una compra de junio cierra el 30/jun y se paga antes del 16/jul.
- **Abonado**: escribe cuánto has abonado (p. ej. con un préstamo). **Saldo** = lo que falta.
- **Estado Pago**: Pagado / Abono parcial / Pendiente.
- **Alerta Pago**: 🟢 al día · 🟠 faltan pocos días · 🔴 vencido. Aparece resumido en el
  **Dashboard** (Total que debo, Pedidos por pagar, Pagos vencidos y la lista de próximos pagos).

> Puedes cambiar el día de cierre (30), el día de pago (16) y "avisar si faltan X días" en `Config`.

### Traslados
`Fecha`, `Código`, `Cantidad`, `Origen` y `Destino`. Mueve unidades entre Casa y
Mercado Libre actualizando ambas existencias.

### Ventas (registro directo)
`Fecha`, `Cliente`, `Código`, `Cantidad`, `Ubicación` y `Valor Recibido`. El sistema
descuenta el inventario y calcula utilidad y margen.
> Las ventas también se generan automáticamente al **guardar una factura**.

### Ajustes
`Fecha`, `Código`, `Cantidad (+/-)`, `Ubicación`, `Motivo` y `Nota`.
Usa cantidades **negativas** para restar (p. ej. `-2` por producto dañado). No altera
compras ni ventas.

### Clientes
`Nombre`, `Documento/NIT`, `Teléfono`, `Correo`, `Dirección`, etc. El `ID` se genera solo.
Para WhatsApp escribe el teléfono en formato internacional sin signos (ej.: `573001112233`).

---

## 5. Facturación

Desde el menú **📊 Gestión → 🧾 Facturación**:

1. **➕ Nueva factura**: asigna número consecutivo y fecha, y limpia la plantilla.
2. Diligencia la **Factura**:
   - Elige el **Cliente** (sus datos se autocompletan).
   - Selecciona **Origen inv.** (Casa / Mercado Libre) y la **forma de pago**.
   - Agrega productos por **Código** (la descripción y el total se completan solos);
     escribe **Cantidad** y **Vr. Unitario**.
   - Subtotal, **IVA (calculado por renglón según la categoría de cada producto)** y Total
     se calculan automáticamente.
3. **💾 Guardar venta y registrar inventario**: vuelca los renglones al módulo **Ventas**,
   descuenta el inventario y avanza el consecutivo.
4. **📄 Generar / Descargar PDF**: crea el PDF (sin cuadrículas, con aspecto de documento
   comercial) en la carpeta **"Facturas"** de tu Drive y muestra enlaces para ver y descargar.
5. **✉️ Enviar por correo**: envía el PDF al correo del cliente.
6. **🟢 Compartir por WhatsApp**: abre WhatsApp con el mensaje y el enlace de la factura.

---

## 5b. Catálogo comercial con fotos (hoja `Catálogo Venta`)

Pensado para enviar a tus clientes un catálogo por categoría (arte, tecnología, etc.),
**con imagen del producto** y mostrando el precio que elijas (al detal o al por mayor).

**Cómo se cargan los datos (una sola vez, en la hoja `Catálogo`):**
- **URL Imagen**: pega el enlace público de la foto del producto (ver nota abajo).
- **Precio Detal**: precio de venta al público.
- **Precio Mayor**: precio al por mayor.
- **Medidas (ej. 30x40)**: úsalo para arte/cuadros. Escribe el tamaño como `20x20`, `30x40`,
  etc. El catálogo ordena los productos de cada categoría **de menor a mayor tamaño**
  automáticamente (calcula el área). Si un producto no tiene medidas, se ordena por nombre.

**Cómo generar el catálogo (en la hoja `Catálogo Venta`):**
1. En **Categoría** elige la categoría que quieres mostrar, o **TODAS** para incluir todo tu
   surtido (lista desplegable).
2. En **Tipo de precio** elige **Detal** o **Mayor**.
3. La hoja arma sola la lista: **foto + nombre + medidas/código + precio**, solo de los
   productos **activos**, ordenados por categoría y por tamaño.
4. Para enviarlo: menú **Archivo → Descargar → PDF**. En las opciones desmarca *Mostrar
   cuadrículas*, elige **Hoja actual** y exporta. Listo: un catálogo con apariencia profesional.

> **Agregar categorías nuevas:** ve a la hoja `Config`, columna **Categorías**, y escribe la
> nueva categoría en una fila vacía (y su IVA % al lado, si aplica). Aparecerá automáticamente
> tanto en el desplegable del `Catálogo` como en el selector del `Catálogo Venta`.

> **¿De dónde sale la URL de la imagen?** La función `IMAGE()` necesita un enlace **público**
> que termine idealmente en `.jpg`/`.png`. Opciones fáciles:
> - Sube la foto a **Google Drive**, dale clic derecho → *Compartir* → "Cualquier persona con
>   el enlace", y usa el enlace.
> - O usa la URL de la imagen tal como aparece en tu publicación de Mercado Libre / tu web
>   (clic derecho sobre la foto → *Copiar dirección de la imagen*).
> Si una celda de imagen sale vacía, casi siempre es porque la URL no es pública o no apunta
> directo al archivo de imagen.

> El catálogo muestra hasta **60 productos por categoría**. Si necesitas más, se puede ampliar.

---

## 6. Personalización (hoja `Config`)

- **Datos de empresa**: nombre, NIT, dirección, teléfono, correo, web y **URL del logo**
  (pega la URL pública de una imagen para que salga en la factura).
- **Facturación**: `Prefijo de factura` (ej. `FAC-`) y `Próximo número de factura`.
- **IVA por categoría** (columnas `Categorías` + `IVA %`): escribe el porcentaje que
  corresponde a cada categoría (p. ej. **19** para Electrónica, **5** para Ropa, en blanco o
  **0** si no aplica). Cada producto del `Catálogo` toma su IVA según su categoría, y la
  **factura calcula el IVA por renglón** sumando lo que corresponde a cada producto. Así
  manejas productos con IVA del 19% y del 5% en la misma factura.
- **ID carpeta Drive para PDF** (opcional): pega el ID de una carpeta para guardar allí los PDF.
- **Listas editables** (columnas a la derecha): `Categorías`, `Ubicaciones`, `Estados`,
  `Motivos de ajuste`. Amplíalas según tu negocio (si agregas una categoría nueva, ponle su IVA %).

> **Precios de venta:** en la hoja `Catálogo` cada producto tiene **Precio Detal** y
> **Precio Mayor** (además del Costo Promedio que se calcula solo con las compras).

> **Moneda:** los valores están formateados en **pesos colombianos (COP) sin decimales**
> (ej. `$ 1.234.567`). El separador de miles lo aplica Google Sheets según la configuración
> regional del archivo; si lo ves distinto, ve a **Archivo → Configuración → Configuración
> regional → Colombia**.

Después de cambiar listas o agregar muchos productos, ejecuta
**🔄 Actualizar Dashboard y gráficos** para refrescar los gráficos.

---

## 7. Notas y límites

- Diseñado para hasta **1.000 productos** y **2.000 movimientos** por hoja (ampliable
  aumentando `CAT_LAST` y `MAX_ROWS` en el código y volviendo a inicializar).
- Funciona en **computador y móvil** (la app de Google Sheets ejecuta el menú y el script).
- "Reconstruir sistema" **recrea encabezados y fórmulas** pero respeta tus datos siempre que
  conserves las columnas; aun así, haz una copia del libro antes de reconstruir si ya tienes
  información cargada.
- El PDF y el correo usan tu cuenta de Google; la primera vez debes autorizar los permisos.

---

## 8. Solución de problemas

| Síntoma | Causa / Solución |
|---------|------------------|
| No aparece el menú **📊 Gestión** | Recarga la hoja (F5). Verifica que pegaste los 4 `.gs`. |
| "No se encontró la hoja…" | Ejecuta **🚀 Inicializar / Reconstruir sistema**. |
| Stock o costo en 0 | Revisa que el `Código` de la compra/venta exista igual en el `Catálogo`. |
| `⚠ Código no existe` | El código escrito no está en el `Catálogo`. |
| El PDF sale con cuadrícula | Usa el botón **Generar PDF** del menú (no "Descargar como PDF" del navegador). |
| No envía correo / no abre WhatsApp | Completa el correo/teléfono del cliente o autoriza los permisos. |

---

## 9. Estructura del proyecto

```
inventario-gsheets/
├── generator/
│   ├── build_xlsx.py                   # Generador del .xlsx (solo librería estándar)
│   └── Sistema_Gestion_Inventario.xlsx # ⭐ ARCHIVO LISTO PARA ABRIR EN GOOGLE SHEETS
├── apps-script/                        # Automatizaciones opcionales (PDF, factura, etc.)
│   ├── appsscript.json     # Manifiesto: zona horaria y permisos
│   ├── Codigo.gs           # Menú, constantes y utilidades
│   ├── Setup.gs            # Construcción de todas las hojas y fórmulas
│   ├── Dashboard.gs        # Gráficos del tablero
│   └── Facturacion.gs      # Plantilla de factura, PDF, correo y WhatsApp
└── README.md
```
