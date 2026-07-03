# 📘 MANUAL DE USO — Sistema de Gestión

Guía paso a paso de **todo** lo que hace la plantilla, con ejemplos para cada situación.
Consúltalo cuando tengas una duda; está organizado por temas.

> **Regla de oro:** solo escribes en las celdas **blancas o amarillas (entrada)**.
> Las **azules claras / con fórmula** se calculan solas — no las toques.

---

## 🟢 PASO 0 — Arrancar el sistema (una sola vez)

1. **Datos de tu empresa:** hoja **Config** → nombre, NIT, dirección, teléfono, correo, logo (URL) y mensaje del pie de factura.
2. **IVA por categoría:** en Config, tabla *Categorías* → ponle a cada categoría su IVA % (19, 5, 0…). Si vendes algo nuevo, agrega la categoría en una fila vacía.
3. **Costos fijos por venta (valor de referencia):** en Config, "Costo fijo por venta: envío (Mary) — referencia" y "bolsa y etiqueta — referencia" (hoy $1.000 c/u). Este valor es solo el que se **precarga** en cada venta nueva; cada venta lo puede cambiar individualmente (ver EJEMPLO 6-B más abajo).
4. **Saldo base de tus billeteras:** hoja **Billeteras** → columna **Saldo Inicial** → escribe cuánto tienes HOY en cada cuenta (Efectivo, Nequi, Bancolombia, etc.). De ahí en adelante el **Saldo Actual** se mueve solo.
   - *Ejemplo:* $200.000 en efectivo, $500.000 en Bancolombia, $150.000 en Nequi → los escribes en Saldo Inicial de cada fila.
5. **Inventario inicial (productos que YA tienes):** crea cada producto en **Catálogo** y luego cárgalo en **Compras** así:
   - Fecha = el día que empiezas a usar el sistema (si no recuerdas cuándo lo compraste, usa hoy).
   - Cantidad = lo que tienes HOY · Costo Unitario = lo que te costó (o tu mejor estimado).
   - **¿A crédito? = No** · **Medio de Pago = "Inventario inicial (ya pagado)"**.
   - 👉 Así carga el stock y el costo **sin descontar plata de tus billeteras** (porque esa plata ya la gastaste hace tiempo). Tampoco genera deuda de tarjeta.
   - *Nota:* si quieres que "Días en Bodega" refleje su antigüedad real, pon una fecha aproximada más vieja; si no recuerdas, no importa, la rotación empieza a contar desde hoy.

---

## 📦 EJEMPLO 1 — Crear un producto (hoja Catálogo)
Escribes una sola vez: **Código, Nombre, Categoría, Estado (Activo), Stock Mínimo, URL Imagen, Precio Detal, Precio Mayor, Medidas** (para arte, ej. `30x40`).
- *Ejemplo:* `T100 | Cuadro Atardecer | Arte | Activo | 2 | (url) | 60000 | 45000 | 30x40`.
- Se calcula solo: Costo Promedio, Stocks, Valor, IVA %, Días en bodega, Vendidas, Reabastecer.

---

## 🛒 EJEMPLO 2 — Compra de CONTADO (débito o efectivo)
Hoja **Compras**: Fecha, Código, Cantidad, Costo Unitario, **Medio de Pago = la billetera** (ej. `Bancolombia` o `Efectivo`), **¿A crédito? = No**, Fecha Llegada.
- *Efecto:* suma stock a Casa, recalcula costo promedio, y **resta el dinero de esa billetera**.

## 💳 EJEMPLO 3 — Compra a CRÉDITO (tarjeta)
Hoja **Compras**: igual, pero **Medio de Pago = la tarjeta** (ej. `TC Bancolombia 1 (1517)`) y **¿A crédito? = Sí**.
- *Efecto:* suma stock y la **deuda aparece en la hoja Tarjetas** (no descuenta billetera todavía).

## 🧾 EJEMPLO 4 — Abonar / pagar una tarjeta (por partes)
Compraste $200.000 a crédito y pagas de a poco. Hoja **Pagos Tarjeta** → una fila por pago: Fecha, Tarjeta, Valor, **Billetera (de dónde salió)**.
- Hoy: `TC Bancolombia 1 (1517) | 80000 | Bancolombia`.
- Otra semana: `TC Bancolombia 1 (1517) | 120000 | Nequi`.
- *Efecto:* la hoja **Tarjetas** baja la deuda sola; **Billeteras** resta de la cuenta usada.

---

## 🔄 EJEMPLO 5 — Trasladar a Mercado Libre
Hoja **Traslados**: Fecha, Código, Cantidad, Origen `Casa`, Destino `Mercado Libre`.
- *Efecto:* baja Stock Casa, sube Stock ML (el total no cambia).

---

## 💰 EJEMPLO 6 — Vender (una sola cosa)
Hoja **Ventas**, una fila: Fecha, **N° Factura** (tú lo inventas, ej. `FAC-100`), Cliente, Código, Cantidad, Ubicación, **Valor Recibido**, **Forma de Cobro**.
- *Efecto:* descuenta inventario, calcula **Utilidad Neta**, y suma la plata a esa billetera.

## 💵 EJEMPLO 6-B — Cambiar el costo de Mary/Papelería en UNA venta, o cobrar publicidad
Al final de la fila de **Ventas** hay 3 columnas editables (por si esa venta en particular tuvo un costo distinto):
- **Costo Mary** → ya viene con $1.000 puesto, pero si ese día el envío costó $1.500, borra el 1.000 y pon 1500 **solo en esa fila**.
- **Costo Papelería** → igual, viene con $1.000, cámbialo si esa venta usó más o menos empaque.
- **Publicidad** → viene **vacía**. Solo la llenas si esa venta específica vino de una pauta/publicidad paga (ej. $5.000 de pauta). Si la venta no tuvo publicidad, la dejas vacía y no resta nada.
- *Efecto:* la columna **Costos Fijos** de esa fila se recalcula sola sumando las 3 (Mary + Papelería + Publicidad), y de ahí sale la **Utilidad Neta** real de esa venta puntual.
- *Ejemplo:* vendiste con pauta de Instagram que te costó $8.000 → en esa fila, Publicidad = 8000. En las demás ventas (sin pauta) la dejas vacía.

## 🛍️ EJEMPLO 7 — Vender VARIOS productos a un cliente
T100 x5, T200 x5, T300 x50 → en **Ventas** pones **una fila por producto, todas con el mismo N° Factura** (ej. `FAC-101`).
- *Efecto:* al elegir `FAC-101` en la Factura, salen los 3 juntos.

## 🌐 EJEMPLO 8 — Venta por Mercado Libre (la plata llega después)
1. Regístrala en **Ventas** con Forma de Cobro = `Mercado Libre` → queda "por cobrar" en la billetera MELI.
2. Cuando **retiras** ese dinero: hoja **Movimientos** → Tipo `Retiro plataforma`, Origen `Mercado Libre`, Destino `Bancolombia`, Valor.
- (Skydrops igual: registras la venta con Forma `Skydrops` y luego el retiro.)

---

## 🧾 EJEMPLO 9 — Hacer la factura (desde una venta ya hecha)
Hoja **Factura** → en **"Venta N°"** elige el número (ej. `FAC-100`). Se autollena todo.
- PDF: **Archivo → Descargar → PDF**.

## 🖼️ EJEMPLO 10 — Catálogo de productos en PDF
Hoja **Catálogo Venta** → elige **Categoría** (o `TODAS`) y **Tipo de precio** (`Detal`/`Mayor`). Muestra foto, medidas, precio y **Disp. Casa (und)**.
- **Se actualiza solo:** un producto solo aparece aquí si está "Activo" **Y** tiene Stock Casa mayor a 0. Si aún no le has hecho ninguna Compra, o si ya se agotó en Casa, desaparece automáticamente de este catálogo — no tienes que marcarlo Inactivo a mano. (Si lo tienes en Mercado Libre pero no en Casa, tampoco sale aquí, porque este catálogo es para lo que puedes vender directo desde Casa.)
- Para enviarlo: selecciona las filas con productos y **Archivo → Descargar → PDF → Celdas seleccionadas**.

---

## ⚠️ EJEMPLO 11 — Ajuste (dañado o perdido)
Hoja **Ajustes**: Fecha, Código, **Cantidad con signo** (`-1`), Ubicación, Motivo.

---

## 🤝 EJEMPLO 12 — Prestar plata
Hoja **Préstamos**: Fecha, Persona, Teléfono, Motivo, **Valor Prestado**, **Billetera origen** (ej. `Nequi`).
- *Efecto:* resta de esa billetera; aparece en el Dashboard.

## 💵 EJEMPLO 13 — Te pagan (mismo día, varias billeteras)
Hoy Carlos paga $30.000 efectivo + $20.000 Nequi + $10.000 Daviplata. Hoja **Abonos Préstamos** → **una fila por cada pago**:

| Fecha | Persona | Valor | Forma de Cobro |
|---|---|---|---|
| HOY | Carlos | 30000 | Efectivo |
| HOY | Carlos | 20000 | Nequi |
| HOY | Carlos | 10000 | Daviplata |

- *Efecto:* cada uno entra a su billetera; el saldo de Carlos baja $60.000.

## 📅 EJEMPLO 14 — Abonos en días/billeteras diferentes
Mañana te abona $20.000 por Bancolombia → otra fila con la fecha de mañana y Forma `Bancolombia`. **No hay límite de abonos.**

---

## 🔁 EJEMPLO 15 — Mover plata entre billeteras / gasto
Hoja **Movimientos**:
- **Transferencia:** Origen `Bancolombia`, Destino `Nequi`, Valor.
- **Gasto:** Origen `Efectivo`, Destino **vacío**, Valor, Nota.
- **Aporte** (meter plata nueva): Origen vacío, Destino la billetera, Valor.

## 🏧 EJEMPLO 15-B — Metiste o sacaste efectivo sin que sea venta ni gasto (aporte de capital, retiro de cajero)
Esto pasa, por ejemplo: metiste capital propio a la empresa, o sacaste plata del banco para tener efectivo en caja.
- **Capital que entra (aporte):** Movimientos → Tipo `Aporte`, Origen **vacío**, Destino la billetera donde entró (ej. `Bancolombia` o `Efectivo`), Valor.
- **Retiro del banco a la caja física:** Movimientos → Tipo `Transferencia`, Origen `Bancolombia`, Destino `Efectivo`, Valor.
- *Efecto en Cierre Diario:* **no toques nada a mano.** El Cuadre de Efectivo tiene 2 líneas automáticas que leen Movimientos solas: "+ Transferencias/Aportes que entraron a Efectivo hoy" y "− ... que salieron de Efectivo hoy". Así, cuando cuentes el efectivo físico, el "esperado" ya incluye ese dinero y el cuadre te da bien sin ajustes manuales.
- *Importante:* ni el Aporte ni la Transferencia interna son ganancia ni gasto — solo mueven dónde está guardada la plata.

---

## 📊 EJEMPLO 16 — Cierre del día
Hoja **Cierre Diario** (fecha en HOY): ingresos por método, MELI aparte, costos fijos, préstamos del día, **movimiento neto**, **ganancia del día** y el **historial** de 30 días.

## 🏦 EJEMPLO 17 — Cuánto tengo y cuánto debo
- **Billeteras:** saldo de cada cuenta. **Tarjetas:** deuda de cada tarjeta. **Dashboard:** resumen del mes.

## 🔥 EJEMPLO 18 — Qué reabastecer
**Catálogo**, columna **Reabastecer**: 🔥 se vende rápido · 🔴 agotado · 🟠 stock bajo.

---

## 💡 Conceptos clave
- **Utilidad Neta (Ventas)** = tu ganancia real de esa venta → lo que separas en el "sobre".
- **Ganancia del mes (Dashboard)** = suma de esas ganancias.
- **Préstamos** = NO son ganancia ni gasto; solo mueven plata que va y vuelve.
- **Billeteras** = lo que tienes · **Tarjetas** = lo que debes.

---

## 🔒 Cómo proteger las fórmulas (para que no se dañen)

Para que nadie (ni tú por error) borre una fórmula, en Google Sheets:

**Método recomendado — "Mostrar advertencia" (no bloquea, solo avisa):**
1. Selecciona las columnas/celdas que tienen fórmulas (las de fondo **azul claro** o de resultados).
2. Menú **Datos → Proteger hojas y rangos**.
3. En el panel de la derecha, clic en **Establecer permisos**.
4. Elige **"Mostrar una advertencia al editar"** → **Listo**.
5. Si algún día intentas cambiarlas, te avisa antes (puedes seguir si de verdad quieres).

**Qué proteger (lo más importante):**
- Hojas que son SOLO resultados (protégelas completas): **Dashboard, Inventario, Tarjetas**.
- Hojas con fórmulas + pocas casillas de entrada: protege las **columnas de fórmula** en **Catálogo** (F a L y de la R en adelante), **Compras** (C, F, I, J, L–N, P), **Ventas** (E, I–N), **Billeteras** (columna Saldo Actual), **Cierre Diario** (todo menos las celdas amarillas).
- No protejas las columnas donde tú escribes (las blancas/amarillas).

> Truco rápido: si te parece mucho, protege primero las 3 hojas de solo-resultados (Dashboard, Inventario, Tarjetas) y la columna **Saldo Actual** de Billeteras. Eso ya evita los daños más comunes.

---

## 📝 Historial de cambios

**3 de julio 2026:**
- **Cierre Diario:** nuevas líneas automáticas que suman/restan las transferencias del día que entran o salen de Efectivo (leyendo Movimientos) — ver EJEMPLO 15-B.
- **Ventas:** 3 columnas nuevas editables al final (Costo Mary, Costo Papelería, Publicidad) — ver EJEMPLO 6-B.
- **Catálogo Venta:** ya no muestra productos con Stock Casa = 0, aunque estén "Activo" — ver EJEMPLO 10.

