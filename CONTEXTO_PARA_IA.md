# 🧠 CONTEXTO DEL PROYECTO — léelo primero (para retomar con una IA desde cero)

Si eres una IA y te comparten este repositorio, **lee este archivo completo antes de hacer nada**.
Aquí está todo el contexto para que entiendas qué es, cómo funciona y cómo ayudar sin confundirte.

---

## 1. Qué es este proyecto
Un **sistema de gestión de inventario y finanzas** para un negocio que vende productos
(incluido **arte/cuadros**, y también tecnología, hogar, etc.). Está hecho como un **libro de
Google Sheets** que se **genera con un script de Python** (`generator/build_xlsx.py`) que
produce un archivo **`.xlsx`**. El `.xlsx` se sube a Google Drive y se abre como Hoja de
cálculo de Google. **No es una app**: todo vive en la hoja de cálculo, con fórmulas.

## 2. Dónde está todo (este repositorio)
- **`generator/build_xlsx.py`** → el GENERADOR. Aquí está TODO el código que arma el libro.
  Para cambiar la plantilla **se edita este archivo y se vuelve a generar**.
- **`generator/Sistema_Gestion_Inventario.xlsx`** → libro CON datos de ejemplo (para practicar).
- **`generator/Sistema_Gestion_Inventario_VACIO.xlsx`** → libro VACÍO (para uso real).
- **`MANUAL.md`** → manual de uso paso a paso (con ~18 ejemplos).
- Rama de trabajo: **`sistema-inventario`**.

## 3. Cómo se genera y se valida
- Ejecutar: `python3 build_xlsx.py` → crea los dos `.xlsx`.
- **No hay internet ni librerías externas** (no openpyxl): el `.xlsx` se construye a mano como
  un ZIP de archivos XML (OOXML) usando solo la librería estándar de Python.
- Validar siempre tras generar: (a) el ZIP abre, (b) cada XML es válido, (c) ningún índice de
  estilo fuera de rango, (d) sin referencias a hojas inexistentes, (e) sin solapes de celdas
  combinadas, (f) `python3 -c "import ast; ast.parse(...)"` para sintaxis.
- **Fórmulas:** se usan funciones compatibles con Google Sheets: `VLOOKUP, SUMIF, SUMIFS, IF,
  IFERROR, COUNTIF/S, EOMONTH, TODAY, FILTER, SORT, QUERY, IMAGE, INDEX, MATCH, MINIFS`.
- **Moneda:** pesos colombianos (COP) **sin decimales**.
- **Región del documento:** comas `,` como separador de argumentos (locale Estados Unidos). Si
  el usuario usa locale Colombia, las fórmulas usan `;`.
- En el generador, `s.formula(...)` guarda la fórmula **sin** el `=` inicial.

## 4. El negocio (contexto que NO se debe perder)
- El **arte** se divide por **medidas** (ej. `30x40`) y el catálogo se ordena por tamaño.
- **Ubicaciones de inventario:** **Casa** (lo que el dueño vende) y **Mercado Libre** (ya está allá).
- **Billeteras (donde hay plata):** Efectivo, Nequi, Bancolombia, Bancolombia MM, Daviplata,
  Davivienda, **Mercado Libre** y **Skydrops**. MELI y Skydrops **pagan después** (se retiran a
  una billetera mediante la hoja Movimientos).
- **Costos fijos por venta:** envío "Mary" $1.000 + bolsa/etiqueta $1.000 (configurables; suben con el tiempo).
- **Tarjetas de crédito (8):** TC Bancolombia 1 (1517), 2 (0554), 3 (4582), 4 (8366), 5 (0623),
  TC Bancolombia MM 6 (6423), MM 7 (2459), TC NU 8 (3750). **Cierran el 30 y se pagan el 16** del mes siguiente.
- A veces **presta dinero** a personas (sale de una billetera) y le pagan en **abonos** (varios, en distintas billeteras/fechas).
- **IVA por categoría** (19%, 5%, 0%…), configurable.

## 5. Las 18 hojas y qué hace cada una
1. **Dashboard** — KPIs del mes, gráficos, alertas, deuda de tarjetas, préstamos. (solo resultados)
2. **Catálogo** — maestro de productos. Tú llenas: Código, Nombre, Categoría, Estado, Stock
   Mínimo, URL Imagen, Precio Detal, Precio Mayor, Medidas. Calcula: costo promedio, stock por
   ubicación, IVA%, días en bodega, vendidas, reabastecer. **Otras hojas dependen de Compras!B,D,F y de las claves de Ventas — no mover esas columnas.**
3. **Compras** — entradas de mercancía. Suma stock a Casa, recalcula costo promedio. Campos de
   pago: Medio de Pago (billetera o tarjeta), ¿A crédito?, fechas de cierre/llegada. "Inventario
   inicial (ya pagado)" = cargar productos viejos sin afectar billeteras.
4. **Ventas** — descuenta inventario; calcula Utilidad Neta (valor − costo − costos fijos);
   tiene Forma de Cobro (billetera) y N° Factura.
5. **Cierre Diario** — cuadre del día: ingresos por método, MELI aparte, gastos, préstamos del
   día, ganancia, cuadre de efectivo e historial de 30 días.
6. **Billeteras** — saldo en tiempo real de cada billetera (Saldo Inicial editable + automático).
7. **Tarjetas** — deuda por tarjeta (compras a crédito − pagos).
8. **Movimientos** — transferencias entre billeteras, retiros de MELI/Skydrops, gastos, aportes.
9. **Pagos Tarjeta** — abonos a tarjetas, indicando de qué billetera salió.
10. **Traslados** — mover unidades Casa ↔ Mercado Libre.
11. **Ajustes** — daños/pérdidas/correcciones (cantidad con signo).
12. **Préstamos** — préstamos a personas, con Billetera origen; saldo y alerta de antigüedad.
13. **Abonos Préstamos** — pagos que devuelven, con fecha y forma de cobro.
14. **Inventario** — reporte consolidado del catálogo (solo lectura).
15. **Clientes** — base de clientes.
16. **Factura** — **se autollena eligiendo una venta (N° Factura)**; PDF con área de impresión.
17. **Catálogo Venta** — catálogo comercial con foto, medidas, precio (detal/mayor) y unidades
    disponibles en Casa; filtra por categoría o TODAS; PDF.
18. **Config** — datos de empresa, IVA por categoría, listas de billeteras/tarjetas/categorías, costos fijos.

## 6. Reglas que NO se deben romper
- El **Catálogo** calcula stock/costo leyendo `Compras` (col B Código, D Cantidad, F Costo Total),
  `Ventas` (clave y cantidad), `Traslados` y `Ajustes`. **No cambiar esas columnas de posición.**
- **Billeteras** (lo que tienes) y **Tarjetas** (lo que debes) son cosas distintas.
- Solo se escribe en celdas de **entrada** (blancas/amarillas); las **azules** son fórmulas.
- La **Factura** y el **Catálogo Venta** tienen **área de impresión** fija para que el PDF salga limpio.

## 7. Qué quiere aprender / objetivo del dueño
Manejar y **entender mejor sus finanzas**: inventario, costos, **ganancias reales**, deudas de
tarjetas, **saldos por billetera**, préstamos y el **cierre diario**. Quiere ir aprendiendo a
moverse solo consultando el **MANUAL** en vez de preguntar cada cosa.

## 8. Cómo ayudar sin confundirse
- Para cambios: **editar `build_xlsx.py` → regenerar → validar → subir a la rama** `sistema-inventario`.
- Mantener fórmulas estándar y moneda COP; no romper dependencias del Catálogo ni de Tesorería.
- Responder **en español**, claro, con ejemplos numéricos y paso a paso.
- El usuario abre el archivo en **Google Sheets** (no Excel), así que FILTER/QUERY/IMAGE funcionan.
- Antes de decir que algo está "mal", verificar: muchas veces es un tema de **fecha** (el Cierre
  filtra por la fecha del día) o de **datos no llenados** (ej. URL de imagen vacía), no un error.
