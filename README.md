# Foliador Automático PAC / POS · ColDev

Aplicación de escritorio autónoma diseñada para la planta pesquera. Elimina la necesidad de revisar manualmente el planillero físico, abrir Word y tipear números de folio correlativos uno a uno.

---

## 💎 Características Principales

1. **Preservación 100% de la Estructura Word Original**:
   - Trabaja directamente sobre las copias binarias de las plantillas oficiales PAC V15, POS V12 y Prerrequisitos V16.
   - Conserva sellos de SERNAPESCA, logos de la planta, tablas, bordes y fuentes sin alterar namespaces de OpenXML.
2. **Folios Centrados y Formateados**:
   - Cada folio sale centrado quirúrgicamente (`<w:jc w:val="center"/>`) en el encabezado oficial.
   - En **PC 3** se consolidó la casilla en un único rectángulo centrado (sin la celda residual de `000104`).
3. **Catálogo Completo (12 Planillas Oficiales)**:
   - **PAC**: PCC 1, PC 2, PC 3, PC 4, PC 5, PC 6, PC 7, Días Productivos V15.
   - **POS**: POS 1, POS 2, POS 3, POS 4.
4. **Impresión Directa**:
   - Envío directo a la impresora predeterminada de Windows (ej. `EPSON43D68B L6270 Series`) al momento de generar o mediante el botón de imprimir lote.
5. **Ejecutable Standalone (.EXE)**:
   - Incluye `FoliadorPAC_POS.exe` listo para funcionar en cualquier PC con Windows sin necesidad de tener Python instalado.

---

## 🚀 Cómo Usar

### Opción 1: Ejecutable Directo (Recomendado para Operadores)
- Hacé doble clic en **`FoliadorPAC_POS.exe`** (o en `iniciar_foliador.bat`).
- Seleccioná la planilla en el desplegable de alto contraste.
- Indicá la cantidad de hojas y hacé clic en **`🚀 GENERAR PLANILLAS FOLIADAS`**.
- Podés activar la casilla `Mandar a imprimir automáticamente a la impresora` para que salgan por la bandeja al instante.

### Opción 2: Sincronizar con el Planillero Físico
- Hacé clic en **`✏️ Sincronizar Físico`**.
- Ingresá el último número de folio que ya está impreso/usado físicamente.
- El sistema se alineará y la próxima planilla continuará el correlativo exacto.

---

## 📦 Publicación en GitHub y Releases

Para publicar este proyecto en un nuevo repositorio de GitHub y subir el ejecutable para descarga pública:

```bash
# 1. Crear el repositorio en GitHub (o usar GitHub CLI: gh repo create)
gh repo create foliador-pac-pos --public --source=. --remote=origin --push

# 2. Crear un Release con el ejecutable .exe adjunto:
gh release create v1.0.0 dist/FoliadorPAC_POS.exe --title "Foliador Oficial PAC/POS v1.0.0" --notes "Versión inicial autónoma con 12 planillas oficiales, centrado de folios e impresión directa Epson."
```

Alternativamente por la web de GitHub:
1. Creá el repositorio en tu cuenta.
2. Hacé `git push -u origin main`.
3. Andá a **Releases** > **Draft a new release**.
4. Arrastrá el archivo `dist/FoliadorPAC_POS.exe` al área de assets y publicá.

---

## 📁 Estructura del Código

- [`app_gui.py`](file:///c:/Users/the_8/Desktop/PERFIL%20JOSE%20COLIVORO%20URIBE%20-COLDEV,%20DESARROLLADOR/02-products/foliador-automatico-pac-pos/app_gui.py): Interfaz gráfica Tkinter de alto contraste.
- [`core/folio_engine.py`](file:///c:/Users/the_8/Desktop/PERFIL%20JOSE%20COLIVORO%20URIBE%20-COLDEV,%20DESARROLLADOR/02-products/foliador-automatico-pac-pos/core/folio_engine.py): Motor central de inyección de folios, control de estado e impresión Word COM.
- [`config/plantillas.json`](file:///c:/Users/the_8/Desktop/PERFIL%20JOSE%20COLIVORO%20URIBE%20-COLDEV,%20DESARROLLADOR/02-products/foliador-automatico-pac-pos/config/plantillas.json): Catálogo maestro de planillas.
- [`data/estado_folios.json`](file:///c:/Users/the_8/Desktop/PERFIL%20JOSE%20COLIVORO%20URIBE%20-COLDEV,%20DESARROLLADOR/02-products/foliador-automatico-pac-pos/data/estado_folios.json): Base de datos local de folios e historial de auditoría.
- [`dist/FoliadorPAC_POS.exe`](file:///c:/Users/the_8/Desktop/PERFIL%20JOSE%20COLIVORO%20URIBE%20-COLDEV,%20DESARROLLADOR/02-products/foliador-automatico-pac-pos/dist/FoliadorPAC_POS.exe): Binario compilado listo para distribución.
