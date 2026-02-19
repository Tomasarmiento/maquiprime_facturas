# Procesador de Facturas MAQUIPRIME (Aplicación de Escritorio)

Aplicación de escritorio en Python (Tkinter) para usuarios no técnicos.

## Funcionalidades
- Interfaz gráfica con botones (sin terminal).
- Procesamiento masivo de CFDI 4.0 XML desde carpetas por mes/empleado.
- Carga en `FICHERO_CONTROL_2026.xlsx` en el sheet mensual correspondiente por fecha.
- Validaciones:
  - RFC receptor igual a `MES2301274X9`.
  - Fecha dentro de 2026.
- Resaltados:
  - Amarillo: factura en carpeta de mes distinto al de su fecha.
  - Rojo: UUID duplicado (incluye el registro previo y el nuevo).
- Orden de cada sheet: Empleado (A-Z), luego Fecha (asc).
- Log de ejecución y errores.
- Si existen filas previas con fecha vacía o inválida en el Excel, el proceso ya no se detiene: se reporta advertencia y esas filas se ordenan al final.

## ¿Dónde está la URL de `FICHERO_CONTROL_2026.xlsx`?
No hay URL. Esta versión es **desktop local** y trabaja con la carpeta sincronizada de Dropbox en la computadora.

Debes seleccionar en la app:
1. La carpeta local `2026`.
2. El archivo `FICHERO_CONTROL_2026.xlsx` dentro de esa misma carpeta.

Ruta ejemplo:
```text
.../Joaquin GL Dropbox/MOLGROUP - URUGUAY/MAQUIPRIME/Gastos MX/2026/FICHERO_CONTROL_2026.xlsx
```

La app ahora incluye botón **Autodetectar Excel** para encontrar ese archivo automáticamente dentro de `2026`.

## Instalación
```bash
python -m venv .venv
source .venv/bin/activate  # En Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Uso
```bash
python app.py
```

Desde la interfaz:
1. Selecciona la carpeta `2026`.
2. (Opcional) usa **Autodetectar Excel**.
3. Verifica que el archivo sea `FICHERO_CONTROL_2026.xlsx`.
4. (Opcional) activa modo simulación.
5. Presiona **Procesar facturas**.

## Estructura esperada
```text
2026/
├── FICHERO_CONTROL_2026.xlsx
├── Enero/
│   ├── Empleado 1/
│   │   └── *.xml
│   └── Empleado 2/
├── Febrero/
└── ...
```

## Notas
- La columna `Comentarios` se deja vacía.
- `Empleado` se toma exactamente del nombre de la carpeta.
- Si `ClaveProdServ` no está mapeado, se guarda el código crudo.
