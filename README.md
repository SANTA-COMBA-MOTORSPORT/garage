# Couto Garage Inventario

App de inventario del Couto Garage: plano de taller, living y foso, vista 3D de cada mueble,
acceso por móvil y PIN, y Gemini para fotos y búsquedas.

- `index.html`: la app (va en GitHub Pages).
- `Code.gs`: el servidor (va en Apps Script, dentro de la hoja de cálculo).
- `CNAME`: el dominio `garage.santacombamotorsport.es`.

## Puesta en marcha (unos 15 minutos, con la cuenta de SC Motorsport)

### 1. Hoja y servidor
1. Crea una hoja de cálculo de Google nueva y llámala `Couto Garage Inventario`.
2. Menú **Extensiones > Apps Script**. Borra lo que haya, pega todo `Code.gs` y guarda.
3. Arriba, elige la función **instalar** y pulsa **Ejecutar**. Acepta los permisos.
   Crea las pestañas, los 19 palets, las 47 cajas y el aviso de los lunes.
4. En la pestaña **Usuarios** rellena móvil y PIN de cada uno (rol `admin` o `usuario`).
   Sin móvil y PIN nadie puede entrar.

### 2. Gemini
1. Crea una clave en Google AI Studio (Get API key).
2. En Apps Script: **Configuración del proyecto > Propiedades de la secuencia de comandos >
   Añadir propiedad**. Nombre `GEMINI_API_KEY`, valor la clave.
3. Vuelve a la hoja, recárgala y usa el menú **Couto Garage > Probar Gemini**.

### 3. Publicar el servidor
1. En Apps Script: **Implementar > Nueva implementación > Aplicación web**.
2. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier usuario**.
3. Copia la URL que termina en `/exec`.

### 4. Publicar la app
1. Abre `index.html`, busca `var API_URL='';` y pega la URL entre las comillas.
2. Sube `index.html` y `CNAME` al repositorio de GitHub y activa **Pages** en la rama principal.
3. En Dinahosting crea un registro **CNAME**: `garage` apuntando a `santa-comba-motorsport.github.io`.
4. Cuando el dominio responda, marca **Enforce HTTPS** en Pages.

## Cuando cambies Code.gs
**Implementar > Gestionar implementaciones > editar > Versión: nueva versión**. La URL no cambia.

## Dónde está cada cosa en la hoja
| Pestaña | Contenido |
|---|---|
| Articulos | Inventario. Se puede editar a mano; la app lo recoge al actualizar |
| Cajas | Cajas y organizadores, y en qué hueco están |
| Palets, Posiciones | Colocación de palets y muebles en el plano |
| Usuarios | Móvil, PIN, nombre, rol y activo (`NO` para dar de baja) |
| Movimientos | Registro de quién cambió qué |
| Fotos | Miniaturas de los artículos |
| Ubicaciones | Lista de códigos (se rellena desde la app, en Ubicaciones) |
