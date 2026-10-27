# Instalar Salidas QR (Google Drive + GitHub)

Hay dos piezas:
- **Servidor y base de datos** en Google Drive (Codigo.gs). Guarda alumnos, familias, fotos y salidas.
- **Página** en GitHub Pages (index.html). Es lo que abren padres, maestros y dirección en el celular.

La instalación se hace una sola vez desde una computadora (unos 20 minutos). Después, la dirección hace todo desde el celular.

Usa la **cuenta de Google de la escuela** (Workspace). Con una cuenta Gmail normal, Google solo permite enviar correos a 100 destinatarios al día; con la de la escuela, 1,500.

## Parte 1 · Google Drive

1. En drive.google.com: **Nuevo → Hojas de cálculo de Google**. Nómbrala **Salidas QR – Base de datos**.
2. En la hoja: **Extensiones → Apps Script**.
3. Borra lo que trae `Código.gs` y pega todo el contenido de **Codigo.gs**. Guarda.
4. Engrane **Configuración del proyecto** → Zona horaria: **(GMT-06:00) Monterrey**.
5. Arriba elige la función **instalar** y toca **Ejecutar**. Acepta los permisos (Configuración avanzada → Ir a … → Permitir).
6. Regresa a la hoja: aparece la **contraseña de dirección**. Anótala.
7. En Apps Script: **Implementar → Nueva implementación → Aplicación web**.
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
8. Toca **Implementar** y copia la **URL de la aplicación web** (termina en `/exec`).

### Probar con datos ficticios (opcional)

Recarga la hoja de cálculo: aparece el menú **Salidas QR**. Elige **Cargar datos de prueba** (o en Apps Script ejecuta la función `cargarDatosDePrueba`). Se crean 116 alumnos, 96 familias, 12 maestros y una semana de salidas, siempre los mismos.

- Dirección: **direccion / direccion2026**
- Maestro Carlos Estala (4°A): **carlos / carlos2026**
- Demás maestros: su primer nombre sin acentos (jorge, patricia, laura…) / **maestro2026**
- Padres: CURP del hijo mayor (pestaña Alumnos) + PIN (pestaña Familias)

Antes de usarla con alumnos reales: menú **Salidas QR → Borrar todos los datos** y cambia la contraseña de dirección desde la página.

## Parte 2 · GitHub (la página y su link)

El código ya está en este repositorio (`index.html` y `servidor/Codigo.gs`). Solo falta:

1. Activar la página: **Settings → Pages**. En "Source" elige **Deploy from a branch**, rama **main**, carpeta **/ (root)**. Toca **Save**.
2. Esperar 1 o 2 minutos y recargar esa pantalla: arriba aparece el link de la página. Ese es el link para padres, maestros y dirección.
3. Conectar la página con Drive: abre **index.html** → lápiz (**Edit**) → en la línea `const API_URL = 'PEGA_AQUI_LA_URL_DE_APPS_SCRIPT';` reemplaza el texto entre comillas por la URL de Apps Script (paso 8 de la Parte 1) → **Commit changes**.
4. Abrir el link en el celular y en el menú del navegador elegir **Agregar a pantalla de inicio**.

**Para actualizar la página:** edita o reemplaza index.html en el repositorio. En un par de minutos se ve el cambio.

**Si después quieres usar tu dominio:** Settings → Pages → Custom domain, escribe tu dominio y sigue las instrucciones de DNS que te da GitHub. Activa **Enforce HTTPS**.

**Qué es público y qué no:** el código de la página es visible para cualquiera, pero no contiene datos. Los alumnos, PIN y fotos viven solo en tu Google Drive, y para verlos hay que iniciar sesión. Nunca subas la hoja de cálculo ni contraseñas al repositorio.

## Parte 3 · Empezar a usarla (desde el celular, como dirección)

1. Entra con usuario **direccion** y la contraseña anotada. Cámbiala en **Más → Contraseña**.
2. **Más → Importar alumnos**: descarga la plantilla, llénala (o usa tu Excel con columnas Nombre, CURP, Grupo, Tutor, Teléfono, Correo, CURP hermano mayor) y súbela. Cada familia nueva recibe un PIN.
3. **Maestros → Agregar maestro**: uno por grupo. Comparte a cada uno su usuario y contraseña.
4. Con las hojas de autorización firmadas: busca al alumno en **Alumnos**, abre su ficha y agrega las **personas autorizadas** con foto (puedes tomarla con el celular).
5. Entrega a cada familia su CURP de acceso (el del hijo mayor) y su PIN. El padre ve el QR en su cuenta.

## Uso diario

- **Maestro o tablet en la puerta:** Escanear → Abrir cámara → apunta al QR → toca la foto de quien recoge → Registrar salida → Escanear siguiente. El tutor recibe un correo.
- También funciona escaneando con la cámara normal del celular: abre la página con el alumno.
- **Otro grupo:** el maestro titular da permiso en Permisos (dura solo ese día).
- **Fin de ciclo:** Más → Cerrar ciclo escolar. Sube a todos de grado, egresa a 6° y pasa la cuenta al siguiente hermano.

## Si cambias el código del servidor

Implementar → Administrar implementaciones → lápiz → Versión: **Nueva versión** → Implementar. Así la URL no cambia y no hay que tocar index.html.

## Seguridad

- No compartas la hoja de cálculo: tiene CURP, PIN y contraseñas.
- La sesión se cierra sola a las 6 horas.
- Si un padre pierde el QR o lo compartió, que lo cambie en su cuenta, o cámbialo tú desde la ficha del alumno.
