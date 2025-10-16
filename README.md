                                                                                           Diagrama del Flujo del Sistema de TI
                                                                                           
Componentes del Sistema
Tu PC (Desarrollador) 💻: Es tu entorno local donde actualizas el archivo Excel y gestionas el código fuente.

GitHub 🐙: Actúa como el repositorio central y la "única fuente de la verdad". Almacena no solo tu código Python, sino también el archivo de datos FullStack_Consolidado.xlsx.

Render 🚀: Es la plataforma de despliegue (PaaS). Su trabajo es ejecutar tu aplicación Dash en la nube y hacerla accesible a través de una URL pública.

Railway 🚄: Es tu proveedor de base de datos como servicio (DBaaS). Aloja tu base de datos MySQL en la nube, haciéndola accesible desde cualquier lugar.

Usuario Final 🧍: La persona que interactúa con tu dashboard a través de su navegador web.

El Flujo de Actualización (Paso a Paso)
Este es el proceso completo que ocurre cada vez que actualizas los datos:

Actualización Local: Tú, como desarrollador, reemplazas el archivo FullStack_Consolidado.xlsx en la carpeta de tu proyecto en tu PC con la nueva versión de los datos.

Subida a GitHub: Usando la terminal Git Bash, ejecutas los comandos git add ., git commit, y git push. Esto sube el nuevo archivo Excel y cualquier otro cambio de código a tu repositorio en GitHub.

Despliegue Automático (CI/CD): Render está configurado para "escuchar" los cambios en tu repositorio de GitHub. En el momento en que detecta el push, inicia un nuevo despliegue automáticamente.

Ejecución de la Aplicación en Render:

a. Build: Render instala todas las librerías de tu archivo requirements.txt.

b. Run: Render ejecuta tu script dashboard_kpi_DB.py. Lo primero que hace tu script al arrancar es llamar a la función actualizar_y_cargar_datos_desde_excel().

Sincronización de Datos:

a. Lectura: La aplicación en Render lee el nuevo archivo Excel que acaba de descargar desde GitHub.

b. Escritura: La aplicación se conecta a tu base de datos en Railway y usa if_exists='replace' para borrar la tabla antigua y cargar los datos frescos del Excel.

Carga del Dashboard:

a. Lectura de la DB: Una vez que los datos están en Railway, la aplicación de Render los lee desde la base de datos para construir las tablas y gráficos. Este paso es muy rápido.

b. Interfaz de Usuario: Render sirve la aplicación Dash, que ya está lista con los datos actualizados.

Acceso del Usuario: El usuario final navega a tu URL (https://dashboard-kpi-fullstack.onrender.com), se autentica, y ve el dashboard con la información más reciente. Su navegador se comunica directamente con la aplicación que se está ejecutando en Render.

<img width="640" height="650" alt="image" src="https://github.com/user-attachments/assets/dfc9bc1d-1df8-42b7-aaab-aee987b77ed3" />
