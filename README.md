# 📊 Dashboard KPI Full-Stack

Este repositorio contiene una aplicación **Dash (Python)** que muestra un panel de control (dashboard) con indicadores clave de rendimiento (KPIs). Los datos se actualizan automáticamente desde un archivo Excel y se almacenan en una base de datos MySQL en la nube. La app está desplegada públicamente y se actualiza cada vez que se sube una nueva versión del archivo de datos.

---

## 🏗️ Arquitectura del Sistema

El sistema sigue una arquitectura full-stack moderna con integración continua y despliegue automático:

| Componente        | Rol                                                                 |
|-------------------|---------------------------------------------------------------------|
| 💻 **Tu PC**       | Entorno de desarrollo local. Aquí editas el archivo Excel y el código. |
| 🐙 **GitHub**      | Única fuente de verdad: almacena tanto el código como `FullStack_Consolidado.xlsx`. |
| 🚀 **Render**      | Plataforma de despliegue (PaaS). Ejecuta la app Dash en la nube.      |
| 🚄 **Railway**     | Base de datos como servicio (DBaaS). Aloja la base de datos MySQL.    |
| 🧍 **Usuario Final**| Accede al dashboard desde cualquier navegador mediante una URL pública. |

🔗 **URL del Dashboard**: [https://dashboard-kpi-fullstack.onrender.com](https://dashboard-kpi-fullstack.onrender.com)

---

## 🔄 Flujo de Actualización de Datos

Cada vez que actualizas los datos, el sistema sigue este proceso automático:

1. **📝 Actualización Local**  
   Reemplazas el archivo `FullStack_Consolidado.xlsx` en tu entorno local.

2. **📤 Subida a GitHub**  
   Ejecutas:
   ```bash
   git add .
   git commit -m "Actualización de datos"
   git push origin main
