# 📅 Gestión de Horarios — IES Virgen de la Calle (Palencia) 2025/26

Aplicación web para la gestión de horarios, ausencias, guardias y actividades del IES Virgen de la Calle de Palencia.

## 🌐 URLs

- **Producción**: https://horarios-iesvc-2526.netlify.app
- **Repositorio**: https://github.com/LuisGP1984/Horarios_IES_VC_25_26 (privado)

## 🗂️ Archivos

| Archivo | Descripción |
|---------|-------------|
| `index.html` | Aplicación completa (HTML + CSS + JS) |
| `horarios.json` | Datos de horarios generados desde Excel con macro VBA |
| `manifest.json` | Configuración PWA para instalar como app |
| `logo.png` | Logo del centro |

## 🔐 Acceso

| Rol | Contraseña | Permisos |
|-----|-----------|----------|
| Docente | `VC2526` | Consulta (horarios, guardianes, ausencias, aulas, ACEX) |
| Dirección | `EGL2530` | Consulta + gestión (ausencias, sustituciones, ACEX, reservas) |

## 📱 Módulos

### ⚡ Ahora
- Sesión actual en curso con reloj en tiempo real
- Guardianes asignados a esta sesión
- Guardias a cubrir (docentes ausentes con clase)
- Buscador de aula libre
- ACEX activas hoy
- Selector de fecha para consultar otro día

### 📅 Horarios
- Horario completo de cualquier docente por día
- Muestra sustitutos con 🔄 si hay sustitución activa
- Sesiones con materia, grupo, aula y tipo

### 🛡️ Guardianes
- Tabla semanal de guardianes por sesión (L-V)
- Contadores por sesión con semáforo de color
- Actualizado automáticamente con sustituciones activas

### 📋 Ausencias *(Dirección)*
- Registro de ausencias desde Supabase (compartido entre dispositivos)
- Descarga del parte diario en Word (.docx) y PDF (2 páginas)
- Muestra nombre del sustituto si hay sustitución activa

### 🔴 Cobertura
- Semáforo semanal: guardianes disponibles vs ausencias a cubrir
- Carga automática desde Supabase al abrir

### 🏫 Aulas
- Ocupación de todas las aulas por fecha y sesión
- Reserva de aulas desde Supabase
- Sesión Vespertina (16:00–21:00)
- Scroll horizontal para ver todas las sesiones

### 🔄 Sustituciones *(Dirección)*
- Sustituciones indefinidas hasta cancelación manual
- El sustituto aparece en todos los desplegables (horarios, ausencias, guardianes)
- El sustituido desaparece de los selectores mientras dure la sustitución

### 🎒 ACEX
- Actividades Complementarias y Extraescolares por semana
- Todas las sesiones marcadas por defecto al crear
- Aparece en el módulo Ahora si hay ACEX activa

### ℹ️ Acerca de
- Información del desarrollador
- Botón "Instalar como app" (PWA)

## 🗄️ Base de datos (Supabase)

```sql
-- Ausencias
ausencias (id, fecha, codigo, nombre, sesiones jsonb, created_at)

-- Reservas de aulas
reservas (id, fecha, aula, sesion, docente, motivo, created_at)

-- Sustituciones activas
sustituciones (id, docente_original, codigo_original, docente_sustituto, 
               fecha_inicio, fecha_fin, activa, created_at)

-- ACEX
acex (id, fecha, sesiones text[], docentes, grupos, observaciones, created_at)
```

## ⏰ Sesiones

| Clave | Nombre | Horario |
|-------|--------|---------|
| 1 | Guardia mañana | 08:00–08:30 |
| 2 | Sesión 1 | 08:30–09:20 |
| 3 | Sesión 2 | 09:25–10:15 |
| 3R | Sesión 3 | 10:20–11:10 |
| 4 | Recreo | 11:10–11:35 |
| 5 | Sesión 4 | 11:35–12:25 |
| 6 | Sesión 5 | 12:30–13:25 |
| 7 | Sesión 6 | 13:30–14:20 |
| 8 | Sesión 7 | 14:25–15:15 |
| V | Vespertino | 16:00–21:00 |

> Viernes no tiene Sesión 7.

## 🛠️ Stack técnico

- **Frontend**: HTML + CSS + JavaScript (vanilla, sin frameworks)
- **Base de datos**: Supabase (PostgreSQL)
- **Hosting**: Netlify
- **Datos de horarios**: JSON generado desde Excel con macro VBA
- **PWA**: manifest.json + meta tags Apple

## 👨‍💻 Autor

**Luis González Posada**  
Jefe de Estudios — IES Virgen de la Calle, Palencia  
luis.gonpos@educa.jcyl.es  
#soydelvirgendelacalle

---
*v2.0 · Junio 2026 · Desarrollado con Claude AI*
