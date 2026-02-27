# 📋 CronoSheets - Sistema de Gestión de Guardias

## 🎯 Descripción General

Aplicación de Google Sheets para gestionar guardias y cronogramas de forma colaborativa con control de permisos mediante **Owner Bridge Pattern**.

### Características
- ✅ Agregar guardias en batch
- ✅ Eliminar guardias con protección
- ✅ Importar desde Excel
- ✅ Hojas individuales protegidas
- ✅ Resumen automatizado
- ✅ **Usuarios con permisos reducidos pueden hacer cambios** (NUEVO)

---

## 🏗️ Arquitectura

El sistema usa **2 proyectos de Google Apps Script**:

```
┌────────────────────────────────────┐
│ PROYECTO MAIN (backend.js)         │
│ ├─ Sidebar.html                    │
│ ├─ Código.js (triggers)            │
│ └─ Usuarios sin permisos ejecutan  │
└────────────────────────────────────┘
              ↓ (HTTP POST con TOKEN)
        [OWNER BRIDGE]
              ↓
┌────────────────────────────────────┐
│ PROYECTO OWNER (/owner)            │
│ ├─ Código.js (endpoints)           │
│ └─ Ejecuta como PROPIETARIO        │
└────────────────────────────────────┘
```

---

## ⚙️ Setup Inicial

### Paso 1: Deployar el Owner Bridge

**En el proyecto `/owner`:**

```javascript
// 1. Genera un token único
generateAndSetOwnerToken()
// Copia el token de los logs
```

Luego:
1. **Extensiones → Apps Script → Implementar (⊙) → Nuevas implementaciones**
2. Tipo: `Web app`
3. Ejecutar como: Tu email
4. Quién tiene acceso: `Cualquier persona`
5. **Copia la URL** que aparece

### Paso 2: Configurar el Cliente

**En el proyecto MAIN:**

```javascript
setOwnerBridgeConfig(
  'https://script.google.com/macros/s/[TU_URL_AQUI]/exec',
  '[TOKEN_QUE_COPIASTE]'
)
```

### Paso 3: Validar Configuración

```javascript
ejecutarDiagnosticoOwnerBridge()
// Debe retornar: success: true, habilitado: true
```

---

## 📊 Cambios Recientes (26 Feb 2026)

### ✅ Problemas Solucionados

| Antes | Ahora |
|-------|-------|
| ❌ Usuarios sin permisos NO podían desproteger hojas | ✅ Owner Bridge maneja desprotecciones |
| ❌ Solo 2 emails podían desproteger | ✅ 3 editores en las protecciones |
| ❌ Logs poco útiles para debugging | ✅ Logs detallados en cada paso |
| ❌ Fallos silenciosos | ✅ Reintentos automáticos |

### 📝 Archivos Modificados

- `backend.js` - Mejorada lógica de desprotección + logs detallados
- `Código.js` - Agregado 3er editor a protecciones
- 📄 Documentación nueva (4 archivos guía)

---

## 🚀 Uso Típico

### Usuario Trabaja Sin Ser Propietario

```
1. Abre Sidebar: ⚙️ Gestión de Guardias → 📋 Abrir Panel
2. Agregar/Eliminar guardias normalmente
3. Sistema automáticamente:
   - Intenta desproteger vía Owner Bridge (preferencia)
   - Si falla, intenta localmente (fallback)
   - Reprotege con 3 editores distribuidos
4. Fin - Usuario no necesita permisos especiales
```

### Verificar que Funciona

```javascript
// Comando rápido para validar TODO:
ejecutarDiagnosticoOwnerBridge()

// Esperado:
{
  config: { habilitado: true, tokenConfigurado: true },
  test: { success: true, message: "Conexión OK" }
}
```

---

## 📖 Documentación Rápida

**⭐ EMPEZAR AQUÍ:**
- [VERIFICACION_RAPIDA.md](VERIFICACION_RAPIDA.md) - Qué fue arreglado (5 min)

**Luego:**
- [PRUEBAS_Y_TROUBLESHOOTING.md](PRUEBAS_Y_TROUBLESHOOTING.md) - Testing paso a paso
- [OWNER_BRIDGE_EXPLICACION.md](OWNER_BRIDGE_EXPLICACION.md) - Cómo funciona realmente
- [GUIA_PERMISOS.md](GUIA_PERMISOS.md) - Guía completa de solución

---

## 🧪 Testing Rápido

### Test 1: Diagnóstico
```javascript
ejecutarDiagnosticoOwnerBridge()
```
✅ Debe retornar `success: true`

### Test 2: Usuario Normal
1. Comparte el sheet con usuario diferente (Editor)
2. Usuario agrega/elimina guardias desde Sidebar
3. Debe funcionar sin errores de permiso

### Test 3: Revisar Logs
```
Extensiones → Apps Script → Ejecuciones → Última ejecución
Busca: "Owner Bridge desprotegió exitosamente"
```

---

## 🆘 Troubleshooting Rápido

| Error | Causa | Solución |
|-------|-------|----------|
| invalid_token | Token no coincide | Regenera: `generateAndSetOwnerToken()` |
| HTTP 403/404 | URL incorrecta | Redeploy Web App en owner |
| No permisos | Owner Bridge sin acceso | Revisa configuración |

**Ver [PRUEBAS_Y_TROUBLESHOOTING.md](PRUEBAS_Y_TROUBLESHOOTING.md) para más detalles**

---

## 📞 Funciones Útiles

```javascript
// Diagnóstico
ejecutarDiagnosticoOwnerBridge()      // Revisa todo
verOwnerBridgeConfig()                 // Solo configuración
testOwnerBridgeConexion()              // Solo conectividad

// Configuración
setOwnerBridgeConfig(url, token)       // Configura
verOwnerBridgeConfig()                 // Verifica

// Owner (ejecutar en /owner)
generateAndSetOwnerToken()             // Genera token único
verToken()                             // Muestra token actual

// Útiles
aplicarProteccionTotal()               // Reprotege todo
repararProteccionResumen()             // Arregla resumen
borrarHojasNoEnResumen()               // Limpia hojas huérfanas
```

---

## 📋 Estructura del Repo

```
/
├── backend.js                      # Lógica principal de agregar/eliminar
├── Código.js                       # Triggers y funciones del sheet
├── Sidebar.html                    # Interfaz del panel
├── Estilos.html                    # CSS del sidebar
├── appsscript.json                 # Config del script MAIN
├── README.md                       # Guía rápida (anterior)
├── README_NUEVO.md                 # Este archivo (mejorado)
├── VERIFICACION_RAPIDA.md          # ⭐ Qué fue arreglado - LEER PRIMERO
├── GUIA_PERMISOS.md               # Guía detallada de solución
├── OWNER_BRIDGE_EXPLICACION.md    # Arquitectura en profundidad
├── PRUEBAS_Y_TROUBLESHOOTING.md   # Testing y resolución
└── owner/                         # Proyecto separado (Owner Bridge)
    ├── Código.js                  # Endpoints del bridge
    └── appsscript.json            # Config del owner
```

---

## 🔐 Seguridad

- ✅ Token único en cada llamada al Owner Bridge
- ✅ Owner Bridge solo ejecuta acciones válidas
- ✅ Permisos de sheet limitados por protecciones
- ✅ Logs de todas las operaciones
- ✅ Web App solo responde a POST con token correcto

---

## 💡 Notas Importantes

1. **Owner Bridge es PREFERENCIA:** El sistema intenta vía Owner Bridge primero, completa la operación, y no intenta localmente si funciona.

2. **3 Editores ahora:** Cualquiera de estos puede desproteger:
   - `desarrollo.it@bacarsa.com.ar`
   - `implementaciones.it@bacarsa.com.ar`
   - El usuario que ejecuta la acción

3. **Logs Detallados:** Cada paso deja un log claro. Revisa `Ejecuciones` si algo falla.

4. **Fallback Automático:** Si Owner Bridge no responde, intenta localmente. Si eso también falla, sabes exactamente por qué.

---

## 📞 Contacto & Soporte

Si encuentras problemas:
1. Lee [VERIFICACION_RAPIDA.md](VERIFICACION_RAPIDA.md) (5 minutos)
2. Ejecuta `ejecutarDiagnosticoOwnerBridge()`
3. Revisa los logs en Ejecuciones
4. Reporta con logs exactos del paso 3

---

**Última actualización:** 26 de febrero de 2026  
**Versión:** 2.0 (Con fixes de permisos)  
**Estado:** ✅ Producción - Todos los usuarios pueden hacer cambios
