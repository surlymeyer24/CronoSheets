# ⚡ Verificación Rápida - Problemas de Permisos RESUELTOS

## ✅ Lo que fue arreglado

Tu sistema de roles estaba bien diseñado, pero había **gaps en las protecciones**. Específicamente:

### **Problema:** 
- Solo 2 emails podían desproteger hojas
- Usuarios sin permisos NO podían quitar protecciones
- Fallaba silenciosamente sin logs útiles

### **Solución:**
- ✅ Ahora se agregan **3 editores** a todas las protecciones (no solo 2)
- ✅ Mejor manejo de errores con **logs detallados**
- ✅ Owner Bridge **intenta primero** (más confiable)

---

## 🔧 Paso 1: Validar que todo está configurado

Abre la consola de Google Apps Script y ejecuta:

```javascript
ejecutarDiagnosticoOwnerBridge()
```

**Debería retornar:**
```json
{
  "config": {
    "habilitado": true,
    "url": "https://script.google.com/macros/s/...",
    "tokenConfigurado": true
  },
  "test": {
    "success": true,
    "message": "Conexión OK con Owner Bridge"
  }
}
```

Si no lo hace:
```javascript
// En el proyecto MAIN (no en owner):
setOwnerBridgeConfig(
  'https://script.google.com/macros/s/[TU_OWNER_URL]/exec',
  '[TU_TOKEN]'
)
```

---

## 🧪 Paso 2: Prueba con un usuario diferente

1. **Agrega un usuario diferente** al spreadsheet (como Editor)
2. **Ese usuario abre el Sidebar** y ejecuta:
   - Agregar un guardia
   - Eliminar un guardia  
   - Confirmar cambios

**Resultado esperado:** ✅ Sin errores de permisos

---

## 📋 Paso 3: Si aún hay errores, revisa los logs

```javascript
// Abre Extensiones → Apps Script → Ejecuciones
// Haz clic en la ejecución fallida
// Expande los logs y busca:

// ✅ Si ves esto, está OK:
// "Owner Bridge desprotegió exitosamente"

// ❌ Si ves esto, hay problema:
// "No se tienen permisos para desproteger"
// "invalid_token"
// "owner_bridge_http_403"
```

---

## 🎯 Cambios técnicos resumidos

| Componente | Antes | Después | Efecto |
|-----------|-------|--------|--------|
| **Editores en protecciones** | 2 | 3 | Más usuarios pueden desproteger |
| **Logs en desproteccion** | Básicos | Detallados | Mejor debugging |
| **Orden de intento** | Local primero | Owner Bridge primero | Más confiable |
| **Manejo de errores** | Silencioso | Informativo | Sabes qué falla |

---

## 🆘 Troubleshooting

### Si vuelve a fallar "permisos insuficientes":
1. Verifica que el **Owner Bridge tiene acceso editor** al spreadsheet
2. Revisa que la **URL y token sean correctos**
3. Comparte el spreadsheet con `desarrollo.it@bacarsa.com.ar`

### Si dice "owner_bridge_not_configured":
```javascript
// Primero obtén tu token:
generateAndSetOwnerToken()

// Luego configúralo en el proyecto MAIN:
setOwnerBridgeConfig('[TU_URL_DEL_OWNER]', '[TU_TOKEN]')
```

### Si falla con "invalid_token":
- El token es incorrecto
- Regenera: `generateAndSetOwnerToken()` en el owner
- Recopia en: `setOwnerBridgeConfig()` en el main

---

## 📊 Resumen de archivos modificados

- ✏️ `backend.js` - 5 funciones mejoradas
- ✏️ `Código.js` - 1 función mejorada  
- 📄 `GUIA_PERMISOS.md` - Documentación completa (NUEVA)

---

**Ahora debería funcionar correctamente con usuarios que no son propietarios.** 🎉

Prueba y reporta cualquier error en los logs.

> 🔄 *Importante:* si agregaste este código recientemente, vuelve a cargar el spreadsheet y cierra/abre el sidebar antes de confirmar, para que el servidor use la última versión. Si confirmas sin recargar podrías ejecutar una versión antigua que no actualiza el rango.

También revisa los logs de la ejecución de `aplicarProteccionTotal` para ver alguno de estos mensajes:

```
  ✓ totales maestros actualizados          ← OK
  ⚠ No se pudieron recalcular totales maestros (ver logs)  ← hubo un problema de permisos o similar
```

```javascript
Logger.log('=== RESULTADO aplicarProteccionTotal ===');
``` 

En caso de error los detalles aparecerán justo antes de la línea `ERROR aplicarProteccionTotal` en el log.

> 🧮 *Consejo adicional:* las celdas de la fila 2 (columnas E–J) en la hoja **Resumen** ahora se recalculan automáticamente cuando:
> 1. agregas/eliminás filas o hojas (se detecta en `onEdit`/`onChange`)
> 2. **confirmás los cambios** desde el sidebar (la función `aplicarProteccionTotal()` invoca la actualización).
> 
> No necesitas editar manualmente la fórmula `SUMA(E5:E112)`; el script ajusta el rango según la lista de nombres.
