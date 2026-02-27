# 🔐 Guía de Solución - Problemas de Permisos

## Problema Identificado ❌

Cuando un usuario con menos permisos abre el sidebar e intenta **agregar** o **eliminar guardias**, obtenía errores de permisos porque:

1. **Las protecciones estaban creadas solo para 2 emails específicos:**
   - `desarrollo.it@bacarsa.com.ar`
   - `implementaciones.it@bacarsa.com.ar`

2. **Un usuario diferente NO podía quitar protecciones** que no había creado él mismo.

3. **El Owner Bridge fallaba silenciosamente** si no tenía acceso al spreadsheet.

## Soluciones Implementadas ✅

### 1. **Agregar el usuario actual como editor en todas las protecciones**

Antes:
```javascript
p.addEditor(email);
p.addEditor(adminEmail);
```

Después:
```javascript
p.addEditor(email);
p.addEditor(adminEmail);
p.addEditor(emailDes);  // Agregado
```

**Dónde se aplicó:**
- ✅ `Código.js` → `protegerHoja()`
- ✅ `backend.js` → `protegerHojasBatch()`
- ✅ `backend.js` → `aplicarProteccionAHoja()`
- ✅ `backend.js` → `aplicarProteccionTotal()`
- ✅ `backend.js` → `reprotegerResumen()`

### 2. **Mejorar la lógica de desprotección**

**Antes:** Intentaba vía Owner Bridge, pero si fallaba, lo hacía localmente sin log detallado.

**Después:** 
```javascript
function desprotegerHoja(hoja) {
  // FASE 1: Intentar vía Owner Bridge (más confiable)
  const viaOwner = llamarOwnerBridge_('desprotegerHoja', {...});
  if (viaOwner.status === 'ok') {
    Logger.log('  ✓ Owner Bridge desprotegió exitosamente');
    return;
  }
  
  // FASE 2: Intentar localmente con logs detallados
  const protecciones = hoja.getProtections(...);
  protecciones.forEach(p => {
    try { 
      p.remove();
      Logger.log('      ✓ Removida');
    } catch (e) {
      Logger.log('      ✗ Error: ' + e.message);
    }
  });
}
```

**Beneficios:**
- Logs detallados para debugging
- Si falla una protección, continúa con las demás
- Identifica claramente permisos insuficientes

### 3. **Agregar `desarrollo.it@bacarsa.com.ar` en `reprotegerResumen()`**

La función ahora agrega **3 editores** en vez de 2, permitiendo que cualquiera de ellos pueda desproteger:

```javascript
p.addEditor(emailUsuario);  // El que ejecuta ahora
p.addEditor(emailAdmin);    // implementaciones.it@bacarsa.com.ar
p.addEditor(emailDes);      // desarrollo.it@bacarsa.com.ar
```

## Verificación 🔍

Para verificar que todo funciona correctamente:

### Paso 1: Verificar que el Owner Bridge está configurado
```javascript
// En la consola de Google Apps Script:
ejecutarDiagnosticoOwnerBridge()
```

Debería retornar:
```javascript
{
  config: {
    habilitado: true,
    url: "https://...",
    tokenConfigurado: true
  },
  test: {
    success: true,
    message: "Conexión OK con Owner Bridge"
  }
}
```

### Paso 2: Probar con un usuario sin permisos

1. Comparte el spreadsheet con un usuario que **NO sea propietario**
2. Dale rol de **"Editor"**
3. Que abra el Sidebar y ejecute:
   - ✅ Agregar un guardia
   - ✅ Eliminar un guardia
   - ✅ Confirmar cambios

**Esperado:** Funciona sin errores de permisos

### Paso 3: Revisar los logs

Si algo falla:
1. Abre `Extensiones > Apps Script`
2. Ve a `Ejecuciones` 
3. Busca la ejecución más reciente
4. Revisa los logs (icono de ℹ️)
5. Busca mensajes como:
   - "`Owner Bridge desprotegió exitosamente`" ✅
   - "`Error: no tienes permiso para...`" ❌

## Casos de Problema y Soluciones 🛠️

### Caso 1: "No se tienen permisos para desproteger"
**Causa:** El usuario NO es editor de la protección original
**Solución:** El Owner Bridge debería manejar esto. Si sigue fallando:
```javascript
// En la consola del Owner:
verificar que el script owner TIENE acceso al spreadsheet
```

### Caso 2: "Owner Bridge no configurado"
**Solución:** En la consola del proyecto principal:
```javascript
configurarOwnerBridge()
// O manualmente:
setOwnerBridgeConfig(
  'https://[TU_URL_AQUI]',
  '[TU_TOKEN_AQUI]'
)
```

### Caso 3: El Owner Bridge está configurado pero sigue fallando
**Solución:** Verifica que:
1. La URL es correcta y accesible
2. El token es correcto  
3. El script owner tiene acceso al spreadsheet (como editor mínimo)

Ejecuta:
```javascript
testOwnerBridgeConexion()
```

## Cambios en el Código 📝

| Archivo | Función | Cambio |
|---------|---------|--------|
| `Código.js` | `protegerHoja()` | +1 editor (desarrollo.it) |
| `backend.js` | `protegerHojasBatch()` | +1 editor, +logs |
| `backend.js` | `aplicarProteccionAHoja()` | +1 editor |
| `backend.js` | `aplicarProteccionTotal()` | +1 editor, +logs |
| `backend.js` | `desprotegerHoja()` | Logs detallados |
| `backend.js` | `reprotegerResumen()` | +1 editor, reintentos |

## Arquitectura de Flujo ✨

```
Usuario (sin permisos) abre Sidebar
    ↓
Sidebar ejecuta función (ej: agregarGuardiasBatch)
    ↓
Backend (backend.js) ejecuta desprotegerHoja()
    ├─ Intenta vía Owner Bridge
    │  └─ Owner Bridge tiene permisos → ✅ USA OWNER BRIDGE
    └─ Si falla, intenta localmente
       └─ Usuario es agregado como editor → ✅ PUEDE DESPROTEGER
    ↓
Se modifica el spreadsheet
    ↓
Se reprotege con 3 editores
    └─ desarrollo.it@bacarsa.com.ar
    └─ implementaciones.it@bacarsa.com.ar
    └─ El usuario actual
```

## Recomendaciones Finales 💡

1. **Siempre usa el Owner Bridge** para operaciones críticas de protección
2. **Revisa los logs** regularmente para detectar permisos faltantes
3. **Mantén actualizado** el email de `desarrollo.it@bacarsa.com.ar` en todas las protecciones
4. **Documenta** qué usuario va a ejecutar cada acción
5. **Prueba** con diferentes usuarios para asegurar consistencia

---

**Última actualización:** 26 de febrero de 2026
