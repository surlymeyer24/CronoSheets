# CronoSheets - Owner Bridge (cómo aplicar los cambios)

Este proyecto usa 2 Apps Script:

- **Cliente** (este repo raíz): corre dentro del Google Sheet donde los usuarios trabajan.
- **Owner** (`/owner`): se publica como Web App ejecutando como propietario para operaciones con protección.

## 1) Deploy del proyecto owner

Desde `owner/`:

1. Publicá/actualizá el Web App con:
   - **Execute as**: `Me`
   - **Who has access**: según política (si usás token, puede ser abierto a cualquiera)
2. Copiá la URL de Web App.
3. En el proyecto owner, ejecutá `generateAndSetOwnerToken()` una vez.
4. Copiá el token desde logs.

## 2) Configurar el cliente

En el Apps Script del cliente, ejecutá:

```javascript
setOwnerBridgeConfig('URL_DEL_WEBAPP_OWNER', 'TOKEN_OWNER');
```

Podés verificar configuración con:

```javascript
verOwnerBridgeConfig();
```

## 3) Validar conexión

Ejecutá:

```javascript
testOwnerBridgeConexion();
```

Debe devolver `success: true` y estado `ok`.

## 4) Probar flujo real con usuario editor

Con un usuario **no propietario**:

1. Agregar guardia (debe desproteger/reproteger Resumen sin error de permisos).
2. Eliminar guardia (debe borrar hoja aunque tenga protección).
3. Botón "🔧 Reparar Protección Resumen".

Si el owner bridge falla, el script deja log y hace fallback local cuando puede.

## 5) Si no funciona

Checklist rápido:

- URL de Web App correcta y corresponde al deploy activo.
- `OWNER_TOKEN` del cliente coincide con el owner.
- El owner tiene acceso al Spreadsheet destino.
- Re-deploy del Web App luego de cambios en `owner/Código.js`.
- Revisar `Executions` en ambos proyectos (cliente/owner).
