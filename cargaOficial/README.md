# Carga oficial – desplegar en 5 proyectos

Los configs para los 5 proyectos están **en la raíz de ministerio** (`.clasp-carga1.json` … `.clasp-carga5.json`). Esta carpeta `cargaOficial/` guarda copias de referencia.

Ejecutar **siempre desde la raíz del proyecto ministerio**:

```bash
cd /ruta/a/ministerio

# Un proyecto (añadir --force si dice "Skipping push" o "already up to date")
clasp push --project .clasp-carga1.json --force
clasp push --project .clasp-carga2.json --force
# ... .clasp-carga3.json, .clasp-carga4.json, .clasp-carga5.json
```

Todos de una vez (Bash):

```bash
for i in 1 2 3 4 5; do clasp push --project ".clasp-carga$i.json" --force; done
```

En PowerShell:

```powershell
1..5 | ForEach-Object { clasp push --project ".clasp-carga$_.json" --force }
```

**Importante:** Si no usas `--force`, clasp puede no subir cambios si detecta que el proyecto está “al día”; en ese caso repite el comando con `--force`.
