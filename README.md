# WSUS_Auditor

## Descripcion General

`WSUS_Auditor.ps1` es una herramienta en PowerShell con interfaz grafica (WinForms) para auditar, diagnosticar y aplicar acciones correctivas sobre equipos gestionados por WSUS y Active Directory.

La version actual mantiene la estructura y el flujo originales, pero deja de depender de un entorno fijo. Al iniciar, muestra una pantalla de configuracion para indicar:

- Servidor WSUS
- Puerto WSUS
- Uso de SSL
- Dominio DNS/FQDN
- OU base o DN base de Active Directory
- Filtro opcional de OUs de primer nivel
- Perfiles guardados en JSON local
- Modo simulacion local con datos mock

Con esto la herramienta puede reutilizarse en distintos dominios y jerarquias de AD sin quedar acoplada a una organizacion concreta.

---

## Caracteristicas Principales

- Configuracion inicial del entorno AD y WSUS al abrir la herramienta.
- Guardado y carga de perfiles en `wsus_auditor.profiles.json`.
- Integracion con Active Directory para mostrar la jerarquia de OUs y equipos.
- Carga perezosa del arbol para mejorar rendimiento en entornos grandes.
- Cache de equipos WSUS para acelerar cruces entre AD y WSUS.
- Listado de equipos con informacion combinada de AD y WSUS.
- Filtro rapido de equipos con errores o incidencias.
- Exportacion de resultados a CSV.
- Acciones correctivas remotas via WMI.
- Consola de logs en tiempo real.
- Validacion offline mediante datos simulados en `wsus_auditor.mock.json`.

---

## Requisitos

- PowerShell 5.1 o superior.
- Modulo `ActiveDirectory` disponible para modo real.
- Acceso a un servidor WSUS para modo real.
- Permisos de lectura sobre AD y WSUS para modo real.
- Permisos administrativos en los equipos remotos para aplicar soluciones.
- Acceso de red a `ADMIN$` y WMI en los equipos cliente.

---

## Flujo de Uso

1. Abrir PowerShell como administrador.
2. Ejecutar:

```powershell
.\WSUS_Auditor.ps1
```

3. Completar la configuracion inicial del entorno o cargar un perfil.
4. Navegar por la jerarquia AD cargada desde la OU base indicada.
5. Seleccionar equipos o ramas completas del arbol.
6. Listar, filtrar, exportar o aplicar soluciones segun sea necesario.

---

## Configuracion Inicial

La ventana inicial permite adaptar la herramienta a otros entornos:

- `Perfil guardado`: permite cargar una configuracion ya almacenada.
- `Nombre del perfil`: nombre con el que se guardara o actualizara el entorno actual.
- `Servidor WSUS`: nombre DNS o NetBIOS del WSUS.
- `Puerto WSUS`: por defecto 8530 o 8531 segun SSL.
- `Usar SSL`: activa la conexion segura al WSUS.
- `Dominio DNS/FQDN`: se usa para completar FQDNs cuando AD no expone `DNSHostName`.
- `OU base / DN base`: punto de partida para cargar el arbol de Active Directory.
- `Filtro de OUs de primer nivel`: opcional. Si se deja vacio, se muestran todas las OUs hijas directas de la base.
- `Modo simulacion local`: evita dependencias a AD/WSUS reales y carga datos desde un JSON local.
- `Fichero de datos mock`: ruta al JSON de simulacion.

---

## Funcionalidades

### 1. Arbol AD

- Carga dinamica de nodos.
- Soporte para seleccionar ramas completas.
- Nodo especial `WSUS (sin AD)` para detectar equipos presentes en WSUS pero fuera del ambito AD configurado.

### 2. Listado de equipos

- Cruce entre AD y WSUS.
- Datos mostrados: nombre, IP, version, marca, modelo, errores, porcentaje de instalacion, ultimo reporte y ultimo contacto.
- Coloreado por estado para priorizar incidencias.

### 3. Filtro y busqueda

- Filtro rapido para mostrar solo equipos problematicos.
- Busqueda por nombre, IP, marca, modelo o error.

### 4. Exportacion

- Exportacion global a CSV.
- Exportacion individual desde el menu contextual de la tabla.

### 5. Soluciones remotas

- Reinicio de servicios Windows Update y BITS.
- Limpieza de cache de Windows Update.
- Forzar deteccion y reporte.
- Re-registro de DLLs.
- Reset completo de Windows Update.

Las acciones se generan en un script temporal, se copian al equipo remoto y se ejecutan por WMI (`Win32_Process`), recuperando despues el log remoto.

---

## Notas Tecnicas

- La aplicacion sigue siendo un script WinForms monolitico, pero ahora parametriza el entorno al inicio.
- Se usan `DNSHostName` y un sufijo DNS configurable para evitar dependencias a dominios hardcodeados.
- Los perfiles se guardan en JSON local junto al script, salvo que se indique otra ruta con `-ProfilesPath`.
- El modo mock reutiliza el mismo flujo funcional mediante datos simulados y un perfil offline.
- Si WSUS no conecta, la herramienta puede seguir mostrando la parte de AD, aunque sin cruces ni estado WSUS.
- El filtro de OUs de primer nivel limita tanto el arbol visible como el calculo de equipos "WSUS sin AD".

---

## Validacion Offline

Puede ejecutarse una comprobacion automatizada sin AD ni WSUS con:

```powershell
.\WSUS_Auditor.ps1 -SelfTest -ProfileName "Offline Demo"
```

El `SelfTest` valida:

- carga del perfil JSON
- carga de datos mock
- construccion del arbol
- seleccion de equipos
- listado y cruce AD/WSUS
- filtrado de incidencias
- exportacion CSV
- lectura y escritura del JSON de perfiles

---

## Archivos Relacionados

- `WSUS_Auditor.ps1`: aplicacion principal.
- `wsus_auditor.profiles.json`: perfiles de entorno.
- `wsus_auditor.mock.json`: datos de simulacion para pruebas offline.
- `wsus_auditor.selftest.csv`: CSV generado por la ultima ejecucion de `SelfTest`.

---

## Siguientes Mejoras Recomendadas

- Guardar las credenciales de forma segura o integrarse con credenciales del sistema.
- Separar UI, acceso a AD/WSUS y remediacion en modulos.
- Añadir validaciones de conectividad antes de abrir la ventana principal.
- Empaquetar la herramienta como ejecutable para despliegue estandar.
