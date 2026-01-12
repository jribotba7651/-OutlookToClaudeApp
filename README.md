# Outlook to Claude Calendar App

Una aplicación WPF moderna que te permite exportar eventos de tu calendario de Outlook directamente a Claude AI usando la API de Claude.

## ✨ Características

- 📅 **Lee eventos de Outlook**: Selecciona un rango de fechas y carga automáticamente tus eventos
- ✅ **Selección granular**: Escoge exactamente qué eventos quieres compartir
- 🌐 **Upload directo a Claude API**: Los eventos se suben a la nube de Claude (no archivos locales)
- 📝 **Preview de Markdown**: Ve cómo se verán tus eventos antes de exportar
- 🔑 **API Key validation**: Prueba tu Claude API key antes de exportar
- 🎨 **UI moderna**: Interfaz limpia y fácil de usar

## 🚀 Cómo usar

### 1. Obtén tu Claude API Key

1. Ve a [https://console.anthropic.com/](https://console.anthropic.com/)
2. Inicia sesión con tu cuenta
3. Ve a **Settings** > **API Keys**
4. Crea una nueva API key y cópiala

### 2. Ejecuta la aplicación

1. Abre `OutlookToClaudeApp.exe`
2. Selecciona el rango de fechas de tus eventos
3. Pega tu Claude API Key
4. Haz clic en **Load Events** para cargar eventos de Outlook

### 3. Selecciona y exporta

1. Selecciona los eventos que quieres compartir con Claude
2. (Opcional) Haz clic en **Preview Markdown** para ver el formato
3. Haz clic en **Export to Claude**
4. ¡Listo! El **File ID** se copia automáticamente al portapapeles

### 4. Usa en Claude

En tu conversación con Claude, simplemente escribe:

```
Review my calendar events @file_abc123xyz
```

(Reemplaza `file_abc123xyz` con el File ID que copiaste)

## 📋 Requisitos

- Windows 10/11
- .NET 8.0 Runtime
- Microsoft Outlook instalado
- Claude API Key (requiere cuenta de Claude)

## 🏗️ Arquitectura

```
OutlookToClaudeApp/
├── Models/
│   ├── CalendarEvent.cs       # Modelo de evento de calendario
│   ├── ApiConfig.cs            # Configuración de API keys
│   └── ExportResult.cs         # Resultado de exportación
│
├── Services/
│   ├── OutlookService.cs       # Integración con Outlook
│   └── ClaudeApiService.cs     # Integración con Claude API
│
└── MainWindow.xaml/cs          # UI principal
```

## 🔧 Desarrollo

### Compilar desde código fuente

```bash
dotnet build
```

### Ejecutar en modo desarrollo

```bash
dotnet run
```

## 📝 Formato de exportación

Los eventos se exportan en Markdown con este formato:

```markdown
# Calendar Events

**Export Date:** 2025-01-18 10:30
**Total Events:** 5

---

## Monday, January 20, 2025

### Team Standup

**Time:** 10:00 AM - 10:30 AM
**Location:** Zoom
**Organizer:** manager@company.com

**Details:**
Weekly standup meeting...

---
```

## 🔐 Seguridad

- Las API keys NO se guardan en disco
- Solo se almacenan en memoria durante la sesión
- Los archivos se suben directamente a Claude vía HTTPS
- No se guardan copias locales de los eventos

## 🐛 Troubleshooting

### "Failed to connect to Outlook"
- Asegúrate que Outlook está instalado y configurado
- Abre Outlook al menos una vez antes de usar la app

### "Invalid API Key"
- Verifica que copiaste la API key completa
- Asegura que la API key no ha expirado

### "No events found"
- Verifica el rango de fechas
- Asegúrate de tener eventos en tu calendario de Outlook

## 📄 Licencia

Proyecto personal - Uso libre

## 🙏 Agradecimientos

- **Anthropic** - Claude API
- **NetOffice** - Outlook COM Interop
- **Microsoft** - WPF Framework
