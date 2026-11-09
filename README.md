# Convertilo · servidor de documentos

Convierte Word, PowerPoint y Excel con LibreOffice, PDF a Word con pdf2docx y PDF a texto con PyMuPDF.
Lo usa la app web **Convertilo** (https://convertilo.pages.dev). Los archivos se borran apenas termina cada conversión.

## Rutas

- `GET /health` → `{"ok": true}`
- `POST /convert` (formulario con `file` y `to`) → devuelve el archivo convertido

## Variables

| Variable | Para qué | Por defecto |
|---|---|---|
| `ALLOWED_ORIGINS` | Dominios que pueden usarlo, separados por coma | `*` |
| `MAX_MB` | Tamaño máximo del documento | `50` |
| `TIMEOUT_SECONDS` | Tiempo máximo por conversión | `180` |
| `MAX_PARALLEL` | Conversiones al mismo tiempo | `2` |

## Publicar

`render.yaml` deja todo listo para Render.com (plan gratis, Docker).
