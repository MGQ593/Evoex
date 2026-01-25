# Iconos del Add-in

Este directorio debe contener los iconos del add-in en los siguientes tamaños:

## Archivos requeridos

| Archivo | Tamaño | Uso |
|---------|--------|-----|
| `icon-16.png` | 16x16 px | Ribbon (pequeño) |
| `icon-32.png` | 32x32 px | Ribbon (mediano) |
| `icon-80.png` | 80x80 px | Ribbon (grande) |
| `icon-128.png` | 128x128 px | Store / Admin Center |

## Recomendaciones

- Usa formato PNG con fondo transparente
- Mantén el diseño simple y reconocible a tamaños pequeños
- Usa colores que contrasten bien con el ribbon de Office
- Considera usar un icono que represente IA/chat (robot, burbuja de chat, etc.)

## Herramientas sugeridas

- [Figma](https://figma.com) - Diseño gratuito
- [Canva](https://canva.com) - Fácil de usar
- [Flaticon](https://flaticon.com) - Iconos gratuitos

## Ejemplo de generación rápida

Si tienes un icono SVG base, puedes convertirlo a los tamaños requeridos usando ImageMagick:

```bash
convert icon.svg -resize 16x16 icon-16.png
convert icon.svg -resize 32x32 icon-32.png
convert icon.svg -resize 80x80 icon-80.png
convert icon.svg -resize 128x128 icon-128.png
```
