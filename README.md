# ZZLabs Next.js v5

## Instalar

```bash
npm install
npm run dev
```

Abre: http://localhost:3000

## Cambios
- botón flotante de WhatsApp: +54 2923 500173
- frase de marca más profesional en el hero
- eliminada la galería visual automática/repetitiva de GitHub
- nueva galería curada desde `public/projects/`
- placeholders incluidos hasta que agregues tus capturas reales
- responsive desktop / tablet / celular
- menú móvil de la v4 conservado

## Agregar fotos reales
Pon tus imágenes en:

`public/projects/`

Luego cambia `image:` dentro del array `featuredProjects` en:

`src/components/home.tsx`
