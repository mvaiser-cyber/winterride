# Route map generator

Builds the inline SVG route map used in the Travel schedule section of `index.html`.

```bash
cd tools/map
npm install us-atlas@3.0.1 topojson-client@3.1.0 d3-geo@3.1.1
node gen.js map.svg
```

Edit the `stops` list in `gen.js` (e.g. the cabin's coordinates once it's booked), regenerate,
and replace the `<svg>…</svg>` inside `<figure class="mapfig">` in `index.html`.
