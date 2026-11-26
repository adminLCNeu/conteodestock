// Ubicaciones disponibles para el conteo.
// Cada una puede tener una foto (archivo de imagen subido junto a la app,
// en la raíz del repo) para identificarla más fácil al elegir sucursal.
// Si "photo" es null, se muestra un ícono genérico en su lugar.
// hiddenProducts (opcional): códigos de productos que no se muestran para
// contar en esa sucursal. Para que un producto aparezca SOLO en algunas
// sucursales, en products.js se le agrega locations: ['Nombre sucursal'].
const LOCATIONS = [
  { name: 'BCN 1 - Space', photo: 'loc-bcn1.jpg' },
  {
    name: 'BCN 2 - Moon', photo: 'loc-bcn2.jpg',
    // Productos (por código) que NO se cuentan en esta sucursal.
    hiddenProducts: [
      // Cannolis
      2386, 2384, 2388, 2385, 2387,
      // Vaschettas grande / pequeña / descartable
      90002, 90003, 90004,
      // Bebidas: Sprite lata, S. Pellegrino 500 vidrio
      90005, 2130,
      // Caja pastelería, vasos plástico pastelería
      2291, 90016,
      // Limpieza: producto WC, Ecogras, V75 ambientador, valletas verdes,
      // trapos para cristales, Speed Natur, For Chlor
      90028, 90029, 90031, 90038, 90040, 90045, 90046,
      // Uniformes azules: camisas por género (se reemplazan por unisex)
      90066, 90067, 90068, 90069, 90070, 90071, 90072, 90073, 90074, 90075,
      // Ice pops: Smile, Crema de maní s/azúcar, Sorbete de frutos rojos
      2235, 2223, 2236,
    ],
  },
  { name: 'Madrid', photo: 'loc-madrid.jpg' },
  { name: 'Málaga 1', photo: 'loc-malaga1.jpg' },
  {
    name: 'Valencia', photo: 'loc-valencia.jpg',
    // Productos (por código) que NO se cuentan en esta sucursal.
    hiddenProducts: [
      // Cucuruchos anteriores (reemplazados por los conos de la planilla de varios)
      2170, 2171, 2172, 2380, 2173, 2174, 2282,
      // Copas chica/media/grande/XL (reemplazadas por Tarrinas Printedcup S/M/L/XL)
      2192, 2193, 2194, 2195,
    ],
  },
  { name: 'Fabrica BCN', photo: 'loc-fabrica.jpg' },
];

// Usuarios habilitados para entrar a la app.
// canDelete: true = puede borrar conteos/desperdicios ya finalizados del historial.
// Por ahora solo "claudios" tiene ese permiso, tal como pediste.
//
// ⚠️ IMPORTANTE: esta app es un sitio estático (sin servidor), así que este
// login es un control simple para que el personal no borre conteos por
// error o entre gente no autorizada — NO es seguridad real. Cualquiera que
// sepa mirar el código fuente de la página puede ver estas contraseñas.
// Si en algún momento necesitás protección de verdad (por ejemplo, más
// sucursales, más gente, datos sensibles), lo ideal sería sumar un backend
// con autenticación real.
//
// Cambiá estas contraseñas antes de subir la app a producción.
// canEditPrices: true = puede entrar a la pantalla de Precios (ver y editar
// Precio 1 / Precio 2 de cada producto, y descargar el products.js actualizado
// + el historial de cambios). Por ahora solo "batodesrets", tal como pediste.
// locations (opcional): si se define, ese usuario SOLO puede contar en esas
// sucursales (deben coincidir exactamente con un "name" de LOCATIONS). Si no
// se define, el usuario puede contar en todas las sucursales.
const USERS = [
  { username: 'batodesrets', password: '100393', canDelete: true, canEditPrices: true },
  { username: 'bautista', password: '100393', canDelete: false, canEditPrices: false },
  { username: 'agostina', password: '123456', canDelete: false, canEditPrices: false, locations: ['BCN 2 - Moon'] },
  { username: 'manuel', password: '123456', canDelete: false, canEditPrices: false, locations: ['BCN 1 - Space'] },
  { username: 'simon', password: '123456', canDelete: false, canEditPrices: false, locations: ['Madrid'] },
  { username: 'julian', password: '123456', canDelete: false, canEditPrices: false },
  { username: 'malaga', password: '123456', canDelete: false, canEditPrices: false, locations: ['Málaga 1'] },
  { username: 'fabrica', password: '123456', canDelete: false, canEditPrices: false },
  { username: 'sergio S', password: '123456', canDelete: false, canEditPrices: false, locations: ['Valencia'] },
];

// Límites de "cantidad razonable" por producto. Si lo contado supera el límite,
// el renglón se marca en naranja al cargarlo (se puede seguir igual) y, al
// finalizar, la app pregunta por cada valor naranja si es correcto.
// Se busca primero por categoría, después por unidad de medida, y si no hay
// nada se usa "default". "local" = sucursales, "fabrica" = Fabrica BCN.
// Si una sucursal te confirma que un valor alto es real, subí el límite acá.
const QTY_LIMITS = {
  byCategory: {
    'ICE POPS CLASSIC':         { local: 600,    fabrica: 5000 },
    'ICE POPS SIN BAÑO':        { local: 600,    fabrica: 5000 },
    'ICE POPS BAÑADOS':         { local: 600,    fabrica: 5000 },
    'ICE POPS LUXURY':          { local: 600,    fabrica: 5000 },
    'ICE POP DUBAI':            { local: 600,    fabrica: 5000 },
    'FRASCOS DDL':              { local: 1500,   fabrica: 5000 },
    'BEBIDAS':                  { local: 600,    fabrica: 600 },
    'PASTELERIA':               { local: 400,    fabrica: 400 },
    'CUCURUCHOS':               { local: 30000,  fabrica: 30000 },
    'POLIPAPEL':                { local: 60000,  fabrica: 60000 },
    'PACKAGING':                { local: 100000, fabrica: 100000 },
    'SERVILLETAS':              { local: 100000, fabrica: 100000 },
    'CUCHARAS':                 { local: 100000, fabrica: 100000 },
    'CAFÉ':                     { local: 20000,  fabrica: 20000 },
    'MILK SHAKE':               { local: 20000,  fabrica: 20000 },
    'VARIOS':                   { local: 10000,  fabrica: 10000 },
    'CAPSULAS CAFÉ':            { local: 5000,   fabrica: 5000 },
  },
  byUnit: {
    kg:     { local: 150,  fabrica: 500 },
    litro:  { local: 100,  fabrica: 500 },
    caja:   { local: 200,  fabrica: 1000 },
    unidad: { local: 1000, fabrica: 5000 },
  },
  default: { local: 1000, fabrica: 5000 },
};
// Helado: el límite es por cada vaschetta pesada (no por el total del sabor).
// Una vaschetta que pese más de esto se marca en naranja.
const HELADO_PESADA_MAX_KG = 10;
const FABRICA_NAME = 'Fabrica BCN';

if (typeof module !== 'undefined') { module.exports = { LOCATIONS, USERS, QTY_LIMITS, FABRICA_NAME, HELADO_PESADA_MAX_KG }; }
