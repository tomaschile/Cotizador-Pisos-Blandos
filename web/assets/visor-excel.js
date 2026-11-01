// Visor de Excel en una pestaña aparte, para vendedores sin Excel instalado.
// Adaptado del mismo visor del cotizador de enchufes (web/assets/visor-excel.js
// en Cotizador-Sinthesi_MK) a la planilla de pinturas: una sola hoja, con la
// tabla "Detalle de compra" cuyas columnas se llaman "SKU" y "Cantidad"
// (en el cotizador de enchufes son "Código MK"/"Cantidad" y solo se buscan en
// la hoja "Resumen" porque ahí hay varias hojas con tablas repetidas; acá no
// hace falta esa restricción porque solo hay una hoja).
// Usa SheetJS (cargado antes que este script) para leer el .xlsx ya
// descargado y volcarlo a una tabla HTML de solo lectura en una ventana
// nueva — no depende de ningún visor externo ni sube el archivo a ningún
// servidor de terceros.
//
// La ventana se abre en dos pasos (abrirVentanaVisorPendiente + luego
// completarVisorExcel) porque el Excel no está listo hasta después de un
// fetch — si se llama a window.open() recién ahí, el navegador lo bloquea
// por no ser ya parte del gesto de click original.
function abrirVentanaVisorPendiente() {
  try {
    var ventana = window.open('', '_blank');
    if (ventana) {
      ventana.document.write(
        '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Generando…</title></head>' +
        '<body style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,sans-serif;' +
        'color:#5b6570;padding:40px">Generando el archivo…</body></html>'
      );
    }
    return ventana;
  } catch (e) {
    return null;
  }
}

var _ETIQUETAS_PIE_TABLA = ['Subtotal', 'TOTAL', 'Total Obra'];
function extraerColumnaPorEncabezado(ws, etiqueta) {
  var filas = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  var colActiva = null;
  var valores = [];
  var terminado = false;
  filas.forEach(function (fila) {
    if (terminado) return;
    var vacia = fila.every(function (v) { return v === '' || v == null; });
    var esPie = !vacia && fila.some(function (v) {
      return typeof v === 'string' && (_ETIQUETAS_PIE_TABLA.indexOf(v) !== -1 || /^IVA\s*\(/.test(v));
    });
    if (vacia || esPie) {
      if (colActiva != null) terminado = true;
      colActiva = null;
      return;
    }
    var idx = fila.indexOf(etiqueta);
    if (idx !== -1) { colActiva = idx; return; }
    if (colActiva != null && fila[colActiva] !== '' && fila[colActiva] != null) {
      valores.push(fila[colActiva]);
    }
  });
  return valores;
}

function completarVisorExcel(ventana, arrayBuffer, nombre) {
  if (!ventana || ventana.closed || typeof XLSX === 'undefined') return;
  try {
    var libro = XLSX.read(arrayBuffer, { type: 'array' });
    var botonesData = [];
    var cuerpo = libro.SheetNames.map(function (hoja, idx) {
      var ws = libro.Sheets[hoja];
      var tabla = XLSX.utils.sheet_to_html(ws, { header: '', footer: '' });

      var idTabla = 'copiar-tabla-' + idx;
      botonesData.push({ id: idTabla, texto: XLSX.utils.sheet_to_csv(ws, { FS: '\t' }) });
      var botonesHtml = '<button type="button" class="copiar" id="' + idTabla + '">Copiar tabla</button>';

      var skus = extraerColumnaPorEncabezado(ws, 'SKU');
      if (skus.length) {
        var idSku = 'copiar-sku-' + idx;
        botonesData.push({ id: idSku, texto: skus.join('\n') });
        botonesHtml += '<button type="button" class="copiar" id="' + idSku + '">Copiar SKU</button>';
      }
      var cantidades = extraerColumnaPorEncabezado(ws, 'Cantidad');
      if (cantidades.length) {
        var idCant = 'copiar-cantidades-' + idx;
        botonesData.push({ id: idCant, texto: cantidades.join('\n') });
        botonesHtml += '<button type="button" class="copiar" id="' + idCant + '">Copiar cantidades</button>';
      }

      return '<div class="hoja-top"><h2>' + hoja + '</h2>' + botonesHtml + '</div>' + tabla;
    }).join('');
    ventana.document.open();
    ventana.document.write(
      '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>' + nombre + '</title>' +
      '<style>' +
      'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;' +
      'margin:24px;color:#1a1d21;background:#fff}' +
      '.aviso{color:#5b6570;font-size:.85rem;margin-bottom:20px}' +
      '.hoja-top{display:flex;flex-wrap:wrap;align-items:center;gap:12px;margin:32px 0 10px}' +
      'h2{font-size:1.05rem;margin:0;color:#1f3864}' +
      '.copiar{font:inherit;font-size:.8rem;cursor:pointer;border:1px solid #1f6feb;' +
      'background:#fff;color:#1f6feb;border-radius:6px;padding:4px 10px}' +
      '.copiar:hover{background:#eaf2fe}' +
      'table{border-collapse:collapse;font-size:.85rem;margin-bottom:8px}' +
      'td{border:1px solid #c9d2db;padding:5px 10px;white-space:nowrap}' +
      'tr:nth-child(even) td{background:#f7f9fb}' +
      'td[data-t="n"]{text-align:right;font-variant-numeric:tabular-nums}' +
      'table tr:first-child td[colspan]{background:#1f2a44;color:#fff;font-weight:700;' +
      'font-size:.95rem;text-align:center;white-space:normal}' +
      '</style></head><body>' +
      '<p class="aviso">Vista de solo lectura de <strong>' + nombre + '</strong> — el archivo ya se ' +
      'descargó a tu equipo si necesitas editarlo en Excel o Google Sheets. Arrastrar el mouse no ' +
      'selecciona columnas sueltas (limitación del navegador) — usa los botones "Copiar SKU" / ' +
      '"Copiar cantidades" para eso, o "Copiar tabla" para pegar todo en una planilla real.</p>' +
      cuerpo +
      '</body></html>'
    );
    ventana.document.close();

    botonesData.forEach(function (b) {
      var btn = ventana.document.getElementById(b.id);
      if (!btn) return;
      var original = btn.textContent;
      btn.addEventListener('click', function () {
        ventana.navigator.clipboard.writeText(b.texto).then(function () {
          btn.textContent = '¡Copiado!';
          setTimeout(function () { btn.textContent = original; }, 1500);
        }).catch(function () {
          btn.textContent = 'No se pudo copiar';
          setTimeout(function () { btn.textContent = original; }, 1500);
        });
      });
    });
  } catch (e) {
    console.error('No se pudo mostrar el visor de Excel:', e);
    try { ventana.close(); } catch (e2) { /* noop */ }
  }
}

function descargarExcelConVisor(url, nombre, ventana) {
  return fetch(url, { credentials: 'same-origin' })
    .then(function (r) {
      if (!r.ok) throw new Error('No se pudo descargar ' + nombre);
      return r.arrayBuffer();
    })
    .then(function (buf) {
      var blobUrl = URL.createObjectURL(new Blob([buf]));
      var a = document.createElement('a');
      a.href = blobUrl; a.download = nombre;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(blobUrl); }, 10000);
      completarVisorExcel(ventana, buf, nombre);
    })
    .catch(function (e) {
      if (ventana && !ventana.closed) ventana.close();
      console.error('No se pudo abrir el visor de Excel:', e);
    });
}
