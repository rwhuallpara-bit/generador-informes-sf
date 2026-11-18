// Base de datos rápida de equipos calibrados (puedes agregar más aquí)
const EQUIPOS_DB = {
  eq1: {
    marca: "DRÄGER",
    modelo: "ALCOTEST 6820",
    serie: "ARKM-0412",
    certificado: "CERT-CAL-2026-089",
    cal_ultima: "15/06/2026",
    cal_proxima: "15/12/2026"
  },
  eq2: {
    marca: "DRÄGER",
    modelo: "ALCOTEST 7510",
    serie: "ARPN-1108",
    certificado: "CERT-CAL-2026-114",
    cal_ultima: "20/08/2026",
    cal_proxima: "20/02/2027"
  }
};

document.addEventListener("DOMContentLoaded", () => {
  const hoy = new Date();
  document.getElementById("fecha_informe").value = hoy.toISOString().split("T")[0];
  document.getElementById("fecha_evento").value = hoy.toISOString().split("T")[0];
  document.getElementById("hora").value = hoy.toTimeString().slice(0, 5);
  cargarDatosEquipo();
});

function cargarDatosEquipo() {
  const sel = document.getElementById("selectorEquipo").value;
  if (EQUIPOS_DB[sel]) {
    const eq = EQUIPOS_DB[sel];
    document.getElementById("marca").value = eq.marca;
    document.getElementById("modelo").value = eq.modelo;
    document.getElementById("serie").value = eq.serie;
    document.getElementById("certificado").value = eq.certificado;
    document.getElementById("cal_ultima").value = eq.cal_ultima;
    document.getElementById("cal_proxima").value = eq.cal_proxima;
  }
}

function cambiarFormulario() {
  const tipo = document.getElementById("tipoInforme").value;
  const esAlcotest = tipo === "alcoholemia";
  
  document.getElementById("bloque-alcoholemia").classList.toggle("hidden", !esAlcotest);
  document.getElementById("bloque-licencia_falsa").classList.toggle("hidden", esAlcotest);
  
  document.getElementById("referencia").value = esAlcotest
    ? "INFORME DE RESULTADO POSITIVO DE ALCOTEST – PLANTA READY MIX PORTILLO"
    : "INF. DE LICENCIA DE CONDUCIR FALSA – PLANTA VIACHA";
}

// Convierte archivo de imagen subido a ArrayBuffer para incrustar en Word
function leerImagen(inputElement) {
  return new Promise((resolve, reject) => {
    const file = inputElement.files[0];
    if (!file) return resolve(null);
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function formatearFechaLatam(fechaISO) {
  if (!fechaISO) return "";
  const [y, m, d] = fechaISO.split("-");
  return `${d}/${m}/${y}`;
}

document.getElementById("informeForm").addEventListener("submit", async function (e) {
  e.preventDefault();
  
  const tipo = document.getElementById("tipoInforme").value;
  const rutaPlantilla = `plantillas/${tipo}.docx`;

  try {
    const response = await fetch(rutaPlantilla);
    if (!response.ok) throw new Error(`No se encontró la plantilla en ${rutaPlantilla}`);
    const content = await response.arrayBuffer();

    const img1Buffer = await leerImagen(document.getElementById("foto_1"));
    const img2Buffer = await leerImagen(document.getElementById("foto_2"));

    // Configuración del módulo de imágenes (ajusta tamaño en píxeles para que encaje en la celda)
    const imageModule = new window.ImageModule({
      centered: true,
      getImage: function (tagValue) { return tagValue; },
      getSize: function () { return [250, 185]; } // Ancho x Alto en px dentro de la tabla
    });

    const zip = new PizZip(content);
    const doc = new window.docxtemplater(zip, {
      modules: [imageModule],
      paragraphLoop: true,
      linebreaks: true,
    });

    const datos = {
      correlativo: document.getElementById("correlativo").value,
      remitente: document.getElementById("remitente").value,
      referencia: document.getElementById("referencia").value.toUpperCase(),
      fecha_informe: formatearFechaLatam(document.getElementById("fecha_informe").value),
      fecha_evento: formatearFechaLatam(document.getElementById("fecha_evento").value),
      hora: document.getElementById("hora").value,
      lugar: document.getElementById("lugar").value,
      puesto: document.getElementById("lugar").value,
      // Campos Alcotest
      nombre: document.getElementById("nombre").value.toUpperCase(),
      ci: document.getElementById("ci").value,
      cargo: document.getElementById("cargo").value,
      empresa: document.getElementById("empresa").value.toUpperCase(),
      resp_soboce: document.getElementById("resp_soboce").value,
      resp_control: document.getElementById("resp_control").value,
      reg_firmado: document.getElementById("reg_firmado").value,
      prueba_pasiva: document.getElementById("prueba_pasiva").value,
      prueba_auto: document.getElementById("prueba_auto").value,
      prueba_manual: document.getElementById("prueba_manual").value,
      aclaraciones: document.getElementById("aclaraciones").value || "Ninguna.",
      marca: document.getElementById("marca").value,
      modelo: document.getElementById("modelo").value,
      serie: document.getElementById("serie").value,
      certificado: document.getElementById("certificado").value,
      cal_ultima: document.getElementById("cal_ultima").value,
      cal_proxima: document.getElementById("cal_proxima").value,
      // Campos Licencia Falsa
      personal_reporta: document.getElementById("personal_reporta").value,
      conductor: document.getElementById("conductor").value.toUpperCase(),
      nro_licencia: document.getElementById("nro_licencia").value,
      categoria: document.getElementById("categoria").value.toUpperCase(),
      sgto_epi: document.getElementById("sgto_epi").value,
      firma_nombre: document.getElementById("firma_nombre").value,
      firma_cargo: document.getElementById("firma_cargo").value.toUpperCase(),
      // Imágenes
      foto_1: img1Buffer,
      foto_2: img2Buffer
    };

    doc.render(datos);

    const out = doc.getZip().generate({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    const sufijo = tipo === "alcoholemia" ? datos.nombre : datos.conductor;
    const nombreArchivo = `SF-${datos.correlativo}_2026_${tipo.toUpperCase()}_${sufijo.replace(/\s+/g, "_")}.docx`;
    saveAs(out, nombreArchivo);

  } catch (error) {
    alert("Error al generar el informe: " + error.message);
    console.error(error);
  }
});
