/* ==========================================================
   Sur Consultores — Agenda Petrinovic
   ========================================================== */

// La misma aplicación de Apps Script que recibe las solicitudes de reunión
// (ver script.js). Entrega los cupos ya tomados y anota las reservas en la
// pestaña "Reservas Petrinovic" de la planilla.
const FORM_ENDPOINT =
  "https://script.google.com/macros/s/AKfycbwfvXfzC4RNslFam_Nv22-R51k2v3U95wz_OinqW7C1gs3BpasNLWV-dzSMzFly0qnPBg/exec";

// Meses hacia adelante que se pueden reservar, contando el actual.
const MESES_VISIBLES = 4;

const origenParam = (new URLSearchParams(location.search).get("origen") || "").trim().toLowerCase();
const ORIGEN = origenParam ? origenParam.slice(0, 40) : "Página web";

document.getElementById("year").textContent = new Date().getFullYear();

/* ---------- Header y menú móvil ---------- */
const header = document.querySelector(".site-header");
window.addEventListener("scroll", () => header.classList.toggle("scrolled", window.scrollY > 8), { passive: true });

const toggle = document.querySelector(".nav-toggle");
const mobileNav = document.getElementById("mobile-nav");
function setMenu(open) {
  toggle.setAttribute("aria-expanded", String(open));
  toggle.setAttribute("aria-label", open ? "Cerrar menú" : "Abrir menú");
  mobileNav.hidden = !open;
}
toggle.addEventListener("click", () => setMenu(mobileNav.hidden));
window.addEventListener("resize", () => { if (window.innerWidth > 980) setMenu(false); });

/* ---------- Fechas ---------- */
function isoLocal(d) {
  return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
}
function desdeIso(iso) { return new Date(iso + "T12:00:00"); }
function esFinDeSemana(d) { return d.getDay() === 0 || d.getDay() === 6; }
function fechaLarga(iso) {
  return desdeIso(iso).toLocaleDateString("es-CL", { weekday: "long", day: "numeric", month: "long" });
}

// Se reserva desde el próximo día hábil hasta el último día del último mes visible.
const hoy = new Date();
hoy.setHours(12, 0, 0, 0);
const primerDia = new Date(hoy);
do primerDia.setDate(primerDia.getDate() + 1); while (esFinDeSemana(primerDia));
const ultimoDia = new Date(hoy.getFullYear(), hoy.getMonth() + MESES_VISIBLES, 0, 12);

/* ---------- Disponibilidad ---------- */
let capacidad = 50;
let reservados = {};
let agendaActiva = false;

function libres(iso) { return Math.max(0, capacidad - (reservados[iso] || 0)); }

async function cargarDisponibilidad() {
  try {
    const res = await fetch(`${FORM_ENDPOINT}?accion=disponibilidad`);
    const datos = await res.json();
    // Una versión anterior del Apps Script no sabe de reservas: sin este dato
    // la agenda no acepta envíos, para no mezclarlos con las solicitudes de reunión.
    if (datos.ok && datos.reservados) {
      capacidad = datos.capacidad || capacidad;
      reservados = datos.reservados;
      agendaActiva = true;
    }
  } catch (err) {
    console.error(err);
  }
  pintarCalendario();
}

/* ---------- Calendario ---------- */
const form = document.getElementById("agenda-form");
const fechaInput = document.getElementById("fecha");
const cuposInput = document.getElementById("cupos");
const cuposHint = document.getElementById("cupos-hint");
const resumen = document.getElementById("agenda-summary");
const diasEl = form.querySelector(".calendar-days");
const tituloEl = form.querySelector(".calendar-title");
const navPrev = form.querySelector('[data-mes="-1"]');
const navNext = form.querySelector('[data-mes="1"]');

let mesVisible = new Date(primerDia.getFullYear(), primerDia.getMonth(), 1, 12);

function pintarCalendario() {
  const anio = mesVisible.getFullYear();
  const mes = mesVisible.getMonth();
  const titulo = mesVisible.toLocaleDateString("es-CL", { month: "long", year: "numeric" });
  tituloEl.textContent = titulo.charAt(0).toUpperCase() + titulo.slice(1);
  navPrev.disabled = anio === primerDia.getFullYear() && mes === primerDia.getMonth();
  navNext.disabled = anio === ultimoDia.getFullYear() && mes === ultimoDia.getMonth();

  diasEl.textContent = "";
  const desfase = (new Date(anio, mes, 1).getDay() + 6) % 7; // semana parte el lunes
  for (let i = 0; i < desfase; i++) {
    diasEl.appendChild(Object.assign(document.createElement("span"), { className: "day-empty" }));
  }

  const totalDias = new Date(anio, mes + 1, 0).getDate();
  for (let n = 1; n <= totalDias; n++) {
    const fecha = new Date(anio, mes, n, 12);
    const iso = isoLocal(fecha);
    const quedan = libres(iso);
    const habil = !esFinDeSemana(fecha) && fecha >= primerDia && fecha <= ultimoDia;

    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "day";
    btn.dataset.fecha = iso;
    btn.setAttribute("role", "radio");
    btn.setAttribute("aria-checked", String(fechaInput.value === iso));
    btn.textContent = n;

    if (habil) {
      const nota = document.createElement("small");
      nota.textContent = quedan === 0 ? "Lleno" : quedan;
      if (quedan > 0) nota.appendChild(Object.assign(document.createElement("span"), { className: "cupos-palabra", textContent: " cupos" }));
      btn.appendChild(nota);
      btn.classList.toggle("lleno", quedan === 0);
      btn.classList.toggle("pocos", quedan > 0 && quedan <= 10);
      btn.setAttribute("aria-label", `${fechaLarga(iso)}, ${quedan === 0 ? "sin cupos" : quedan + " cupos libres"}`);
    }
    btn.disabled = !habil || quedan === 0;
    diasEl.appendChild(btn);
  }
}

navPrev.addEventListener("click", () => { mesVisible.setMonth(mesVisible.getMonth() - 1); pintarCalendario(); });
navNext.addEventListener("click", () => { mesVisible.setMonth(mesVisible.getMonth() + 1); pintarCalendario(); });

diasEl.addEventListener("click", (e) => {
  const btn = e.target.closest(".day");
  if (!btn || btn.disabled) return;
  fechaInput.value = btn.dataset.fecha;
  setError(fechaInput, "");
  pintarCalendario();
  actualizarCupos();
});

function actualizarCupos() {
  if (!fechaInput.value) return;
  const quedan = libres(fechaInput.value);
  cuposInput.max = quedan;
  cuposHint.textContent = `Quedan ${quedan} de ${capacidad} cupos el ${fechaLarga(fechaInput.value)}.`;
  if (cuposInput.value) validateInput(cuposInput);
  actualizarResumen();
}

function actualizarResumen() {
  const servicio = form.querySelector('input[name="servicio"]:checked');
  const cupos = Number(cuposInput.value);
  const listo = servicio && fechaInput.value && cupos >= 1 && cupos <= libres(fechaInput.value);
  resumen.hidden = !listo;
  if (listo) {
    resumen.textContent = `Reservarás ${cupos} ${cupos === 1 ? "cupo" : "cupos"} de ${servicio.value.toLowerCase()} para el ${fechaLarga(fechaInput.value)}.`;
  }
}

/* ---------- RUT ---------- */
function limpiarRut(valor) { return valor.replace(/[^0-9kK]/g, "").toUpperCase(); }

function rutValido(valor) {
  const rut = limpiarRut(valor);
  if (rut.length < 8) return false;
  const cuerpo = rut.slice(0, -1);
  const dv = rut.slice(-1);
  let suma = 0;
  let factor = 2;
  for (let i = cuerpo.length - 1; i >= 0; i--) {
    suma += Number(cuerpo[i]) * factor;
    factor = factor === 7 ? 2 : factor + 1;
  }
  const esperado = 11 - (suma % 11);
  return dv === (esperado === 11 ? "0" : esperado === 10 ? "K" : String(esperado));
}

function formatearRut(valor) {
  const rut = limpiarRut(valor);
  if (rut.length < 2) return valor;
  return rut.slice(0, -1).replace(/\B(?=(\d{3})+(?!\d))/g, ".") + "-" + rut.slice(-1);
}

/* ---------- Validación ---------- */
const messages = {
  servicio: "Elige un servicio.",
  fecha: "Elige un día en el calendario.",
  cupos: "Indica cuántos cupos necesitas.",
  empresa: "Ingresa el nombre de la empresa.",
  rut: "Ingresa un RUT válido, por ejemplo 76.123.456-7.",
  nombre: "Ingresa el nombre de contacto.",
  telefono: "Ingresa un teléfono de contacto.",
  email: "Ingresa un correo válido.",
};

function fieldOf(el) { return el.closest(".field"); }

function setError(el, msg) {
  const field = fieldOf(el);
  field.classList.toggle("invalid", Boolean(msg));
  field.querySelector(".error").textContent = msg || "";
}

function validateInput(el) {
  let valid = el.name === "fecha" ? Boolean(el.value) : el.checkValidity();
  let msg = messages[el.name];

  if (el.name === "telefono" && valid) valid = el.value.replace(/\D/g, "").length >= 8;
  if (el.name === "rut" && valid) valid = rutValido(el.value);
  if (el.name === "cupos" && el.value && fechaInput.value) {
    const quedan = libres(fechaInput.value);
    const n = Number(el.value);
    valid = Number.isInteger(n) && n >= 1 && n <= quedan;
    if (!valid) msg = `Puedes reservar entre 1 y ${quedan} cupos para ese día.`;
  }

  setError(el, valid ? "" : msg);
  return valid;
}

function validateForm() {
  let firstInvalid = null;
  const checked = new Set();

  form.querySelectorAll("[required]").forEach((el) => {
    if (el.type === "radio") {
      if (checked.has(el.name)) return;
      checked.add(el.name);
      const ok = Boolean(form.querySelector(`input[name="${el.name}"]:checked`));
      setError(el, ok ? "" : messages[el.name]);
      if (!ok && !firstInvalid) firstInvalid = el;
      return;
    }
    if (!validateInput(el) && !firstInvalid) firstInvalid = el;
  });

  if (firstInvalid) {
    const destino = firstInvalid.type === "hidden" ? diasEl.querySelector(".day:not(:disabled)") : firstInvalid;
    if (destino) destino.focus({ preventScroll: true });
    fieldOf(firstInvalid).scrollIntoView({ behavior: "smooth", block: "center" });
  }
  return !firstInvalid;
}

form.addEventListener("input", (e) => {
  const el = e.target;
  if (el.name === "servicio") setError(el, "");
  else if (messages[el.name] && fieldOf(el).classList.contains("invalid")) validateInput(el);
  actualizarResumen();
});
form.addEventListener("change", (e) => {
  if (e.target.name === "cupos") validateInput(e.target);
});
document.getElementById("rut").addEventListener("blur", (e) => {
  if (e.target.value.trim()) {
    e.target.value = formatearRut(e.target.value);
    validateInput(e.target);
  }
});

/* ---------- Envío ---------- */
form.addEventListener("submit", async (e) => {
  e.preventDefault();
  if (!validateForm()) return;

  if (!agendaActiva) {
    alert("La agenda todavía no está recibiendo reservas. Escríbenos a fllancafil@surconsultores.org o por WhatsApp al +56 9 4818 4418.");
    return;
  }

  const button = form.querySelector('button[type="submit"]');
  const data = Object.fromEntries(new FormData(form));
  data.tipo = "reserva_petrinovic";
  data.cupos = Number(data.cupos);
  data.origen = ORIGEN;
  button.disabled = true;
  button.textContent = "Enviando…";

  try {
    // Texto plano para evitar la consulta previa (preflight), igual que en script.js.
    const res = await fetch(FORM_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(data),
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const respuesta = await res.json();

    if (respuesta.motivo === "sin_cupos") {
      // Otra empresa reservó mientras se llenaba el formulario.
      reservados[data.fecha] = capacidad - (respuesta.disponibles || 0);
      pintarCalendario();
      actualizarCupos();
      validateInput(cuposInput);
      cuposInput.focus();
      button.disabled = false;
      button.textContent = "Reservar cupos";
      return;
    }
    if (!respuesta.ok || !respuesta.reserva) throw new Error(respuesta.error || "Respuesta no válida");

    form.querySelector("[data-name]").textContent = data.nombre.trim().split(" ")[0];
    form.querySelector("[data-detalle]").textContent =
      `${data.cupos} ${data.cupos === 1 ? "cupo" : "cupos"} de ${data.servicio.toLowerCase()} para el ${fechaLarga(data.fecha)}`;
    form.querySelector(".form-body").hidden = true;
    const success = form.querySelector(".form-success");
    success.hidden = false;
    success.focus();
  } catch (err) {
    console.error(err);
    button.disabled = false;
    button.textContent = "Reservar cupos";
    alert("No pudimos registrar tu reserva. Inténtalo nuevamente o escríbenos a fllancafil@surconsultores.org.");
  }
});

pintarCalendario();
cargarDisponibilidad();
