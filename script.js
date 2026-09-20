/* ==========================================================
   Sur Consultores — Interacciones
   ========================================================== */

// Endpoint que recibe las solicitudes de reunión (ej. Formspree, Getform,
// un webhook de Make/Zapier o tu propio backend). Si queda vacío, el
// formulario solo muestra el mensaje de confirmación.
const FORM_ENDPOINT = "";

document.getElementById("year").textContent = new Date().getFullYear();

/* ---------- Header con borde al hacer scroll ---------- */
const header = document.querySelector(".site-header");
const mobileCta = document.querySelector(".mobile-cta");
const hero = document.querySelector(".hero");
const booking = document.getElementById("agendar");

function onScroll() {
  header.classList.toggle("scrolled", window.scrollY > 8);

  // CTA fija en móvil: visible después del hero y oculta en la sección de agenda
  const pastHero = hero.getBoundingClientRect().bottom < 0;
  const b = booking.getBoundingClientRect();
  const inBooking = b.top < window.innerHeight && b.bottom > 0;
  const show = pastHero && !inBooking;
  mobileCta.classList.toggle("show", show);
  mobileCta.setAttribute("aria-hidden", String(!show));
  mobileCta.tabIndex = show ? 0 : -1;
}
window.addEventListener("scroll", onScroll, { passive: true });
onScroll();

/* ---------- Menú móvil ---------- */
const toggle = document.querySelector(".nav-toggle");
const mobileNav = document.getElementById("mobile-nav");

function setMenu(open) {
  toggle.setAttribute("aria-expanded", String(open));
  toggle.setAttribute("aria-label", open ? "Cerrar menú" : "Abrir menú");
  mobileNav.hidden = !open;
}
toggle.addEventListener("click", () => setMenu(mobileNav.hidden));
mobileNav.querySelectorAll("a").forEach((a) => a.addEventListener("click", () => setMenu(false)));
window.addEventListener("resize", () => { if (window.innerWidth > 820) setMenu(false); });

/* ---------- Animación de aparición ---------- */
const revealTargets = document.querySelectorAll(
  ".section-head, .challenge-list li, .service, .industry-list li, .steps li, .booking-form, .faq-list details"
);
if ("IntersectionObserver" in window) {
  const io = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) {
        entry.target.classList.add("visible");
        io.unobserve(entry.target);
      }
    });
  }, { rootMargin: "0px 0px -8% 0px" });

  revealTargets.forEach((el) => {
    el.classList.add("reveal");
    io.observe(el);
  });
}

/* ---------- Formulario de agenda ---------- */
const form = document.getElementById("booking-form");
const dateInput = document.getElementById("fecha");

// La fecha mínima es el próximo día hábil
(function setMinDate() {
  const d = new Date();
  d.setDate(d.getDate() + 1);
  while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1);
  const iso = new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  dateInput.min = iso;
})();

const messages = {
  nombre: "Ingresa tu nombre.",
  empresa: "Ingresa el nombre de tu empresa.",
  email: "Ingresa un correo válido.",
  telefono: "Ingresa un teléfono de contacto.",
  industria: "Selecciona una industria.",
  dotacion: "Selecciona un rango.",
  tema: "Elige un tema principal.",
  fecha: "Elige una fecha hábil a partir de mañana.",
};

function fieldOf(el) { return el.closest(".field"); }

function setError(el, msg) {
  const field = fieldOf(el);
  field.classList.toggle("invalid", Boolean(msg));
  field.querySelector(".error").textContent = msg || "";
}

function validateInput(el) {
  let valid = el.checkValidity();

  if (el.name === "telefono" && valid) {
    valid = el.value.replace(/\D/g, "").length >= 8;
  }
  if (el.name === "fecha" && valid && el.value) {
    const day = new Date(el.value + "T12:00:00").getDay();
    valid = day !== 0 && day !== 6;
  }

  setError(el, valid ? "" : messages[el.name]);
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
    firstInvalid.focus({ preventScroll: true });
    fieldOf(firstInvalid).scrollIntoView({ behavior: "smooth", block: "center" });
  }
  return !firstInvalid;
}

// Validación en vivo después del primer intento
form.addEventListener("input", (e) => {
  const el = e.target;
  if (!el.name || !messages[el.name]) return;
  if (el.type === "radio") return setError(el, "");
  if (fieldOf(el).classList.contains("invalid")) validateInput(el);
});
form.addEventListener("change", (e) => {
  if (e.target.tagName === "SELECT" || e.target.type === "date") validateInput(e.target);
});

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  if (!validateForm()) return;

  const button = form.querySelector('button[type="submit"]');
  const data = Object.fromEntries(new FormData(form));
  button.disabled = true;
  button.textContent = "Enviando…";

  try {
    if (FORM_ENDPOINT) {
      const res = await fetch(FORM_ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
    } else {
      console.info("[Sur Consultores] Configura FORM_ENDPOINT en script.js. Datos:", data);
      await new Promise((r) => setTimeout(r, 600));
    }

    form.querySelector("[data-name]").textContent = data.nombre.trim().split(" ")[0];
    form.querySelector(".form-body").hidden = true;
    const success = form.querySelector(".form-success");
    success.hidden = false;
    success.focus();
  } catch (err) {
    console.error(err);
    button.disabled = false;
    button.textContent = "Solicitar reunión";
    alert("No pudimos enviar tu solicitud. Inténtalo nuevamente o escríbenos a contacto@surconsultores.cl.");
  }
});
