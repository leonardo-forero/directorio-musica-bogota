let data = [];
let markers = [];
let map;
let markerById = new Map();

document.addEventListener("DOMContentLoaded", () => {
    initializeMap();
    document.getElementById("search").addEventListener("input", applyFilters);
    document.getElementById("typeFilter").addEventListener("change", applyFilters);
    document.getElementById("localidadFilter").addEventListener("change", applyFilters);
    document.getElementById("clearFilters").addEventListener("click", clearFilters);
    loadExcel();
});

function initializeMap() {
    map = L.map("map").setView([4.65, -74.08], 11);

    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
        maxZoom: 19,
        attribution: '&copy; OpenStreetMap contributors'
    }).addTo(map);
}

async function loadExcel() {
    setStatus("Cargando datos…");

    try {
        const response = await fetch("data/directorio.xlsx");

        if (!response.ok) {
            throw new Error(`No se pudo cargar data/directorio.xlsx (${response.status})`);
        }

        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });

        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];

        data = XLSX.utils.sheet_to_json(sheet, {
            defval: "",
            raw: false
        });

        data = data
            .map(normalizeRecord)
            .filter(item => item.Visible.toLowerCase() !== "no");

        initializeFilters();
        renderDirectory(data);
        renderMap(data);

        setStatus(`${data.length} registros visibles`);
    } catch (error) {
        console.error(error);
        setStatus("No fue posible cargar la matriz.");
        document.getElementById("directoryList").innerHTML = `
            <div class="empty">
                <strong>No se pudieron cargar los datos.</strong>
                <p>Verifica que el archivo exista exactamente en <code>data/directorio.xlsx</code>.</p>
                <p>Si abriste el archivo HTML directamente desde tu computador, recuerda que
                <strong>fetch()</strong> puede estar bloqueado. En GitHub Pages funcionará correctamente.</p>
            </div>
        `;
    }
}

function normalizeRecord(item) {
    return {
        ID: String(item.ID ?? "").trim(),
        Nombre: String(item.Nombre ?? "").trim(),
        Tipo: String(item.Tipo ?? "").trim(),
        Agrupacion: String(item.Agrupacion ?? "").trim(),
        Correo: String(item.Correo ?? "").trim(),
        Telefono: String(item.Telefono ?? "").trim(),
        Direccion: String(item.Direccion ?? "").trim(),
        Localidad: String(item.Localidad ?? "").trim(),
        Descripcion: String(item.Descripcion ?? "").trim(),
        Latitud: String(item.Latitud ?? "").replace(",", ".").trim(),
        Longitud: String(item.Longitud ?? "").replace(",", ".").trim(),
        Visible: String(item.Visible ?? "Sí").trim()
    };
}

function initializeFilters() {
    const typeFilter = document.getElementById("typeFilter");
    const localidadFilter = document.getElementById("localidadFilter");

    typeFilter.innerHTML = '<option value="">Todos los tipos</option>';
    localidadFilter.innerHTML = '<option value="">Todas las localidades</option>';

    uniqueSorted(data.map(item => item.Tipo)).forEach(type => {
        const option = document.createElement("option");
        option.value = type;
        option.textContent = type;
        typeFilter.appendChild(option);
    });

    uniqueSorted(data.map(item => item.Localidad)).forEach(localidad => {
        const option = document.createElement("option");
        option.value = localidad;
        option.textContent = localidad;
        localidadFilter.appendChild(option);
    });
}

function uniqueSorted(values) {
    return [...new Set(values.filter(Boolean))].sort((a, b) =>
        a.localeCompare(b, "es")
    );
}

function applyFilters() {
    const search = document.getElementById("search").value.trim().toLowerCase();
    const type = document.getElementById("typeFilter").value;
    const localidad = document.getElementById("localidadFilter").value;

    const filtered = data.filter(item => {
        const searchable = [
            item.ID,
            item.Nombre,
            item.Tipo,
            item.Agrupacion,
            item.Localidad,
            item.Descripcion
        ].join(" ").toLowerCase();

        return (
            searchable.includes(search) &&
            (!type || item.Tipo === type) &&
            (!localidad || item.Localidad === localidad)
        );
    });

    renderDirectory(filtered);
    renderMap(filtered);

    setStatus(`${filtered.length} de ${data.length} registros`);
}

function clearFilters() {
    document.getElementById("search").value = "";
    document.getElementById("typeFilter").value = "";
    document.getElementById("localidadFilter").value = "";

    renderDirectory(data);
    renderMap(data);
    setStatus(`${data.length} registros visibles`);
}

function renderDirectory(items) {
    const container = document.getElementById("directoryList");
    const counter = document.getElementById("counter");

    container.innerHTML = "";
    counter.textContent = `${items.length} ${items.length === 1 ? "registro" : "registros"}`;

    if (!items.length) {
        container.innerHTML = '<div class="empty">No hay registros que coincidan con los filtros.</div>';
        return;
    }

    items.forEach(item => {
        const card = document.createElement("article");
        card.className = "person-card";

        card.innerHTML = `
           <h3>${escapeHtml(item.Nombre || "Sin nombre")}</h3>
           ${item.Tipo ? `<div class="type">${escapeHtml(item.Tipo)}</div>` : ""}
           ${item.Agrupacion ? `<div class="group">${escapeHtml(item.Agrupacion)}</div>` : ""}
           ${item.Localidad ? `<div class="localidad">📍 ${escapeHtml(item.Localidad)}</div>` : ""}
           ${item.Correo ? `<div class="contact">✉️ <a href="mailto:${escapeHtml(item.Correo)}" onclick="event.stopPropagation()">${escapeHtml(item.Correo)}</a></div>` : ""}
           ${item.Descripcion ? `<p>${escapeHtml(item.Descripcion)}</p>` : ""}
       `;

        card.addEventListener("click", () => focusOnMap(item));
        container.appendChild(card);
    });
}

function renderMap(items) {
    markers.forEach(marker => map.removeLayer(marker));
    markers = [];
    markerById.clear();

    const validPoints = [];

    items.forEach(item => {
        const lat = parseFloat(item.Latitud);
        const lng = parseFloat(item.Longitud);

        if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;

        const marker = L.marker([lat, lng]).addTo(map);

        const popup = `
           <div class="popup-title">${escapeHtml(item.Nombre || "Sin nombre")}</div>
           ${item.Tipo ? `<div class="popup-type">${escapeHtml(item.Tipo)}</div>` : ""}
           ${item.Agrupacion ? `<div><strong>${escapeHtml(item.Agrupacion)}</strong></div>` : ""}
           ${item.Localidad ? `<div>📍 ${escapeHtml(item.Localidad)}</div>` : ""}
           ${item.Correo ? `<div class="popup-contact">✉️ <a href="mailto:${escapeHtml(item.Correo)}">${escapeHtml(item.Correo)}</a></div>` : ""}
           ${item.Descripcion ? `<div class="popup-description">${escapeHtml(item.Descripcion)}</div>` : ""}
       `;

        marker.bindPopup(popup);
        marker.itemData = item;

        markers.push(marker);
        markerById.set(item.ID, marker);
        validPoints.push([lat, lng]);
    });

    if (validPoints.length > 0 && items.length === data.length) {
        map.fitBounds(validPoints, { padding: [30, 30], maxZoom: 13 });
    }
}

function focusOnMap(item) {
    const lat = parseFloat(item.Latitud);
    const lng = parseFloat(item.Longitud);

    if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
        setStatus("Este registro todavía no tiene coordenadas.");
        return;
    }

    map.setView([lat, lng], 16, { animate: true });

    const marker = markerById.get(item.ID);
    if (marker) marker.openPopup();
}

function setStatus(message) {
    document.getElementById("status").textContent = message;
}

function escapeHtml(value) {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}
