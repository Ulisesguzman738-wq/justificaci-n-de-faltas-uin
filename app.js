// Justifaltas - Core Application JavaScript (Excel Sincronizado, Límites, Contraseñas y Soporte macOS/file://)

// Google Sheets Integration Configuration
// Coloca aquí la URL de la Web App obtenida al implementar tu Google Apps Script.
// Si está vacía, el sistema operará en "Modo Local" (usando localStorage).
const GOOGLE_SCRIPT_URL = "";

// Helper para actualizar el indicador visual de sincronización en la barra lateral
function showSyncStatus(status, text) {
    const dot = document.getElementById('sync-dot');
    const textEl = document.getElementById('sync-text');
    if (!dot || !textEl) return;
    
    // Remover animación
    dot.style.animation = 'none';
    
    switch(status) {
        case 'local':
            dot.style.background = '#94a3b8'; // Gris
            dot.style.boxShadow = '0 0 4px rgba(148,163,184,0.5)';
            textEl.innerText = text || 'Modo Local';
            break;
        case 'sincronizando':
            dot.style.background = '#f97316'; // Naranja
            dot.style.boxShadow = '0 0 8px rgba(249,115,22,0.8)';
            textEl.innerText = text || 'Sincronizando...';
            dot.style.animation = 'pulse 1.2s infinite alternate';
            break;
        case 'subiendo':
            dot.style.background = '#3b82f6'; // Azul
            dot.style.boxShadow = '0 0 8px rgba(59,130,246,0.8)';
            textEl.innerText = text || 'Subiendo archivo...';
            dot.style.animation = 'pulse 1.2s infinite alternate';
            break;
        case 'sincronizado':
            dot.style.background = '#22c55e'; // Verde
            dot.style.boxShadow = '0 0 8px rgba(34,197,94,0.8)';
            textEl.innerText = text || 'Sincronizado';
            break;
        case 'error':
            dot.style.background = '#ef4444'; // Rojo
            dot.style.boxShadow = '0 0 8px rgba(239,68,68,0.8)';
            textEl.innerText = text || 'Error de sincronización';
            break;
    }
}

// Inyectar estilos para animación de parpadeo si no existen
if (!document.getElementById('sync-style-animation')) {
    const style = document.createElement('style');
    style.id = 'sync-style-animation';
    style.innerHTML = `
        @keyframes pulse {
            0% { opacity: 0.4; }
            100% { opacity: 1; }
        }
    `;
    document.head.appendChild(style);
}

// 1. DATABASE SYSTEM (localStorage representation of DB schema with in-memory fallback for file:// compatibility)
const DB_KEY = 'justifaltas_db_final_v2';

// Seed Data directly synchronized with the Excel sheet (with corrections and passwords)
const officialDB = {
    usuarios: [
        { ID_Usuario: 'sim_user_epy1yk2', Correo_Electronico: 'Prueba@uin.edu.mx', Nombre_Completo: 'Usuario de Pruebas UIN', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: '2026-06-14T00:00:00Z', Activo: 1 },
        { ID_Usuario: 'u1', Correo_Electronico: 'ulisega03s24a@uinteramericana.edu.mx', Nombre_Completo: 'Ulise Guzman Alvaro', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: '2026-05-10T00:00:00Z', Activo: 1 },
        { ID_Usuario: 'u2', Correo_Electronico: 'syanethh03s24a@uinteramericana.edu.mx', Nombre_Completo: 'Sandra Yanet Hernández Hernández', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: '2026-05-11T00:00:00Z', Activo: 1 },
        { ID_Usuario: 'u3', Correo_Electronico: 'irenehh03s24a@uinteramericana.edu.mx', Nombre_Completo: 'Irene Hernández Hernández', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 'u4', Correo_Electronico: 'iyuleidals03s24a@uinteramericana.edu.mx', Nombre_Completo: 'Itzel Yuleida Lila Sanchez', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 'u5', Correo_Electronico: 'dantonioie03s24a@uinteramericana.edu.mx', Nombre_Completo: 'Diego Antonio Iracheta Escareño', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 't1', Correo_Electronico: 'mdavila0311@uinteramericana.edu.mx', Nombre_Completo: 'Miguel Angel Davila Calzada', Rol: 'maestro', Contrasena: '12345', Fecha_Registro: '2026-05-01T00:00:00Z', Activo: 1 },
        { ID_Usuario: 't2', Correo_Electronico: 'ana.hernandez03ea26@uinteramericana.edu.mx', Nombre_Completo: 'Oscar Alavez Reyes', Rol: 'maestro', Contrasena: '12345', Fecha_Registro: '2026-05-01T00:00:00Z', Activo: 1 },
        { ID_Usuario: 't3', Correo_Electronico: 'oscar.alavez03ea26p@uinteramericana.edu.mx', Nombre_Completo: 'Andres Simón Treviño', Rol: 'maestro', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 't4', Correo_Electronico: 'andres.simon03ea26@uinteramericana.edu.mx', Nombre_Completo: 'Ana Eloisa Hernández Jimenez', Rol: 'maestro', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 't5', Correo_Electronico: 'direccionsc_sd25o@uinteramericana.edu.mx', Nombre_Completo: 'Lic. Sandra Silva', Rol: 'coordinacion', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 't6', Correo_Electronico: 'escolarsc@universidadinteramericana.edu.mx', Nombre_Completo: 'Escolar Santa Catarina', Rol: 'coordinacion', Contrasena: '12345', Fecha_Registro: null, Activo: 1 }
    ],
    justificaciones: [
        {
            ID_Justificante: 'j_demo_001',
            ID_Alumno: 'u4',
            Fecha_Falta: '2026-06-14',
            Periodo_Tetra: 'Mayo – Agosto',
            Parcial: 'Parcial 2',
            Motivo: 'Consulta médica',
            Descripcion: 'Ejemplo de solicitud para probar el flujo completo de Justifaltas.',
            Estado: 'Pendiente',
            Fecha_Registro: '2026-06-14T10:00:00Z',
            ID_Coordinador_Revisor: null,
            Fecha_Revision: null
        }
    ],
    archivos_adjuntos: [
        {
            ID_Archivo: 'a_demo_001',
            ID_Justificante: 'j_demo_001',
            Nombre_Archivo: 'evidencia_demo.pdf',
            Ruta_Archivo: 'data:application/pdf;base64,JVBERi0xLjQKJdPr6gogMSAwIG9iagogIDw8IC9UeXBlIC9DYXRhbG9nIC9QYWdlcyAyIDAgUiA+PiBlbmRvYmoKMiAwIG9iagogIDw8IC9UeXBlIC9QYWdlcyAvS2lkcyBbIDMgMCBSIF0gL0NvdW50IDEgPj4gZW5kb2JqCjMgMCBSIHhtbA==',
            Tipo_Archivo: 'PDF',
            Fecha_Carga: '2026-06-14T10:05:00Z',
            Tamano_KB: 12
        }
    ],
    justificante_maestro: [
        {
            ID: 'jm_demo_001',
            ID_Justificante: 'j_demo_001',
            ID_Maestro: 't1',
            Estado_Maestro: 'Pendiente',
            Fecha_Notificacion: '2026-06-14T10:05:00Z',
            Fecha_Justificacion: null
        }
    ],
    observaciones: [
        {
            ID_Observacion: 'obs_demo_001',
            ID_Justificante: 'j_demo_001',
            ID_Usuario: 't5',
            Comentario: 'Pendiente de revisión por coordinación.',
            Tipo: 'Revision',
            Fecha: '2026-06-14T10:10:00Z'
        }
    ],
    notificaciones: [
        {
            ID_Notificacion: 'n_demo_001',
            ID_Usuario: 't1',
            ID_Justificante: 'j_demo_001',
            Mensaje: 'Nueva solicitud pendiente de aprobación por coordinación.',
            Leida: 0,
            Fecha: '2026-06-14T10:10:00Z'
        }
    ]
};

let DB = null;
let inMemorySessionEmail = null; // Session backup for file:// protocol

function loadDatabase() {
    try {
        const data = localStorage.getItem(DB_KEY);
        if (!data) {
            localStorage.setItem(DB_KEY, JSON.stringify(officialDB));
            DB = JSON.parse(JSON.stringify(officialDB)); // Deep clone
        } else {
            DB = JSON.parse(data);
        }
    } catch (e) {
        // Fallback for Safari/Chrome on file:// where localStorage might throw a SecurityError
        console.warn("localStorage is restricted under file:// protocol. Falling back to in-memory database.");
        if (!DB) {
            DB = JSON.parse(JSON.stringify(officialDB));
        }
    }
}

async function syncDatabaseFromSheets() {
    if (!GOOGLE_SCRIPT_URL) {
        showSyncStatus('local', 'Modo Local');
        return;
    }
    
    showSyncStatus('sincronizando', 'Sincronizando con Google Sheets...');
    
    try {
        const response = await fetch(GOOGLE_SCRIPT_URL);
        const resData = await response.json();
        
        if (resData.success && resData.data) {
            DB = resData.data;
            try {
                localStorage.setItem(DB_KEY, JSON.stringify(DB));
            } catch(e) {}
            
            showSyncStatus('sincronizado', 'Sincronizado con Google Sheets');
            
            if (currentUser) {
                const refreshedUser = DB.usuarios.find(u => u.ID_Usuario === currentUser.ID_Usuario || u.Correo_Electronico.toLowerCase() === currentUser.Correo_Electronico.toLowerCase());
                if (refreshedUser) {
                    currentUser = refreshedUser;
                }
                
                if (currentRole === 'alumno') renderAlumnoDashboard();
                else if (currentRole === 'coordinacion') renderCoordinacionDashboard();
                else if (currentRole === 'maestro') renderMaestroDashboard();
            }
        } else {
            showSyncStatus('error', 'Error al leer Google Sheets');
            console.error("Error reading sheets:", resData.error);
        }
    } catch(e) {
        showSyncStatus('error', 'Error de red al sincronizar');
        console.error("Network error reading sheets:", e);
    }
}

function saveDatabase() {
    try {
        localStorage.setItem(DB_KEY, JSON.stringify(DB));
    } catch (e) {
        console.warn("Could not save to localStorage. Changes will remain in-memory for this session.");
    }
    
    if (GOOGLE_SCRIPT_URL) {
        showSyncStatus('sincronizando', 'Guardando cambios...');
        
        fetch(GOOGLE_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'save_db',
                data: DB
            }),
            headers: {
                'Content-Type': 'text/plain'
            }
        })
        .then(response => response.json())
        .then(resData => {
            if (resData.success) {
                showSyncStatus('sincronizado', 'Sincronizado con Google Sheets');
            } else {
                showSyncStatus('error', 'Error al guardar cambios');
                console.error("Error saving DB to Sheets:", resData.error);
            }
        })
        .catch(err => {
            showSyncStatus('error', 'Error de red al guardar');
            console.error("Network error saving DB to Sheets:", err);
        });
    }
}

// 2. STATE MANAGEMENT & AUTH
let currentUser = null;
let currentRole = null;
let activeRequestForReview = null;

function initApp() {
    loadDatabase();
    
    // Iniciar sincronización si hay URL configurada
    if (GOOGLE_SCRIPT_URL) {
        syncDatabaseFromSheets();
    } else {
        showSyncStatus('local', 'Modo Local');
    }
    
    // Default Tetra value initialization based on current system month
    const currentMonth = new Date().getMonth() + 1; // 1-12
    let initialTetra = "Mayo – Agosto";
    if (currentMonth >= 1 && currentMonth <= 4) initialTetra = "Enero – Abril";
    else if (currentMonth >= 9 && currentMonth <= 12) initialTetra = "Septiembre – Diciembre";
    
    const filterSelect = document.getElementById('global-tetra-filter');
    filterSelect.value = initialTetra;
    
    // Listener for Global Tetra Filter
    filterSelect.addEventListener('change', () => {
        if (currentRole === 'alumno') renderAlumnoDashboard();
        else if (currentRole === 'coordinacion') renderCoordinacionDashboard();
        else if (currentRole === 'maestro') renderMaestroDashboard();
    });

    let savedSession = null;
    try {
        savedSession = localStorage.getItem('justifaltas_session_secured');
    } catch (e) {
        savedSession = inMemorySessionEmail;
    }

    if (savedSession) {
        const sessionUser = DB.usuarios.find(u => u.Correo_Electronico === savedSession);
        if (sessionUser) {
            login(sessionUser);
            return;
        }
    }
    
    showScreen('login');
}

function login(user) {
    currentUser = user;
    currentRole = user.Rol;
    
    try {
        localStorage.setItem('justifaltas_session_secured', user.Correo_Electronico);
    } catch (e) {
        inMemorySessionEmail = user.Correo_Electronico;
    }
    
    document.getElementById('user-name').innerText = user.Nombre_Completo;
    document.getElementById('user-avatar').innerText = user.Nombre_Completo.charAt(0);
    document.getElementById('user-role-badge').innerText = user.Rol === 'coordinacion' ? 'Coordinación' : user.Rol;
    document.getElementById('role-display-badge').innerText = user.Rol === 'coordinacion' ? 'Coordinación' : user.Rol;
    
    setupSidebarMenu();
    showScreen('app');
    
    if (currentRole === 'alumno') {
        switchTab('panel-alumno');
        renderAlumnoDashboard();
    } else if (currentRole === 'coordinacion') {
        switchTab('panel-coordinacion');
        renderCoordinacionDashboard();
    } else if (currentRole === 'maestro') {
        switchTab('panel-maestro');
        renderMaestroDashboard();
    }
}

function logout() {
    currentUser = null;
    currentRole = null;
    try {
        localStorage.removeItem('justifaltas_session_secured');
    } catch (e) {
        inMemorySessionEmail = null;
    }
    showScreen('login');
}

// 3. UI HELPER FUNCTIONS
function showScreen(screen) {
    const loginScreen = document.getElementById('login-screen');
    const appContainer = document.getElementById('app-container');
    
    if (screen === 'login') {
        loginScreen.style.display = 'flex';
        appContainer.style.display = 'none';
        toggleLoginMode('login');
    } else {
        loginScreen.style.display = 'none';
        appContainer.style.display = 'flex';
    }
}

function switchTab(tabId) {
    const panels = document.querySelectorAll('.view-panel');
    panels.forEach(p => p.classList.remove('active'));
    
    const target = document.getElementById(tabId);
    if (target) target.classList.add('active');
    
    const menuItems = document.querySelectorAll('.menu-item');
    menuItems.forEach(item => {
        if (item.getAttribute('data-target') === tabId) {
            item.classList.add('active');
        } else {
            item.classList.remove('active');
        }
    });
}

function setupSidebarMenu() {
    const menuList = document.getElementById('sidebar-menu');
    menuList.innerHTML = '';
    
    if (currentRole === 'alumno') {
        menuList.innerHTML = `
            <li><a class="menu-item active" data-target="panel-alumno" onclick="switchTab('panel-alumno')">📊 Dashboard Alumno</a></li>
        `;
    } else if (currentRole === 'coordinacion') {
        menuList.innerHTML = `
            <li><a class="menu-item active" data-target="panel-coordinacion" onclick="switchTab('panel-coordinacion')">⚖️ Revisar Solicitudes</a></li>
        `;
    } else if (currentRole === 'maestro') {
        menuList.innerHTML = `
            <li><a class="menu-item active" data-target="panel-maestro" onclick="switchTab('panel-maestro')">🎓 Bandeja de Clases</a></li>
        `;
    }
}

// 4. TETRA & PARCIAL CALCULATION LOGIC
function getTetraAndParcial(dateStr) {
    if (!dateStr) return { tetra: '', parcial: '' };
    
    const parts = dateStr.split('-');
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10); // 1-12
    
    let tetra = "";
    let monthInTetra = 1;
    
    if (month >= 1 && month <= 4) {
        tetra = "Enero – Abril";
        monthInTetra = month;
    } else if (month >= 5 && month <= 8) {
        tetra = "Mayo – Agosto";
        monthInTetra = month - 4;
    } else {
        tetra = "Septiembre – Diciembre";
        monthInTetra = month - 8;
    }
    
    let parcial = "";
    if (monthInTetra === 1) {
        parcial = "Parcial 1";
    } else if (monthInTetra === 2) {
        parcial = "Parcial 2";
    } else {
        parcial = "Parcial 3";
    }
    
    return { tetra, parcial };
}

// 5. ALUMNO DASHBOARD LOGIC
function renderAlumnoDashboard() {
    const activeTetra = document.getElementById('global-tetra-filter').value;
    document.getElementById('welcome-title').innerText = `¡Hola, ${currentUser.Nombre_Completo}!`;
    document.getElementById('welcome-subtitle').innerText = 'Registra tus faltas y consulta el estatus de tus justificantes.';
    document.getElementById('student-history-title').innerText = `Mis Solicitudes (${activeTetra})`;
    
    const studentId = currentUser.ID_Usuario;
    
    const studentRequests = DB.justificaciones.filter(j => j.ID_Alumno === studentId && j.Periodo_Tetra === activeTetra);
    
    // Metrics
    const total = studentRequests.length;
    const pending = studentRequests.filter(j => j.Estado === 'Pendiente').length;
    const justified = studentRequests.filter(j => j.Estado === 'Enterada por Maestro').length;
    
    document.getElementById('metric-student-total').innerText = total;
    document.getElementById('metric-student-pending').innerText = pending;
    document.getElementById('metric-student-justified').innerText = justified;
    
    // Populate Teachers
    const formTeachers = document.getElementById('form-teachers');
    formTeachers.innerHTML = '';
    DB.usuarios.filter(u => u.Rol === 'maestro').forEach(m => {
        const option = document.createElement('option');
        option.value = m.ID_Usuario;
        option.innerText = m.Nombre_Completo;
        formTeachers.appendChild(option);
    });
    
    // Render Table
    const tableBody = document.getElementById('student-requests-table');
    tableBody.innerHTML = '';
    
    if (studentRequests.length === 0) {
        tableBody.innerHTML = `<tr><td colspan="7" class="empty-state">No hay solicitudes registradas para el periodo: ${activeTetra}.</td></tr>`;
        return;
    }
    
    studentRequests.sort((a,b) => new Date(b.Fecha_Registro) - new Date(a.Fecha_Registro));
    
    studentRequests.forEach(j => {
        const mappings = DB.justificante_maestro.filter(jm => jm.ID_Justificante === j.ID_Justificante);
        let teachersStatusHTML = '';
        mappings.forEach(m => {
            const maestroUser = DB.usuarios.find(u => u.ID_Usuario === m.ID_Maestro);
            const name = maestroUser ? maestroUser.Nombre_Completo : 'Maestro';
            if (m.Estado_Maestro === 'Enterada por Maestro') {
                teachersStatusHTML += `<div style="font-size:0.72rem; color:var(--success); font-weight:600;">✔️ ${name} (Confirmó: ${formatDateTimeString(m.Fecha_Justificacion)})</div>`;
            } else {
                teachersStatusHTML += `<div style="font-size:0.72rem; color:var(--text-muted);">⏳ ${name} (Pendiente)</div>`;
            }
        });
        
        let reviewLogHTML = '';
        if (j.ID_Coordinador_Revisor) {
            const coordUser = DB.usuarios.find(u => u.ID_Usuario === j.ID_Coordinador_Revisor);
            const coordName = coordUser ? coordUser.Nombre_Completo : 'Coordinación';
            const actionVerb = j.Estado === 'Rechazada' ? 'Rechazado' : 'Aprobado';
            reviewLogHTML = `<div style="font-size:0.72rem; color:var(--primary); font-weight:600; margin-top:4px;">🛡️ ${actionVerb} por: ${coordName}</div>`;
        }
        
        const obs = DB.observaciones.filter(o => o.ID_Justificante === j.ID_Justificante).reverse();
        const obsText = obs.length > 0 ? obs[0].Comentario : null;
        
        let badgeClass = 'badge-pending';
        let statusLabel = j.Estado;
        if (j.Estado === 'Aprobada') badgeClass = 'badge-approved';
        else if (j.Estado === 'Rechazada') badgeClass = 'badge-rejected';
        else if (j.Estado === 'Enterada por Maestro') {
            badgeClass = 'badge-finalized';
            statusLabel = 'Enterada';
        }
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><small style="color:var(--text-muted); font-family:monospace;">${j.ID_Justificante}</small></td>
            <td><strong>${formatDateString(j.Fecha_Falta)}</strong></td>
            <td><span class="badge badge-approved" style="background:#f3f4f6; color:var(--text-main);">${j.Parcial || 'Sin Asignar'}</span></td>
            <td>${j.Motivo}</td>
            <td>
                ${teachersStatusHTML}
                ${reviewLogHTML}
            </td>
            <td>
                <a href="#" class="evidence-link" onclick="viewEvidenceModal('${j.ID_Justificante}')">📎 Ver Archivo</a>
            </td>
            <td>
                <span class="badge ${badgeClass}">${statusLabel}</span>
                ${obsText ? `<br><small style="color: var(--text-muted); font-style: italic;">Obs: ${obsText}</small>` : ''}
            </td>
        `;
        tableBody.appendChild(row);
    });
}

// 6. COORDINACION DASHBOARD LOGIC
function renderCoordinacionDashboard() {
    const activeTetra = document.getElementById('global-tetra-filter').value;
    document.getElementById('welcome-title').innerText = `Panel de Coordinación`;
    document.getElementById('welcome-subtitle').innerText = 'Portal de validación y control escolar.';
    
    const myJustifications = DB.justificaciones.filter(j => j.Periodo_Tetra === activeTetra);
    
    const totalPending = myJustifications.filter(j => j.Estado === 'Pendiente').length;
    const totalApproved = myJustifications.filter(j => j.Estado === 'Aprobada' || j.Estado === 'Enterada por Maestro').length;
    const totalRejected = myJustifications.filter(j => j.Estado === 'Rechazada').length;
    
    document.getElementById('metric-coord-pending').innerText = totalPending;
    document.getElementById('metric-coord-approved').innerText = totalApproved;
    document.getElementById('metric-coord-rejected').innerText = totalRejected;
    
    const tableBody = document.getElementById('coord-pending-table');
    tableBody.innerHTML = '';
    
    const pendingRequests = myJustifications.filter(j => j.Estado === 'Pendiente');
    
    if (pendingRequests.length === 0) {
        tableBody.innerHTML = `<tr><td colspan="4" class="empty-state">No hay solicitudes pendientes en la bandeja de entrada para ${activeTetra}.</td></tr>`;
        document.getElementById('coord-review-details').style.display = 'none';
        document.getElementById('coord-empty-details').style.display = 'block';
        resetEvidenceViewer();
        return;
    }
    
    pendingRequests.sort((a,b) => new Date(a.Fecha_Registro) - new Date(a.Fecha_Registro));
    
    pendingRequests.forEach(j => {
        const studentUser = DB.usuarios.find(u => u.ID_Usuario === j.ID_Alumno);
        
        const row = document.createElement('tr');
        row.style.cursor = 'pointer';
        row.onclick = () => selectRequestForReview(j.ID_Justificante);
        
        row.innerHTML = `
            <td>
                <strong>${studentUser ? studentUser.Nombre_Completo : 'Alumno'}</strong><br>
                <small style="color: var(--text-muted);">${studentUser ? studentUser.Correo_Electronico : ''}</small>
            </td>
            <td>${formatDateString(j.Fecha_Falta)}</td>
            <td><span class="badge badge-pending">${j.Motivo}</span></td>
            <td><button class="btn btn-primary btn-sm">Evaluar</button></td>
        `;
        tableBody.appendChild(row);
    });
    
    if (activeRequestForReview) {
        const req = DB.justificaciones.find(j => j.ID_Justificante === activeRequestForReview.ID_Justificante);
        if (req && req.Estado === 'Pendiente' && req.Periodo_Tetra === activeTetra) {
            selectRequestForReview(req.ID_Justificante);
        } else {
            activeRequestForReview = null;
        }
    }
}

function selectRequestForReview(reqId) {
    const req = DB.justificaciones.find(j => j.ID_Justificante === reqId);
    if (!req) return;
    
    activeRequestForReview = req;
    const studentUser = DB.usuarios.find(u => u.ID_Usuario === req.ID_Alumno);
    
    document.getElementById('coord-empty-details').style.display = 'none';
    document.getElementById('coord-review-details').style.display = 'block';
    
    document.getElementById('review-student').innerText = studentUser ? studentUser.Nombre_Completo : 'N/A';
    document.getElementById('review-matricula').innerText = studentUser ? studentUser.ID_Usuario : 'N/A';
    document.getElementById('review-correo').innerText = studentUser ? studentUser.Correo_Electronico : 'N/A';
    document.getElementById('review-date').innerText = formatDateString(req.Fecha_Falta);
    document.getElementById('review-reason').innerText = req.Motivo;
    document.getElementById('review-description').innerText = req.Descripcion;
    document.getElementById('coord-comments').value = '';
    
    const parcialSelect = document.getElementById('review-parcial-select');
    parcialSelect.value = req.Parcial || getTetraAndParcial(req.Fecha_Falta).parcial || 'Parcial 1';
    
    const container = document.getElementById('evidence-viewer-container');
    container.innerHTML = '';
    
    const fileRecord = DB.archivos_adjuntos.find(a => a.ID_Justificante === req.ID_Justificante);
    
    if (!fileRecord) {
        container.innerHTML = `<div class="empty-state">No se cargó archivo de evidencia para este justificante.</div>`;
        return;
    }
    
    if (fileRecord.Ruta_Archivo.startsWith('http')) {
        let previewUrl = fileRecord.Ruta_Archivo;
        if (previewUrl.includes('/view')) {
            previewUrl = previewUrl.replace('/view', '/preview');
        }
        const frame = document.createElement('iframe');
        frame.src = previewUrl;
        frame.className = 'evidence-frame';
        container.appendChild(frame);
        
        const linkBtn = document.createElement('a');
        linkBtn.href = fileRecord.Ruta_Archivo;
        linkBtn.target = '_blank';
        linkBtn.className = 'btn btn-primary';
        linkBtn.style.marginTop = '10px';
        linkBtn.style.display = 'inline-block';
        linkBtn.style.textDecoration = 'none';
        linkBtn.innerHTML = '🔗 Abrir evidencia en pestaña nueva';
        container.appendChild(linkBtn);
    } else if (fileRecord.Ruta_Archivo.startsWith('data:application/pdf') || fileRecord.Tipo_Archivo === 'PDF') {
        const frame = document.createElement('iframe');
        frame.src = fileRecord.Ruta_Archivo;
        frame.className = 'evidence-frame';
        container.appendChild(frame);
    } else {
        const img = document.createElement('img');
        img.src = fileRecord.Ruta_Archivo;
        img.className = 'evidence-img';
        img.alt = 'Evidencia';
        container.appendChild(img);
    }
}

function resetEvidenceViewer() {
    const container = document.getElementById('evidence-viewer-container');
    container.innerHTML = `
        <div class="evidence-placeholder" id="evidence-view-placeholder">
            <p>Selecciona una solicitud para visualizar su documento soporte.</p>
        </div>
    `;
}

// 7. MAESTRO DASHBOARD LOGIC
function renderMaestroDashboard() {
    const activeTetra = document.getElementById('global-tetra-filter').value;
    document.getElementById('welcome-title').innerText = `Panel del Docente`;
    document.getElementById('welcome-subtitle').innerText = 'Revisa notificaciones de justificantes escolares y confirma su recepción.';
    
    const teacherId = currentUser.ID_Usuario;
    const myMappings = DB.justificante_maestro.filter(jm => jm.ID_Maestro === teacherId);
    
    const myJustifications = [];
    myMappings.forEach(m => {
        const j = DB.justificaciones.find(x => x.ID_Justificante === m.ID_Justificante);
        if (j && j.Periodo_Tetra === activeTetra && (j.Estado === 'Aprobada' || j.Estado === 'Enterada por Maestro')) {
            myJustifications.push({
                justificacion: j,
                maestro_mapping: m
            });
        }
    });
    
    // Metrics
    const totalApproved = myJustifications.length;
    const pendingConfirm = myJustifications.filter(x => x.maestro_mapping.Estado_Maestro === 'Pendiente').length;
    const confirmedCount = myJustifications.filter(x => x.maestro_mapping.Estado_Maestro === 'Enterada por Maestro').length;
    
    document.getElementById('metric-teacher-approved').innerText = totalApproved;
    document.getElementById('metric-teacher-pending').innerText = pendingConfirm;
    document.getElementById('metric-teacher-justified').innerText = confirmedCount;
    
    const notificationsArea = document.getElementById('teacher-notifications-area');
    notificationsArea.innerHTML = '';
    
    const unreadNotifications = DB.notificaciones.filter(n => n.ID_Usuario === teacherId && n.Leida === 0);
    if (unreadNotifications.length > 0) {
        notificationsArea.innerHTML = `
            <div class="notification-banner">
                <p>🔔 Tienes <strong>${unreadNotifications.length} nueva(s) justificación(es)</strong> pendientes de confirmar recepción.</p>
                <button class="btn btn-primary btn-sm" onclick="clearTeacherNotifications()">Marcar como Leídas</button>
            </div>
        `;
    }
    
    const tableBody = document.getElementById('teacher-requests-table');
    tableBody.innerHTML = '';
    
    if (myJustifications.length === 0) {
        tableBody.innerHTML = `<tr><td colspan="8" class="empty-state">No tienes justificantes notificados en el periodo ${activeTetra}.</td></tr>`;
        return;
    }
    
    myJustifications.sort((a,b) => {
        if (a.maestro_mapping.Estado_Maestro === 'Pendiente' && b.maestro_mapping.Estado_Maestro === 'Enterada por Maestro') return -1;
        if (a.maestro_mapping.Estado_Maestro === 'Enterada por Maestro' && b.maestro_mapping.Estado_Maestro === 'Pendiente') return 1;
        return new Date(b.justificacion.Fecha_Registro) - new Date(a.justificacion.Fecha_Registro);
    });
    
    myJustifications.forEach(item => {
        const j = item.justificacion;
        const m = item.maestro_mapping;
        const studentUser = DB.usuarios.find(u => u.ID_Usuario === j.ID_Alumno);
        
        let statusBadge = `<span class="badge badge-approved">Pendiente Confirmar</span>`;
        let actionBtn = `<button class="btn btn-success btn-sm" onclick="confirmTeacherReceipt('${m.ID}')">📋 Confirmar recepción</button>`;
        
        if (m.Estado_Maestro === 'Enterada por Maestro') {
            statusBadge = `<span class="badge badge-finalized">Confirmada</span>`;
            actionBtn = `<span style="color:var(--success); font-weight:600; font-size:0.82rem;">Confirmado: ${formatDateTimeString(m.Fecha_Justificacion)}</span>`;
        }
        
        const coordObs = DB.observaciones.filter(o => o.ID_Justificante === j.ID_Justificante && o.Tipo === 'Revision').reverse();
        const coordComment = coordObs.length > 0 ? coordObs[0].Comentario : 'Aprobado por Coordinación';
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><strong>${studentUser ? studentUser.Nombre_Completo : 'Alumno'}</strong></td>
            <td><small>${studentUser ? studentUser.Correo_Electronico : 'N/A'}</small></td>
            <td><strong>${formatDateString(j.Fecha_Falta)}</strong></td>
            <td><span class="badge badge-approved" style="background:#f3f4f6; color:var(--text-main);">${j.Parcial}</span></td>
            <td style="font-size: 0.85rem; font-style: italic; color: var(--primary);">
                "${coordComment}"
            </td>
            <td>
                <a href="#" class="evidence-link" onclick="viewEvidenceModal('${j.ID_Justificante}')">📎 Evidencia</a>
            </td>
            <td>${statusBadge}</td>
            <td>${actionBtn}</td>
        `;
        tableBody.appendChild(row);
    });
}

function clearTeacherNotifications() {
    const teacherId = currentUser.ID_Usuario;
    DB.notificaciones.forEach(n => {
        if (n.ID_Usuario === teacherId) {
            n.Leida = 1;
        }
    });
    saveDatabase();
    renderMaestroDashboard();
}

// 8. DATE SELECTION AND BUSINESS VALIDATION
function validateDateInput() {
    const dateInput = document.getElementById('form-date');
    const dateVal = dateInput.value;
    
    const timeLimitWarning = document.getElementById('form-time-limit-warning');
    const limitWarning = document.getElementById('form-parcial-limit-warning');
    const submitBtn = document.getElementById('form-submit-btn');
    const calculatedTetra = document.getElementById('form-calculated-tetra');
    const calculatedParcial = document.getElementById('form-calculated-parcial');
    
    if (!dateVal) {
        timeLimitWarning.style.display = 'none';
        limitWarning.style.display = 'none';
        calculatedTetra.value = '';
        calculatedParcial.value = '';
        submitBtn.disabled = false;
        return;
    }
    
    const { tetra, parcial } = getTetraAndParcial(dateVal);
    calculatedTetra.value = tetra;
    calculatedParcial.value = parcial;
    
    const now = new Date();
    const absenceDate = new Date(dateVal + 'T23:59:59');
    const diffMs = now.getTime() - absenceDate.getTime();
    const diffHours = diffMs / (1000 * 60 * 60);
    
    let isTimeExpired = diffHours > 24;
    timeLimitWarning.style.display = isTimeExpired ? 'block' : 'none';
    
    const studentId = currentUser.ID_Usuario;
    const existingCount = DB.justificaciones.filter(j => 
        j.ID_Alumno === studentId && 
        j.Estado !== 'Rechazada' &&
        j.Periodo_Tetra === tetra &&
        j.Parcial === parcial
    ).length;
    
    let isLimitReached = existingCount >= 2;
    limitWarning.style.display = isLimitReached ? 'block' : 'none';
    
    if (isTimeExpired || isLimitReached) {
        submitBtn.disabled = true;
        submitBtn.style.opacity = '0.5';
        submitBtn.style.cursor = 'not-allowed';
    } else {
        submitBtn.disabled = false;
        submitBtn.style.opacity = '1';
        submitBtn.style.cursor = 'pointer';
    }
}

document.getElementById('form-date').addEventListener('change', validateDateInput);
document.getElementById('form-date').addEventListener('input', validateDateInput);

// 9. NEW JUSTIFICATION SUBMISSION (Alumno)
document.getElementById('new-justification-form').addEventListener('submit', function(e) {
    e.preventDefault();
    
    const date = document.getElementById('form-date').value;
    const reason = document.getElementById('form-reason').value;
    const desc = document.getElementById('form-description').value;
    
    const teacherSelect = document.getElementById('form-teachers');
    const selectedTeachers = Array.from(teacherSelect.selectedOptions).map(opt => opt.value);
    
    if (selectedTeachers.length === 0) {
        alert('Debes seleccionar al menos un maestro.');
        return;
    }
    
    const fileInput = document.getElementById('form-file');
    const file = fileInput.files[0];
    
    if (!file) {
        alert('Debe cargar una evidencia física obligatoria (PDF/JPG/PNG).');
        return;
    }
    
    const { tetra, parcial } = getTetraAndParcial(date);
    const existingCount = DB.justificaciones.filter(j => 
        j.ID_Alumno === currentUser.ID_Usuario && 
        j.Estado !== 'Rechazada' &&
        j.Periodo_Tetra === tetra &&
        j.Parcial === parcial
    ).length;
    
    if (existingCount >= 2) {
        alert(`Límite alcanzado: Ya tienes registradas 2 justificaciones en el ${parcial} del periodo ${tetra}.`);
        return;
    }
    
    const reader = new FileReader();
    reader.onload = async function(event) {
        const base64Data = event.target.result;
        const newJustId = 'just_' + Math.random().toString(36).substr(2, 9);
        
        let fileUrl = base64Data;
        
        if (GOOGLE_SCRIPT_URL) {
            showSyncStatus('subiendo', 'Subiendo evidencia a Google Drive...');
            try {
                const response = await fetch(GOOGLE_SCRIPT_URL, {
                    method: 'POST',
                    body: JSON.stringify({
                        action: 'upload_file',
                        fileName: file.name,
                        base64Data: base64Data
                    }),
                    headers: { 'Content-Type': 'text/plain' }
                });
                const resData = await response.json();
                if (resData.success && resData.url) {
                    fileUrl = resData.url;
                    showSyncStatus('sincronizando', 'Archivo subido. Registrando...');
                } else {
                    console.error("Error upload_file script:", resData.error);
                    alert("Advertencia: No se pudo subir el archivo a Google Drive. Se guardará localmente.");
                }
            } catch (err) {
                console.error("Error upload_file fetch:", err);
                alert("Advertencia: Falló la subida del archivo a Google Drive. Se guardará localmente.");
            }
        }
        
        const newJustification = {
            ID_Justificante: newJustId,
            ID_Alumno: currentUser.ID_Usuario,
            Fecha_Falta: date,
            Periodo_Tetra: tetra,
            Parcial: parcial,
            Motivo: reason,
            Descripcion: desc,
            Estado: 'Pendiente',
            Fecha_Registro: new Date().toISOString(),
            ID_Coordinador_Revisor: null,
            Fecha_Revision: null
        };
        
        const fileExt = file.name.split('.').pop().toUpperCase();
        const newFile = {
            ID_Archivo: 'file_' + Math.random().toString(36).substr(2, 9),
            ID_Justificante: newJustId,
            Nombre_Archivo: file.name,
            Ruta_Archivo: fileUrl,
            Tipo_Archivo: ['PDF','JPG','PNG'].includes(fileExt) ? fileExt : 'PNG',
            Fecha_Carga: new Date().toISOString(),
            Tamano_KB: Math.round(file.size / 1024)
        };
        
        DB.justificaciones.push(newJustification);
        DB.archivos_adjuntos.push(newFile);
        
        selectedTeachers.forEach(maestroId => {
            DB.justificante_maestro.push({
                ID: 'jm_' + Math.random().toString(36).substr(2, 9),
                ID_Justificante: newJustId,
                ID_Maestro: maestroId,
                Estado_Maestro: 'Pendiente',
                Fecha_Notificacion: new Date().toISOString(),
                Fecha_Justificacion: null
            });
        });
        
        DB.observaciones.push({
            ID_Observacion: 'obs_' + Math.random().toString(36).substr(2, 9),
            ID_Justificante: newJustId,
            ID_Usuario: currentUser.ID_Usuario,
            Comentario: 'Solicitud registrada e ingresada al sistema.',
            Tipo: 'Registro',
            Fecha: new Date().toISOString()
        });
        
        saveDatabase();
        
        document.getElementById('new-justification-form').reset();
        document.getElementById('file-name-indicator').style.display = 'none';
        
        alert('¡Solicitud registrada correctamente! Queda en estado "Pendiente" en Coordinación.');
        renderAlumnoDashboard();
    };
    
    reader.readAsDataURL(file);
});

const dropzone = document.getElementById('file-dropzone');
const fileInput_field = document.getElementById('form-file');
const nameIndicator = document.getElementById('file-name-indicator');

dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.style.borderColor = 'var(--primary)';
});
dropzone.addEventListener('dragleave', () => {
    dropzone.style.borderColor = 'var(--border-color)';
});
dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.style.borderColor = 'var(--border-color)';
    if (e.dataTransfer.files.length) {
        fileInput_field.files = e.dataTransfer.files;
        updateFileNameIndicator();
    }
});
fileInput_field.addEventListener('change', updateFileNameIndicator);

function updateFileNameIndicator() {
    if (fileInput_field.files.length) {
        nameIndicator.innerText = `📄 ${fileInput_field.files[0].name} (${(fileInput_field.files[0].size / 1024 / 1024).toFixed(2)} MB)`;
        nameIndicator.style.display = 'inline-block';
    } else {
        nameIndicator.style.display = 'none';
    }
}

document.getElementById('btn-approve-request').addEventListener('click', function() {
    if (!activeRequestForReview) return;
    
    const comments = document.getElementById('coord-comments').value.trim();
    const assignedParcial = document.getElementById('review-parcial-select').value;
    
    activeRequestForReview.Estado = 'Aprobada';
    activeRequestForReview.Parcial = assignedParcial;
    activeRequestForReview.ID_Coordinador_Revisor = currentUser.ID_Usuario;
    activeRequestForReview.Fecha_Revision = new Date().toISOString();
    
    DB.observaciones.push({
        ID_Observacion: 'obs_' + Math.random().toString(36).substr(2, 9),
        ID_Justificante: activeRequestForReview.ID_Justificante,
        ID_Usuario: currentUser.ID_Usuario,
        Comentario: comments || `Aprobada por Coordinación Escolar. Asignado a: ${assignedParcial}.`,
        Tipo: 'Revision',
        Fecha: new Date().toISOString()
    });
    
    const linkedTeachers = DB.justificante_maestro.filter(jm => jm.ID_Justificante === activeRequestForReview.ID_Justificante);
    linkedTeachers.forEach(t => {
        t.Fecha_Notificacion = new Date().toISOString();
        
        DB.notificaciones.push({
            ID_Notificacion: 'not_' + Math.random().toString(36).substr(2, 9),
            ID_Usuario: t.ID_Maestro,
            ID_Justificante: activeRequestForReview.ID_Justificante,
            Mensaje: `Se ha aprobado la justificación del alumno ${DB.usuarios.find(u => u.ID_Usuario === activeRequestForReview.ID_Alumno)?.Nombre_Completo}. Por favor, confirma la recepción.`,
            Leida: 0,
            Fecha: new Date().toISOString()
        });
    });
    
    saveDatabase();
    alert('Justificación aprobada por Coordinación. Se ha notificado a los maestros para confirmar la recepción.');
    
    activeRequestForReview = null;
    renderCoordinacionDashboard();
});

document.getElementById('btn-reject-request').addEventListener('click', function() {
    if (!activeRequestForReview) return;
    
    const comments = document.getElementById('coord-comments').value.trim();
    if (!comments) {
        alert('Debes ingresar las Observaciones / Motivo de Rechazo obligatoriamente.');
        return;
    }
    
    activeRequestForReview.Estado = 'Rechazada';
    activeRequestForReview.ID_Coordinador_Revisor = currentUser.ID_Usuario;
    activeRequestForReview.Fecha_Revision = new Date().toISOString();
    
    DB.observaciones.push({
        ID_Observacion: 'obs_' + Math.random().toString(36).substr(2, 9),
        ID_Justificante: activeRequestForReview.ID_Justificante,
        ID_Usuario: currentUser.ID_Usuario,
        Comentario: comments,
        Tipo: 'Revision',
        Fecha: new Date().toISOString()
    });
    
    saveDatabase();
    alert('Justificación rechazada.');
    
    activeRequestForReview = null;
    renderCoordinacionDashboard();
});

window.confirmTeacherReceipt = function(mappingId) {
    const mapping = DB.justificante_maestro.find(jm => jm.ID === mappingId);
    if (!mapping) return;
    
    mapping.Estado_Maestro = 'Enterada por Maestro';
    mapping.Fecha_Justificacion = new Date().toISOString();
    
    const relatedMappings = DB.justificante_maestro.filter(jm => jm.ID_Justificante === mapping.ID_Justificante);
    const allConfirmed = relatedMappings.every(m => m.Estado_Maestro === 'Enterada por Maestro');
    
    if (allConfirmed) {
        const parentJust = DB.justificaciones.find(j => j.ID_Justificante === mapping.ID_Justificante);
        if (parentJust) {
            parentJust.Estado = 'Enterada por Maestro';
            
            DB.observaciones.push({
                ID_Observacion: 'obs_' + Math.random().toString(36).substr(2, 9),
                ID_Justificante: mapping.ID_Justificante,
                ID_Usuario: currentUser.ID_Usuario,
                Comentario: 'Notificación de justificante enterada y confirmada por el total de los docentes.',
                Tipo: 'Cierre',
                Fecha: new Date().toISOString()
            });
        }
    }
    
    saveDatabase();
    alert('Has confirmado la recepción del justificante. El estatus ha cambiado a "Enterada por Maestro".');
    
    renderMaestroDashboard();
};

function viewEvidenceModal(justId) {
    const j = DB.justificaciones.find(x => x.ID_Justificante === justId);
    const fileRecord = DB.archivos_adjuntos.find(a => a.ID_Justificante === justId);
    if (!j || !fileRecord) return;
    
    const modalBackdrop = document.createElement('div');
    modalBackdrop.style.position = 'fixed';
    modalBackdrop.style.top = '0';
    modalBackdrop.style.left = '0';
    modalBackdrop.style.width = '100%';
    modalBackdrop.style.height = '100%';
    modalBackdrop.style.background = 'rgba(0,0,0,0.6)';
    modalBackdrop.style.backdropFilter = 'blur(4px)';
    modalBackdrop.style.display = 'flex';
    modalBackdrop.style.alignItems = 'center';
    modalBackdrop.style.justifyContent = 'center';
    modalBackdrop.style.zIndex = '100000';
    modalBackdrop.onclick = () => document.body.removeChild(modalBackdrop);
    
    const modalContent = document.createElement('div');
    modalContent.style.background = 'var(--bg-card)';
    modalContent.style.borderRadius = 'var(--radius-lg)';
    modalContent.style.padding = '2rem';
    modalContent.style.width = '80%';
    modalContent.style.maxWidth = '750px';
    modalContent.style.height = '85%';
    modalContent.style.display = 'flex';
    modalContent.style.flexDirection = 'column';
    modalContent.onclick = (e) => e.stopPropagation();
    
    const title = document.createElement('div');
    title.className = 'card-title';
    title.style.marginBottom = '1rem';
    title.innerHTML = `
        <span>Evidencia: ${fileRecord.Nombre_Archivo}</span>
        <div style="display: flex; gap: 8px; align-items: center;">
            ${fileRecord.Ruta_Archivo.startsWith('http') ? `<a href="${fileRecord.Ruta_Archivo}" target="_blank" class="btn btn-primary btn-sm" style="text-decoration:none;">🔗 Abrir en Drive</a>` : ''}
            <button class="btn btn-secondary btn-sm" onclick="this.closest('.card-title').parentElement.parentElement.remove()">Cerrar</button>
        </div>
    `;
    modalContent.appendChild(title);
    
    const body = document.createElement('div');
    body.style.flex = '1';
    body.style.overflow = 'hidden';
    body.style.background = 'hsl(220, 20%, 94%)';
    body.style.borderRadius = 'var(--radius-md)';
    body.style.display = 'flex';
    body.style.alignItems = 'center';
    body.style.justifyContent = 'center';
    body.style.padding = '1rem';
    
    if (fileRecord.Ruta_Archivo.startsWith('http')) {
        let previewUrl = fileRecord.Ruta_Archivo;
        if (previewUrl.includes('/view')) {
            previewUrl = previewUrl.replace('/view', '/preview');
        }
        const frame = document.createElement('iframe');
        frame.src = previewUrl;
        frame.style.width = '100%';
        frame.style.height = '100%';
        frame.style.border = 'none';
        frame.style.background = 'white';
        body.appendChild(frame);
    } else if (fileRecord.Ruta_Archivo.startsWith('data:application/pdf')) {
        const frame = document.createElement('iframe');
        frame.src = fileRecord.Ruta_Archivo;
        frame.style.width = '100%';
        frame.style.height = '100%';
        frame.style.border = 'none';
        frame.style.background = 'white';
        body.appendChild(frame);
    } else {
        const img = document.createElement('img');
        img.src = fileRecord.Ruta_Archivo;
        img.style.maxWidth = '100%';
        img.style.maxHeight = '100%';
        img.style.objectFit = 'contain';
        body.appendChild(img);
    }
    
    modalContent.appendChild(body);
    modalBackdrop.appendChild(modalContent);
    document.body.appendChild(modalBackdrop);
}

function formatDateString(isoString) {
    if (!isoString) return 'N/A';
    const date = new Date(isoString + 'T00:00:00');
    return date.toLocaleDateString('es-MX', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
    });
}

function formatDateTimeString(isoString) {
    if (!isoString) return 'N/A';
    const date = new Date(isoString);
    return date.toLocaleString('es-MX', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
    });
}

document.getElementById('login-form').addEventListener('submit', function(e) {
    e.preventDefault();
    const email = document.getElementById('login-email').value.trim();
    const password = document.getElementById('login-password').value;
    
    const user = DB.usuarios.find(u => u.Correo_Electronico.toLowerCase() === email.toLowerCase());
    
    if (user && user.Contrasena === password) {
        login(user);
    } else {
        alert('Credenciales inválidas. Verifica tu correo institucional y contraseña.');
    }
});

document.getElementById('logout-button').addEventListener('click', logout);

window.toggleLoginMode = function(mode) {
    const loginFormView = document.getElementById('login-form-view');
    const changeFormView = document.getElementById('login-change-pwd-view');
    
    if (mode === 'change') {
        loginFormView.style.display = 'none';
        changeFormView.style.display = 'block';
        document.getElementById('login-change-pwd-form').reset();
    } else {
        loginFormView.style.display = 'block';
        changeFormView.style.display = 'none';
        document.getElementById('login-form').reset();
    }
};

document.getElementById('login-change-pwd-form').addEventListener('submit', function(e) {
    e.preventDefault();
    const email = document.getElementById('change-email').value.trim();
    const currentPwd = document.getElementById('change-pwd-current').value;
    const newPwd = document.getElementById('change-pwd-new').value;
    const confirmPwd = document.getElementById('change-pwd-confirm').value;
    
    const user = DB.usuarios.find(u => u.Correo_Electronico.toLowerCase() === email.toLowerCase());
    
    if (!user) {
        alert('El correo ingresado no está registrado.');
        return;
    }
    
    if (user.Contrasena !== currentPwd) {
        alert('La contraseña actual es incorrecta.');
        return;
    }
    
    if (newPwd.length < 5) {
        alert('La nueva contraseña debe tener al menos 5 caracteres.');
        return;
    }
    
    if (newPwd !== confirmPwd) {
        alert('La confirmación no coincide con la nueva contraseña.');
        return;
    }
    
    user.Contrasena = newPwd;
    saveDatabase();
    alert('Contraseña cambiada exitosamente. Ahora puedes iniciar sesión con tu nueva contraseña.');
    toggleLoginMode('login');
});

window.openPasswordModal = function() {
    document.getElementById('password-modal').style.display = 'flex';
    document.getElementById('change-password-form').reset();
};

window.closePasswordModal = function() {
    document.getElementById('password-modal').style.display = 'none';
};

document.getElementById('change-password-form').addEventListener('submit', function(e) {
    e.preventDefault();
    const currentPwd = document.getElementById('pwd-current').value;
    const newPwd = document.getElementById('pwd-new').value;
    const confirmPwd = document.getElementById('pwd-confirm').value;
    
    if (currentUser.Contrasena !== currentPwd) {
        alert('La contraseña actual es incorrecta.');
        return;
    }
    
    if (newPwd.length < 5) {
        alert('La nueva contraseña debe tener al menos 5 caracteres.');
        return;
    }
    
    if (newPwd !== confirmPwd) {
        alert('La confirmación no coincide con la nueva contraseña.');
        return;
    }
    
    currentUser.Contrasena = newPwd;
    const dbUser = DB.usuarios.find(u => u.ID_Usuario === currentUser.ID_Usuario);
    if (dbUser) {
        dbUser.Contrasena = newPwd;
    }
    
    saveDatabase();
    alert('Contraseña cambiada exitosamente.');
    closePasswordModal();
});

// INITIALIZATION
window.onload = initApp;
