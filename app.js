// Justifaltas - Core Application JavaScript (Excel Sincronizado, Límites, Contraseñas y Soporte macOS/file://)

// Google Sheets Integration Configuration
// Coloca aquí la URL de la Web App obtenida al implementar tu Google Apps Script.
// Si está vacía, el sistema operará en "Modo Local" (usando localStorage).
const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwo-m3rR6vx2yQVuEcbKbIM_tooOW5uGD2ffrMdYMoNNCDfEvSd8njEGLU98FDkbVU0/exec";

// Override native window.alert with custom premium HTML modal
window.alert = function(message) {
    if (document.activeElement) {
        try {
            document.activeElement.blur();
        } catch(e) {}
    }
    const modal = document.getElementById('custom-alert-modal');
    const msgEl = document.getElementById('custom-alert-message');
    const iconEl = document.getElementById('custom-alert-icon');
    if (!modal || !msgEl) {
        console.warn("Custom alert modal not found. Falling back to console:", message);
        return;
    }
    
    // Auto-detect warning vs success based on message content
    if (message.toLowerCase().includes('éxito') || message.toLowerCase().includes('correctamente') || message.toLowerCase().includes('exitosamente') || message.toLowerCase().includes('aprobada') || message.toLowerCase().includes('aprobado') || message.toLowerCase().includes('confirmada')) {
        iconEl.innerText = '✔️';
    } else {
        iconEl.innerText = '⚠️';
    }
    
    msgEl.innerText = message;
    modal.style.display = 'flex';
};

window.closeCustomAlert = function() {
    const modal = document.getElementById('custom-alert-modal');
    if (modal) {
        modal.style.display = 'none';
    }
};

async function hashPassword(password) {
    const encoder = new TextEncoder();
    const data = encoder.encode(password);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
    return hashHex;
}

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
        { ID_Usuario: 'Uguzman', Correo_Electronico: 'ulisega03s24a@uinteramericana.edu.mx', Nombre_Completo: 'Ulise Guzman Alvaro', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: '2026-05-10T00:00:00Z', Activo: 1 },
        { ID_Usuario: 'Syanet', Correo_Electronico: 'syanethh03s24a@uinteramericana.edu.mx', Nombre_Completo: 'Sandra Yanet Hernández Hernández', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: '2026-05-11T00:00:00Z', Activo: 1 },
        { ID_Usuario: 'Ihernandez', Correo_Electronico: 'irenehh03s24a@uinteramericana.edu.mx', Nombre_Completo: 'Irene Hernández Hernández', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 'Yulila', Correo_Electronico: 'iyuleidals03s24a@uinteramericana.edu.mx', Nombre_Completo: 'Itzel Yuleida Lila Sanchez', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 'Dantonio', Correo_Electronico: 'dantonioie03s24a@uinteramericana.edu.mx', Nombre_Completo: 'Diego Antonio Iracheta Escareño', Rol: 'alumno', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 'Mangel', Correo_Electronico: 'mdavila0311@uinteramericana.edu.mx', Nombre_Completo: 'Miguel Angel Davila Calzada', Rol: 'maestro', Contrasena: '12345', Fecha_Registro: '2026-05-01T00:00:00Z', Activo: 1 },
        { ID_Usuario: 'Aeloisa', Correo_Electronico: 'ana.hernandez03ea26@uinteramericana.edu.mx', Nombre_Completo: 'Ana Eloisa Hernández Jimenez', Rol: 'maestro', Contrasena: '12345', Fecha_Registro: '2026-05-01T00:00:00Z', Activo: 1 },
        { ID_Usuario: 'Oalavez', Correo_Electronico: 'oscar.alavez03ea26p@uinteramericana.edu.mx', Nombre_Completo: 'Oscar Alavez Reyes', Rol: 'maestro', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 'Asimon', Correo_Electronico: 'andres.simon03ea26@uinteramericana.edu.mx', Nombre_Completo: 'Andres Simón Treviño', Rol: 'maestro', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 'SanSilva', Correo_Electronico: 'direccionsc_sd25o@uinteramericana.edu.mx', Nombre_Completo: 'Lic. Sandra Silva', Rol: 'coordinacion', Contrasena: '12345', Fecha_Registro: null, Activo: 1 },
        { ID_Usuario: 'EscSanta', Correo_Electronico: 'escolarsc@universidadinteramericana.edu.mx', Nombre_Completo: 'Escolar Santa Catarina', Rol: 'coordinacion', Contrasena: '12345', Fecha_Registro: null, Activo: 1 }
    ],
    justificaciones: [
        {
            ID_Justificante: 'j_demo_001',
            ID_Alumno: 'Yulila',
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
            ID_Maestro: 'Mangel',
            Estado_Maestro: 'Pendiente',
            Fecha_Notificacion: '2026-06-14T10:05:00Z',
            Fecha_Justificacion: null
        }
    ],
    observaciones: [
        {
            ID_Observacion: 'obs_demo_001',
            ID_Justificante: 'j_demo_001',
            ID_Usuario: 'SanSilva',
            Comentario: 'Pendiente de revisión por coordinación.',
            Tipo: 'Revision',
            Fecha: '2026-06-14T10:10:00Z'
        }
    ],
    notificaciones: [
        {
            ID_Notificacion: 'n_demo_001',
            ID_Usuario: 'Mangel',
            ID_Justificante: 'j_demo_001',
            Message: 'Nueva solicitud pendiente de aprobación por coordinación.',
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
        console.warn("localStorage is restricted. Falling back to in-memory database.");
        DB = JSON.parse(JSON.stringify(officialDB));
    }
    
    // Garantizar que la base de datos y sus tablas no sean nulas o indefinidas (Esquema Defensivo)
    if (!DB || typeof DB !== 'object') {
        DB = JSON.parse(JSON.stringify(officialDB));
    }
    if (!DB.usuarios) DB.usuarios = [];
    if (!DB.justificaciones) DB.justificaciones = [];
    if (!DB.archivos_adjuntos) DB.archivos_adjuntos = [];
    if (!DB.justificante_maestro) DB.justificante_maestro = [];
    if (!DB.observaciones) DB.observaciones = [];
    if (!DB.notificaciones) DB.notificaciones = [];
    
    // Sincronizar usuarios nuevos del código con la base de datos de caché local (localStorage)
    let updated = false;
    officialDB.usuarios.forEach(offU => {
        const exists = DB.usuarios.some(u => u.Correo_Electronico.toLowerCase() === offU.Correo_Electronico.toLowerCase());
        if (!exists) {
            DB.usuarios.push(offU);
            updated = true;
        }
    });
    
    if (updated) {
        try {
            localStorage.setItem(DB_KEY, JSON.stringify(DB));
        } catch(e) {}
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
        
        if (resData.success && resData.data && resData.data.usuarios && resData.data.usuarios.length > 0) {
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

window.manualSyncDatabase = function() {
    if (GOOGLE_SCRIPT_URL) {
        syncDatabaseFromSheets().then(() => {
            alert("Sincronización con Google Sheets completada de forma exitosa.");
        }).catch(() => {
            alert("No se pudo conectar con Google Sheets. Se mantendrán los datos locales temporales.");
        });
    } else {
        alert("El portal está configurado en Modo Local sin URL de Google Sheets conectada.");
    }
};

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
                data: DB,
                callerRole: currentUser ? getNormalizedRole(currentUser.Rol) : 'guest',
                callerId: currentUser ? currentUser.ID_Usuario : ''
            }),
            headers: {
                'Content-Type': 'text/plain'
            }
        })
        .then(response => response.json())
        .then(resData => {
            if (resData.success) {
                showSyncStatus('sincronizado', 'Sincronizado con Google Sheets');
                if (resData.data) {
                    DB = resData.data;
                    try {
                        localStorage.setItem(DB_KEY, JSON.stringify(DB));
                    } catch(e) {}
                }
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
function getNormalizedRole(role) {
    if (!role) return '';
    const r = role.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    if (r.includes('alumno') || r.includes('estudiante')) {
        return 'alumno';
    }
    if (r.includes('maestro') || r.includes('docente') || r.includes('profesor')) {
        return 'maestro';
    }
    if (r.includes('coordinac') || r.includes('direc') || r.includes('admin')) {
        return 'coordinacion';
    }
    return r;
}

function normalizeTetra(tetraStr) {
    if (!tetraStr) return '';
    return tetraStr.replace(/–/g, '-').replace(/\s+/g, ' ').trim().toLowerCase();
}

let currentUser = null;
let currentRole = null;
let activeRequestForReview = null;

async function initApp() {
    const savedTheme = localStorage.getItem('selected-theme') || 'uin-blue';
    applySelectedTheme(savedTheme);
    
    loadDatabase();
    
    // Iniciar sincronización de forma prioritaria con Google Sheets antes de iniciar sesion o renderizar
    if (GOOGLE_SCRIPT_URL) {
        try {
            await syncDatabaseFromSheets();
        } catch (e) {
            console.warn("Error prioritario de sincronizacion en inicio:", e);
        }
    } else {
        showSyncStatus('local', 'Modo Local');
    }
    
    // Default Tetra value initialization based on current system month
    const currentMonth = new Date().getMonth() + 1; // 1-12
    let initialTetra = "Mayo – Agosto";
    if (currentMonth >= 1 && currentMonth <= 4) initialTetra = "Enero – Abril";
    else if (currentMonth >= 9 && currentMonth <= 12) initialTetra = "Septiembre – Diciembre";
    
    const filterSelect = document.getElementById('global-tetra-filter');
    if (filterSelect) {
        filterSelect.value = initialTetra;
        
        // Listener for Global Tetra Filter
        filterSelect.addEventListener('change', () => {
            if (currentRole === 'alumno') renderAlumnoDashboard();
            else if (currentRole === 'coordinacion') renderCoordinacionDashboard();
            else if (currentRole === 'maestro') renderMaestroDashboard();
        });
    }

    let savedSession = null;
    try {
        savedSession = localStorage.getItem('justifaltas_session_secured');
    } catch (e) {
        savedSession = inMemorySessionEmail;
    }

    if (savedSession) {
        const sessionUser = DB.usuarios.find(u => u.Correo_Electronico.toLowerCase() === savedSession.toLowerCase());
        if (sessionUser) {
            login(sessionUser);
            return;
        }
    }
    
    // Periodic synchronization from Google Sheets in the background every 45 seconds (only when a user is logged in)
    setInterval(async () => {
        if (currentUser && GOOGLE_SCRIPT_URL) {
            try {
                await syncDatabaseFromSheets();
            } catch (e) {
                console.warn("Background sync error:", e);
            }
        }
    }, 45000);
    
    showScreen('login');
}

function login(user) {
    currentUser = user;
    currentRole = getNormalizedRole(user.Rol);
    
    resetInactivityTimer();
    
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
    
    // Clear inactivity timer
    if (inactivityTimer) {
        clearTimeout(inactivityTimer);
        inactivityTimer = null;
    }
    
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
    
    if (tabId === 'panel-gestion-usuarios') {
        renderManagementUsers();
    }
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
            <li><a class="menu-item" data-target="panel-gestion-usuarios" onclick="switchTab('panel-gestion-usuarios')">👥 Gestión de Usuarios</a></li>
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
    document.getElementById('welcome-title').innerHTML = `¡Hola, <span class="text-bold">${currentUser.Nombre_Completo}</span>!`;
    document.getElementById('welcome-subtitle').innerText = 'Registra tus faltas y consulta el estatus de tus justificantes.';
    document.getElementById('student-history-title').innerText = `Mis Solicitudes (${activeTetra})`;
    
    const studentId = currentUser.ID_Usuario;
    
    const studentRequests = DB.justificaciones.filter(j => j.ID_Alumno === studentId && normalizeTetra(j.Periodo_Tetra) === normalizeTetra(activeTetra));
    
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
    DB.usuarios.filter(u => getNormalizedRole(u.Rol) === 'maestro').forEach(m => {
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
            <td data-label="ID"><small style="color:var(--text-muted); font-family:monospace;">${j.ID_Justificante}</small></td>
            <td data-label="Fecha Falta"><strong>${formatDateString(j.Fecha_Falta)}</strong></td>
            <td data-label="Parcial"><span class="badge badge-approved" style="background: var(--primary-light); color: var(--primary); border: 1px solid var(--border-color); font-weight: 600;">${j.Parcial || 'Sin Asignar'}</span></td>
            <td data-label="Motivo">${j.Motivo}</td>
            <td data-label="Docentes Asignados">
                ${teachersStatusHTML}
                ${reviewLogHTML}
            </td>
            <td data-label="Evidencia">
                <a href="#" class="evidence-link" onclick="viewEvidenceModal('${j.ID_Justificante}')">📎 Ver Archivo</a>
            </td>
            <td data-label="Estado / Seguimiento">
                <span class="badge ${badgeClass}">${statusLabel}</span>
                ${j.Estado === 'Pendiente' ? `<button class="btn btn-danger btn-sm" style="padding: 2px 6px; font-size: 0.72rem; margin-top: 6px; display: block;" onclick="deleteStudentRequest('${j.ID_Justificante}')">🗑️ Eliminar</button>` : ''}
                ${obsText ? `<br><small style="color: var(--text-muted); font-style: italic;">Obs: ${obsText}</small>` : ''}
            </td>
        `;
        tableBody.appendChild(row);
    });
}

// 6. COORDINACION DASHBOARD LOGIC
function renderCoordinacionDashboard() {
    const activeTetra = document.getElementById('global-tetra-filter').value;
    document.getElementById('welcome-title').innerHTML = `Panel de <span class="text-bold">Coordinación</span>`;
    document.getElementById('welcome-subtitle').innerText = 'Portal de validación y control escolar.';
    
    const myJustifications = DB.justificaciones.filter(j => normalizeTetra(j.Periodo_Tetra) === normalizeTetra(activeTetra));
    
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
            <td data-label="Alumno">
                <strong>${studentUser ? studentUser.Nombre_Completo : 'Alumno'}</strong><br>
                <small style="color: var(--text-muted);">${studentUser ? studentUser.Correo_Electronico : ''}</small>
            </td>
            <td data-label="Fecha Falta">${formatDateString(j.Fecha_Falta)}</td>
            <td data-label="Motivo"><span class="badge badge-pending">${j.Motivo}</span></td>
            <td data-label="Acción"><button class="btn btn-primary btn-sm">Evaluar</button></td>
        `;
        tableBody.appendChild(row);
    });
    
    if (activeRequestForReview) {
        const req = DB.justificaciones.find(j => j.ID_Justificante === activeRequestForReview.ID_Justificante);
        if (req && req.Estado === 'Pendiente' && normalizeTetra(req.Periodo_Tetra) === normalizeTetra(activeTetra)) {
            selectRequestForReview(req.ID_Justificante);
        } else {
            activeRequestForReview = null;
            document.getElementById('coord-review-details').style.display = 'none';
            document.getElementById('coord-empty-details').style.display = 'block';
            resetEvidenceViewer();
        }
    } else {
        document.getElementById('coord-review-details').style.display = 'none';
        document.getElementById('coord-empty-details').style.display = 'block';
        resetEvidenceViewer();
    }
}

function selectRequestForReview(reqId) {
    const req = DB.justificaciones.find(j => j.ID_Justificante === reqId);
    if (!req) return;
    
    activeRequestForReview = req;
    const studentUser = DB.usuarios.find(u => u.ID_Usuario === req.ID_Alumno);
    
    const detailsPanel = document.getElementById('coord-review-details');
    detailsPanel.setAttribute('data-request-id', reqId);
    document.getElementById('coord-empty-details').style.display = 'none';
    detailsPanel.style.display = 'block';
    
    document.getElementById('review-student').innerText = studentUser ? studentUser.Nombre_Completo : 'N/A';
    document.getElementById('review-matricula').innerText = studentUser ? studentUser.ID_Usuario : 'N/A';
    document.getElementById('review-correo').innerText = studentUser ? studentUser.Correo_Electronico : 'N/A';
    document.getElementById('review-date').innerText = formatDateString(req.Fecha_Falta);
    document.getElementById('review-reason').innerText = req.Motivo;
    document.getElementById('review-description').innerText = req.Descripcion;
    document.getElementById('coord-comments').value = '';
    
    const parcialSelect = document.getElementById('review-parcial-select');
    parcialSelect.value = req.Parcial || getTetraAndParcial(req.Fecha_Falta).parcial || 'Parcial 1';
    
    const tetraSelect = document.getElementById('review-tetra-select');
    if (tetraSelect) {
        tetraSelect.value = req.Periodo_Tetra || getTetraAndParcial(req.Fecha_Falta).tetra || 'Mayo – Agosto';
    }
    
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
    document.getElementById('welcome-title').innerHTML = `Panel del <span class="text-bold">Docente</span>`;
    document.getElementById('welcome-subtitle').innerText = 'Revisa notificaciones de justificantes escolares y confirma su recepción.';
    
    const teacherId = currentUser.ID_Usuario;
    const myMappings = DB.justificante_maestro.filter(jm => jm.ID_Maestro === teacherId);
    
    const myJustifications = [];
    myMappings.forEach(m => {
        const j = DB.justificaciones.find(x => x.ID_Justificante === m.ID_Justificante);
        if (j && normalizeTetra(j.Periodo_Tetra) === normalizeTetra(activeTetra) && j.Estado !== 'Rechazada') {
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
        
        let statusBadge = '';
        let actionBtn = '';
        
        if (j.Estado === 'Pendiente') {
            statusBadge = `<span class="badge badge-pending" style="background: rgba(255, 159, 10, 0.12); color: var(--warning); border: 1px solid rgba(255, 159, 10, 0.2); font-weight: 600;">Pendiente Escolar</span>`;
            actionBtn = `<span style="color: var(--warning); font-weight: 500; font-size: 0.8rem;">Esperando Aprobación</span>`;
        } else if (m.Estado_Maestro === 'Enterada por Maestro') {
            statusBadge = `<span class="badge badge-finalized">Confirmada</span>`;
            actionBtn = `<span style="color:var(--success); font-weight:600; font-size:0.82rem;">Confirmado: ${formatDateTimeString(m.Fecha_Justificacion)}</span>`;
        } else {
            statusBadge = `<span class="badge badge-approved">Pendiente Confirmar</span>`;
            actionBtn = `<button class="btn btn-success btn-sm" onclick="confirmTeacherReceipt('${m.ID}')">📋 Confirmar recepción</button>`;
        }
        
        const coordObs = DB.observaciones.filter(o => o.ID_Justificante === j.ID_Justificante && o.Tipo === 'Revision').reverse();
        const coordComment = j.Estado === 'Pendiente' ? 'En espera de revisión por Coordinación Escolar' : (coordObs.length > 0 ? coordObs[0].Comentario : 'Aprobado por Coordinación');
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td data-label="Alumno"><strong>${studentUser ? studentUser.Nombre_Completo : 'Alumno'}</strong></td>
            <td data-label="Correo"><small>${studentUser ? studentUser.Correo_Electronico : 'N/A'}</small></td>
            <td data-label="Fecha Falta"><strong>${formatDateString(j.Fecha_Falta)}</strong></td>
            <td data-label="Parcial"><span class="badge badge-approved" style="background: var(--primary-light); color: var(--primary); border: 1px solid var(--border-color); font-weight: 600;">${j.Parcial}</span></td>
            <td data-label="Obs. Coordinación" style="font-size: 0.85rem; font-style: italic; color: var(--primary);">
                "${coordComment}"
            </td>
            <td data-label="Evidencia">
                <a href="#" class="evidence-link" onclick="viewEvidenceModal('${j.ID_Justificante}')">📎 Evidencia</a>
            </td>
            <td data-label="Estado">${statusBadge}</td>
            <td data-label="Acción">${actionBtn}</td>
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
let selectedDates = [];

window.removeSelectedDate = function(dateVal) {
    selectedDates = selectedDates.filter(d => d !== dateVal);
    renderSelectedDates();
};

function renderSelectedDates() {
    const container = document.getElementById('selected-dates-container');
    if (!container) return;
    container.innerHTML = '';
    
    selectedDates.sort().forEach(dateVal => {
        const pill = document.createElement('span');
        pill.style.background = 'var(--primary)';
        pill.style.color = '#fff';
        pill.style.padding = '0.4rem 0.8rem';
        pill.style.borderRadius = '20px';
        pill.style.fontSize = '0.8rem';
        pill.style.fontWeight = '600';
        pill.style.display = 'inline-flex';
        pill.style.alignItems = 'center';
        pill.style.gap = '6px';
        pill.innerHTML = `
            <span>${formatDateString(dateVal)}</span>
            <span style="cursor:pointer; font-weight:800; font-size: 1.1rem; line-height: 1; padding: 2px 8px; margin-left: 2px; display: inline-block;" onclick="removeSelectedDate('${dateVal}')">×</span>
        `;
        container.appendChild(pill);
    });
    
    validateMultipleDates();
}

function validateMultipleDates() {
    const timeLimitWarning = document.getElementById('form-time-limit-warning');
    const limitWarning = document.getElementById('form-parcial-limit-warning');
    const submitBtn = document.getElementById('form-submit-btn');
    const calculatedTetra = document.getElementById('form-calculated-tetra');
    const calculatedParcial = document.getElementById('form-calculated-parcial');
    
    if (selectedDates.length === 0) {
        timeLimitWarning.style.display = 'none';
        limitWarning.style.display = 'none';
        calculatedTetra.value = '';
        calculatedParcial.value = '';
        submitBtn.disabled = false;
        submitBtn.style.opacity = '1';
        submitBtn.style.cursor = 'pointer';
        return;
    }
    
    let tetra = calculatedTetra.value;
    let parcial = calculatedParcial.value;
    
    // Auto-compute if values are not set
    if (!tetra || !parcial) {
        const firstDate = selectedDates[0];
        const computed = getTetraAndParcial(firstDate);
        if (!tetra) {
            tetra = computed.tetra;
            calculatedTetra.value = tetra;
        }
        if (!parcial) {
            parcial = computed.parcial;
            calculatedParcial.value = parcial;
        }
    }
    
    let anyExpired = false;
    const now = new Date();
    selectedDates.forEach(d => {
        const parts = d.split('-');
        if (parts.length >= 3) {
            const year = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const day = parseInt(parts[2], 10);
            const absenceDate = new Date(year, month, day, 23, 59, 59);
            
            const diffMs = now.getTime() - absenceDate.getTime();
            const diffHours = diffMs / (1000 * 60 * 60);
            if (diffHours > 24) {
                anyExpired = true;
            }
        }
    });
    
    const studentId = currentUser.ID_Usuario;
    const existingCount = DB.justificaciones.filter(j => 
        j.ID_Alumno === studentId && 
        j.Estado !== 'Rechazada' &&
        j.Periodo_Tetra === tetra &&
        j.Parcial === parcial
    ).length;
    
    let isLimitReached = existingCount >= 2;
    
    timeLimitWarning.style.display = anyExpired ? 'block' : 'none';
    limitWarning.style.display = isLimitReached ? 'block' : 'none';
    
    if (anyExpired || isLimitReached) {
        submitBtn.disabled = true;
        submitBtn.style.opacity = '0.5';
        submitBtn.style.cursor = 'not-allowed';
    } else {
        submitBtn.disabled = false;
        submitBtn.style.opacity = '1';
        submitBtn.style.cursor = 'pointer';
    }
}

function parseManualDate(str) {
    str = str.trim();
    // Regex for DD/MM/YYYY or DD-MM-YYYY
    let regex1 = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/;
    let match1 = str.match(regex1);
    if (match1) {
        let day = parseInt(match1[1], 10);
        let month = parseInt(match1[2], 10);
        let year = parseInt(match1[3], 10);
        
        if (month < 1 || month > 12) return null;
        let maxDays = new Date(year, month, 0).getDate();
        if (day < 1 || day > maxDays) return null;
        
        return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
    
    // Regex for YYYY-MM-DD
    let regex2 = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/;
    let match2 = str.match(regex2);
    if (match2) {
        let year = parseInt(match2[1], 10);
        let month = parseInt(match2[2], 10);
        let day = parseInt(match2[3], 10);
        
        if (month < 1 || month > 12) return null;
        let maxDays = new Date(year, month, 0).getDate();
        if (day < 1 || day > maxDays) return null;
        
        return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
    
    return null;
}

function handleDateChange() {
    const dateInput = document.getElementById('form-date');
    const inputVal = dateInput.value.trim();
    if (!inputVal) return;
    
    const parsedDate = parseManualDate(inputVal);
    if (!parsedDate) {
        alert('Formato de fecha inválido. Por favor, ingresa una fecha válida en formato DD/MM/AAAA (ejemplo: 12/07/2026).');
        return;
    }
    
    if (!selectedDates.includes(parsedDate)) {
        selectedDates.push(parsedDate);
        renderSelectedDates();
    } else {
        alert('Esta fecha ya ha sido seleccionada.');
    }
    dateInput.value = '';
}

// Bind manual date handlers
document.getElementById('btn-add-date').addEventListener('click', handleDateChange);
document.getElementById('form-date').addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault(); // Prevent submitting the whole form
        handleDateChange();
    }
});

// Bind Periodo and Parcial dropdown changes to re-run limit validations
document.getElementById('form-calculated-tetra').addEventListener('change', validateMultipleDates);
document.getElementById('form-calculated-parcial').addEventListener('change', validateMultipleDates);

// Bind reason dropdown changes to toggle manual custom reason field
document.getElementById('form-reason').addEventListener('change', function() {
    const otherContainer = document.getElementById('form-reason-other-container');
    if (this.value === 'Otro') {
        otherContainer.style.display = 'block';
        document.getElementById('form-reason-other').focus();
    } else {
        otherContainer.style.display = 'none';
        document.getElementById('form-reason-other').value = '';
    }
});

// 9. NEW JUSTIFICATION SUBMISSION (Alumno)
function showFormAlert(type, message) {
    if (document.activeElement) {
        try {
            document.activeElement.blur();
        } catch(e) {}
    }
    const errorDiv = document.getElementById('form-error-message');
    const successDiv = document.getElementById('form-success-message');
    if (!errorDiv || !successDiv) return;
    
    errorDiv.style.display = 'none';
    successDiv.style.display = 'none';
    
    if (type === 'error') {
        errorDiv.innerHTML = '⚠️ ' + message;
        errorDiv.style.display = 'block';
        errorDiv.scrollIntoView({ behavior: 'smooth', block: 'center' });
    } else if (type === 'success') {
        successDiv.innerHTML = '✔️ ' + message;
        successDiv.style.display = 'block';
        successDiv.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}

function setSubmitButtonLoading(isLoading) {
    const submitBtn = document.getElementById('form-submit-btn');
    if (!submitBtn) return;
    if (isLoading) {
        submitBtn.disabled = true;
        submitBtn.style.opacity = '0.5';
        submitBtn.style.cursor = 'not-allowed';
        submitBtn.setAttribute('data-original-text', submitBtn.innerText);
        submitBtn.innerText = 'Enviando solicitud...';
    } else {
        submitBtn.disabled = false;
        submitBtn.style.opacity = '1';
        submitBtn.style.cursor = 'pointer';
        const origText = submitBtn.getAttribute('data-original-text') || 'Enviar Solicitud';
        submitBtn.innerText = origText;
    }
}

// 9. NEW JUSTIFICATION SUBMISSION (Alumno)
document.getElementById('new-justification-form').addEventListener('submit', function(e) {
    e.preventDefault();
    console.log("Iniciando envío de justificación...");
    
    // Clear any previous alerts
    const errorDiv = document.getElementById('form-error-message');
    const successDiv = document.getElementById('form-success-message');
    if (errorDiv) errorDiv.style.display = 'none';
    if (successDiv) successDiv.style.display = 'none';
    
    try {
        if (selectedDates.length === 0) {
            showFormAlert('error', 'Debes seleccionar al menos una fecha en el calendario.');
            return;
        }
        
        const date = selectedDates.join(', ');
        let reason = document.getElementById('form-reason').value;
        const desc = document.getElementById('form-description').value;
        
        if (!reason) {
            showFormAlert('error', 'Por favor, selecciona un motivo de la lista.');
            return;
        }
        
        if (reason === 'Otro') {
            const customReason = document.getElementById('form-reason-other').value.trim();
            if (!customReason) {
                showFormAlert('error', 'Por favor, escribe el motivo personalizado de tu falta.');
                return;
            }
            reason = customReason;
        }
        
        if (!desc.trim()) {
            showFormAlert('error', 'Por favor, escribe una descripción detallada del motivo.');
            return;
        }
        
        const teacherSelect = document.getElementById('form-teachers');
        const selectedTeachers = [];
        for (let i = 0; i < teacherSelect.options.length; i++) {
            if (teacherSelect.options[i].selected) {
                selectedTeachers.push(teacherSelect.options[i].value);
            }
        }
        
        if (selectedTeachers.length === 0) {
            showFormAlert('error', 'Debes seleccionar al menos un maestro de la lista.');
            return;
        }
        
        const fileInput = document.getElementById('form-file');
        const file = fileInput.files[0];
        
        if (!file) {
            showFormAlert('error', 'Debe cargar un archivo de evidencia física obligatoriamente.');
            return;
        }
        
        const tetra = document.getElementById('form-calculated-tetra').value;
        const parcial = document.getElementById('form-calculated-parcial').value;
        
        if (!tetra) {
            showFormAlert('error', 'Por favor, selecciona un Periodo (Tetra) para la falta.');
            return;
        }
        
        if (!parcial) {
            showFormAlert('error', 'Por favor, selecciona un Parcial para la falta.');
            return;
        }
        
        if (!DB.justificaciones) {
            DB.justificaciones = [];
        }
        const existingCount = DB.justificaciones.filter(j => 
            j.ID_Alumno === currentUser.ID_Usuario && 
            j.Estado !== 'Rechazada' &&
            j.Periodo_Tetra === tetra &&
            j.Parcial === parcial
        ).length;
        
        if (existingCount >= 2) {
            showFormAlert('error', `Límite alcanzado: Ya tienes registradas 2 justificaciones en el ${parcial} del periodo ${tetra}.`);
            return;
        }
        
        // Disable submit button during upload process to prevent double submission
        setSubmitButtonLoading(true);
        
        const reader = new FileReader();
        reader.onerror = function() {
            showFormAlert('error', 'Error al leer el archivo de evidencia física.');
            setSubmitButtonLoading(false);
        };
        reader.onload = async function(event) {
            try {
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
                            showFormAlert('error', 'Advertencia: No se pudo subir a Google Drive, pero se guardará de forma local.');
                        }
                    } catch (err) {
                        console.error("Error upload_file fetch:", err);
                        showFormAlert('error', 'Advertencia: Error de red con Google Drive, guardando de forma local.');
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
                
                if (!DB.justificaciones) DB.justificaciones = [];
                if (!DB.archivos_adjuntos) DB.archivos_adjuntos = [];
                if (!DB.justificante_maestro) DB.justificante_maestro = [];
                if (!DB.observaciones) DB.observaciones = [];
                
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
                selectedDates = [];
                renderSelectedDates();
                
                showFormAlert('success', '¡Solicitud registrada correctamente! Queda en estado "Pendiente" en Coordinación.');
                renderAlumnoDashboard();
            } catch (innerError) {
                showFormAlert('error', 'Error interno al procesar el envío: ' + innerError.message);
                console.error(innerError);
            } finally {
                // Re-enable button after process completion
                setSubmitButtonLoading(false);
            }
        };
        
        reader.readAsDataURL(file);
    } catch (error) {
        showFormAlert('error', 'Error al iniciar el envío: ' + error.message);
        console.error(error);
        setSubmitButtonLoading(false);
    }
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
    const detailsPanel = document.getElementById('coord-review-details');
    const reqId = detailsPanel.getAttribute('data-request-id');
    const req = DB.justificaciones.find(j => j.ID_Justificante === reqId);
    if (!req) {
        alert("Error: No hay ninguna solicitud seleccionada.");
        return;
    }
    
    const comments = document.getElementById('coord-comments').value.trim();
    const assignedParcial = document.getElementById('review-parcial-select').value;
    const assignedTetra = document.getElementById('review-tetra-select').value;
    
    if (!assignedTetra) {
        alert('Error: Por favor, selecciona un periodo (tetra) para la justificación.');
        return;
    }
    
    if (!assignedParcial) {
        alert('Error: Por favor, selecciona un parcial para la justificación.');
        return;
    }
    
    req.Estado = 'Aprobada';
    req.Periodo_Tetra = assignedTetra;
    req.Parcial = assignedParcial;
    req.ID_Coordinador_Revisor = currentUser.ID_Usuario;
    req.Fecha_Revision = new Date().toISOString();
    
    DB.observaciones.push({
        ID_Observacion: 'obs_' + Math.random().toString(36).substr(2, 9),
        ID_Justificante: req.ID_Justificante,
        ID_Usuario: currentUser.ID_Usuario,
        Comentario: comments || `Aprobada por Coordinación Escolar. Asignado a: ${assignedParcial}.`,
        Tipo: 'Revision',
        Fecha: new Date().toISOString()
    });
    
    const linkedTeachers = DB.justificante_maestro.filter(jm => jm.ID_Justificante === req.ID_Justificante);
    linkedTeachers.forEach(t => {
        t.Fecha_Notificacion = new Date().toISOString();
        
        DB.notificaciones.push({
            ID_Notificacion: 'not_' + Math.random().toString(36).substr(2, 9),
            ID_Usuario: t.ID_Maestro,
            ID_Justificante: req.ID_Justificante,
            Mensaje: `Se ha aprobado la justificación del alumno ${DB.usuarios.find(u => u.ID_Usuario === req.ID_Alumno)?.Nombre_Completo}. Por favor, confirma la recepción.`,
            Leida: 0,
            Fecha: new Date().toISOString()
        });
    });
    
    saveDatabase();
    alert('Justificación aprobada por Coordinación. Se ha notificado a los maestros para confirmar la recepción.');
    
    activeRequestForReview = null;
    detailsPanel.removeAttribute('data-request-id');
    renderCoordinacionDashboard();
});

document.getElementById('btn-reject-request').addEventListener('click', function() {
    const detailsPanel = document.getElementById('coord-review-details');
    const reqId = detailsPanel.getAttribute('data-request-id');
    const req = DB.justificaciones.find(j => j.ID_Justificante === reqId);
    if (!req) {
        alert("Error: No hay ninguna solicitud seleccionada.");
        return;
    }
    
    const comments = document.getElementById('coord-comments').value.trim();
    if (!comments) {
        alert('Debes ingresar las Observaciones / Motivo de Rechazo obligatoriamente.');
        return;
    }
    
    req.Estado = 'Rechazada';
    req.ID_Coordinador_Revisor = currentUser.ID_Usuario;
    req.Fecha_Revision = new Date().toISOString();
    
    DB.observaciones.push({
        ID_Observacion: 'obs_' + Math.random().toString(36).substr(2, 9),
        ID_Justificante: req.ID_Justificante,
        ID_Usuario: currentUser.ID_Usuario,
        Comentario: comments,
        Tipo: 'Revision',
        Fecha: new Date().toISOString()
    });
    
    saveDatabase();
    alert('Justificación rechazada.');
    
    activeRequestForReview = null;
    detailsPanel.removeAttribute('data-request-id');
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
    modalContent.style.padding = window.innerWidth <= 480 ? '1.25rem' : '2rem';
    modalContent.style.width = window.innerWidth <= 480 ? '92%' : '80%';
    modalContent.style.maxWidth = '750px';
    modalContent.style.height = window.innerWidth <= 480 ? '80%' : '85%';
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
    if (isoString.includes(',')) {
        return isoString.split(',').map(s => {
            const parts = s.trim().split('-');
            if (parts.length < 3) return 'N/A';
            const year = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const day = parseInt(parts[2], 10);
            const date = new Date(year, month, day);
            return date.toLocaleDateString('es-MX', {
                month: 'short',
                day: 'numeric'
            });
        }).join(', ');
    }
    const parts = isoString.split('-');
    if (parts.length < 3) return 'N/A';
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    const date = new Date(year, month, day);
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

document.getElementById('login-form').addEventListener('submit', async function(e) {
    e.preventDefault();
    const email = document.getElementById('login-email').value.trim();
    const password = document.getElementById('login-password').value;
    
    if (!email || !password) {
        alert('Por favor, ingresa tu correo y contraseña.');
        return;
    }
    
    const user = DB.usuarios.find(u => u.Correo_Electronico.toLowerCase() === email.toLowerCase());
    
    if (user) {
        const isSha256 = /^[a-fA-F0-9]{64}$/.test(user.Contrasena);
        let isMatch = false;
        
        if (isSha256) {
            const hashedInput = await hashPassword(password);
            isMatch = (user.Contrasena === hashedInput);
        } else {
            isMatch = (String(user.Contrasena) === String(password));
            if (isMatch) {
                console.log("Auto-cifrando contraseña heredada a SHA-256...");
                const hashedPwd = await hashPassword(password);
                user.Contrasena = hashedPwd;
                saveDatabase();
            }
        }
        
        if (isMatch) {
            login(user);
        } else {
            alert('Credenciales inválidas. Verifica tu correo institucional y contraseña.');
        }
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

document.getElementById('login-change-pwd-form').addEventListener('submit', async function(e) {
    e.preventDefault();
    const email = document.getElementById('change-email').value.trim();
    const currentPwd = document.getElementById('change-pwd-current').value;
    const newPwd = document.getElementById('change-pwd-new').value;
    const confirmPwd = document.getElementById('change-pwd-confirm').value;
    
    if (!email || !currentPwd || !newPwd || !confirmPwd) {
        alert('Por favor, completa todos los campos.');
        return;
    }
    
    const user = DB.usuarios.find(u => u.Correo_Electronico.toLowerCase() === email.toLowerCase());
    
    if (!user) {
        alert('El correo ingresado no está registrado.');
        return;
    }
    
    const isSha256 = /^[a-fA-F0-9]{64}$/.test(user.Contrasena);
    let isCurrentPwdCorrect = false;
    if (isSha256) {
        const hashedCurrent = await hashPassword(currentPwd);
        isCurrentPwdCorrect = (user.Contrasena === hashedCurrent);
    } else {
        isCurrentPwdCorrect = (String(user.Contrasena) === String(currentPwd));
    }
    
    if (!isCurrentPwdCorrect) {
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
    
    const hashedNew = await hashPassword(newPwd);
    user.Contrasena = hashedNew;
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

document.getElementById('change-password-form').addEventListener('submit', async function(e) {
    e.preventDefault();
    const currentPwd = document.getElementById('pwd-current').value;
    const newPwd = document.getElementById('pwd-new').value;
    const confirmPwd = document.getElementById('pwd-confirm').value;
    
    if (!currentPwd || !newPwd || !confirmPwd) {
        alert('Por favor, completa todos los campos obligatorios.');
        return;
    }
    
    const isSha256 = /^[a-fA-F0-9]{64}$/.test(currentUser.Contrasena);
    let isCurrentPwdCorrect = false;
    if (isSha256) {
        const hashedCurrent = await hashPassword(currentPwd);
        isCurrentPwdCorrect = (currentUser.Contrasena === hashedCurrent);
    } else {
        isCurrentPwdCorrect = (String(currentUser.Contrasena) === String(currentPwd));
    }
    
    if (!isCurrentPwdCorrect) {
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
    
    const hashedNew = await hashPassword(newPwd);
    currentUser.Contrasena = hashedNew;
    const dbUser = DB.usuarios.find(u => u.ID_Usuario === currentUser.ID_Usuario);
    if (dbUser) {
        dbUser.Contrasena = hashedNew;
    }
    
    saveDatabase();
    alert('Contraseña cambiada exitosamente.');
    closePasswordModal();
});

window.resetLocalCache = function() {
    if (confirm("¿Estás seguro de que deseas restablecer la base de datos local de pruebas? Esto cerrará tu sesión, borrará la caché local del navegador y cargará los datos de fábrica (incluyendo los nuevos usuarios).")) {
        localStorage.clear();
        location.reload();
    }
};

window.deleteRequest = function(reqId) {
    // 1. Eliminar de justificaciones
    DB.justificaciones = DB.justificaciones.filter(j => j.ID_Justificante !== reqId);
    
    // 2. Eliminar de justificante_maestro
    DB.justificante_maestro = DB.justificante_maestro.filter(jm => jm.ID_Justificante !== reqId);
    
    // 3. Eliminar de archivos_adjuntos
    DB.archivos_adjuntos = DB.archivos_adjuntos.filter(a => a.ID_Justificante !== reqId);
    
    // 4. Eliminar de observaciones
    DB.observaciones = DB.observaciones.filter(o => o.ID_Justificante !== reqId);
    
    // 5. Eliminar de notificaciones
    DB.notificaciones = DB.notificaciones.filter(n => n.ID_Justificante !== reqId);
    
    saveDatabase();
    
    alert("La solicitud ha sido eliminada permanentemente de la base de datos.");
    
    if (currentRole === 'alumno') {
        renderAlumnoDashboard();
    } else if (currentRole === 'coordinacion') {
        activeRequestForReview = null;
        const detailsPanel = document.getElementById('coord-review-details');
        if (detailsPanel) detailsPanel.removeAttribute('data-request-id');
        renderCoordinacionDashboard();
    }
};

window.deleteStudentRequest = function(reqId) {
    if (confirm("¿Estás seguro de que deseas eliminar permanentemente esta solicitud? Esta acción no se puede deshacer y borrará la evidencia cargada.")) {
        window.deleteRequest(reqId);
    }
};

window.deleteCoordRequest = function() {
    const detailsPanel = document.getElementById('coord-review-details');
    if (!detailsPanel) return;
    const reqId = detailsPanel.getAttribute('data-request-id');
    if (!reqId) {
        alert("Error: No hay ninguna solicitud seleccionada para eliminar.");
        return;
    }
    if (confirm("¿Estás seguro de que deseas eliminar permanentemente esta solicitud de la base de datos general? Esta acción no se puede deshacer y afectará los registros del alumno y docentes.")) {
        window.deleteRequest(reqId);
    }
};

window.applySelectedTheme = function(themeName) {
    const root = document.documentElement;
    localStorage.setItem('selected-theme', themeName);
    
    // Reset body background inline styles
    document.body.style.backgroundImage = '';
    document.body.style.backgroundColor = '';
    
    if (themeName === 'uin-blue') {
        // UIN Blue & Cosmic Dark Theme (Default)
        root.style.setProperty('--primary', 'hsl(210, 100%, 50%)'); // UIN Blue
        root.style.setProperty('--primary-light', 'rgba(0, 113, 227, 0.12)');
        root.style.setProperty('--primary-dark', 'hsl(210, 100%, 42%)');
        root.style.setProperty('--btn-text', '#ffffff');
        root.style.setProperty('--btn-shadow', 'rgba(0, 113, 227, 0.25)');
        root.style.setProperty('--bg-main', '#000000');
        root.style.setProperty('--bg-card', 'rgba(22, 22, 23, 0.65)');
        root.style.setProperty('--bg-sidebar', 'rgba(18, 18, 19, 0.85)');
        root.style.setProperty('--border-color', 'rgba(0, 113, 227, 0.12)');
        root.style.setProperty('--text-main', 'hsl(0, 0%, 96%)');
        root.style.setProperty('--text-muted', 'hsl(0, 0%, 65%)');
        
        document.body.style.setProperty('background-image', 
            'radial-gradient(circle at 85% 85%, rgba(0, 113, 227, 0.08), transparent 50%), ' +
            'radial-gradient(circle at 15% 15%, rgba(138, 43, 226, 0.05), transparent 45%), ' +
            'linear-gradient(rgba(255, 255, 255, 0.012) 1px, transparent 1px), ' +
            'linear-gradient(90deg, rgba(255, 255, 255, 0.012) 1px, transparent 1px)', 'important');
            
    } else if (themeName === 'aurora') {
        // Neon Aurora Boreal Dark Theme
        root.style.setProperty('--primary', 'hsl(280, 100%, 65%)');
        root.style.setProperty('--primary-light', 'rgba(168, 85, 247, 0.15)');
        root.style.setProperty('--primary-dark', 'hsl(280, 100%, 55%)');
        root.style.setProperty('--btn-text', '#000000');
        root.style.setProperty('--btn-shadow', 'rgba(168, 85, 247, 0.25)');
        root.style.setProperty('--bg-main', '#04020a');
        root.style.setProperty('--bg-card', 'rgba(15, 12, 25, 0.65)');
        root.style.setProperty('--bg-sidebar', 'rgba(10, 8, 16, 0.85)');
        root.style.setProperty('--border-color', 'rgba(168, 85, 247, 0.12)');
        root.style.setProperty('--text-main', 'hsl(250, 40%, 96%)');
        root.style.setProperty('--text-muted', 'hsl(250, 20%, 65%)');
        
        document.body.style.setProperty('background-image', 
            'radial-gradient(circle at 85% 85%, rgba(6, 182, 212, 0.08), transparent 50%), ' +
            'radial-gradient(circle at 15% 15%, rgba(168, 85, 247, 0.08), transparent 45%), ' +
            'linear-gradient(rgba(168, 85, 247, 0.015) 1px, transparent 1px), ' +
            'linear-gradient(90deg, rgba(168, 85, 247, 0.015) 1px, transparent 1px)', 'important');
            
    } else if (themeName === 'emerald') {
        // Emerald Glass Dark Theme
        root.style.setProperty('--primary', 'hsl(150, 80%, 45%)');
        root.style.setProperty('--primary-light', 'rgba(16, 185, 129, 0.15)');
        root.style.setProperty('--primary-dark', 'hsl(150, 80%, 38%)');
        root.style.setProperty('--btn-text', '#000000');
        root.style.setProperty('--btn-shadow', 'rgba(16, 185, 129, 0.25)');
        root.style.setProperty('--bg-main', '#010603');
        root.style.setProperty('--bg-card', 'rgba(10, 22, 15, 0.65)');
        root.style.setProperty('--bg-sidebar', 'rgba(6, 16, 10, 0.85)');
        root.style.setProperty('--border-color', 'rgba(16, 185, 129, 0.12)');
        root.style.setProperty('--text-main', 'hsl(140, 30%, 96%)');
        root.style.setProperty('--text-muted', 'hsl(140, 15%, 65%)');
        
        document.body.style.setProperty('background-image', 
            'radial-gradient(circle at 85% 85%, rgba(52, 211, 153, 0.07), transparent 50%), ' +
            'radial-gradient(circle at 15% 15%, rgba(16, 185, 129, 0.05), transparent 45%), ' +
            'linear-gradient(rgba(16, 185, 129, 0.012) 1px, transparent 1px), ' +
            'linear-gradient(90deg, rgba(16, 185, 129, 0.012) 1px, transparent 1px)', 'important');
            
    } else if (themeName === 'light') {
        // Premium Apple Light Mode (Claro)
        root.style.setProperty('--primary', 'hsl(210, 100%, 50%)');
        root.style.setProperty('--primary-light', 'rgba(0, 122, 255, 0.1)');
        root.style.setProperty('--primary-dark', 'hsl(210, 100%, 45%)');
        root.style.setProperty('--btn-text', '#ffffff');
        root.style.setProperty('--btn-shadow', 'rgba(0, 122, 255, 0.25)');
        root.style.setProperty('--bg-main', '#f5f5f7');
        root.style.setProperty('--bg-card', 'rgba(255, 255, 255, 0.8)');
        root.style.setProperty('--bg-sidebar', 'rgba(245, 245, 247, 0.9)');
        root.style.setProperty('--border-color', 'rgba(0, 0, 0, 0.06)');
        root.style.setProperty('--text-main', '#1d1d1f');
        root.style.setProperty('--text-muted', '#86868b');
        
        document.body.style.setProperty('background-image', 
            'radial-gradient(circle at 85% 85%, rgba(0, 122, 255, 0.03), transparent 50%), ' +
            'radial-gradient(circle at 15% 15%, rgba(168, 85, 247, 0.03), transparent 45%), ' +
            'linear-gradient(rgba(0, 0, 0, 0.01) 1px, transparent 1px), ' +
            'linear-gradient(90deg, rgba(0, 0, 0, 0.01) 1px, transparent 1px)', 'important');
    }
    
    // Update metric blue backgrounds to align with theme accent
    const metricBlue = document.querySelectorAll('.metric-icon.blue');
    metricBlue.forEach(el => {
        el.style.backgroundColor = 'var(--primary-light)';
        el.style.color = 'var(--primary)';
    });
    
    const selector = document.getElementById('theme-selector');
    if (selector) selector.value = themeName;
};

// Inactivity Session Timeout Manager (5 Minutes)
let inactivityTimer = null;
const INACTIVITY_TIMEOUT = 5 * 60 * 1000; // 5 minutes in milliseconds

window.resetInactivityTimer = function() {
    if (inactivityTimer) {
        clearTimeout(inactivityTimer);
    }
    // Only arm timer if a user is logged in
    if (currentUser) {
        inactivityTimer = setTimeout(autoLogout, INACTIVITY_TIMEOUT);
    }
};

function autoLogout() {
    if (currentUser) {
        logout();
        alert("Tu sesión ha expirado por inactividad de 5 minutos por seguridad.");
    }
}

// Bind user activity events to keep session alive
['click', 'mousemove', 'keypress', 'scroll', 'touchstart'].forEach(eventName => {
    document.addEventListener(eventName, resetInactivityTimer, { passive: true });
});

// ================= USER DIRECTORY MANAGEMENT (Coordinación Only) =================
window.renderManagementUsers = function() {
    const tableBody = document.getElementById('management-users-table');
    if (!tableBody) return;
    tableBody.innerHTML = '';
    
    if (!DB || !DB.usuarios) {
        tableBody.innerHTML = `<tr><td colspan="4" class="empty-state">No hay usuarios cargados en la base de datos.</td></tr>`;
        return;
    }
    
    // Sort users safely handling null, undefined or non-string values
    const sortedUsers = [...DB.usuarios].filter(u => u && u.Nombre_Completo && u.Correo_Electronico).sort((a, b) => {
        const rolA = String(a.Rol || '').toLowerCase().trim();
        const rolB = String(b.Rol || '').toLowerCase().trim();
        if (rolA !== rolB) return rolA.localeCompare(rolB);
        
        const nameA = String(a.Nombre_Completo || '').trim();
        const nameB = String(b.Nombre_Completo || '').trim();
        return nameA.localeCompare(nameB);
    });
    
    sortedUsers.forEach(u => {
        const row = document.createElement('tr');
        
        let roleBadge = '';
        const role = getNormalizedRole(u.Rol);
        if (role === 'coordinacion') {
            roleBadge = `<span class="badge badge-pending" style="background: rgba(255, 159, 10, 0.12); color: var(--warning);">Coordinador</span>`;
        } else if (role === 'maestro') {
            roleBadge = `<span class="badge badge-approved" style="background: var(--primary-light); color: var(--primary);">Docente</span>`;
        } else {
            roleBadge = `<span class="badge badge-approved" style="background: rgba(48, 209, 88, 0.12); color: var(--success);">Alumno</span>`;
        }
        
        const isSelf = currentUser && currentUser.ID_Usuario === u.ID_Usuario;
        const deleteBtn = isSelf 
            ? `<span style="font-size: 0.8rem; color: var(--text-muted); font-style: italic;">Sesión Activa</span>`
            : `<button class="btn btn-danger" style="padding: 0.3rem 0.75rem; font-size: 0.75rem;" onclick="deleteUserRecord('${u.ID_Usuario}')">Eliminar</button>`;
            
        row.innerHTML = `
            <td data-label="Nombre"><strong>${u.Nombre_Completo}</strong></td>
            <td data-label="Correo"><small>${u.Correo_Electronico}</small></td>
            <td data-label="Rol">${roleBadge}</td>
            <td data-label="Acciones">${deleteBtn}</td>
        `;
        tableBody.appendChild(row);
    });
};

window.deleteUserRecord = function(userId) {
    if (confirm("¿Estás seguro de que deseas eliminar permanentemente a este usuario de la base de datos de la universidad?")) {
        DB.usuarios = DB.usuarios.filter(u => u.ID_Usuario !== userId);
        
        // Clean up mappings and notifications from local DB representation as well
        if (DB.justificante_maestro) {
            DB.justificante_maestro = DB.justificante_maestro.filter(jm => jm.ID_Maestro !== userId);
        }
        if (DB.notificaciones) {
            DB.notificaciones = DB.notificaciones.filter(n => n.ID_Usuario !== userId);
        }
        
        // Save local changes
        try {
            localStorage.setItem(DB_KEY, JSON.stringify(DB));
        } catch(e) {}
        
        // Request explicit deletion to Apps Script
        if (GOOGLE_SCRIPT_URL) {
            showSyncStatus('sincronizando', 'Eliminando usuario...');
            fetch(GOOGLE_SCRIPT_URL, {
                method: 'POST',
                body: JSON.stringify({
                    action: 'delete_user',
                    userId: userId
                }),
                headers: {
                    'Content-Type': 'text/plain'
                }
            })
            .then(response => response.json())
            .then(resData => {
                if (resData.success) {
                    showSyncStatus('sincronizado', 'Sincronizado con Google Sheets');
                    if (resData.data) {
                        DB = resData.data;
                        try {
                            localStorage.setItem(DB_KEY, JSON.stringify(DB));
                        } catch(e) {}
                        renderManagementUsers();
                    }
                } else {
                    showSyncStatus('error', 'Error al eliminar usuario');
                }
            })
            .catch(err => {
                showSyncStatus('error', 'Error de red al eliminar');
            });
        }
        
        renderManagementUsers();
        alert("El usuario ha sido eliminado exitosamente de la base de datos.");
    }
};

document.getElementById('add-user-form').addEventListener('submit', async function(e) {
    e.preventDefault();
    const name = document.getElementById('add-user-name').value.trim();
    const email = document.getElementById('add-user-email').value.trim();
    const role = document.getElementById('add-user-role').value;
    const pwd = document.getElementById('add-user-pwd').value.trim() || '12345';
    
    if (!name || !email) {
        alert("Por favor, completa el nombre y el correo institucional.");
        return;
    }
    
    if (!email.endsWith('@uinteramericana.edu.mx') && !email.endsWith('@universidadinteramericana.edu.mx')) {
        alert("El correo electrónico debe ser una cuenta institucional de la Universidad Interamericana (@uinteramericana.edu.mx o @universidadinteramericana.edu.mx).");
        return;
    }
    
    const exists = DB.usuarios.some(u => u.Correo_Electronico.toLowerCase() === email.toLowerCase());
    if (exists) {
        alert("Error: Este correo institucional ya se encuentra registrado.");
        return;
    }
    
    const hashedPwd = await hashPassword(pwd);
    
    const newUser = {
        ID_Usuario: 'usr_' + Math.random().toString(36).substr(2, 9),
        Correo_Electronico: email,
        Nombre_Completo: name,
        Rol: role === 'coordinacion' ? 'coordinacion' : (role === 'maestro' ? 'maestro' : 'alumno'),
        Contrasena: hashedPwd,
        Fecha_Registro: new Date().toISOString(),
        Activo: 1
    };
    
    DB.usuarios.push(newUser);
    saveDatabase();
    
    // Reset Form
    document.getElementById('add-user-form').reset();
    document.getElementById('add-user-pwd').value = '12345';
    
    renderManagementUsers();
    alert(`Usuario registrado exitosamente en la base de datos.\nNombre: ${name}\nRol: ${newUser.Rol.toUpperCase()}`);
});

// INITIALIZATION
window.onload = initApp;
