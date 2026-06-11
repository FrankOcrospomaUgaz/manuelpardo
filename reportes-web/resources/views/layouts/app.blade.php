<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="csrf-token" content="{{ csrf_token() }}">
    <title>{{ $title ?? 'Sistema de Control de Personal' }}</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg: #eef2f6;
            --surface: #ffffff;
            --surface-soft: #f8fafc;
            --line: #d7dee8;
            --line-strong: #b9c3d0;
            --text: #111827;
            --muted: #64748b;
            --accent: #0f766e;
            --accent-dark: #0b5f59;
            --accent-soft: #dff5f2;
            --warn: #b42318;
            --warn-bg: #fff1ee;
            --shadow: 0 18px 40px rgba(15, 23, 42, .08);
            --sidebar-bg: #101827;
            --sidebar-hover: rgba(255,255,255,.07);
            --radius: 8px;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Inter', system-ui, -apple-system, sans-serif; font-size: 14px; color: var(--text); background: var(--bg); }
        a { color: inherit; text-decoration: none; }

        /* Shell Grid */
        .app-shell {
            display: grid;
            grid-template-columns: 286px minmax(0, 1fr);
            min-height: 100vh;
        }

        /* Sidebar */
        .sidebar {
            background: var(--sidebar-bg);
            color: #dbe5f0;
            padding: 18px 14px;
            position: sticky;
            top: 0;
            height: 100vh;
            overflow-y: auto;
            display: flex;
            flex-direction: column;
        }
        .sidebar::-webkit-scrollbar { width: 4px; }
        .sidebar::-webkit-scrollbar-thumb { background: rgba(255,255,255,.15); border-radius: 4px; }

        .brand {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 4px 8px 18px;
            border-bottom: 1px solid rgba(255,255,255,.10);
            margin-bottom: 14px;
        }
        .brand-mark {
            width: 34px;
            height: 34px;
            border-radius: 8px;
            display: grid;
            place-items: center;
            background: var(--accent);
            color: #fff;
            font-weight: 800;
        }
        .brand strong { display: block; color: #fff; }
        .brand span { display: block; font-size: 12px; color: #8fa1b6; margin-top: 2px; }

        .nav-label {
            color: #8fa1b6;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            margin: 14px 8px 8px;
        }
        .nav-list { display: grid; gap: 3px; }

        .nav-group { margin-bottom: 3px; }
        .nav-group-btn {
            display: flex;
            align-items: center;
            width: 100%;
            background: none;
            border: none;
            color: #dbe5f0;
            padding: 9px 10px;
            border-radius: 7px;
            cursor: pointer;
            font: inherit;
            font-size: 14px;
            font-weight: 600;
            gap: 10px;
            transition: all .15s;
            text-align: left;
        }
        .nav-group-btn:hover { background: var(--sidebar-hover); color: #fff; }
        .nav-group-btn .arrow { margin-left: auto; font-size: 10px; transition: transform .2s; }
        .nav-group.open .arrow { transform: rotate(90deg); }
        .nav-group-btn .icon { width: 16px; text-align: center; font-size: 14px; }

        .nav-children { max-height: 0; overflow: hidden; transition: max-height .25s ease; padding-left: 12px; }
        .nav-group.open .nav-children { max-height: 800px; }
        
        .nav-child {
            display: block;
            padding: 8px 10px 8px 24px;
            color: #8fa1b6;
            font-size: 13px;
            border-radius: 6px;
            transition: all .12s;
        }
        .nav-child:hover { background: var(--sidebar-hover); color: #fff; }
        .nav-child.active {
            background: var(--accent);
            color: #fff;
            font-weight: 600;
        }

        .nav-direct {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 9px 10px;
            color: #dbe5f0;
            border-radius: 7px;
            font-size: 14px;
            font-weight: 600;
            transition: all .15s;
        }
        .nav-direct:hover { background: var(--sidebar-hover); color: #fff; }
        .nav-direct.active {
            background: var(--accent);
            color: #fff;
        }

        .sidebar-footer { margin-top: auto; padding-top: 14px; border-top: 1px solid rgba(255,255,255,.08); }
        .user-pill { display: flex; align-items: center; gap: 10px; padding: 8px 10px; border-radius: 8px; background: rgba(255,255,255,.04); }
        .user-avatar { width: 32px; height: 32px; border-radius: 8px; background: var(--accent); display: grid; place-items: center; color: #fff; font-weight: 800; font-size: 13px; }
        .user-info { flex: 1; min-width: 0; }
        .user-info strong { display: block; color: #e2e8f0; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .user-info span { color: #64748b; font-size: 11px; }
        .logout-btn { background: none; border: none; color: #f04438; cursor: pointer; font-size: 16px; padding: 4px; border-radius: 6px; }
        .logout-btn:hover { background: rgba(240,68,56,.1); }

        /* Content Area */
        .content-area {
            min-width: 0;
            display: grid;
            grid-template-rows: auto 1fr;
            overflow: hidden;
        }

        .topbar {
            height: 64px;
            padding: 0 28px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            background: rgba(255,255,255,.82);
            border-bottom: 1px solid var(--line);
            backdrop-filter: blur(10px);
            position: sticky;
            top: 0;
            z-index: 5;
        }
        .topbar-title { font-weight: 800; color: var(--accent-dark); font-size: 16px; }
        .topbar-subtitle { color: var(--muted); font-size: 13px; margin-top: 2px; }

        main { width: 100%; max-width: 100%; padding: 28px; overflow: hidden; }

        /* Custom Styles & Utilities */
        .page-header { display: flex; align-items: flex-end; justify-content: space-between; gap: 18px; margin-bottom: 18px; }
        .page-header h1 { margin: 0; font-size: 28px; letter-spacing: 0; }
        .page-header .subtitle { color: var(--muted); font-size: 13px; margin-top: 7px; }

        .btn, button, .button {
            min-height: 42px;
            border: 1px solid var(--accent);
            border-radius: 7px;
            background: var(--accent);
            color: #fff;
            padding: 10px 15px;
            text-decoration: none;
            cursor: pointer;
            font-weight: 800;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            font-size: 13px;
            transition: all .15s;
        }
        .btn:hover, button:hover, .button:hover { background: var(--accent-dark); border-color: var(--accent-dark); }
        
        .btn-secondary, .button.secondary {
            background: #fff;
            color: var(--accent-dark);
            border-color: var(--line-strong);
        }
        .btn-secondary:hover, .button.secondary:hover { border-color: var(--accent); background: var(--accent-soft); }
        
        .btn-danger { background: var(--warn); border-color: var(--warn); }
        .btn-danger:hover { background: #9b1c1c; border-color: #9b1c1c; }
        
        .btn-sm { min-height: 32px; padding: 6px 12px; font-size: 12px; }

        .card, .panel {
            background: var(--surface);
            border: 1px solid var(--line);
            border-radius: 8px;
            box-shadow: var(--shadow);
        }
        .card { padding: 17px; }
        .panel { padding: 18px; margin-bottom: 18px; }
        .filter-panel { border-top: 3px solid var(--accent); }
        .metric-strip { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 18px; }
        .metric { background: var(--surface); border: 1px solid var(--line); border-radius: 8px; padding: 14px; box-shadow: 0 8px 24px rgba(15,23,42,.04); }
        .metric span { color: var(--muted); font-size: 12px; }
        .metric strong { display: block; margin-top: 5px; font-size: 20px; }

        /* Legacy Forms & Inputs */
        form.filters { display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 16px; align-items: end; }
        label { display: grid; gap: 7px; color: #475569; font-size: 12px; font-weight: 700; }
        input, select { width: 100%; border: 1px solid var(--line-strong); border-radius: 7px; padding: 10px 11px; color: var(--text); background: #fff; min-height: 42px; font: inherit; }
        select[multiple] { min-height: 136px; padding: 6px; }
        select[multiple] option { padding: 8px 10px; border-radius: 5px; margin-bottom: 3px; }
        input:focus, select:focus { outline: 3px solid var(--accent-soft); border-color: var(--accent); }
        .field-hint { color: var(--muted); font-size: 12px; font-weight: 500; }
        .checkbox { display: flex; align-items: center; gap: 8px; min-height: 42px; }
        .checkbox input { width: auto; min-height: auto; }

        /* Forms (Modern Layouts) */
        .form-card { padding: 24px; }
        .form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 18px; }
        .form-group { display: flex; flex-direction: column; gap: 6px; }
        .form-group.full { grid-column: 1 / -1; }
        .form-label { font-size: 12px; font-weight: 700; color: #475569; }
        .form-label .required { color: var(--warn); }
        .form-control { width: 100%; border: 1px solid var(--line-strong); border-radius: 7px; padding: 10px 11px; font: inherit; font-size: 13px; background: #fff; transition: all .15s; min-height: 42px; }
        .form-control:focus { outline: 3px solid var(--accent-soft); border-color: var(--accent); }
        .form-hint { font-size: 11px; color: var(--muted); }
        .form-actions { display: flex; gap: 10px; margin-top: 24px; padding-top: 18px; border-top: 1px solid var(--line); }

        /* Tables (Legacy & Modern) */
        .table-wrap { overflow: auto; border: 1px solid var(--line); border-radius: 8px; background: #fff; box-shadow: var(--shadow); max-width: 100%; }
        table { width: 100%; border-collapse: collapse; min-width: 900px; }
        th, td { border-bottom: 1px solid var(--line); padding: 10px 12px; text-align: left; white-space: nowrap; }
        th { background: #e9f5f3; font-size: 12px; color: #12423e; text-transform: uppercase; position: sticky; top: 0; z-index: 1; }
        tr:hover td { background: #fbfefe; }
        tr:last-child td { border-bottom: 0; }

        .table-card { overflow: hidden; }
        .table-toolbar { padding: 16px 20px; display: flex; align-items: center; gap: 12px; border-bottom: 1px solid var(--line); flex-wrap: wrap; background: #fff; }
        .table-search { flex: 1; min-width: 200px; position: relative; }
        .table-search input { width: 100%; border: 1px solid var(--line-strong); border-radius: 7px; padding: 9px 12px 9px 36px; font: inherit; background: #fff; }
        .table-search input:focus { outline: none; border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }
        .table-search::before { content: '🔍'; position: absolute; left: 12px; top: 50%; transform: translateY(-50%); font-size: 13px; }
        .table-count { color: var(--muted); font-size: 12px; font-weight: 600; }

        .data-table { width: 100%; border-collapse: collapse; min-width: 900px; }
        .data-table th { background: #e9f5f3; padding: 10px 12px; text-align: left; font-size: 12px; font-weight: 700; color: #12423e; text-transform: uppercase; border-bottom: 1px solid var(--line); position: sticky; top: 0; z-index: 1; }
        .data-table td { padding: 10px 12px; border-bottom: 1px solid var(--line); font-size: 13px; white-space: nowrap; }
        .data-table tr:hover td { background: #fbfefe; }
        .data-table tr:last-child td { border-bottom: none; }
        .table-actions { display: flex; gap: 6px; }

        .badge { display: inline-flex; padding: 3px 10px; border-radius: 20px; font-size: 11px; font-weight: 700; }
        .badge-success { background: var(--accent-soft); color: var(--accent-dark); }
        .badge-danger { background: var(--warn-bg); color: var(--warn); }
        .badge-warn { background: #fffbeb; color: #b45309; }

        /* Modal */
        .modal-overlay { position: fixed; inset: 0; background: rgba(15,23,42,.5); backdrop-filter: blur(4px); z-index: 100; display: none; align-items: center; justify-content: center; }
        .modal-overlay.show { display: flex; }
        .modal { background: #fff; border-radius: 10px; box-shadow: var(--shadow); width: 90%; max-width: 440px; overflow: hidden; }
        .modal-header { padding: 18px 20px; border-bottom: 1px solid var(--line); font-weight: 700; display: flex; justify-content: space-between; align-items: center; }
        .modal-body { padding: 20px; }
        .modal-footer { padding: 14px 20px; border-top: 1px solid var(--line); display: flex; gap: 10px; justify-content: flex-end; }
        .modal-close { background: none; border: none; font-size: 20px; cursor: pointer; color: var(--muted); }

        /* Alerts */
        .alert { padding: 12px 16px; border-radius: 8px; font-size: 13px; margin-bottom: 16px; display: flex; align-items: center; gap: 10px; border: 1px solid #f04438; background: var(--warn-bg); color: var(--warn); }
        .alert-success { background: var(--accent-soft); color: var(--accent-dark); border-color: var(--accent); }

        /* Empty state */
        .empty-state { text-align: center; padding: 48px 20px; color: var(--muted); background: var(--surface); border: 1px dashed var(--line-strong); border-radius: 8px; }
        
        /* Responsive */
        @media (max-width: 1024px) {
            .app-shell { grid-template-columns: 1fr; }
            .sidebar { position: relative; height: auto; display: block; }
            .nav-list { grid-template-columns: repeat(auto-fill, minmax(210px, 1fr)); }
            .topbar { position: static; }
            main { padding: 18px; }
        }
    </style>
</head>
<body>
    <div class="app-shell">
        {{-- Sidebar --}}
        <aside class="sidebar" id="sidebar">
            <div class="brand">
                <div class="brand-mark">CP</div>
                <div>
                    <strong>Control de Personal</strong>
                    <span>Huella Digital v3.0</span>
                </div>
            </div>

            <nav class="nav-list">
                {{-- Dashboard --}}
                <a class="nav-direct {{ request()->is('/') ? 'active' : '' }}" href="{{ url('/') }}">
                    <span class="icon">📊</span> Inicio
                </a>

                {{-- Mantenimientos --}}
                <div class="nav-label">Administración</div>
                <div class="nav-group {{ request()->is('mantenimiento/*') ? 'open' : '' }}">
                    <button class="nav-group-btn" onclick="this.parentElement.classList.toggle('open')">
                        <span class="icon">⚙️</span> Mantenimientos <span class="arrow">▶</span>
                    </button>
                    <div class="nav-children">
                        <a class="nav-child {{ request()->is('mantenimiento/areas*') ? 'active' : '' }}" href="{{ url('mantenimiento/areas') }}">Áreas</a>
                        <a class="nav-child {{ request()->is('mantenimiento/cargos*') ? 'active' : '' }}" href="{{ url('mantenimiento/cargos') }}">Cargos</a>
                        <a class="nav-child {{ request()->is('mantenimiento/categorias*') ? 'active' : '' }}" href="{{ url('mantenimiento/categorias') }}">Categorías</a>
                        <a class="nav-child {{ request()->is('mantenimiento/tipos-hora*') ? 'active' : '' }}" href="{{ url('mantenimiento/tipos-hora') }}">Tipos de Hora</a>
                        <a class="nav-child {{ request()->is('mantenimiento/motivos*') ? 'active' : '' }}" href="{{ url('mantenimiento/motivos') }}">Motivos Justificación</a>
                        <a class="nav-child {{ request()->is('mantenimiento/feriados*') ? 'active' : '' }}" href="{{ url('mantenimiento/feriados') }}">Feriados</a>
                        <a class="nav-child {{ request()->is('mantenimiento/grupos*') ? 'active' : '' }}" href="{{ url('mantenimiento/grupos') }}">Grupos</a>
                        <a class="nav-child {{ request()->is('mantenimiento/trabajadores*') ? 'active' : '' }}" href="{{ url('mantenimiento/trabajadores') }}">Trabajadores</a>
                        <a class="nav-child {{ request()->is('mantenimiento/usuarios*') ? 'active' : '' }}" href="{{ url('mantenimiento/usuarios') }}">Usuarios</a>
                    </div>
                </div>

                {{-- Movimientos --}}
                <div class="nav-group {{ request()->is('movimiento/*') ? 'open' : '' }}">
                    <button class="nav-group-btn" onclick="this.parentElement.classList.toggle('open')">
                        <span class="icon">📋</span> Movimientos <span class="arrow">▶</span>
                    </button>
                    <div class="nav-children">
                        <a class="nav-child {{ request()->is('movimiento/horarios*') ? 'active' : '' }}" href="{{ url('movimiento/horarios') }}">Horarios por Trabajador</a>
                        <a class="nav-child {{ request()->is('movimiento/justificaciones*') ? 'active' : '' }}" href="{{ url('movimiento/justificaciones') }}">Justificaciones</a>
                        <a class="nav-child {{ request()->is('movimiento/suspensiones*') ? 'active' : '' }}" href="{{ url('movimiento/suspensiones') }}">Suspensiones</a>
                        <a class="nav-child {{ request()->is('movimiento/permisos*') ? 'active' : '' }}" href="{{ url('movimiento/permisos') }}">Permisos</a>
                        <a class="nav-child {{ request()->is('movimiento/marcacion*') ? 'active' : '' }}" href="{{ url('movimiento/marcacion') }}">Marcación</a>
                        <a class="nav-child {{ request()->is('movimiento/consolidado*') ? 'active' : '' }}" href="{{ url('movimiento/consolidado') }}">Consolidado Mensual</a>
                    </div>
                </div>

                {{-- Reportes --}}
                <div class="nav-group {{ request()->is('reportes*') ? 'open' : '' }}">
                    <button class="nav-group-btn" onclick="this.parentElement.classList.toggle('open')">
                        <span class="icon">📈</span> Reportes <span class="arrow">▶</span>
                    </button>
                    <div class="nav-children">
                        @foreach (config('reports') as $rSlug => $rData)
                            <a class="nav-child {{ request()->is('reportes/'.$rSlug) ? 'active' : '' }}" href="{{ route('reports.show', $rSlug) }}">{{ $rData['title'] }}</a>
                        @endforeach
                    </div>
                </div>
            </nav>

            {{-- User Footer --}}
            <div class="sidebar-footer">
                <div class="user-pill">
                    <div class="user-avatar">{{ strtoupper(substr(session('usuario_nombre', 'U'), 0, 1)) }}</div>
                    <div class="user-info">
                        <strong>{{ session('usuario_nombre', 'Usuario') }}</strong>
                        <span>Conectado</span>
                    </div>
                    <form method="POST" action="{{ url('logout') }}" style="display:inline">
                        @csrf
                        <button type="submit" class="logout-btn" title="Cerrar sesión">✕</button>
                    </form>
                </div>
            </div>
        </aside>

        {{-- Content Area --}}
        <div class="content-area">
            <header class="topbar">
                <div>
                    <div class="topbar-title">@yield('topbar-title', 'Sistema de Control de Personal')</div>
                    <div class="topbar-subtitle">@yield('topbar-subtitle', 'Panel de administración y control')</div>
                </div>
                <div>
                    @yield('topbar-actions')
                </div>
            </header>
            <main>
                @if (session('success'))
                    <div class="alert alert-success">✅ {{ session('success') }}</div>
                @endif
                @if (session('error'))
                    <div class="alert">❌ {{ session('error') }}</div>
                @endif
                @yield('content')
            </main>
        </div>
    </div>

    {{-- Delete confirmation modal --}}
    <div class="modal-overlay" id="deleteModal">
        <div class="modal">
            <div class="modal-header">
                Confirmar eliminación
                <button class="modal-close" onclick="closeDeleteModal()">×</button>
            </div>
            <div class="modal-body">
                <p>¿Está seguro que desea eliminar este registro? Esta acción no se puede deshacer.</p>
            </div>
            <div class="modal-footer">
                <button class="btn btn-secondary btn-sm" onclick="closeDeleteModal()">Cancelar</button>
                <form id="deleteForm" method="POST" style="display:inline">
                    @csrf @method('DELETE')
                    <button type="submit" class="btn btn-danger btn-sm">Eliminar</button>
                </form>
            </div>
        </div>
    </div>

    <script>
        function confirmDelete(url) {
            document.getElementById('deleteForm').action = url;
            document.getElementById('deleteModal').classList.add('show');
        }
        function closeDeleteModal() {
            document.getElementById('deleteModal').classList.remove('show');
        }
        // Table search filtering
        document.addEventListener('DOMContentLoaded', function() {
            const searchInput = document.getElementById('searchInput');
            if (searchInput) {
                searchInput.addEventListener('input', function() {
                    const term = this.value.toLowerCase();
                    document.querySelectorAll('.data-table tbody tr').forEach(row => {
                        row.style.display = row.textContent.toLowerCase().includes(term) ? '' : 'none';
                    });
                });
            }
        });
    </script>
    @stack('scripts')
</body>
</html>
