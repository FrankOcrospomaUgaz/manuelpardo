<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{{ $title ?? 'Reportes' }}</title>
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
        }
        * { box-sizing: border-box; }
        body {
            margin: 0;
            color: var(--text);
            background: var(--bg);
            font-family: "Segoe UI", Arial, Helvetica, sans-serif;
            font-size: 14px;
            overflow-x: hidden;
        }
        a { color: inherit; }
        .app-shell {
            display: grid;
            grid-template-columns: 286px minmax(0, 1fr);
            min-height: 100vh;
        }
        .sidebar {
            background: #101827;
            color: #dbe5f0;
            padding: 18px 14px;
            position: sticky;
            top: 0;
            height: 100vh;
            overflow: auto;
        }
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
        .nav-item {
            display: block;
            text-decoration: none;
            color: #dbe5f0;
            border-radius: 7px;
            padding: 9px 10px;
            line-height: 1.2;
            border: 1px solid transparent;
        }
        .nav-item:hover { background: rgba(255,255,255,.07); }
        .nav-item.active {
            background: var(--accent);
            color: #fff;
            box-shadow: 0 10px 24px rgba(15, 118, 110, .28);
        }
        .content {
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
        .topbar-title { font-weight: 800; color: var(--accent-dark); }
        .topbar-subtitle { color: var(--muted); font-size: 13px; }
        main { width: 100%; max-width: 100%; padding: 28px; overflow: hidden; }
        .page-head {
            display: flex;
            align-items: flex-end;
            justify-content: space-between;
            gap: 18px;
            margin-bottom: 18px;
        }
        h1 { margin: 0; font-size: 28px; letter-spacing: 0; }
        .eyebrow {
            color: var(--muted);
            margin: 7px 0 0;
            font-size: 13px;
        }
        .muted { color: var(--muted); }
        .metric-strip {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
            gap: 12px;
            margin-bottom: 18px;
        }
        .metric {
            background: var(--surface);
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 14px;
            box-shadow: 0 8px 24px rgba(15,23,42,.04);
        }
        .metric span { color: var(--muted); font-size: 12px; }
        .metric strong { display: block; margin-top: 5px; font-size: 20px; }
        .grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
            gap: 14px;
        }
        .card, .panel {
            background: var(--surface);
            border: 1px solid var(--line);
            border-radius: 8px;
            box-shadow: var(--shadow);
        }
        .card {
            padding: 17px;
            display: block;
            text-decoration: none;
            min-height: 104px;
        }
        .card:hover { border-color: var(--accent); transform: translateY(-1px); }
        .card strong { display: block; font-size: 16px; margin-bottom: 8px; }
        .card span { display: block; font-size: 12px; line-height: 1.4; }
        .panel { padding: 18px; margin-bottom: 18px; }
        .filter-panel { border-top: 3px solid var(--accent); }
        form.filters {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
            gap: 16px;
            align-items: end;
        }
        label { display: grid; gap: 7px; color: #475569; font-size: 12px; font-weight: 700; }
        input, select {
            width: 100%;
            border: 1px solid var(--line-strong);
            border-radius: 7px;
            padding: 10px 11px;
            color: var(--text);
            background: #fff;
            min-height: 42px;
            font: inherit;
        }
        select[multiple] {
            min-height: 136px;
            padding: 6px;
        }
        select[multiple] option {
            padding: 8px 10px;
            border-radius: 5px;
            margin-bottom: 3px;
        }
        input:focus, select:focus {
            outline: 3px solid var(--accent-soft);
            border-color: var(--accent);
        }
        .field-hint { color: var(--muted); font-size: 12px; font-weight: 500; }
        .checkbox { display: flex; align-items: center; gap: 8px; min-height: 42px; }
        .checkbox input { width: auto; min-height: auto; }
        .actions {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            align-items: end;
        }
        button, .button {
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
        }
        .button.secondary {
            background: #fff;
            color: var(--accent-dark);
            border-color: var(--line-strong);
        }
        .button.secondary:hover { border-color: var(--accent); background: var(--accent-soft); }
        .file-icon {
            width: 22px;
            height: 22px;
            border-radius: 5px;
            display: inline-grid;
            place-items: center;
            color: #fff;
            font-size: 12px;
            font-weight: 900;
        }
        .file-icon.excel { background: #15803d; }
        .file-icon.pdf { background: #dc2626; }
        .file-icon.csv { background: #2563eb; }
        .table-wrap {
            overflow: auto;
            border: 1px solid var(--line);
            border-radius: 8px;
            background: #fff;
            box-shadow: var(--shadow);
            max-width: 100%;
        }
        table { width: 100%; border-collapse: collapse; min-width: 900px; }
        th, td {
            border-bottom: 1px solid var(--line);
            padding: 10px 12px;
            text-align: left;
            white-space: nowrap;
        }
        th {
            background: #e9f5f3;
            font-size: 12px;
            color: #12423e;
            text-transform: uppercase;
            position: sticky;
            top: 0;
            z-index: 1;
        }
        tr:hover td { background: #fbfefe; }
        tr:last-child td { border-bottom: 0; }
        .empty-state {
            background: var(--surface);
            border: 1px dashed var(--line-strong);
            border-radius: 8px;
            padding: 28px;
            color: var(--muted);
            text-align: center;
        }
        .alert {
            border: 1px solid #f04438;
            background: var(--warn-bg);
            color: var(--warn);
            border-radius: 8px;
            padding: 13px 15px;
            margin-bottom: 18px;
            line-height: 1.45;
        }
        @media (max-width: 980px) {
            .app-shell { grid-template-columns: 1fr; }
            .sidebar {
                position: relative;
                height: auto;
                display: block;
            }
            .nav-list { grid-template-columns: repeat(auto-fill, minmax(210px, 1fr)); }
            .topbar { position: static; }
            main { padding: 18px; }
            .page-head { align-items: flex-start; flex-direction: column; }
        }
    </style>
</head>
<body>
    @php
        $menuReports = config('reports');
        $activeSlug = $slug ?? null;
    @endphp
    <div class="app-shell">
        <aside class="sidebar">
            <div class="brand">
                <div class="brand-mark">R</div>
                <div>
                    <strong>Reportes Web</strong>
                    <span>Huella Digital VI</span>
                </div>
            </div>
            <div class="nav-label">Reportes</div>
            <nav class="nav-list">
                @foreach ($menuReports as $menuSlug => $menuReport)
                    <a class="nav-item {{ $activeSlug === $menuSlug ? 'active' : '' }}" href="{{ route('reports.show', $menuSlug) }}">
                        {{ $menuReport['title'] }}
                    </a>
                @endforeach
            </nav>
        </aside>

        <div class="content">
            <header class="topbar">
                <div>
                    <div class="topbar-title">Sistema de Reportes</div>
                    <div class="topbar-subtitle">Consultas legacy SQL Server con exportacion profesional</div>
                </div>
                <a class="button secondary" href="{{ route('reports.index') }}">Menu</a>
            </header>
            <main>
                @yield('content')
            </main>
        </div>
    </div>
</body>
</html>
