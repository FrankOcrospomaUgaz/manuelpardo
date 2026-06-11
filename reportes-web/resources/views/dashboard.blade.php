@extends('layouts.app', ['title' => 'Dashboard'])

@section('topbar-title', 'Panel Principal')
@section('breadcrumb', 'Inicio')

@section('content')
    <div class="page-header">
        <div>
            <h1>📊 Panel Principal</h1>
            <p class="subtitle">Bienvenido al Sistema de Control de Personal — {{ now()->translatedFormat('l d \\d\\e F, Y') }}</p>
        </div>
    </div>

    {{-- Stats --}}
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;margin-bottom:28px;">
        <div class="card" style="padding:20px;border-left:4px solid var(--accent)">
            <div style="color:var(--muted);font-size:12px;font-weight:600">TRABAJADORES</div>
            <div style="font-size:28px;font-weight:900;margin-top:6px;color:var(--accent-dark)">{{ $stats['trabajadores'] }}</div>
            <div style="color:var(--muted-light);font-size:11px;margin-top:4px">Registrados en el sistema</div>
        </div>
        <div class="card" style="padding:20px;border-left:4px solid var(--accent2)">
            <div style="color:var(--muted);font-size:12px;font-weight:600">ÁREAS</div>
            <div style="font-size:28px;font-weight:900;margin-top:6px;color:#7c3aed">{{ $stats['areas'] }}</div>
            <div style="color:var(--muted-light);font-size:11px;margin-top:4px">Áreas activas</div>
        </div>
        <div class="card" style="padding:20px;border-left:4px solid var(--success)">
            <div style="color:var(--muted);font-size:12px;font-weight:600">GRUPOS</div>
            <div style="font-size:28px;font-weight:900;margin-top:6px;color:#059669">{{ $stats['grupos'] }}</div>
            <div style="color:var(--muted-light);font-size:11px;margin-top:4px">Grupos configurados</div>
        </div>
        <div class="card" style="padding:20px;border-left:4px solid var(--warn)">
            <div style="color:var(--muted);font-size:12px;font-weight:600">USUARIOS</div>
            <div style="font-size:28px;font-weight:900;margin-top:6px;color:#d97706">{{ $stats['usuarios'] }}</div>
            <div style="color:var(--muted-light);font-size:11px;margin-top:4px">Usuarios del sistema</div>
        </div>
    </div>

    {{-- Quick access --}}
    <h2 style="font-size:18px;font-weight:800;margin-bottom:16px">Acceso Rápido</h2>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px;">
        @php
            $shortcuts = [
                ['url' => 'mantenimiento/trabajadores', 'icon' => '👤', 'title' => 'Trabajadores', 'desc' => 'Gestionar personal registrado'],
                ['url' => 'mantenimiento/areas', 'icon' => '🏢', 'title' => 'Áreas', 'desc' => 'Administrar áreas organizacionales'],
                ['url' => 'mantenimiento/grupos', 'icon' => '👥', 'title' => 'Grupos', 'desc' => 'Gestionar grupos de trabajo'],
                ['url' => 'mantenimiento/feriados', 'icon' => '🎉', 'title' => 'Feriados', 'desc' => 'Configurar días feriados'],
                ['url' => 'mantenimiento/usuarios', 'icon' => '🔑', 'title' => 'Usuarios', 'desc' => 'Administrar accesos al sistema'],
                ['url' => 'reportes/trabajadores', 'icon' => '📈', 'title' => 'Reportes', 'desc' => 'Generar reportes del sistema'],
                ['url' => 'movimiento/justificaciones', 'icon' => '📝', 'title' => 'Justificaciones', 'desc' => 'Registrar justificaciones'],
                ['url' => 'movimiento/suspensiones', 'icon' => '⛔', 'title' => 'Suspensiones', 'desc' => 'Gestionar suspensiones'],
            ];
        @endphp
        @foreach ($shortcuts as $sc)
            <a href="{{ url($sc['url']) }}" class="card" style="padding:18px;display:flex;gap:14px;align-items:center;transition:all .15s;cursor:pointer" onmouseover="this.style.borderColor='var(--accent)';this.style.transform='translateY(-2px)'" onmouseout="this.style.borderColor='';this.style.transform=''">
                <div style="width:42px;height:42px;border-radius:10px;background:var(--accent-soft);display:grid;place-items:center;font-size:20px;flex-shrink:0">{{ $sc['icon'] }}</div>
                <div>
                    <div style="font-weight:700;font-size:14px">{{ $sc['title'] }}</div>
                    <div style="color:var(--muted);font-size:12px;margin-top:2px">{{ $sc['desc'] }}</div>
                </div>
            </a>
        @endforeach
    </div>
@endsection
