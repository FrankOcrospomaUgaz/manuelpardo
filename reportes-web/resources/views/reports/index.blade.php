@extends('reports.layout', ['title' => 'Reportes'])

@section('content')
    <div class="page-head">
        <div>
            <h1>Reportes</h1>
            <p class="eyebrow">Menu web equivalente al modulo de reportes del sistema de escritorio.</p>
        </div>
    </div>

    <section class="metric-strip">
        <div class="metric">
            <span>Reportes disponibles</span>
            <strong>{{ count($reports) }}</strong>
        </div>
        <div class="metric">
            <span>Conexion</span>
            <strong>SQL Server</strong>
        </div>
        <div class="metric">
            <span>Exportacion</span>
            <strong>Excel / PDF</strong>
        </div>
    </section>

    <section class="grid">
        @foreach ($reports as $slug => $report)
            <a class="card" href="{{ route('reports.show', $slug) }}">
                <strong>{{ $report['title'] }}</strong>
                <span class="muted">{{ $report['legacy_form'] }} / {{ $report['legacy_report'] }}</span>
            </a>
        @endforeach
    </section>
@endsection
