<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <style>
        body { font-family: Arial, Helvetica, sans-serif; font-size: 10px; color: #111827; }
        h1 { font-size: 18px; margin: 0 0 4px; text-align: center; }
        .meta { color: #4b5563; margin-bottom: 12px; text-align: center; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #9ca3af; padding: 5px; }
        th { background: #146c94; color: #fff; font-weight: bold; }
        tr:nth-child(even) td { background: #f3f6f8; }
    </style>
</head>
<body>
    <h1>{{ mb_strtoupper($report['title']) }}</h1>
    <div class="meta">Generado: {{ now()->format('d/m/Y H:i') }}</div>

    @if ($rows === [])
        <p>Sin datos para mostrar.</p>
    @else
        <table>
            <thead>
                <tr>
                    @foreach (array_keys($rows[0]) as $column)
                        <th>{{ $column }}</th>
                    @endforeach
                </tr>
            </thead>
            <tbody>
                @foreach ($rows as $row)
                    <tr>
                        @foreach ($row as $value)
                            <td>{{ is_scalar($value) || $value === null ? $value : json_encode($value) }}</td>
                        @endforeach
                    </tr>
                @endforeach
            </tbody>
        </table>
    @endif
</body>
</html>
