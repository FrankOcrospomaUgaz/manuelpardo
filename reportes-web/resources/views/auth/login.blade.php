<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Iniciar Sesión — Control de Personal</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Inter', sans-serif; min-height: 100vh; display: flex; background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%); overflow: hidden; }
        .login-left { flex: 1; display: flex; flex-direction: column; justify-content: center; align-items: center; color: #fff; padding: 40px; position: relative; }
        .login-left::before { content: ''; position: absolute; inset: 0; background: radial-gradient(circle at 30% 40%, rgba(14,165,233,.15) 0%, transparent 60%), radial-gradient(circle at 70% 80%, rgba(139,92,246,.1) 0%, transparent 50%); }
        .login-left .content { position: relative; z-index: 1; text-align: center; max-width: 400px; }
        .login-left h1 { font-size: 32px; font-weight: 900; margin-bottom: 12px; background: linear-gradient(135deg, #0ea5e9, #8b5cf6); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
        .login-left p { color: #94a3b8; font-size: 15px; line-height: 1.6; }
        .floating-shapes { position: absolute; inset: 0; overflow: hidden; pointer-events: none; }
        .shape { position: absolute; border-radius: 50%; opacity: .06; }
        .shape-1 { width: 300px; height: 300px; background: #0ea5e9; top: -50px; right: -80px; animation: float 8s ease-in-out infinite; }
        .shape-2 { width: 200px; height: 200px; background: #8b5cf6; bottom: 10%; left: -40px; animation: float 6s ease-in-out infinite reverse; }
        .shape-3 { width: 150px; height: 150px; background: #10b981; bottom: -30px; right: 20%; animation: float 7s ease-in-out infinite 1s; }
        @keyframes float { 0%, 100% { transform: translateY(0) rotate(0deg); } 50% { transform: translateY(-20px) rotate(5deg); } }

        .login-right { width: 480px; display: flex; align-items: center; justify-content: center; padding: 40px; }
        .login-card { background: #fff; border-radius: 20px; padding: 40px; width: 100%; box-shadow: 0 30px 80px rgba(0,0,0,.3); }
        .login-icon { width: 56px; height: 56px; border-radius: 14px; background: linear-gradient(135deg, #0ea5e9, #8b5cf6); display: grid; place-items: center; color: #fff; font-size: 24px; font-weight: 900; margin-bottom: 24px; }
        .login-card h2 { font-size: 22px; font-weight: 800; color: #0f172a; margin-bottom: 4px; }
        .login-card .sub { color: #64748b; font-size: 13px; margin-bottom: 28px; }

        .field { margin-bottom: 18px; }
        .field label { display: block; font-size: 12px; font-weight: 700; color: #475569; margin-bottom: 7px; }
        .field input { width: 100%; padding: 12px 14px; border: 1.5px solid #e2e8f0; border-radius: 10px; font: inherit; font-size: 14px; transition: all .15s; background: #f8fafc; }
        .field input:focus { outline: none; border-color: #0ea5e9; box-shadow: 0 0 0 3px rgba(14,165,233,.12); background: #fff; }
        .login-btn { width: 100%; padding: 13px; border: none; border-radius: 10px; background: linear-gradient(135deg, #0ea5e9, #0284c7); color: #fff; font: inherit; font-size: 14px; font-weight: 800; cursor: pointer; transition: all .2s; margin-top: 8px; }
        .login-btn:hover { transform: translateY(-1px); box-shadow: 0 8px 25px rgba(14,165,233,.35); }
        .login-error { background: #fee2e2; color: #991b1b; padding: 10px 14px; border-radius: 8px; font-size: 13px; margin-bottom: 16px; border: 1px solid #fecaca; }
        @media (max-width: 900px) { .login-left { display: none; } .login-right { width: 100%; } }
    </style>
</head>
<body>
    <div class="login-left">
        <div class="floating-shapes">
            <div class="shape shape-1"></div>
            <div class="shape shape-2"></div>
            <div class="shape shape-3"></div>
        </div>
        <div class="content">
            <h1>Sistema de Control de Personal</h1>
            <p>Gestión integral de asistencia, horarios, marcaciones y reportes del personal. Versión Web 3.0</p>
        </div>
    </div>
    <div class="login-right">
        <div class="login-card">
            <div class="login-icon">CP</div>
            <h2>Iniciar sesión</h2>
            <p class="sub">Ingrese sus credenciales para acceder al sistema</p>

            @if (session('error'))
                <div class="login-error">{{ session('error') }}</div>
            @endif

            <form method="POST" action="{{ url('login') }}">
                @csrf
                <div class="field">
                    <label>Usuario</label>
                    <input type="text" name="usuario" value="{{ old('usuario') }}" required autofocus placeholder="Ingrese su usuario">
                </div>
                <div class="field">
                    <label>Contraseña</label>
                    <input type="password" name="clave" required placeholder="Ingrese su contraseña">
                </div>
                <button type="submit" class="login-btn">Ingresar al Sistema</button>
            </form>
        </div>
    </div>
</body>
</html>
