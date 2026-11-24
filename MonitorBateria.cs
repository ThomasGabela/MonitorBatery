// MonitorBateria: registra cada segundo la batería, la energía, las temperaturas, la GPU, los FPS y
// los procesos, para diagnosticar apagados del portátil. No arranca con Windows: se abre a mano.
// Con permisos de administrador también mide los FPS (ETW), todos los procesos al instante (WMI)
// y los datos que publica el controlador del equipo de MSI (WMI).
// Está escrito en C# 5 para compilar con el csc.exe que ya trae Windows (ver compilar.bat).
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;

namespace MonitorBateria
{
    static class Program
    {
        internal static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        internal static readonly CultureInfo Es = new CultureInfo("es-ES");
        const int CadaProcesos = 3;      // procesos.csv y GPU por proceso: cada 3 muestras
        const int CadaInfoBateria = 5;   // capacidad completa de la batería: cada 5 muestras
        const int SegundosMsi = 10;      // datos de MSI por WMI: cada 10 s
        const int SegundosBateriaMsi = 2;   // salvo la batería de MSI: cada 2 s
        internal static int Intervalo = 1;   // segundos entre muestras
        static int duracionMax;              // solo para pruebas: --duracion=N

        internal static volatile bool Salir, TerminadoPorDuracion;
        internal static volatile string MotivoFin;
        static readonly object lockFin = new object();
        static bool finalizado;

        internal static bool EsAdmin;
        internal static string DirRegistros, DirSesion;
        static string archivoActiva;
        static Registro logSistema, logProcesos, logEventos, logNuevos, logFps, logMsi;
        static readonly LinkedList<string> eventosRecientes = new LinkedList<string>();
        static readonly object lockEventos = new object();
        internal static volatile int VersionEventos;
        internal static volatile Lectura UltimaLectura;
        static long muestras, ultimaMuestraTicks;
        internal static DateTime Inicio;
        static int pidPropio;

        static Bateria bat;
        static Contadores cont;
        static Nvidia nv;
        static GpuMotores gpuMot;
        static Procesos procs;
        static Fps fps;
        static EventLogWatcher vigilante;
        static ManagementEventWatcher wmiInicio, wmiFin;
        static readonly Dictionary<int, DateTime> vivos = new Dictionary<int, DateTime>();
        internal static volatile string EstadoFps = "", EstadoProcesos = "", EstadoMsi = "", Pantallas = "";
        // Últimos valores del controlador de MSI ya traducidos (se sustituye el objeto entero)
        internal static volatile DatosMsi Msi = new DatosMsi();

        // Estado anterior, para registrar los cambios como eventos
        static int prevAc = -1, prevCargando = -1, prevPct = -1, prevLimPasivo = 100, prevAhorro = -1;
        static uint prevCompleta;
        static double prevVolt = double.NaN;
        static bool prevBatOk = true;
        static string prevFrenosGpu = "", prevPlan = "", prevModo = "";
        static HashSet<string> juegosPrev = new HashSet<string>();
        static volatile string juegoActual = "";
        static volatile Dictionary<int, UsoGpu> gpuPorPid = new Dictionary<int, UsoGpu>();
        internal static volatile string Plan = "", Modo = "";

        // Episodios en los que la batería aporta energía estando enchufada
        static int deficitMuestras, deficitSinAporte;
        static DateTime deficitInicio;
        static double deficitSumaW, deficitPicoW, deficitCapIni;

        // Consumo del propio monitor (incluye nvidia-smi)
        static long propioCpuPrev;
        static DateTime propioHoraPrev = DateTime.MinValue;
        internal static volatile string PropioTexto = "midiendo...";

        internal static readonly Dictionary<string, string> Juegos = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "RocketLeague.exe", "Rocket League" },
            { "AoE2DE_s.exe", "Age of Empires II DE" },
            { "AoE2DE.exe", "Age of Empires II DE" },
            { "AoE3DE_s.exe", "Age of Empires III DE" },
            { "RelicCardinal.exe", "Age of Empires IV" },
            { "Frostpunk2-Win64-Shipping.exe", "Frostpunk 2" },
        };

        // Procesos que se registran siempre (sumando todas sus instancias), aunque no estén en el top
        static readonly HashSet<string> Vigilados = new HashSet<string>(new[]
        {
            "RocketLeague.exe", "AoE2DE_s.exe", "AoE2DE.exe", "AoE3DE_s.exe", "RelicCardinal.exe", "Frostpunk2-Win64-Shipping.exe",
            "Discord.exe", "chrome.exe", "wallpaper64.exe", "wallpaper32.exe", "webwallpaper32.exe", "steam.exe", "steamwebhelper.exe",
            "msedgewebview2.exe", "NVIDIA Overlay.exe", "nvcontainer.exe", "MsMpEng.exe", "dwm.exe", "claude.exe", "audiodg.exe",
            "HoldemManager.Server.exe", "postgres.exe", "Razer Synapse 3.exe", "vgc.exe", "MonitorBateria.exe", "nvidia-smi.exe",
        }, StringComparer.OrdinalIgnoreCase);

        // Clases WMI de MSI (solo con administrador). No está documentado qué contiene cada una:
        // se guardan en crudo para buscar después la temperatura de la batería, ventiladores, etc.
        static readonly string[] ClasesMsi =
            { "MSI_Master_Battery", "MSI_Slave_Battery", "MSI_Power", "MSI_System", "MSI_CPU", "MSI_VGA", "MSI_AP", "MSI_Device", "MSI_Software" };

        const string ProveedoresXPath = "Provider[@Name='Microsoft-Windows-Kernel-Power' or @Name='Microsoft-Windows-Kernel-Processor-Power'" +
            " or @Name='Microsoft-Windows-Kernel-General' or @Name='EventLog' or @Name='User32' or @Name='Display' or @Name='nvlddmkm'" +
            " or @Name='Microsoft-Windows-WHEA-Logger' or @Name='Microsoft-Windows-ACPI' or @Name='Microsoft-Windows-UserModePowerService']";

        const string CabeceraSistema = "fecha_hora,cargador,cargando,descargando,critico,ahorro_bateria,bateria_pct_windows,bateria_pct_calc," +
            "capacidad_mwh,capacidad_completa_mwh,tasa_mw,voltaje_mv,temp_bateria_c,cpu_uso_pct,cpu_rendimiento_pct,cpu_mhz," +
            "cpu_paquete_w,cpu_nucleos_w,igpu_w,dram_w,temp_acpi_c,limite_pasivo_pct,razones_limite_termico,gpu_w,gpu_w_instantaneo," +
            "gpu_limite_w,gpu_uso_pct,gpu_mem_uso_pct,gpu_temp_c,gpu_mhz,gpu_mem_mhz,gpu_pstate,gpu_razones_hex,gpu_razones,gpu_vram_mb," +
            "cpu_mas_gpu_w,ram_libre_mb,ram_comprometida_pct,juego,juego_fps,juego_frametime_prom_ms,juego_frametime_max_ms," +
            "juego_frametime_p99_ms,plan_energia,modo_energia,bateria_ok,msi_bat_ma,msi_bat_temp_c,msi_bat_restante_mah," +
            "msi_bat_completa_mah,msi_cpu_c,msi_gpu_c,msi_ventilador_cpu_pct,msi_ventilador_gpu_pct";

        [STAThread]
        static int Main(string[] args)
        {
            bool esperar = false, sinVentana = false;
            string captura = null;
            foreach (string a in args)
            {
                int s;
                if (a.StartsWith("--duracion=") && int.TryParse(a.Substring(11), out s)) duracionMax = s;
                else if (a.StartsWith("--captura=")) captura = a.Substring(10);
                else if (a == "--esperar") esperar = true;
                else if (a == "--sin-ventana") sinVentana = true;
                else if (int.TryParse(a, out s) && s >= 1 && s <= 60) Intervalo = s;
            }
            if (!sinVentana)
            {
                Nativo.SetProcessDPIAware();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += (s, e) => Evento("ERROR", "Ventana: " + e.Exception.Message);
            }

            // Una sola instancia. Al reiniciar como administrador, la nueva espera a que cierre la anterior.
            Mutex mutex = null;
            bool unico;
            try { mutex = new Mutex(false, "Local\\MonitorBateria"); unico = mutex.WaitOne(esperar ? 20000 : 0); }
            catch (AbandonedMutexException) { unico = true; }
            catch (UnauthorizedAccessException) { unico = false; }   // la tiene una instancia de administrador
            if (!unico)
            {
                if (!sinVentana)
                    MessageBox.Show("El monitor ya está abierto. Busca su icono junto al reloj de Windows.", "Monitor de batería",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 1;
            }

            try
            {
                Iniciar();
                var hilo = new Thread(Bucle);
                hilo.IsBackground = true;
                hilo.Name = "muestreo";
                hilo.Start();
                if (sinVentana) hilo.Join();
                else
                {
                    Application.Run(new Ventana(captura));
                    Salir = true;
                    hilo.Join(3000);
                }
                Finalizar(MotivoFin ?? "cierre del programa");
            }
            finally
            {
                try { mutex.ReleaseMutex(); } catch { }
            }
            return 0;
        }

        static void Iniciar()
        {
            // Prioridad baja: si el juego necesita la CPU, el monitor cede
            try { Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.BelowNormal; } catch { }
            pidPropio = Process.GetCurrentProcess().Id;
            EsAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Evento("ERROR_FATAL", Convert.ToString(e.ExceptionObject));

            // Fuera de OneDrive: un archivo que cambia cada segundo haría que OneDrive lo subiera sin parar
            DirRegistros = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "MonitorBateria", "registros");
            Directory.CreateDirectory(DirRegistros);
            Inicio = DateTime.Now;
            DirSesion = Path.Combine(DirRegistros, Inicio.ToString("yyyy-MM-dd_HH-mm-ss", Inv));
            Directory.CreateDirectory(DirSesion);
            archivoActiva = Path.Combine(DirRegistros, "sesion_activa.txt");

            logEventos = new Registro(Path.Combine(DirSesion, "eventos.csv"), "fecha_hora,tipo,detalle");
            Evento("INICIO", "Monitor iniciado " + (EsAdmin ? "como administrador" : "sin permisos de administrador") +
                "; una muestra cada " + Intervalo + " s");
            RevisarSesionAnterior();
            // Nombre relativo: sigue sirviendo si la carpeta se mueve
            File.WriteAllText(archivoActiva, Path.GetFileName(DirSesion));

            bat = new Bateria();
            if (!bat.Abrir()) Evento("AVISO", "No se pudo abrir la batería");
            cont = new Contadores();
            nv = new Nvidia(CrearTrabajo());
            nv.Asegurar(Intervalo * 1000);
            gpuMot = new GpuMotores();
            procs = new Procesos();
            Pantallas = TextoPantallas();

            logSistema = new Registro(Path.Combine(DirSesion, "sistema.csv"), CabeceraSistema);
            logProcesos = new Registro(Path.Combine(DirSesion, "procesos.csv"),
                "fecha_hora,tipo,pid,nombre,instancias,cpu_pct,ram_mb,privada_mb,hilos,gpu_pct,gpu_motor,gpu_adaptador");
            logNuevos = new Registro(Path.Combine(DirSesion, "procesos_inicio_fin.csv"),
                "fecha_hora,evento,pid,pid_padre,padre,nombre,ruta,codigo_salida,vida_s");
            logFps = new Registro(Path.Combine(DirSesion, "fps.csv"),
                "fecha_hora,pid,nombre,cuadros,fps,frametime_prom_ms,frametime_max_ms,frametime_p99_ms");

            IniciarFuncionesAdmin();
            EscribirInfo();
            IniciarVigilante();
        }

        static void IniciarFuncionesAdmin()
        {
            if (!EsAdmin)
            {
                EstadoFps = "requiere administrador";
                EstadoProcesos = "revisados cada " + Intervalo + " s";
                EstadoMsi = "requiere administrador";
                return;
            }

            fps = new Fps();
            EstadoFps = fps.Iniciar() ? "activo" : "no disponible: " + fps.Error;
            Evento("FPS", "Medición de FPS (ETW DXGI/D3D9): " + EstadoFps);

            try
            {
                wmiInicio = new ManagementEventWatcher(new WqlEventQuery("SELECT * FROM Win32_ProcessStartTrace"));
                wmiInicio.EventArrived += AlIniciarProceso;
                wmiInicio.Start();
                wmiFin = new ManagementEventWatcher(new WqlEventQuery("SELECT * FROM Win32_ProcessStopTrace"));
                wmiFin.EventArrived += AlTerminarProceso;
                wmiFin.Start();
                EstadoProcesos = "todos al instante";
            }
            catch (Exception ex)
            {
                PararWmi();
                EstadoProcesos = "revisados cada " + Intervalo + " s (WMI falló: " + ex.Message + ")";
            }
            Evento("PROCESOS", "Inicio y fin de procesos: " + EstadoProcesos);

            logMsi = new Registro(Path.Combine(DirSesion, "msi_wmi.csv"), "fecha_hora,clase,instancia,propiedad,valor");
            EstadoMsi = "leyendo...";
            var h = new Thread(BucleMsi);
            h.IsBackground = true;
            h.Priority = ThreadPriority.BelowNormal;
            h.Name = "msi";
            h.Start();
        }

        // ------------------------------------------------------------------ bucle principal

        static void Bucle()
        {
            var reloj = Stopwatch.StartNew();
            long n = 0, k = 0;
            DateTime anterior = DateTime.MinValue;
            while (!Salir)
            {
                DateTime ahora = DateTime.Now;
                if (anterior != DateTime.MinValue && (ahora - anterior).TotalSeconds > Intervalo * 5 + 2)
                    Evento("SALTO", string.Format(Inv, "{0:F0} s sin muestras (suspensión, bloqueo o equipo saturado)", (ahora - anterior).TotalSeconds));
                anterior = ahora;
                try { Muestra(k, ahora); }
                catch (Exception ex) { Evento("ERROR", ex.GetType().Name + ": " + ex.Message); }
                k++;
                if (duracionMax > 0 && reloj.Elapsed.TotalSeconds >= duracionMax)
                {
                    MotivoFin = "fin de la duración de prueba";
                    TerminadoPorDuracion = true;
                    break;
                }

                n++;
                long espera = n * Intervalo * 1000L - reloj.ElapsedMilliseconds;
                if (espera < 0)
                {
                    // Nos retrasamos (equipo saturado o suspensión): volver a alinear con el reloj
                    n = reloj.ElapsedMilliseconds / (Intervalo * 1000L) + 1;
                    espera = n * Intervalo * 1000L - reloj.ElapsedMilliseconds;
                }
                while (espera > 0 && !Salir)
                {
                    int t = (int)Math.Min(espera, 100);
                    Thread.Sleep(t);
                    espera -= t;
                }
            }
        }

        static void Muestra(long k, DateTime ahora)
        {
            var l = new Lectura();
            l.Hora = ahora;

            // Batería (IOCTL directo al controlador: lo mismo que lee Windows, sin pasar por WMI)
            uint est, cap, volt;
            int tasa;
            l.BatOk = bat.LeerEstado(out est, out cap, out volt, out tasa);
            if (l.BatOk && (k % CadaInfoBateria == 0 || bat.Completa == 0)) bat.LeerInfo();
            SYSTEM_POWER_STATUS sps;
            Nativo.GetSystemPowerStatus(out sps);
            l.Ac = l.BatOk ? (est & 1) != 0 : sps.ACLineStatus == 1;
            l.Descargando = l.BatOk && (est & 2) != 0;
            l.Cargando = l.BatOk && (est & 4) != 0;
            l.Critico = l.BatOk && (est & 8) != 0;
            l.AhorroBateria = sps.SystemStatusFlag == 1;
            l.PctWin = sps.BatteryLifePercent == 255 ? -1 : sps.BatteryLifePercent;
            if (l.BatOk)
            {
                if (cap != 0xFFFFFFFF) l.CapMwh = cap;
                if (bat.Completa > 0) l.CompletaMwh = bat.Completa;
                if (cap != 0xFFFFFFFF && bat.Completa > 0) l.PctCalc = cap * 100.0 / bat.Completa;
                if (volt != 0xFFFFFFFF) l.VoltMv = volt;
                if (tasa != int.MinValue) l.TasaMw = tasa;
                l.TempBat = bat.Temperatura();
            }

            // CPU, energía (RAPL) y temperatura (zona térmica ACPI)
            l.CpuUso = cont.CpuUso();
            l.CpuRend = cont.CpuRendimiento();
            l.CpuMhz = l.CpuRend * cont.MhzBase / 100.0;
            l.WPaq = cont.Potencia(0);
            l.WNuc = cont.Potencia(1);
            l.WIgpu = cont.Potencia(2);
            l.WDram = cont.Potencia(3);
            l.TempAcpi = cont.TempAcpi();
            l.LimPasivo = cont.LimitePasivo();
            l.RazTerm = cont.RazonesTermicas();
            l.RamLibre = cont.RamLibre();
            l.RamComp = cont.RamComprometida();

            // GPU NVIDIA (nvidia-smi queda abierto en modo bucle: no se lanza un proceso por muestra)
            nv.Asegurar(Intervalo * 1000);
            string[] g = nv.Leer();
            if (g != null)
            {
                l.GW = Num(g[0]); l.GLim = Num(g[1]); l.GUso = Num(g[2]); l.GMemUso = Num(g[3]); l.GTemp = Num(g[4]);
                l.GMhz = Num(g[5]); l.GMemMhz = Num(g[6]); l.GP = g[7]; l.GRazHex = g[8]; l.GVram = Num(g[9]);
                if (g.Length > 10) l.GWInst = Num(g[10]);
                l.GRaz = RazonesGpu(l.GRazHex, false);
            }
            if (!double.IsNaN(l.WPaq) || !double.IsNaN(l.GW))
                l.CpuMasGpu = (double.IsNaN(l.WPaq) ? 0 : l.WPaq) + (double.IsNaN(l.GW) ? 0 : l.GW);

            // Procesos y FPS
            procs.Muestrear();
            RegistrarProcesosPorSondeo(ahora);
            ActualizarJuegos();
            if (k % CadaProcesos == 0) gpuPorPid = gpuMot.Leer();
            var msi = Msi;
            if (msi.Vigente) l.Msi = msi;
            if (fps != null && fps.Activo) MedirFps(l, ahora);
            else l.FpsTexto = EsAdmin ? EstadoFps : "requiere administrador";
            if (k % 5 == 0) LeerPlan();

            logSistema.Linea(FilaSistema(l));
            logSistema.Volcar();   // al disco en cada muestra: sobrevive a un corte de luz

            DetectarCambios(l);
            if (k % CadaProcesos == 0)
            {
                EscribirProcesos(ahora);
                logProcesos.Volcar();
                logNuevos.Volcar();
            }
            if (k % 30 == 0 || k == 5) MedirPropio(ahora, k);
            UltimaLectura = l;
            Interlocked.Increment(ref muestras);
            Interlocked.Exchange(ref ultimaMuestraTicks, DateTime.UtcNow.Ticks);
        }

        static string FilaSistema(Lectura l)
        {
            return string.Join(",", Hora(l.Hora), B(l.Ac), B(l.Cargando), B(l.Descargando), B(l.Critico), B(l.AhorroBateria),
                l.PctWin < 0 ? "" : l.PctWin.ToString(Inv), F(l.PctCalc, 1), F(l.CapMwh, 0), F(l.CompletaMwh, 0), F(l.TasaMw, 0),
                F(l.VoltMv, 0), F(l.TempBat, 1), F(l.CpuUso, 1), F(l.CpuRend, 1), F(l.CpuMhz, 0), F(l.WPaq, 2), F(l.WNuc, 2),
                F(l.WIgpu, 2), F(l.WDram, 2), F(l.TempAcpi, 1), F(l.LimPasivo, 0), F(l.RazTerm, 0), F(l.GW, 2), F(l.GWInst, 2),
                F(l.GLim, 0), F(l.GUso, 0), F(l.GMemUso, 0), F(l.GTemp, 1), F(l.GMhz, 0), F(l.GMemMhz, 0), Q(l.GP), Q(l.GRazHex),
                Q(l.GRaz), F(l.GVram, 0), F(l.CpuMasGpu, 1), F(l.RamLibre, 0), F(l.RamComp, 1), Q(juegoActual), F(l.Fps, 1),
                F(l.FtProm, 2), F(l.FtMax, 2), F(l.FtP99, 2), Q(Plan), Q(Modo), B(l.BatOk),
                l.Msi == null ? "" : F(l.Msi.Ma, 0), l.Msi == null ? "" : F(l.Msi.TempC, 1), l.Msi == null ? "" : F(l.Msi.RestanteMah, 0),
                l.Msi == null ? "" : F(l.Msi.CompletaMah, 0), l.Msi == null ? "" : F(l.Msi.CpuC, 0), l.Msi == null ? "" : F(l.Msi.GpuC, 0),
                l.Msi == null ? "" : F(l.Msi.VentCpu, 0), l.Msi == null ? "" : F(l.Msi.VentGpu, 0));
        }

        static void DetectarCambios(Lectura l)
        {
            int ac = l.Ac ? 1 : 0, carg = l.Cargando ? 1 : 0, ahorro = l.AhorroBateria ? 1 : 0;
            if (prevAc != -1 && ac != prevAc) Evento("CARGADOR", (l.Ac ? "Conectado" : "DESCONECTADO") + " con " + PctTxt(l));
            if (prevCargando != -1 && carg != prevCargando) Evento("CARGA", (l.Cargando ? "Empieza a cargar" : "Deja de cargar") + " en " + PctTxt(l));
            if (prevAhorro != -1 && ahorro != prevAhorro) Evento("AHORRO", l.AhorroBateria ? "Ahorro de batería activado" : "Ahorro de batería desactivado");
            prevAc = ac;
            prevCargando = carg;
            prevAhorro = ahorro;

            if (l.PctWin >= 0 && prevPct >= 0 && l.PctWin != prevPct)
                Evento("BATERIA", prevPct + " % -> " + l.PctWin + " %" + (l.Ac ? " (enchufado)" : " (sin cargador)"));
            if (l.PctWin >= 0) prevPct = l.PctWin;

            if (l.BatOk != prevBatOk) Evento("BATERIA", l.BatOk ? "Vuelve a responder" : "NO RESPONDE: el controlador no da datos de la batería");
            prevBatOk = l.BatOk;

            if (bat.Completa != 0 && prevCompleta != 0 && bat.Completa != prevCompleta)
                Evento("CAPACIDAD", "Capacidad completa " + prevCompleta + " -> " + bat.Completa + " mWh (" + SaludTxt() + " de la de diseño)");
            if (bat.Completa != 0) prevCompleta = bat.Completa;

            if (!double.IsNaN(l.VoltMv) && !double.IsNaN(prevVolt) && prevVolt - l.VoltMv > 400)
                Evento("VOLTAJE", string.Format(Inv, "Caída brusca {0:F0} -> {1:F0} mV en {2} s", prevVolt, l.VoltMv, Intervalo));
            prevVolt = l.VoltMv;

            // La batería aporta energía estando enchufada: el cargador no da abasto en ese momento.
            // En este equipo Windows siempre informa una tasa de 0, así que si hay datos de MSI se usa
            // la corriente real del chip de la batería.
            double aporteW = AporteBateriaW(l);
            if (!double.IsNaN(aporteW) && aporteW > 0)
            {
                if (deficitMuestras == 0) { deficitInicio = l.Hora; deficitSumaW = 0; deficitPicoW = 0; deficitCapIni = l.CapMwh; }
                deficitMuestras++;
                deficitSinAporte = 0;
                deficitSumaW += aporteW;
                if (aporteW > deficitPicoW) deficitPicoW = aporteW;
                if (deficitMuestras == 2)
                    Evento("DEFICIT", string.Format(Inv, "La batería empieza a aportar {0:F1} W estando enchufado (CPU+GPU {1:F0} W){2}", aporteW, l.CpuMasGpu, JuegoTxt()));
            }
            else if (deficitMuestras > 0 && ++deficitSinAporte >= 10)   // 10 s seguidos sin aporte cierran el episodio
            {
                if (deficitMuestras >= 2)
                    Evento("DEFICIT", string.Format(Inv, "Fin: duró {0:F0} s, media {1:F2} W, pico {2:F2} W, bajó {3:F0} mWh según Windows",
                        (l.Hora - deficitInicio).TotalSeconds - deficitSinAporte, deficitSumaW / deficitMuestras, deficitPicoW, deficitCapIni - l.CapMwh));
                deficitMuestras = 0;
                deficitSinAporte = 0;
            }

            string frenos = RazonesGpu(l.GRazHex, true);
            if (frenos != prevFrenosGpu)
                Evento("GPU", frenos.Length > 0 ? "Limitación por hardware: " + frenos : "Fin de la limitación por hardware");
            prevFrenosGpu = frenos;

            int lp = double.IsNaN(l.LimPasivo) ? 100 : (int)Math.Round(l.LimPasivo);
            if (lp != prevLimPasivo)
                Evento("CPU", lp < 100 ? "Limitación térmica ACPI al " + lp + " %" : "Fin de la limitación térmica ACPI");
            prevLimPasivo = lp;
        }

        // Vatios que aporta la batería estando enchufada (0 si no aporta; NaN si no se sabe)
        internal static double AporteBateriaW(Lectura l)
        {
            if (!l.Ac) return 0;
            if (l.Msi != null && !double.IsNaN(l.Msi.Ma))
            {
                double v = double.IsNaN(l.VoltMv) ? 12000 : l.VoltMv;
                return l.Msi.Ma < -20 ? -l.Msi.Ma * v / 1e6 : 0;
            }
            if (double.IsNaN(l.TasaMw)) return double.NaN;
            return l.TasaMw < -1000 ? -l.TasaMw / 1000.0 : 0;
        }

        static void EscribirProcesos(DateTime ahora)
        {
            string h = Hora(ahora);
            var gpu = gpuPorPid;
            var lista = procs.Actual.Values.Where(x => x.Pid != 0).ToList();
            var elegidos = new Dictionary<int, InfoProc>();
            foreach (var p in lista.OrderByDescending(x => x.CpuPct).Take(12)) elegidos[p.Pid] = p;
            foreach (var p in lista.Where(x => gpu.ContainsKey(x.Pid) && gpu[x.Pid].Pct >= 1).OrderByDescending(x => gpu[x.Pid].Pct).Take(6))
                elegidos[p.Pid] = p;
            foreach (var p in elegidos.Values.OrderByDescending(x => x.CpuPct))
            {
                UsoGpu u;
                gpu.TryGetValue(p.Pid, out u);
                logProcesos.Linea(FilaProceso(h, "top", p.Pid, p.Nombre, 1, p.CpuPct, p.WorkingSet, p.Privada, p.Hilos, u));
            }
            foreach (var grupo in lista.Where(x => Vigilados.Contains(x.Nombre)).GroupBy(x => x.Nombre, StringComparer.OrdinalIgnoreCase))
            {
                UsoGpu mejor = null;
                foreach (var p in grupo)
                {
                    UsoGpu u;
                    if (gpu.TryGetValue(p.Pid, out u) && (mejor == null || u.Pct > mejor.Pct)) mejor = u;
                }
                logProcesos.Linea(FilaProceso(h, "grupo", 0, grupo.Key, grupo.Count(), grupo.Sum(x => x.CpuPct),
                    grupo.Sum(x => x.WorkingSet), grupo.Sum(x => x.Privada), grupo.Sum(x => x.Hilos), mejor));
            }
            logProcesos.Linea(FilaProceso(h, "total", 0, "(todos)", lista.Count, lista.Sum(x => x.CpuPct),
                lista.Sum(x => x.WorkingSet), lista.Sum(x => x.Privada), lista.Sum(x => x.Hilos), null));
        }

        static string FilaProceso(string h, string tipo, int pid, string nombre, int inst, double cpu, long ws, long priv, int hilos, UsoGpu u)
        {
            return string.Join(",", h, tipo, pid, Q(nombre), inst, F(cpu, 1), F(ws / 1048576.0, 0), F(priv / 1048576.0, 0), hilos,
                u == null ? "" : F(u.Pct, 1), u == null ? "" : u.Motor, u == null ? "" : (u.Dgpu ? "dGPU" : "iGPU"));
        }

        // Sin administrador: inicios y fines detectados comparando la lista de procesos de cada muestra
        static void RegistrarProcesosPorSondeo(DateTime ahora)
        {
            if (wmiInicio != null) return;
            string h = Hora(ahora);
            foreach (var p in procs.Nuevos)
            {
                InfoProc padre;
                string np = procs.Actual.TryGetValue(p.PidPadre, out padre) ? padre.Nombre : "";
                logNuevos.Linea(string.Join(",", h, "inicio", p.Pid, p.PidPadre, Q(np), Q(p.Nombre), Q(RutaProceso(p.Pid)), "", ""));
            }
            foreach (var p in procs.Terminados)
            {
                double vida = (ahora - DateTime.FromFileTime(p.Creacion)).TotalSeconds;
                logNuevos.Linea(string.Join(",", h, "fin", p.Pid, p.PidPadre, "", Q(p.Nombre), "", "", F(vida, 1)));
            }
        }

        // Con administrador: WMI avisa de cada proceso, aunque viva solo unos milisegundos
        static void AlIniciarProceso(object s, EventArrivedEventArgs e)
        {
            try
            {
                ManagementBaseObject o = e.NewEvent;
                int pid = Convert.ToInt32(o["ProcessID"]), ppid = Convert.ToInt32(o["ParentProcessID"]);
                DateTime t = TiempoWmi(o);
                lock (vivos) vivos[pid] = t;
                InfoProc padre;
                string np = procs.Actual.TryGetValue(ppid, out padre) ? padre.Nombre : "";
                logNuevos.Linea(string.Join(",", Hora(t), "inicio", pid, ppid, Q(np), Q(Convert.ToString(o["ProcessName"])), Q(RutaProceso(pid)), "", ""));
                logNuevos.Volcar();
            }
            catch { }
        }

        static void AlTerminarProceso(object s, EventArrivedEventArgs e)
        {
            try
            {
                ManagementBaseObject o = e.NewEvent;
                int pid = Convert.ToInt32(o["ProcessID"]), ppid = Convert.ToInt32(o["ParentProcessID"]);
                DateTime t = TiempoWmi(o), ini;
                string vida = "";
                lock (vivos)
                    if (vivos.TryGetValue(pid, out ini)) { vida = F((t - ini).TotalSeconds, 3); vivos.Remove(pid); }
                uint codigo = Convert.ToUInt32(o["ExitStatus"]);
                logNuevos.Linea(string.Join(",", Hora(t), "fin", pid, ppid, "", Q(Convert.ToString(o["ProcessName"])), "", "0x" + codigo.ToString("X8"), vida));
                logNuevos.Volcar();
            }
            catch { }
        }

        static DateTime TiempoWmi(ManagementBaseObject o)
        {
            try { return DateTime.FromFileTime(Convert.ToInt64(o["TIME_CREATED"])); }
            catch { return DateTime.Now; }
        }

        static void MedirFps(Lectura l, DateTime ahora)
        {
            List<Fps.Estad> est = fps.Tomar();
            var actual = procs.Actual;
            string h = Hora(ahora);
            foreach (var e in est.Take(6))
                logFps.Linea(string.Join(",", h, e.Pid, Q(NombreDe(actual, e.Pid)), e.Cuadros, F(e.Fps, 1), F(e.PromMs, 2), F(e.MaxMs, 2), F(e.P99Ms, 2)));
            if (est.Count > 0) logFps.Volcar();
            EstadoFps = "activo · " + fps.Eventos.ToString("N0", Es) + " cuadros contados";

            Fps.Estad juego = est.FirstOrDefault(x => Juegos.ContainsKey(NombreDe(actual, x.Pid)));
            if (juego != null)
            {
                l.Fps = juego.Fps; l.FtProm = juego.PromMs; l.FtMax = juego.MaxMs; l.FtP99 = juego.P99Ms;
                l.FpsTexto = string.Format(Es, "{0}: {1:N0} FPS · cuadro medio {2} · peor {3} · 1 % peores {4}",
                    Juegos[NombreDe(actual, juego.Pid)], juego.Fps, D(juego.PromMs, 1, " ms"), D(juego.MaxMs, 1, " ms"), D(juego.P99Ms, 1, " ms"));
            }
            else if (juegoActual != "") l.FpsTexto = juegoActual + ": sin cuadros en el último segundo (menú, pausa o minimizado)";
            else if (est.Count > 0) l.FpsTexto = string.Format(Es, "sin juego abierto · {0}: {1:N0} FPS", NombreDe(actual, est[0].Pid), est[0].Fps);
            else l.FpsTexto = "sin juego abierto";
        }

        static string NombreDe(Dictionary<int, InfoProc> actual, int pid)
        {
            InfoProc p;
            return actual.TryGetValue(pid, out p) ? p.Nombre : "pid " + pid;
        }

        static void ActualizarJuegos()
        {
            var ahora = new HashSet<string>();
            foreach (var p in procs.Actual.Values)
            {
                string j;
                if (Juegos.TryGetValue(p.Nombre, out j)) ahora.Add(j);
            }
            string hora = DateTime.Now.ToString("HHmmss", Inv);
            foreach (string j in ahora)
                if (!juegosPrev.Contains(j))
                {
                    Evento("JUEGO", "Iniciado: " + j);
                    CopiarConfigJuego(j, Path.Combine(DirSesion, "juegos", Seguro(j) + "_config_" + hora));
                }
            foreach (string j in juegosPrev)
                if (!ahora.Contains(j))
                {
                    Evento("JUEGO", "Cerrado: " + j);
                    CopiarRegistrosJuego(j, Path.Combine(DirSesion, "juegos", Seguro(j) + "_registros_" + hora));
                }
            juegosPrev = ahora;
            juegoActual = string.Join(" + ", ahora.ToArray());
        }

        static void LeerPlan()
        {
            const string clave = @"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes";
            string plan = NombrePlan(Reg(clave, "ActivePowerScheme"));
            string modo = "enchufado: " + NombreModo(Reg(clave, "ActiveOverlayAcPowerScheme")) + " / batería: " + NombreModo(Reg(clave, "ActiveOverlayDcPowerScheme"));
            if (prevPlan != "" && plan != prevPlan) Evento("PLAN", prevPlan + " -> " + plan);
            if (prevModo != "" && modo != prevModo) Evento("MODO", prevModo + " -> " + modo);
            prevPlan = Plan = plan;
            prevModo = Modo = modo;
        }

        static void MedirPropio(DateTime ahora, long k)
        {
            long cpu = 0, ram = 0;
            InfoProc p;
            var actual = procs.Actual;
            if (actual.TryGetValue(pidPropio, out p)) { cpu += p.Cpu; ram += p.WorkingSet; }
            int pn = nv.Pid;
            if (pn != 0 && actual.TryGetValue(pn, out p)) { cpu += p.Cpu; ram += p.WorkingSet; }
            if (propioHoraPrev != DateTime.MinValue)
            {
                double pct = Math.Max(0, (cpu - propioCpuPrev) * 100.0 / ((ahora - propioHoraPrev).Ticks * (double)Environment.ProcessorCount));
                PropioTexto = string.Format(Es, "CPU {0:N2} % del total · RAM {1:N0} MB (incluye nvidia-smi)", pct, ram / 1048576.0);
                if (k % 600 == 0) Evento("CONSUMO_MONITOR", string.Format(Inv, "CPU {0:F2} %, RAM {1:F0} MB", pct, ram / 1048576.0));
            }
            propioCpuPrev = cpu;
            propioHoraPrev = ahora;
        }

        // ------------------------------------------------------------------ datos de MSI (administrador)

        static void BucleMsi()
        {
            var previo = new Dictionary<string, string>();
            var errores = new HashSet<string>();
            var opciones = new EnumerationOptions();
            opciones.Timeout = TimeSpan.FromSeconds(5);
            var ambito = new ManagementScope(@"\\.\root\WMI");
            long vuelta = 0;
            while (!Salir)
            {
                // La batería cada 2 s (su corriente real es el dato clave); el resto cada 10 s
                bool todas = vuelta % (SegundosMsi / SegundosBateriaMsi) == 0;
                var nuevo = Msi.Copia();
                string h = Hora(DateTime.Now);
                int ok = 0;
                foreach (string clase in ClasesMsi)
                {
                    if (!todas && clase != "MSI_Master_Battery") continue;
                    try
                    {
                        using (var buscador = new ManagementObjectSearcher(ambito, new ObjectQuery("SELECT * FROM " + clase), opciones))
                        using (var col = buscador.Get())
                        {
                            foreach (ManagementBaseObject o in col)
                                using (o)
                                {
                                    string inst = Convert.ToString(o["InstanceName"]);
                                    int indice;
                                    int.TryParse(inst.Substring(inst.LastIndexOf('_') + 1), out indice);
                                    foreach (PropertyData p in o.Properties)
                                    {
                                        if (p.Name == "InstanceName" || p.Name == "Active") continue;
                                        string v = ValorWmi(p.Value), clave = clase + "|" + inst + "|" + p.Name, ant;
                                        double num;
                                        if (double.TryParse(v, NumberStyles.Float, Inv, out num)) nuevo.Asignar(clase, indice, num);
                                        if (previo.TryGetValue(clave, out ant) && ant == v) continue;   // solo los cambios
                                        previo[clave] = v;
                                        logMsi.Linea(string.Join(",", h, clase, Q(inst), Q(p.Name), Q(v)));
                                    }
                                }
                            ok++;
                        }
                    }
                    catch (Exception ex)
                    {
                        if (errores.Add(clase)) logMsi.Linea(string.Join(",", h, clase, "", "(error)", Q(ex.Message.Trim())));
                    }
                }
                nuevo.Ticks = DateTime.UtcNow.Ticks;
                Msi = nuevo;
                logMsi.Volcar();
                if (todas) EstadoMsi = ok + " de " + ClasesMsi.Length + " clases · " + previo.Count + " valores";
                vuelta++;
                for (int i = 0; i < SegundosBateriaMsi * 10 && !Salir; i++) Thread.Sleep(100);
            }
        }

        static string ValorWmi(object v)
        {
            if (v == null) return "";
            var bytes = v as byte[];
            if (bytes != null) return BitConverter.ToString(bytes).Replace("-", " ");
            var obj = v as ManagementBaseObject;
            if (obj != null) return "{" + string.Join("; ", obj.Properties.Cast<PropertyData>().Select(p => p.Name + "=" + ValorWmi(p.Value))) + "}";
            var arr = v as Array;
            if (arr != null) return string.Join(" ", arr.Cast<object>().Select(x => ValorWmi(x)));
            var f = v as IFormattable;
            return f != null ? f.ToString(null, Inv) : v.ToString();
        }

        // ------------------------------------------------------------------ eventos y sesiones

        internal static void Evento(string tipo, string detalle)
        {
            DateTime t = DateTime.Now;
            detalle = (detalle ?? "").Replace("\r", " ").Replace("\n", " ");
            var r = logEventos;
            if (r != null)
            {
                r.Linea(string.Join(",", Hora(t), tipo, Q(detalle)));
                r.Volcar();
            }
            lock (lockEventos)
            {
                eventosRecientes.AddFirst(t.ToString("HH:mm:ss", Inv) + "  " + tipo.PadRight(10) + " " + detalle);
                while (eventosRecientes.Count > 200) eventosRecientes.RemoveLast();
                VersionEventos++;
            }
        }

        internal static string[] EventosRecientes()
        {
            lock (lockEventos) return eventosRecientes.ToArray();
        }

        internal static long Muestras { get { return Interlocked.Read(ref muestras); } }

        internal static double SegundosDesdeUltimaMuestra()
        {
            long t = Interlocked.Read(ref ultimaMuestraTicks);
            return t == 0 ? (DateTime.Now - Inicio).TotalSeconds : (DateTime.UtcNow.Ticks - t) / 1e7;
        }

        internal static string TextoPantallas()
        {
            try
            {
                var s = Screen.AllScreens;
                return s.Length + ": " + string.Join(" + ", s.Select(x => x.Bounds.Width + "×" + x.Bounds.Height + (x.Primary ? " (principal)" : "")));
            }
            catch { return "?"; }
        }

        static bool EventoInteresante(string prov, int id)
        {
            switch (prov)
            {
                case "Microsoft-Windows-Kernel-Power": return id != 566 && id != 172 && id != 506 && id != 507 && id != 125 && id != 187;
                case "Microsoft-Windows-Kernel-General": return id == 12 || id == 13;
                case "EventLog": return id == 6005 || id == 6006 || id == 6008;
                case "User32": return id == 1074;
                default: return true;
            }
        }

        static void IniciarVigilante()
        {
            try
            {
                vigilante = new EventLogWatcher(new EventLogQuery("System", PathType.LogName, "*[System[" + ProveedoresXPath + "]]"));
                vigilante.EventRecordWritten += (s, e) =>
                {
                    if (e.EventRecord == null) return;
                    using (var r = e.EventRecord)
                    {
                        if (!EventoInteresante(r.ProviderName, r.Id)) return;
                        Evento("WINDOWS", r.ProviderName.Replace("Microsoft-Windows-", "") + " " + r.Id + ": " + Descripcion(r));
                    }
                };
                vigilante.Enabled = true;
            }
            catch (Exception ex) { Evento("AVISO", "No se pudo vigilar el registro de eventos de Windows: " + ex.Message); }
        }

        static string Descripcion(EventRecord r)
        {
            try
            {
                string d = r.FormatDescription() ?? "";
                d = d.Replace("\r", " ").Replace("\n", " ").Trim();
                return d.Length > 220 ? d.Substring(0, 220) + "…" : d;
            }
            catch { return ""; }
        }

        sealed class EvWin { public DateTime T; public string Prov; public int Id; public string Texto; }

        static List<EvWin> EventosWindows(DateTime desde)
        {
            var res = new List<EvWin>();
            string q = "*[System[" + ProveedoresXPath + " and TimeCreated[@SystemTime>='" +
                desde.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", Inv) + "']]]";
            try
            {
                using (var lector = new EventLogReader(new EventLogQuery("System", PathType.LogName, q)))
                {
                    EventRecord e;
                    while ((e = lector.ReadEvent()) != null && res.Count < 300)
                        using (e)
                        {
                            if (!EventoInteresante(e.ProviderName, e.Id)) continue;
                            DateTime t = e.TimeCreated.HasValue ? e.TimeCreated.Value : DateTime.MinValue;
                            res.Add(new EvWin { T = t, Prov = e.ProviderName, Id = e.Id, Texto = Hora(t) + "  " + e.ProviderName + " " + e.Id + "  " + Descripcion(e) });
                        }
                }
            }
            catch (Exception ex) { res.Add(new EvWin { T = DateTime.MaxValue, Prov = "", Id = 0, Texto = "No se pudo leer el registro de eventos: " + ex.Message }); }
            return res.OrderBy(x => x.T).ToList();
        }

        // Lo primero que pasó después de la última muestra decide cómo terminó la sesión
        static string Clasificar(List<EvWin> eventos, DateTime ultima)
        {
            foreach (var e in eventos.Where(x => x.T >= ultima.AddSeconds(-5)))
            {
                if (e.Prov == "Microsoft-Windows-Kernel-Power" && e.Id == 41)
                    return "APAGADO BRUSCO: Windows registró una pérdida de energía (Kernel-Power 41)";
                if ((e.Prov == "User32" && e.Id == 1074) || (e.Prov == "Microsoft-Windows-Kernel-General" && e.Id == 13) ||
                    (e.Prov == "EventLog" && e.Id == 6006) || (e.Prov == "Microsoft-Windows-Kernel-Power" && e.Id == 109))
                    return "Apagado o reinicio normal de Windows con el monitor abierto";
                if (e.Prov == "Microsoft-Windows-Kernel-Power" && e.Id == 42)
                    return "El equipo entró en suspensión o hibernación";
            }
            return "El monitor se cerró sin aviso (proceso terminado o ventana cerrada a la fuerza)";
        }

        static void RevisarSesionAnterior()
        {
            if (!File.Exists(archivoActiva)) return;
            string prev = "";
            try { prev = File.ReadAllText(archivoActiva).Trim(); File.Delete(archivoActiva); } catch { }
            if (prev.Length == 0) return;
            prev = Path.Combine(DirRegistros, Path.GetFileName(prev));
            if (!Directory.Exists(prev)) return;

            DateTime ultima;
            string[] cola = Cola(Path.Combine(prev, "sistema.csv"), 60, out ultima);
            if (ultima == DateTime.MinValue) ultima = Directory.GetLastWriteTime(prev);
            var eventos = EventosWindows(ultima.AddMinutes(-10));
            string clasif = Clasificar(eventos, ultima);

            var sb = new StringBuilder();
            sb.AppendLine("La sesión del monitor NO se cerró de forma limpia.");
            sb.AppendLine("Clasificación: " + clasif);
            sb.AppendLine("Última muestra registrada: " + Hora(ultima));
            sb.AppendLine("Revisado al volver a abrir el monitor: " + Hora(DateTime.Now));
            sb.AppendLine();
            sb.AppendLine("== Eventos de Windows desde 10 minutos antes de la última muestra ==");
            foreach (var e in eventos) sb.AppendLine(e.Texto);
            sb.AppendLine();
            sb.AppendLine("== Últimas muestras de sistema.csv ==");
            sb.AppendLine(CabeceraSistema);
            foreach (string c in cola) sb.AppendLine(c);
            try { File.WriteAllText(Path.Combine(prev, "FIN_DE_SESION.txt"), sb.ToString(), new UTF8Encoding(true)); } catch { }

            // Los registros de los juegos aún son los de la partida que se cortó (hasta que se vuelva a abrir el juego)
            foreach (string j in Juegos.Values.Distinct()) CopiarRegistrosJuego(j, Path.Combine(prev, "juegos_al_reabrir", Seguro(j)));
            Evento("SESION_ANTERIOR", Path.GetFileName(prev) + ": " + clasif + " (última muestra " + Hora(ultima) + ")");
        }

        static string[] Cola(string ruta, int n, out DateTime ultima)
        {
            ultima = DateTime.MinValue;
            try
            {
                using (var fs = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    fs.Seek(Math.Max(0, fs.Length - 128 * 1024), SeekOrigin.Begin);
                    string texto = new StreamReader(fs, Encoding.UTF8).ReadToEnd();
                    var lineas = texto.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .Where(x => x.Length > 23 && char.IsDigit(x[0])).ToList();
                    if (lineas.Count == 0) return new string[0];
                    DateTime.TryParseExact(lineas[lineas.Count - 1].Substring(0, 23), "yyyy-MM-dd HH:mm:ss.fff", Inv, DateTimeStyles.None, out ultima);
                    return lineas.Skip(Math.Max(0, lineas.Count - n)).ToArray();
                }
            }
            catch { return new string[0]; }
        }

        static void PararWmi()
        {
            foreach (var w in new[] { wmiInicio, wmiFin })
                if (w != null) try { w.Stop(); w.Dispose(); } catch { }
            wmiInicio = wmiFin = null;
        }

        internal static void Finalizar(string motivo)
        {
            lock (lockFin)
            {
                if (finalizado) return;
                finalizado = true;
                Salir = true;
                try { if (vigilante != null) vigilante.Enabled = false; } catch { }
                PararWmi();
                if (fps != null) fps.Detener();
                Evento("FIN", "Monitor cerrado (" + motivo + ")");
                try
                {
                    File.WriteAllText(Path.Combine(DirSesion, "FIN_DE_SESION.txt"),
                        "Cierre limpio del monitor: " + motivo + "\r\nHora: " + Hora(DateTime.Now) + "\r\n", new UTF8Encoding(true));
                }
                catch { }
                foreach (var r in new[] { logSistema, logProcesos, logNuevos, logFps, logMsi, logEventos }) if (r != null) r.Cerrar();
                try { File.Delete(archivoActiva); } catch { }
                if (nv != null) nv.Parar();
            }
        }

        // ------------------------------------------------------------------ información inicial

        static void EscribirInfo()
        {
            const string bios = @"HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS";
            const string so = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion";
            LeerPlan();
            var sb = new StringBuilder();
            sb.AppendLine("Sesión iniciada: " + Hora(Inicio));
            sb.AppendLine("Permisos: " + (EsAdmin ? "administrador" : "usuario normal (sin FPS ni datos de MSI; procesos por sondeo)"));
            sb.AppendLine("Muestreo: cada " + Intervalo + " s; procesos y GPU por proceso cada " + Intervalo * CadaProcesos + " s");
            sb.AppendLine("FPS: " + EstadoFps + " · procesos: " + EstadoProcesos + " · datos MSI: " + EstadoMsi);
            sb.AppendLine("Equipo: " + Reg(bios, "SystemManufacturer") + " " + Reg(bios, "SystemProductName") + " (placa " + Reg(bios, "BaseBoardProduct") + ")");
            sb.AppendLine("BIOS: " + Reg(bios, "BIOSVersion") + " del " + Reg(bios, "BIOSReleaseDate"));
            sb.AppendLine("Windows: " + Reg(so, "DisplayVersion") + ", compilación " + Reg(so, "CurrentBuild") + "." + Reg(so, "UBR"));
            sb.AppendLine("CPU: " + Reg(@"HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString").Trim() +
                " (" + Environment.ProcessorCount + " hilos, base " + F(cont.MhzBase, 0) + " MHz)");
            double t = bat.Temperatura();
            sb.AppendLine("Batería: química " + bat.Quimica + ", diseño " + bat.Disenada + " mWh, completa " + bat.Completa + " mWh (salud " + SaludTxt() +
                "), ciclos " + bat.Ciclos + ", temperatura " + (double.IsNaN(t) ? "no la informa" : F(t, 1) + " °C"));
            sb.AppendLine("Pantallas: " + Pantallas);
            sb.AppendLine("Plan de energía: " + Plan + "; modo " + Modo);
            string info = Path.Combine(DirSesion, "info.txt");
            File.WriteAllText(info, sb.ToString(), new UTF8Encoding(true));

            // Lo que necesita lanzar programas tarda unos segundos: en segundo plano, para empezar a medir ya
            var hilo = new Thread(() =>
            {
                var extra = new StringBuilder();
                extra.AppendLine();
                extra.AppendLine("== GPU (nvidia-smi) ==");
                extra.AppendLine(Ejecutar("nvidia-smi.exe", "--query-gpu=name,driver_version,vbios_version,power.limit,power.default_limit,power.max_limit --format=csv").Trim());
                extra.AppendLine();
                extra.AppendLine("== Turbo del CPU (PERFBOOSTMODE) y estado máximo (PROCTHROTTLEMAX) en el plan activo ==");
                foreach (string ajuste in new[] { "PERFBOOSTMODE", "PROCTHROTTLEMAX" })
                    foreach (string linea in Ejecutar("powercfg.exe", "/qh SCHEME_CURRENT SUB_PROCESSOR " + ajuste).Split('\n'))
                        if (linea.Contains("0x") && !linea.Contains("posible") && !linea.Contains("Possible")) extra.AppendLine(ajuste + ": " + linea.Trim());
                try { File.AppendAllText(info, extra.ToString(), Encoding.UTF8); } catch { }

                string tmp = Path.Combine(Path.GetTempPath(), "monitorbateria_informe.xml");
                Ejecutar("powercfg.exe", "/batteryreport /xml /output \"" + tmp + "\"");
                try { File.Copy(tmp, Path.Combine(DirSesion, "informe_bateria_inicio.xml"), true); File.Delete(tmp); } catch { }
            });
            hilo.IsBackground = true;
            hilo.Priority = ThreadPriority.BelowNormal;
            hilo.Start();
        }

        // ------------------------------------------------------------------ juegos

        static string Docs { get { return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); } }
        static string Local { get { return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData); } }
        static string CarpetaAoe2 { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @"Games\Age of Empires 2 DE"); } }

        // Configuración gráfica al abrir el juego (resolución, calidad, límite de FPS...)
        static void CopiarConfigJuego(string juego, string destino)
        {
            try
            {
                if (juego == "Rocket League")
                    CopiarArchivos(Path.Combine(Docs, @"My Games\Rocket League\TAGame\Config"), "*.ini", destino, false);
                else if (juego == "Age of Empires II DE" && Directory.Exists(CarpetaAoe2))
                    foreach (string d in Directory.GetDirectories(CarpetaAoe2))
                        CopiarArchivos(Path.Combine(d, "profile"), "*.nfp", Path.Combine(destino, Path.GetFileName(d)), false);
                else if (juego == "Frostpunk 2")
                    CopiarArchivos(Path.Combine(Local, @"Frostpunk2\Saved\Config\Windows"), "*.ini", destino, false);
            }
            catch { }
        }

        // Registros que el propio juego escribe durante la partida
        static void CopiarRegistrosJuego(string juego, string destino)
        {
            try
            {
                if (juego == "Rocket League")
                    CopiarArchivos(Path.Combine(Docs, @"My Games\Rocket League\TAGame\Logs"), "Launch.log", destino, true);
                else if (juego == "Age of Empires II DE")
                {
                    string logs = Path.Combine(CarpetaAoe2, "logs");
                    if (!Directory.Exists(logs)) return;
                    var ult = new DirectoryInfo(logs).GetDirectories().Where(x => x.Name != "SlowLog").OrderByDescending(x => x.LastWriteTime).FirstOrDefault();
                    if (ult != null) CopiarArchivos(ult.FullName, "*.txt", Path.Combine(destino, ult.Name), false);
                }
                else if (juego == "Frostpunk 2")
                    CopiarArchivos(Path.Combine(Local, @"Frostpunk2\Saved\Logs"), "*.log", destino, true);
            }
            catch { }
        }

        static void CopiarArchivos(string origen, string patron, string destino, bool soloReciente)
        {
            if (!Directory.Exists(origen)) return;
            var archivos = new DirectoryInfo(origen).GetFiles(patron).Where(f => f.Length < 30L * 1024 * 1024)
                .OrderByDescending(f => f.LastWriteTime).ToList();
            if (soloReciente) archivos = archivos.Take(1).ToList();
            if (archivos.Count == 0) return;
            Directory.CreateDirectory(destino);
            foreach (var f in archivos)
                try { File.Copy(f.FullName, Path.Combine(destino, f.Name), true); } catch { }
        }

        // ------------------------------------------------------------------ utilidades

        static IntPtr CrearTrabajo()
        {
            IntPtr job = Nativo.CreateJobObject(IntPtr.Zero, null);
            if (job == IntPtr.Zero) return IntPtr.Zero;
            const int tam = 144;   // JOBOBJECT_EXTENDED_LIMIT_INFORMATION en x64
            IntPtr info = Marshal.AllocHGlobal(tam);
            try
            {
                for (int i = 0; i < tam; i++) Marshal.WriteByte(info, i, 0);
                Marshal.WriteInt32(info, 16, 0x2000);   // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE: nvidia-smi muere con el monitor
                Nativo.SetInformationJobObject(job, 9, info, tam);
            }
            finally { Marshal.FreeHGlobal(info); }
            return job;
        }

        static string Ejecutar(string exe, string argumentos)
        {
            try
            {
                var si = new ProcessStartInfo(exe, argumentos);
                si.UseShellExecute = false;
                si.RedirectStandardOutput = true;
                si.RedirectStandardError = true;
                si.CreateNoWindow = true;
                si.StandardOutputEncoding = Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage);
                using (var p = Process.Start(si))
                {
                    p.ErrorDataReceived += (s, e) => { };
                    p.BeginErrorReadLine();
                    string o = p.StandardOutput.ReadToEnd();
                    p.WaitForExit(30000);
                    return o;
                }
            }
            catch (Exception ex) { return "(error: " + ex.Message + ")"; }
        }

        static string RutaProceso(int pid)
        {
            IntPtr h = Nativo.OpenProcess(0x1000, false, pid);   // PROCESS_QUERY_LIMITED_INFORMATION
            if (h == IntPtr.Zero) return "";
            try
            {
                var sb = new StringBuilder(1024);
                int n = sb.Capacity;
                return Nativo.QueryFullProcessImageName(h, 0, sb, ref n) ? sb.ToString() : "";
            }
            finally { Nativo.CloseHandle(h); }
        }

        internal static string RazonesGpu(string hex, bool soloFrenos)
        {
            ulong v;
            if (string.IsNullOrEmpty(hex) || !ulong.TryParse(hex.Replace("0x", ""), NumberStyles.HexNumber, Inv, out v)) return "";
            var l = new List<string>();
            if (!soloFrenos && (v & 0x1) != 0) l.Add("reposo");
            if (!soloFrenos && (v & 0x4) != 0) l.Add("límite de potencia");
            if ((v & 0x8) != 0) l.Add("freno de hardware");
            if ((v & 0x20) != 0) l.Add("temperatura (software)");
            if ((v & 0x40) != 0) l.Add("temperatura (hardware)");
            if ((v & 0x80) != 0) l.Add("freno de la fuente de alimentación");
            return string.Join(" + ", l);
        }

        static string NombrePlan(string g)
        {
            switch (g.ToLowerInvariant())
            {
                case "381b4222-f694-41f0-9685-ff5bb260df2e": return "Equilibrado";
                case "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c": return "Alto rendimiento";
                case "95a996c0-1edc-4885-a394-62159565636e": return "Máximo rendimiento";
                case "a1841308-3541-4fab-bc81-f71556f20b4a": return "Economizador";
                case "e9a42b02-d5df-448d-aa00-03f14749eb61": return "Rendimiento máximo";
                default: return g == "" ? "?" : g;
            }
        }

        static string NombreModo(string g)
        {
            switch (g.ToLowerInvariant())
            {
                case "961cc777-2547-4f9d-8174-7d86181b8a7a": return "mejor eficiencia energética";
                case "": case "00000000-0000-0000-0000-000000000000": return "equilibrado";
                case "3af9b8d9-7c97-431d-ad78-34a8bfea439f": return "mejor rendimiento";
                case "ded574b5-45a0-4f42-8737-46345c09c238": return "máximo rendimiento";
                default: return g;
            }
        }

        static string Reg(string clave, string valor)
        {
            try { object o = Registry.GetValue(clave, valor, null); return o == null ? "" : o.ToString(); }
            catch { return ""; }
        }

        internal static string SaludTxt() { return bat.Disenada > 0 && bat.Completa > 0 ? (bat.Completa * 100.0 / bat.Disenada).ToString("N1", Es) + " %" : "n/d"; }
        static string PctTxt(Lectura l) { return l.PctWin >= 0 ? l.PctWin + " %" : F(l.PctCalc, 0) + " %"; }
        static string JuegoTxt() { return juegoActual == "" ? "" : " (jugando a " + juegoActual + ")"; }
        internal static string JuegoActual { get { return juegoActual; } }
        static string Seguro(string s) { return new string(s.Select(c => char.IsLetterOrDigit(c) ? c : '_').ToArray()); }
        static string Hora(DateTime t) { return t.ToString("yyyy-MM-dd HH:mm:ss.fff", Inv); }
        static string B(bool b) { return b ? "1" : "0"; }
        static string F(double v, int dec) { return double.IsNaN(v) || double.IsInfinity(v) ? "" : v.ToString("F" + dec, Inv); }
        internal static string D(double v, int dec, string unidad) { return double.IsNaN(v) || double.IsInfinity(v) ? "n/d" : v.ToString("N" + dec, Es) + unidad; }

        static double Num(string s)
        {
            double v;
            return double.TryParse(s, NumberStyles.Float, Inv, out v) ? v : double.NaN;
        }

        static string Q(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (s.IndexOfAny(new[] { ',', '"' }) < 0) return s;
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }
    }

    // ---------------------------------------------------------------------- ventana

    sealed class Ventana : Form
    {
        static readonly Color Verde = Color.FromArgb(21, 128, 61), Ambar = Color.FromArgb(180, 83, 9), Rojo = Color.FromArgb(185, 28, 28);
        static readonly string[] Filas =
            { "Batería", "Carga / descarga", "Temp. batería", "CPU", "GPU", "CPU + GPU", "FPS", "Juego", "Plan de energía", "Pantallas", "Consumo del monitor", "Registros" };
        readonly Panel cabecera;
        readonly Label lblEstado, lblDetalle, lblModo, lblAviso;
        readonly Dictionary<string, Label> valores = new Dictionary<string, Label>();
        readonly ListBox lstEventos;
        readonly NotifyIcon bandeja;
        readonly System.Windows.Forms.Timer reloj;
        readonly Icon icVerde, icAmbar, icRojo;
        readonly string captura;
        int versionEventos = -1, ticks;
        bool latido, cerrarSinPreguntar;
        Color colorActual = Verde;

        public Ventana(string captura)
        {
            this.captura = captura;
            float esc;
            using (var g = CreateGraphics()) esc = g.DpiX / 96f;
            Func<int, int> S = v => (int)Math.Round(v * esc);

            Text = "Monitor de batería y energía";
            Font = new Font("Segoe UI", 9.5f);
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            if (captura != null) { StartPosition = FormStartPosition.Manual; Location = new Point(-20000, -20000); }   // prueba: fuera de la pantalla
            ClientSize = new Size(S(900), S(720));
            MinimumSize = new Size(S(640), S(480));
            icVerde = Circulo(Color.FromArgb(34, 197, 94));
            icAmbar = Circulo(Color.FromArgb(245, 158, 11));
            icRojo = Circulo(Color.FromArgb(239, 68, 68));
            Icon = icVerde;

            var raiz = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 7, Margin = Padding.Empty, Padding = Padding.Empty };
            raiz.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            raiz.RowStyles.Add(new RowStyle(SizeType.Absolute, S(74)));
            for (int i = 0; i < 4; i++) raiz.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            raiz.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            raiz.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // Cabecera: verde grabando, ámbar si la batería aporta energía enchufado, rojo si algo falla
            cabecera = new Panel { Dock = DockStyle.Fill, BackColor = Verde, Margin = Padding.Empty };
            lblEstado = new Label { AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 16f), Location = new Point(S(14), S(6)), Text = "● GRABANDO" };
            lblDetalle = new Label { AutoSize = true, ForeColor = Color.White, Location = new Point(S(17), S(46)) };
            cabecera.Controls.Add(lblEstado);
            cabecera.Controls.Add(lblDetalle);
            raiz.Controls.Add(cabecera, 0, 0);

            var modo = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, Padding = new Padding(S(10), S(6), S(10), 0), Margin = Padding.Empty };
            lblModo = new Label { AutoSize = true, Margin = new Padding(0, S(7), S(10), S(4)) };
            var btnAdmin = new Button { Text = "Reiniciar como administrador", AutoSize = true, Visible = !Program.EsAdmin, FlatStyle = FlatStyle.System };
            btnAdmin.Click += (s, e) => ReiniciarComoAdmin();
            modo.Controls.Add(lblModo);
            modo.Controls.Add(btnAdmin);
            raiz.Controls.Add(modo, 0, 1);

            lblAviso = new Label
            {
                AutoSize = true, Dock = DockStyle.Fill, Visible = false, BackColor = Color.FromArgb(254, 243, 199), ForeColor = Color.FromArgb(120, 53, 15),
                Padding = new Padding(S(10), S(6), S(10), S(6)), Margin = new Padding(S(12), S(6), S(12), 0), Font = new Font("Segoe UI Semibold", 9.5f)
            };
            raiz.Controls.Add(lblAviso, 0, 2);

            var tabla = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(S(12), S(8), S(12), S(4)), Margin = Padding.Empty };
            tabla.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tabla.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            foreach (string f in Filas)
            {
                tabla.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tabla.Controls.Add(new Label { Text = f, AutoSize = true, ForeColor = Color.FromArgb(105, 105, 105), Margin = new Padding(0, S(3), S(14), S(3)) });
                var v = new Label { AutoSize = true, Margin = new Padding(0, S(3), 0, S(3)), Text = "…" };
                valores[f] = v;
                tabla.Controls.Add(v);
            }
            tabla.Resize += (s, e) =>
            {
                int ancho = Math.Max(S(200), tabla.ClientSize.Width - S(190));
                foreach (var v in valores.Values) v.MaximumSize = new Size(ancho, 0);
            };
            raiz.Controls.Add(tabla, 0, 3);

            raiz.Controls.Add(new Label { Text = "Últimos eventos", AutoSize = true, Font = new Font("Segoe UI Semibold", 9.5f), Margin = new Padding(S(12), S(8), 0, S(2)) }, 0, 4);
            lstEventos = new ListBox
            {
                Dock = DockStyle.Fill, IntegralHeight = false, Font = new Font("Consolas", 9f), Margin = new Padding(S(12), 0, S(12), 0),
                BorderStyle = BorderStyle.FixedSingle, HorizontalScrollbar = true
            };
            raiz.Controls.Add(lstEventos, 0, 5);

            var botones = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(S(10), S(8), S(10), S(10)), Margin = Padding.Empty };
            var btnDetener = new Button { Text = "Detener y cerrar", AutoSize = true, FlatStyle = FlatStyle.System };
            btnDetener.Click += (s, e) => Close();
            var btnCarpeta = new Button { Text = "Abrir carpeta de registros", AutoSize = true, FlatStyle = FlatStyle.System };
            btnCarpeta.Click += (s, e) => AbrirCarpeta();
            botones.Controls.Add(btnDetener);
            botones.Controls.Add(btnCarpeta);
            raiz.Controls.Add(botones, 0, 6);
            Controls.Add(raiz);

            // Icono junto al reloj: se ve que está grabando aunque la ventana esté detrás del juego
            var menu = new ContextMenuStrip();
            menu.Items.Add("Mostrar ventana", null, (s, e) => Mostrar());
            menu.Items.Add("Abrir carpeta de registros", null, (s, e) => AbrirCarpeta());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Detener monitor", null, (s, e) => { Mostrar(); Close(); });
            bandeja = new NotifyIcon { Icon = icVerde, Text = "Monitor de batería", Visible = true, ContextMenuStrip = menu };
            bandeja.DoubleClick += (s, e) => Mostrar();

            SystemEvents.DisplaySettingsChanged += AlCambiarPantallas;
            reloj = new System.Windows.Forms.Timer { Interval = 1000 };
            reloj.Tick += (s, e) => Actualizar();
            reloj.Start();
            Actualizar();
        }

        void Actualizar()
        {
            ticks++;
            var l = Program.UltimaLectura;
            double hace = Program.SegundosDesdeUltimaMuestra();
            bool parado = hace > Math.Max(6, Program.Intervalo * 5);
            double aporte = l == null ? double.NaN : Program.AporteBateriaW(l);
            bool deficit = !double.IsNaN(aporte) && aporte > 0;
            bool batMal = l != null && !l.BatOk;

            Color c = parado || batMal ? Rojo : deficit ? Ambar : Verde;
            if (c != colorActual)
            {
                colorActual = c;
                cabecera.BackColor = c;
                Icon ic = c == Verde ? icVerde : c == Ambar ? icAmbar : icRojo;
                Icon = ic;
                bandeja.Icon = ic;
            }
            string tip = "Monitor de batería: " + (parado ? "SIN MUESTRAS" : "grabando") +
                (l != null && l.PctWin >= 0 ? " · " + l.PctWin + " %" : "") + (l != null ? " · " + Program.D(l.CpuMasGpu, 0, " W") : "");
            bandeja.Text = tip.Length > 63 ? tip.Substring(0, 63) : tip;
            if (Program.TerminadoPorDuracion) { cerrarSinPreguntar = true; Close(); return; }
            if (WindowState == FormWindowState.Minimized) return;   // minimizada: solo el icono, menos trabajo

            latido = !latido;   // el punto parpadea: si deja de parpadear, la ventana está colgada
            lblEstado.Text = parado ? "■ SIN MUESTRAS" : (latido ? "● " : "○ ") + "GRABANDO";
            TimeSpan d = DateTime.Now - Program.Inicio;
            lblDetalle.Text = string.Format(Program.Es, "{0:00}:{1:mm\\:ss} grabando · {2:N0} muestras · última hace {3:N0} s · una cada {4} s",
                (int)d.TotalHours, d, Program.Muestras, hace, Program.Intervalo);
            lblModo.Text = Program.EsAdmin
                ? "Administrador ✓   FPS: " + Program.EstadoFps + "   ·   procesos: " + Program.EstadoProcesos + "   ·   datos MSI: " + Program.EstadoMsi
                : "Sin administrador: no se miden los FPS ni los datos de MSI, y los procesos se revisan cada " + Program.Intervalo + " s.";

            if (parado) MostrarAviso("El monitor no registra muestras desde hace " + hace.ToString("N0", Program.Es) + " s.");
            else if (batMal) MostrarAviso("La batería no responde: el controlador no da datos.");
            else if (deficit) MostrarAviso(string.Format(Program.Es, "La batería está aportando {0:N1} W estando enchufado: el cargador no da abasto en este momento.", aporte));
            else if (lblAviso.Visible) lblAviso.Visible = false;

            if (l != null)
            {
                string estado = l.Cargando ? "cargando" : l.Descargando ? "descargando" : "ni carga ni descarga";
                V("Batería", string.Format(Program.Es, "cargador {0} · {1} · {2} ({3} / {4} mWh) · salud {5}{6}",
                    l.Ac ? "CONECTADO" : "DESCONECTADO", estado, l.PctWin >= 0 ? l.PctWin + " %" : Program.D(l.PctCalc, 0, " %"),
                    Program.D(l.CapMwh, 0, ""), Program.D(l.CompletaMwh, 0, ""), Program.SaludTxt(), l.AhorroBateria ? " · ahorro de batería activo" : ""));
                string carga = (double.IsNaN(l.TasaMw) ? "n/d" : (l.TasaMw >= 0 ? "+" : "") + (l.TasaMw / 1000).ToString("N1", Program.Es) + " W") +
                    " según Windows · voltaje " + Program.D(l.VoltMv / 1000, 2, " V");
                if (l.Msi != null && !double.IsNaN(l.Msi.Ma))
                    carga = string.Format(Program.Es, "corriente real {0:N0} mA (≈ {1:N1} W; negativa = se descarga) · ", l.Msi.Ma,
                        l.Msi.Ma * (double.IsNaN(l.VoltMv) ? 12000 : l.VoltMv) / 1e6) + carga;
                V("Carga / descarga", carga);
                string temp = "la batería no la informa a Windows" + (Program.EsAdmin ? "" : " (con administrador se lee del chip)");
                if (l.Msi != null && !double.IsNaN(l.Msi.TempC))
                    temp = Program.D(l.Msi.TempC, 1, " °C") + string.Format(Program.Es, " según el chip · capacidad del chip {0:N0} de {1:N0} mAh ({2:N1} % de salud)",
                        l.Msi.CompletaMah, l.Msi.DisenoMah, 100 * l.Msi.CompletaMah / l.Msi.DisenoMah);
                else if (!double.IsNaN(l.TempBat)) temp = Program.D(l.TempBat, 1, " °C");
                V("Temp. batería", temp);
                V("CPU", string.Format(Program.Es, "uso {0} · {1} ({2}) · {3} (núcleos {4} · iGPU {5} · RAM {6}) · {7} · límite térmico {8}",
                    Program.D(l.CpuUso, 0, " %"), Program.D(l.CpuMhz, 0, " MHz"), Program.D(l.CpuRend, 0, " %"), Program.D(l.WPaq, 1, " W"),
                    Program.D(l.WNuc, 1, ""), Program.D(l.WIgpu, 1, ""), Program.D(l.WDram, 1, ""), Program.D(l.TempAcpi, 0, " °C"), Program.D(l.LimPasivo, 0, " %")));
                V("GPU", string.Format(Program.Es, "uso {0} · {1} (instantáneo {2}) de {3} · {4} · {5} · {6} · {7}",
                    Program.D(l.GUso, 0, " %"), Program.D(l.GW, 1, " W"), Program.D(l.GWInst, 1, " W"), Program.D(l.GLim, 0, " W"), Program.D(l.GTemp, 0, " °C"),
                    Program.D(l.GMhz, 0, " MHz"), l.GP == "" ? "n/d" : l.GP, l.GRaz == "" ? "sin limitación" : "limitada por: " + l.GRaz));
                V("CPU + GPU", string.Format(Program.Es, "{0} · RAM libre {1} · memoria comprometida {2}",
                    Program.D(l.CpuMasGpu, 1, " W"), Program.D(l.RamLibre, 0, " MB"), Program.D(l.RamComp, 0, " %")));
                V("FPS", l.FpsTexto ?? "");
            }
            string juego = Program.JuegoActual;
            V("Juego", juego == "" ? "ninguno detectado" : juego);
            V("Plan de energía", Program.Plan + " · " + Program.Modo);
            V("Pantallas", Program.Pantallas);
            V("Consumo del monitor", Program.PropioTexto);
            V("Registros", Program.DirSesion);

            int ver = Program.VersionEventos;
            if (ver != versionEventos)
            {
                versionEventos = ver;
                lstEventos.BeginUpdate();
                lstEventos.Items.Clear();
                foreach (string e in Program.EventosRecientes()) lstEventos.Items.Add(e);
                lstEventos.EndUpdate();
            }

            if (captura != null && ticks == 8)
                try
                {
                    using (var bmp = new Bitmap(Width, Height))
                    {
                        DrawToBitmap(bmp, new Rectangle(0, 0, Width, Height));
                        bmp.Save(captura);
                    }
                }
                catch { }
        }

        // En la prueba de captura la ventana no le quita el foco a lo que se esté usando
        protected override bool ShowWithoutActivation { get { return captura != null; } }

        void V(string fila, string texto)
        {
            var lbl = valores[fila];
            if (lbl.Text != texto) lbl.Text = texto;
        }

        void MostrarAviso(string texto)
        {
            if (lblAviso.Text != texto) lblAviso.Text = texto;
            if (!lblAviso.Visible) lblAviso.Visible = true;
        }

        void Mostrar()
        {
            Show();
            if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
            Activate();
        }

        void AbrirCarpeta()
        {
            try { Process.Start("explorer.exe", "\"" + Program.DirSesion + "\""); } catch { }
        }

        void AlCambiarPantallas(object s, EventArgs e)
        {
            Program.Pantallas = Program.TextoPantallas();
            Program.Evento("PANTALLA", "Cambió la configuración de pantallas: " + Program.Pantallas);
        }

        void ReiniciarComoAdmin()
        {
            var si = new ProcessStartInfo(Application.ExecutablePath, "--esperar" + (Program.Intervalo != 1 ? " " + Program.Intervalo : ""));
            si.UseShellExecute = true;
            si.Verb = "runas";
            try { Process.Start(si); }
            catch (Win32Exception) { return; }   // se canceló el aviso de Windows: sigue como está
            Program.MotivoFin = "reiniciado como administrador";
            cerrarSinPreguntar = true;
            Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing && !cerrarSinPreguntar)
            {
                var r = MessageBox.Show(this, "¿Detener el monitor?\n\nDejará de grabar datos hasta que lo vuelvas a abrir.", "Monitor de batería",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (r != DialogResult.Yes) { e.Cancel = true; return; }
                Program.MotivoFin = "cerrado por el usuario";
            }
            else if (e.CloseReason == CloseReason.WindowsShutDown) Program.MotivoFin = "apagado, reinicio o cierre de sesión de Windows";
            else if (e.CloseReason == CloseReason.TaskManagerClosing) Program.MotivoFin = "cerrado desde el Administrador de tareas";
            base.OnFormClosing(e);
            if (e.Cancel) return;
            reloj.Stop();
            SystemEvents.DisplaySettingsChanged -= AlCambiarPantallas;
            bandeja.Visible = false;
            bandeja.Dispose();
            // Se cierran los registros ya: si Windows se está apagando puede no dar tiempo a más
            Program.Finalizar(Program.MotivoFin ?? "ventana cerrada");
        }

        static Icon Circulo(Color c)
        {
            using (var bmp = new Bitmap(32, 32))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.Transparent);
                    using (var b = new SolidBrush(c)) g.FillEllipse(b, 1, 1, 30, 30);
                    using (var b = new SolidBrush(Color.White)) g.FillEllipse(b, 11, 11, 10, 10);
                }
                return Icon.FromHandle(bmp.GetHicon());
            }
        }
    }

    // ---------------------------------------------------------------------- lectura de cada muestra

    sealed class Lectura
    {
        public DateTime Hora;
        public bool BatOk, Ac, Cargando, Descargando, Critico, AhorroBateria;
        public int PctWin = -1;
        public double PctCalc = double.NaN, CapMwh = double.NaN, CompletaMwh = double.NaN, TasaMw = double.NaN, VoltMv = double.NaN, TempBat = double.NaN;
        public double CpuUso, CpuRend, CpuMhz, WPaq, WNuc, WIgpu, WDram, TempAcpi, LimPasivo, RazTerm, RamLibre, RamComp;
        public double GW = double.NaN, GWInst = double.NaN, GLim = double.NaN, GUso = double.NaN, GMemUso = double.NaN, GTemp = double.NaN;
        public double GMhz = double.NaN, GMemMhz = double.NaN, GVram = double.NaN, CpuMasGpu = double.NaN;
        public double Fps = double.NaN, FtProm = double.NaN, FtMax = double.NaN, FtP99 = double.NaN;
        public string GP = "", GRazHex = "", GRaz = "", FpsTexto = "";
        public DatosMsi Msi;   // null si no hay datos de MSI recientes
    }

    // Valores del controlador de MSI. MSI no documenta las clases WMI: el significado de cada índice
    // está deducido comparándolos con los datos de Windows (voltaje, %, temperaturas de nvidia-smi).
    sealed class DatosMsi
    {
        public double Ma = double.NaN, TempC = double.NaN, RestanteMah = double.NaN, CompletaMah = double.NaN, DisenoMah = double.NaN, Pct = double.NaN;
        public double CpuC = double.NaN, GpuC = double.NaN, VentCpu = double.NaN, VentGpu = double.NaN;
        public long Ticks;

        public bool Vigente { get { return Ticks != 0 && DateTime.UtcNow.Ticks - Ticks < TimeSpan.FromSeconds(15).Ticks; } }

        public DatosMsi Copia() { return (DatosMsi)MemberwiseClone(); }

        public void Asignar(string clase, int indice, double v)
        {
            if (clase == "MSI_Master_Battery")
            {
                switch (indice)
                {
                    case 1: DisenoMah = v; break;
                    case 6: Pct = v; break;
                    case 7: CompletaMah = v; break;
                    case 8: Ma = v; break;              // negativa: la batería se descarga
                    case 9: RestanteMah = v; break;
                    case 11: if (v > 2000) TempC = v / 10.0 - 273.15; break;   // décimas de kelvin
                }
            }
            else if (clase == "MSI_CPU") { if (indice == 1) CpuC = v; else if (indice == 2) VentCpu = v; }
            else if (clase == "MSI_VGA") { if (indice == 1) GpuC = v; else if (indice == 2) VentGpu = v; }
        }
    }

    // Archivo CSV que se vuelca al disco en cada escritura importante
    sealed class Registro
    {
        readonly FileStream fs;
        readonly StreamWriter w;
        readonly object l = new object();
        bool cerrado;

        public Registro(string ruta, string cabecera)
        {
            bool nuevo = !File.Exists(ruta);
            fs = new FileStream(ruta, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
            w = new StreamWriter(fs, new UTF8Encoding(nuevo));
            if (nuevo) w.WriteLine(cabecera);
            Volcar();
        }

        public void Linea(string s) { lock (l) { if (!cerrado) w.WriteLine(s); } }

        // Flush(true) fuerza la escritura física: si se corta la luz, lo ya escrito no se pierde
        public void Volcar() { lock (l) { if (cerrado) return; try { w.Flush(); fs.Flush(true); } catch { } } }

        public void Cerrar() { lock (l) { if (cerrado) return; Volcar(); cerrado = true; try { w.Close(); } catch { } } }
    }

    // ---------------------------------------------------------------------- FPS

    // Cuenta los eventos Present de DXGI (DirectX 10/11/12) y D3D9 de cada proceso con una sesión
    // ETW en tiempo real, igual que PresentMon. Necesita permisos de administrador.
    sealed class Fps
    {
        const string NombreSesion = "MonitorBateriaFPS";
        const int TamPropiedades = 120 + 1024;   // EVENT_TRACE_PROPERTIES + nombre de la sesión
        const int TamLogfile = 448;               // EVENT_TRACE_LOGFILEW en x64
        static readonly Guid Dxgi = new Guid("CA11C036-0102-4A2D-A6AD-F03CFED5D3C9");
        static readonly Guid D3d9 = new Guid("783ACA0A-790E-4D7F-8451-AA850511C6B9");
        ulong sesion, consumidor = ulong.MaxValue;
        IntPtr props = IntPtr.Zero, logfile = IntPtr.Zero, nombre = IntPtr.Zero;
        Nativo.EventRecordCallback alEvento;   // referencia viva para que el GC no la recoja
        readonly object l = new object();
        Dictionary<int, Cuadros> cuadros = new Dictionary<int, Cuadros>();
        readonly Stopwatch reloj = Stopwatch.StartNew();
        double ultimaToma;
        public volatile bool Activo;
        public string Error = "";
        public long Eventos;

        sealed class Cuadros { public long Ultimo; public int N; public readonly List<double> Ms = new List<double>(); }
        public sealed class Estad { public int Pid, Cuadros; public double Fps, PromMs, MaxMs, P99Ms; }

        static IntPtr Propiedades()
        {
            IntPtr p = Marshal.AllocHGlobal(TamPropiedades);
            for (int i = 0; i < TamPropiedades; i++) Marshal.WriteByte(p, i, 0);
            Marshal.WriteInt32(p, 0, TamPropiedades);    // Wnode.BufferSize
            Marshal.WriteInt32(p, 40, 1);                // Wnode.ClientContext: reloj QPC (alta resolución)
            Marshal.WriteInt32(p, 44, 0x00020000);       // Wnode.Flags: WNODE_FLAG_TRACED_GUID
            Marshal.WriteInt32(p, 48, 64);               // BufferSize (KB)
            Marshal.WriteInt32(p, 52, 4);                // MinimumBuffers
            Marshal.WriteInt32(p, 56, 32);               // MaximumBuffers
            Marshal.WriteInt32(p, 64, 0x100);            // LogFileMode: EVENT_TRACE_REAL_TIME_MODE
            Marshal.WriteInt32(p, 68, 1);                // FlushTimer: 1 s
            Marshal.WriteInt32(p, 116, 120);             // LoggerNameOffset
            return p;
        }

        public bool Iniciar()
        {
            try
            {
                // Si el monitor se cerró a la fuerza, la sesión anterior puede seguir viva
                IntPtr vieja = Propiedades();
                try { Nativo.ControlTrace(0, NombreSesion, vieja, 1); } finally { Marshal.FreeHGlobal(vieja); }

                props = Propiedades();
                uint r = Nativo.StartTrace(out sesion, NombreSesion, props);
                if (r != 0) { Error = r == 5 ? "acceso denegado" : "StartTrace devolvió " + r; return false; }
                Guid g = Dxgi;
                r = Nativo.EnableTraceEx2(sesion, ref g, 1, 4, ulong.MaxValue, 0, 0, IntPtr.Zero);
                if (r != 0) { Error = "EnableTraceEx2 (DXGI) devolvió " + r; Detener(); return false; }
                g = D3d9;
                Nativo.EnableTraceEx2(sesion, ref g, 1, 4, ulong.MaxValue, 0, 0, IntPtr.Zero);

                logfile = Marshal.AllocHGlobal(TamLogfile);
                for (int i = 0; i < TamLogfile; i++) Marshal.WriteByte(logfile, i, 0);
                nombre = Marshal.StringToHGlobalUni(NombreSesion);
                Marshal.WriteIntPtr(logfile, 8, nombre);                  // LoggerName
                Marshal.WriteInt32(logfile, 28, 0x100 | 0x10000000);     // ProcessTraceMode: REAL_TIME | EVENT_RECORD
                alEvento = AlEvento;
                Marshal.WriteIntPtr(logfile, 424, Marshal.GetFunctionPointerForDelegate(alEvento));   // EventRecordCallback
                consumidor = Nativo.OpenTrace(logfile);
                if (consumidor == ulong.MaxValue) { Error = "OpenTrace falló (" + Marshal.GetLastWin32Error() + ")"; Detener(); return false; }
                ulong c = consumidor;
                var hilo = new Thread(() => { try { Nativo.ProcessTrace(new[] { c }, 1, IntPtr.Zero, IntPtr.Zero); } catch { } });
                hilo.IsBackground = true;
                hilo.Name = "etw";
                hilo.Start();
                Activo = true;
                return true;
            }
            catch (Exception ex) { Error = ex.Message; Detener(); return false; }
        }

        void AlEvento(IntPtr ev)
        {
            try
            {
                // EVENT_HEADER: ProcessId +12, TimeStamp +16, ProviderId +24, EventDescriptor.Id +40
                int prov = Marshal.ReadInt32(ev, 24);
                int id = Marshal.ReadInt16(ev, 40) & 0xFFFF;
                bool present = (prov == unchecked((int)0xCA11C036) && id == 42) || (prov == 0x783ACA0A && id == 1);   // Present_Start
                if (!present) return;
                int pid = Marshal.ReadInt32(ev, 12);
                long ts = Marshal.ReadInt64(ev, 16);   // FILETIME (100 ns)
                lock (l)
                {
                    Eventos++;
                    Cuadros c;
                    if (!cuadros.TryGetValue(pid, out c)) { c = new Cuadros(); cuadros[pid] = c; }
                    if (c.Ultimo != 0)
                    {
                        double ms = (ts - c.Ultimo) / 10000.0;
                        if (ms >= 0 && ms < 10000) c.Ms.Add(ms);
                    }
                    c.Ultimo = ts;
                    c.N++;
                }
            }
            catch { }
        }

        // Cuadros por proceso desde la llamada anterior
        public List<Estad> Tomar()
        {
            double ahora = reloj.Elapsed.TotalSeconds, seg = ahora - ultimaToma;
            ultimaToma = ahora;
            var res = new List<Estad>();
            lock (l)
            {
                var quietos = new List<int>();
                foreach (var kv in cuadros)
                {
                    Cuadros c = kv.Value;
                    if (c.N == 0) { quietos.Add(kv.Key); continue; }
                    var e = new Estad { Pid = kv.Key, Cuadros = c.N, Fps = seg > 0 ? c.N / seg : 0, PromMs = double.NaN, MaxMs = double.NaN, P99Ms = double.NaN };
                    if (c.Ms.Count > 0)
                    {
                        c.Ms.Sort();
                        e.PromMs = c.Ms.Average();
                        e.MaxMs = c.Ms[c.Ms.Count - 1];
                        e.P99Ms = c.Ms[Math.Max(0, Math.Min(c.Ms.Count - 1, (int)Math.Ceiling(c.Ms.Count * 0.99) - 1))];
                    }
                    res.Add(e);
                    c.N = 0;
                    c.Ms.Clear();
                }
                foreach (int pid in quietos) cuadros.Remove(pid);
            }
            return res.OrderByDescending(x => x.Cuadros).ToList();
        }

        public void Detener()
        {
            Activo = false;
            try { if (props != IntPtr.Zero && sesion != 0) Nativo.ControlTrace(sesion, null, props, 1); } catch { }   // EVENT_TRACE_CONTROL_STOP
            try { if (consumidor != ulong.MaxValue) Nativo.CloseTrace(consumidor); } catch { }
            sesion = 0;
            consumidor = ulong.MaxValue;
        }
    }

    // ---------------------------------------------------------------------- batería

    // Batería leída con las IOCTL del controlador de baterías de Windows
    sealed class Bateria
    {
        const uint IoctlTag = 0x294040, IoctlInfo = 0x294044, IoctlEstado = 0x29404C;
        SafeFileHandle h;
        uint tag;
        readonly IntPtr bufIn = Marshal.AllocHGlobal(64), bufOut = Marshal.AllocHGlobal(64);
        bool tempSoportada = true;
        public uint Disenada, Completa, Ciclos;
        public string Quimica = "";

        public bool Abierta { get { return h != null && !h.IsInvalid && tag != 0; } }

        public bool Abrir()
        {
            Cerrar();
            Guid clase = new Guid("72631E54-78A4-11D0-BCF7-00AA00B7B32A");   // GUID_DEVCLASS_BATTERY
            IntPtr devs = Nativo.SetupDiGetClassDevs(ref clase, IntPtr.Zero, IntPtr.Zero, 0x12);   // PRESENT | DEVICEINTERFACE
            if (devs == new IntPtr(-1)) return false;
            try
            {
                var d = new SP_DEVICE_INTERFACE_DATA();
                d.cbSize = Marshal.SizeOf(typeof(SP_DEVICE_INTERFACE_DATA));
                if (!Nativo.SetupDiEnumDeviceInterfaces(devs, IntPtr.Zero, ref clase, 0, ref d)) return false;
                int req;
                Nativo.SetupDiGetDeviceInterfaceDetail(devs, ref d, IntPtr.Zero, 0, out req, IntPtr.Zero);
                if (req <= 0) return false;
                IntPtr det = Marshal.AllocHGlobal(req);
                try
                {
                    Marshal.WriteInt32(det, IntPtr.Size == 8 ? 8 : 6);
                    if (!Nativo.SetupDiGetDeviceInterfaceDetail(devs, ref d, det, req, out req, IntPtr.Zero)) return false;
                    string ruta = Marshal.PtrToStringUni(new IntPtr(det.ToInt64() + 4));
                    h = Nativo.CreateFile(ruta, 0xC0000000, 3, IntPtr.Zero, 3, 0x80, IntPtr.Zero);
                }
                finally { Marshal.FreeHGlobal(det); }
            }
            finally { Nativo.SetupDiDestroyDeviceInfoList(devs); }
            if (h == null || h.IsInvalid) return false;
            Marshal.WriteInt32(bufIn, 0);
            int r;
            if (!Nativo.DeviceIoControl(h, IoctlTag, bufIn, 4, bufOut, 4, out r, IntPtr.Zero)) return false;
            tag = (uint)Marshal.ReadInt32(bufOut);
            return tag != 0 && LeerInfo();
        }

        bool Consultar(int nivel, int tamSalida)
        {
            Marshal.WriteInt32(bufIn, 0, (int)tag);
            Marshal.WriteInt32(bufIn, 4, nivel);
            Marshal.WriteInt32(bufIn, 8, 0);
            int r;
            return Nativo.DeviceIoControl(h, IoctlInfo, bufIn, 12, bufOut, tamSalida, out r, IntPtr.Zero);
        }

        public bool LeerInfo()
        {
            if (!Abierta || !Consultar(0, 36)) return false;   // BATTERY_INFORMATION (36 bytes)
            var q = new byte[4];
            Marshal.Copy(new IntPtr(bufOut.ToInt64() + 8), q, 0, 4);
            Quimica = Encoding.ASCII.GetString(q).TrimEnd('\0', ' ');
            Disenada = (uint)Marshal.ReadInt32(bufOut, 12);
            Completa = (uint)Marshal.ReadInt32(bufOut, 16);
            Ciclos = (uint)Marshal.ReadInt32(bufOut, 32);
            return true;
        }

        // °C, o NaN si la batería no la informa (muchas no lo hacen)
        public double Temperatura()
        {
            if (!tempSoportada || !Abierta) return double.NaN;
            if (!Consultar(2, 4)) { tempSoportada = false; return double.NaN; }
            uint t = (uint)Marshal.ReadInt32(bufOut);
            if (t == 0) { tempSoportada = false; return double.NaN; }
            return t / 10.0 - 273.15;
        }

        public bool LeerEstado(out uint estado, out uint capacidad, out uint voltaje, out int tasa)
        {
            estado = 0; capacidad = 0; voltaje = 0; tasa = 0;
            if (!Abierta && !Abrir()) return false;
            for (int i = 0; i < 20; i += 4) Marshal.WriteInt32(bufIn, i, 0);
            Marshal.WriteInt32(bufIn, 0, (int)tag);   // BATTERY_WAIT_STATUS con espera 0
            int r;
            if (!Nativo.DeviceIoControl(h, IoctlEstado, bufIn, 20, bufOut, 16, out r, IntPtr.Zero))
            {
                Cerrar();   // el tag cambia si la batería "se reconecta": se reabre en la próxima muestra
                return false;
            }
            estado = (uint)Marshal.ReadInt32(bufOut, 0);
            capacidad = (uint)Marshal.ReadInt32(bufOut, 4);
            voltaje = (uint)Marshal.ReadInt32(bufOut, 8);
            tasa = Marshal.ReadInt32(bufOut, 12);
            return true;
        }

        public void Cerrar()
        {
            if (h != null) { h.Dispose(); h = null; }
            tag = 0;
        }
    }

    sealed class Contadores
    {
        readonly PerformanceCounter cpuUso, cpuRend, tempZona, limPasivo, razones, ramLibre, ramComp;
        readonly PerformanceCounter[] energia = new PerformanceCounter[4];
        public readonly double MhzBase;

        public Contadores()
        {
            cpuUso = Crear("Processor Information", "% Processor Utility", "_Total");
            cpuRend = Crear("Processor Information", "% Processor Performance", "_Total");
            MhzBase = Valor(Crear("Processor Information", "Processor Frequency", "_Total"));
            string[] rapl = { "RAPL_Package0_PKG", "RAPL_Package0_PP0", "RAPL_Package0_PP1", "RAPL_Package0_DRAM" };
            for (int i = 0; i < rapl.Length; i++) energia[i] = Crear("Energy Meter", "Power", rapl[i]);
            string zona = PrimeraInstancia("Thermal Zone Information");
            if (zona != null)
            {
                tempZona = Crear("Thermal Zone Information", "High Precision Temperature", zona);
                limPasivo = Crear("Thermal Zone Information", "% Passive Limit", zona);
                razones = Crear("Thermal Zone Information", "Throttle Reasons", zona);
            }
            ramLibre = Crear("Memory", "Available MBytes", "");
            ramComp = Crear("Memory", "% Committed Bytes In Use", "");
        }

        static PerformanceCounter Crear(string cat, string nombre, string inst)
        {
            try { var p = new PerformanceCounter(cat, nombre, inst, true); p.NextValue(); return p; }
            catch { return null; }
        }

        static string PrimeraInstancia(string cat)
        {
            try { return new PerformanceCounterCategory(cat).GetInstanceNames().OrderBy(x => x).FirstOrDefault(); }
            catch { return null; }
        }

        static double Valor(PerformanceCounter p)
        {
            if (p == null) return double.NaN;
            try { return p.NextValue(); } catch { return double.NaN; }
        }

        public double CpuUso() { return Valor(cpuUso); }
        public double CpuRendimiento() { return Valor(cpuRend); }
        public double Potencia(int i) { return Valor(energia[i]) / 1000.0; }          // mW -> W
        public double TempAcpi() { return Valor(tempZona) / 10.0 - 273.15; }          // décimas de kelvin -> °C
        public double LimitePasivo() { return Valor(limPasivo); }
        public double RazonesTermicas() { return Valor(razones); }
        public double RamLibre() { return Valor(ramLibre); }
        public double RamComprometida() { return Valor(ramComp); }
    }

    // nvidia-smi en modo bucle (-lms): un solo proceso que escribe una línea por intervalo
    sealed class Nvidia
    {
        const string Campos = "power.draw,enforced.power.limit,utilization.gpu,utilization.memory,temperature.gpu," +
            "clocks.gr,clocks.mem,pstate,clocks_event_reasons.active,memory.used";
        const string Instantaneo = ",power.draw.instant";   // no existe en controladores antiguos
        readonly IntPtr job;
        Process p;
        volatile string ultima;
        long ultimaTicks;
        int lineas;
        bool conInstantaneo = true;
        DateTime proximoIntento = DateTime.MinValue;

        public Nvidia(IntPtr job) { this.job = job; }

        public void Asegurar(int intervaloMs)
        {
            try { if (p != null && !p.HasExited) return; } catch { }
            if (p != null && lineas == 0) conInstantaneo = false;   // terminó sin dar datos: probar sin el campo nuevo
            if (DateTime.Now < proximoIntento) return;
            proximoIntento = DateTime.Now.AddSeconds(30);
            try
            {
                string campos = Campos + (conInstantaneo ? Instantaneo : "");
                var si = new ProcessStartInfo("nvidia-smi.exe", "--query-gpu=" + campos + " --format=csv,noheader,nounits -lms " + intervaloMs);
                si.UseShellExecute = false;
                si.RedirectStandardOutput = true;
                si.RedirectStandardError = true;
                si.CreateNoWindow = true;
                var nuevo = new Process();
                nuevo.StartInfo = si;
                lineas = 0;
                nuevo.OutputDataReceived += (s, e) =>
                {
                    if (string.IsNullOrEmpty(e.Data)) return;
                    ultima = e.Data;
                    lineas++;
                    Interlocked.Exchange(ref ultimaTicks, DateTime.Now.Ticks);
                };
                nuevo.ErrorDataReceived += (s, e) => { };
                nuevo.Start();
                try { nuevo.PriorityClass = ProcessPriorityClass.BelowNormal; } catch { }
                if (job != IntPtr.Zero) Nativo.AssignProcessToJobObject(job, nuevo.Handle);
                nuevo.BeginOutputReadLine();
                nuevo.BeginErrorReadLine();
                p = nuevo;
            }
            catch { p = null; }
        }

        // Campos de la última lectura, o null si es vieja (más de 6 s) o está incompleta
        public string[] Leer()
        {
            string l = ultima;
            if (l == null || (DateTime.Now.Ticks - Interlocked.Read(ref ultimaTicks)) > TimeSpan.FromSeconds(6).Ticks) return null;
            string[] c = l.Split(',').Select(x => x.Trim()).ToArray();
            return c.Length >= 10 ? c : null;
        }

        public int Pid { get { try { return p != null && !p.HasExited ? p.Id : 0; } catch { return 0; } } }

        public void Parar() { try { if (p != null && !p.HasExited) p.Kill(); } catch { } }
    }

    sealed class UsoGpu { public double Pct; public string Motor; public bool Dgpu; }

    // Uso de GPU por proceso con los contadores "GPU Engine" (los mismos que el Administrador de tareas)
    sealed class GpuMotores
    {
        readonly PerformanceCounterCategory motores, memoria;
        Dictionary<string, long> previo = new Dictionary<string, long>();
        long previoTs;
        string luidDgpu = "";
        DateTime proximoLuid = DateTime.MinValue;

        public GpuMotores()
        {
            try { motores = new PerformanceCounterCategory("GPU Engine"); memoria = new PerformanceCounterCategory("GPU Adapter Memory"); }
            catch { }
        }

        // El adaptador con más memoria dedicada en uso es la NVIDIA; la Intel apenas tiene
        void ActualizarLuid()
        {
            if (memoria == null || DateTime.Now < proximoLuid) return;
            proximoLuid = DateTime.Now.AddMinutes(1);
            try
            {
                long max = -1;
                foreach (InstanceData d in memoria.ReadCategory()["dedicated usage"].Values)
                    if (d.RawValue > max) { max = d.RawValue; luidDgpu = Luid(d.InstanceName); }
            }
            catch { }
        }

        static string Luid(string inst)
        {
            int i = inst.IndexOf("luid_", StringComparison.OrdinalIgnoreCase);
            if (i < 0) return "";
            string[] p = inst.Substring(i + 5).Split('_');
            return p.Length >= 2 ? p[1].ToLowerInvariant() : "";
        }

        public Dictionary<int, UsoGpu> Leer()
        {
            var res = new Dictionary<int, UsoGpu>();
            if (motores == null) return res;
            ActualizarLuid();
            InstanceDataCollection col;
            try { col = motores.ReadCategory()["running time"]; }
            catch { return res; }
            var actual = new Dictionary<string, long>(col.Count);
            var sumas = new Dictionary<string, long>();
            long ts = 0;
            foreach (InstanceData d in col.Values)
            {
                string n = d.InstanceName;
                actual[n] = d.RawValue;
                ts = d.Sample.TimeStamp100nSec;
                long ant;
                if (!previo.TryGetValue(n, out ant) || d.RawValue <= ant) continue;
                int e = n.IndexOf("_engtype_", StringComparison.OrdinalIgnoreCase);
                int finPid = n.IndexOf('_', 4);
                if (!n.StartsWith("pid_", StringComparison.OrdinalIgnoreCase) || e < 0 || finPid < 0) continue;
                string clave = n.Substring(4, finPid - 4) + "|" + Luid(n) + "|" + n.Substring(e + 9);
                long s;
                sumas.TryGetValue(clave, out s);
                sumas[clave] = s + (d.RawValue - ant);   // "Running Time" va en unidades de 100 ns
            }
            long dt = ts - previoTs;
            bool valido = previoTs != 0 && dt > 0;
            previo = actual;
            previoTs = ts;
            if (!valido) return res;
            foreach (var kv in sumas)
            {
                string[] p = kv.Key.Split('|');
                int pid;
                if (!int.TryParse(p[0], out pid)) continue;
                double pct = Math.Min(100, kv.Value * 100.0 / dt);
                UsoGpu u;
                if (!res.TryGetValue(pid, out u) || pct > u.Pct)
                    res[pid] = new UsoGpu { Pct = pct, Motor = p[2], Dgpu = p[1] == luidDgpu };
            }
            return res;
        }
    }

    sealed class InfoProc
    {
        public int Pid, PidPadre, Hilos;
        public string Nombre;
        public long Creacion, Cpu, WorkingSet, Privada;
        public double CpuPct;
    }

    // Lee todos los procesos de una sola vez con NtQuerySystemInformation, como el Administrador de
    // tareas: no abre cada proceso, así que es barato y también ve los procesos protegidos.
    sealed class Procesos
    {
        IntPtr buf = IntPtr.Zero;
        int tam;
        Dictionary<int, InfoProc> previo = new Dictionary<int, InfoProc>();
        long previoTicks;
        public volatile Dictionary<int, InfoProc> Actual = new Dictionary<int, InfoProc>();
        public readonly List<InfoProc> Nuevos = new List<InfoProc>(), Terminados = new List<InfoProc>();

        public void Muestrear()
        {
            List<InfoProc> lista = Leer();
            if (lista.Count == 0) return;
            long ahora = DateTime.UtcNow.Ticks;
            double dt = previoTicks == 0 ? 0 : (double)(ahora - previoTicks) * Environment.ProcessorCount;
            var dic = new Dictionary<int, InfoProc>(lista.Count);
            Nuevos.Clear();
            Terminados.Clear();
            foreach (InfoProc p in lista)
            {
                InfoProc a;
                if (previo.TryGetValue(p.Pid, out a) && a.Creacion == p.Creacion)
                {
                    if (dt > 0) p.CpuPct = Math.Max(0, (p.Cpu - a.Cpu) * 100.0 / dt);
                }
                else if (previoTicks != 0) Nuevos.Add(p);
                dic[p.Pid] = p;
            }
            if (previoTicks != 0)
                foreach (InfoProc a in previo.Values)
                {
                    InfoProc n;
                    if (!dic.TryGetValue(a.Pid, out n) || n.Creacion != a.Creacion) Terminados.Add(a);
                }
            previo = dic;
            Actual = dic;
            previoTicks = ahora;
        }

        List<InfoProc> Leer()
        {
            var lista = new List<InfoProc>(400);
            if (buf == IntPtr.Zero) { tam = 1 << 20; buf = Marshal.AllocHGlobal(tam); }
            int ret, st;
            while ((st = Nativo.NtQuerySystemInformation(5, buf, tam, out ret)) == unchecked((int)0xC0000004))   // STATUS_INFO_LENGTH_MISMATCH
            {
                Marshal.FreeHGlobal(buf);
                tam = Math.Max(tam * 2, ret + 65536);
                buf = Marshal.AllocHGlobal(tam);
            }
            if (st != 0) return lista;
            long off = 0;
            while (true)
            {
                // Desplazamientos de SYSTEM_PROCESS_INFORMATION en x64
                IntPtr p = new IntPtr(buf.ToInt64() + off);
                var ip = new InfoProc();
                int siguiente = Marshal.ReadInt32(p, 0x00);
                ip.Hilos = Marshal.ReadInt32(p, 0x04);
                ip.Creacion = Marshal.ReadInt64(p, 0x20);
                ip.Cpu = Marshal.ReadInt64(p, 0x28) + Marshal.ReadInt64(p, 0x30);   // usuario + núcleo, en 100 ns
                int largo = Marshal.ReadInt16(p, 0x38) & 0xFFFF;
                IntPtr nombre = Marshal.ReadIntPtr(p, 0x40);
                ip.Pid = (int)Marshal.ReadIntPtr(p, 0x50).ToInt64();
                ip.PidPadre = (int)Marshal.ReadIntPtr(p, 0x58).ToInt64();
                ip.WorkingSet = Marshal.ReadIntPtr(p, 0x90).ToInt64();
                ip.Privada = Marshal.ReadIntPtr(p, 0xC8).ToInt64();
                ip.Nombre = nombre != IntPtr.Zero && largo > 0 ? Marshal.PtrToStringUni(nombre, largo / 2) : (ip.Pid == 0 ? "Idle" : "System");
                lista.Add(ip);
                if (siguiente == 0) break;
                off += siguiente;
            }
            return lista;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    struct SP_DEVICE_INTERFACE_DATA { public int cbSize; public Guid InterfaceClassGuid; public int Flags; public IntPtr Reserved; }

    [StructLayout(LayoutKind.Sequential)]
    struct SYSTEM_POWER_STATUS { public byte ACLineStatus, BatteryFlag, BatteryLifePercent, SystemStatusFlag; public int BatteryLifeTime, BatteryFullLifeTime; }

    static class Nativo
    {
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        public delegate void EventRecordCallback(IntPtr evento);

        [DllImport("setupapi.dll", SetLastError = true)]
        public static extern IntPtr SetupDiGetClassDevs(ref Guid clase, IntPtr enumerador, IntPtr ventana, uint flags);
        [DllImport("setupapi.dll", SetLastError = true)]
        public static extern bool SetupDiEnumDeviceInterfaces(IntPtr devs, IntPtr info, ref Guid clase, uint indice, ref SP_DEVICE_INTERFACE_DATA datos);
        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool SetupDiGetDeviceInterfaceDetail(IntPtr devs, ref SP_DEVICE_INTERFACE_DATA datos, IntPtr detalle, int tam, out int requerido, IntPtr info);
        [DllImport("setupapi.dll")]
        public static extern bool SetupDiDestroyDeviceInfoList(IntPtr devs);
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern SafeFileHandle CreateFile(string nombre, uint acceso, uint compartir, IntPtr seg, uint disp, uint flags, IntPtr plantilla);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool DeviceIoControl(SafeFileHandle h, uint codigo, IntPtr entrada, int tamEntrada, IntPtr salida, int tamSalida, out int devuelto, IntPtr overlapped);
        [DllImport("kernel32.dll")]
        public static extern bool GetSystemPowerStatus(out SYSTEM_POWER_STATUS estado);
        [DllImport("ntdll.dll")]
        public static extern int NtQuerySystemInformation(int clase, IntPtr buffer, int tam, out int devuelto);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr OpenProcess(uint acceso, bool heredar, int pid);
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool QueryFullProcessImageName(IntPtr h, int flags, StringBuilder ruta, ref int tam);
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr h);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr CreateJobObject(IntPtr seg, string nombre);
        [DllImport("kernel32.dll")]
        public static extern bool SetInformationJobObject(IntPtr job, int clase, IntPtr info, int tam);
        [DllImport("kernel32.dll")]
        public static extern bool AssignProcessToJobObject(IntPtr job, IntPtr proceso);
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "StartTraceW")]
        public static extern uint StartTrace(out ulong sesion, string nombre, IntPtr propiedades);
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "ControlTraceW")]
        public static extern uint ControlTrace(ulong sesion, string nombre, IntPtr propiedades, uint control);
        [DllImport("advapi32.dll")]
        public static extern uint EnableTraceEx2(ulong sesion, ref Guid proveedor, uint control, byte nivel, ulong algunaPalabra, ulong todasPalabras, uint espera, IntPtr parametros);
        [DllImport("advapi32.dll", SetLastError = true, EntryPoint = "OpenTraceW")]
        public static extern ulong OpenTrace(IntPtr logfile);
        [DllImport("advapi32.dll")]
        public static extern uint ProcessTrace(ulong[] sesiones, uint cuantas, IntPtr inicio, IntPtr fin);
        [DllImport("advapi32.dll")]
        public static extern uint CloseTrace(ulong sesion);
    }
}
