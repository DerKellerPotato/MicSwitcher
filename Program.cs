// Program.cs
// ALLES IN EINER DATEI, Version 14 - Config-Speicherort korrigiert & Live-Neustart der Erfassung

#nullable disable // Deaktiviert strenge Null-Prüfungen für diese Datei

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces; // Wichtig für IMMNotificationClient
using System.Windows.Forms;
using System.Drawing;

namespace MicSwitcher
{
    // ####################################################################
    // ## Logger-Klasse (Unverändert)
    // ####################################################################
    public static class Log
    {
        private static string _logFile = Path.Combine(AppContext.BaseDirectory, "MicSwitcher.log");
        private static readonly object _lock = new object();
        static Log()
        {
            try { File.WriteAllText(_logFile, $"--- MicSwitcher Log --- {DateTime.Now} ---\n"); }
            catch (Exception ex) { Debug.WriteLine($"Konnte Log-Datei nicht erstellen: {ex.Message}"); }
        }
        public static void Info(string message)
        {
            try { lock (_lock) { File.AppendAllText(_logFile, $"[INFO] {message}\n"); } } catch { }
        }
        public static void Error(string message, Exception ex = null)
        {
            try
            {
                lock (_lock)
                {
                    File.AppendAllText(_logFile, $"[ERROR] {message}\n");
                    if (ex != null)
                    {
                        File.AppendAllText(_logFile, $"        {ex.Message}\n");
                        File.AppendAllText(_logFile, $"        {ex.StackTrace}\n");
                    }
                }
            }
            catch { }
        }
    }

    // ####################################################################
    // ## DATEI 1: AudioLogic.cs (Mit Fixes)
    // ####################################################################

    public class AppSettings
    {
        public string? StandMicId { get; set; }
        public List<string> HeadsetMicIds { get; set; } = new List<string>();
        public bool Autostart { get; set; }
    }

    public class MicSwitcherLogic : IDisposable
    {
        public event Action<string> OnDeviceSwitched;
        private MMDeviceEnumerator _deviceEnumerator = new MMDeviceEnumerator();
        private MuteNotificationClient _notificationClient;
        private AppSettings _settings = new AppSettings();
        private string _configFilePath;
        private string? _lastActiveHeadset = null;
        private bool _isCurrentlySwitching = false;

        public MicSwitcherLogic()
        {
            Log.Info("AudioLogic: Initialisierung...");

            // KORREKTUR: Verwende einen sicheren Speicherort im Benutzerprofil
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appFolderPath = Path.Combine(appDataPath, "MicSwitcher");
            Directory.CreateDirectory(appFolderPath); // Stellt sicher, dass der Ordner existiert
            _configFilePath = Path.Combine(appFolderPath, "MicSwitcher_config.json");
            Log.Info($"AudioLogic: Config-Pfad ist: {_configFilePath}");

            _notificationClient = new MuteNotificationClient(this);
            _deviceEnumerator.RegisterEndpointNotificationCallback(_notificationClient);
            Log.Info("AudioLogic: Initialisierung abgeschlossen.");
        }

        public void LoadSettings()
        {
            try
            {
                if (File.Exists(_configFilePath))
                {
                    string json = File.ReadAllText(_configFilePath);
                    _settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                    Log.Info("AudioLogic: Einstellungen geladen.");
                }
                else
                {
                    Log.Info("AudioLogic: Keine Konfigurationsdatei gefunden, verwende Standard.");
                }
            }
            catch (Exception ex)
            {
                Log.Error("Fehler beim Laden der Config", ex);
                _settings = new AppSettings();
            }
        }

        public void SaveSettings(AppSettings settings)
        {
            try
            {
                _settings = settings;
                string json = JsonSerializer.Serialize(_settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_configFilePath, json);
                SetAutostart(_settings.Autostart);
                Log.Info("AudioLogic: Einstellungen gespeichert.");
            }
            catch (Exception ex)
            {
                Log.Error("Fehler beim Speichern der Config", ex);
            }
        }

        public AppSettings GetCurrentSettings() => _settings;

        public Dictionary<string, string> GetAllMicrophones()
        {
            var devices = new Dictionary<string, string>();
            try
            {
                foreach (var dev in _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
                {
                    devices[dev.ID] = dev.FriendlyName;
                }
            }
            catch (Exception ex)
            {
                Log.Error("Fehler beim Auflisten der Mikrofone", ex);
            }
            return devices;
        }

        // NEU: Funktion zum Stoppen der Überwachung
        public void StopMonitoring()
        {
            Log.Info("AudioLogic: Stoppe Statuserfassung (Deregistrierung)...");
            foreach (string headsetId in _settings.HeadsetMicIds)
            {
                try
                {
                    var device = _deviceEnumerator.GetDevice(headsetId);
                    device.AudioEndpointVolume.OnVolumeNotification -= OnAudioEndpointVolumeChanged;
                    Log.Info($"AudioLogic: Listener für '{device.FriendlyName}' entfernt.");
                }
                catch (Exception ex)
                {
                    // Gerät vielleicht nicht mehr angeschlossen, das ist OK.
                    Log.Error($"Fehler beim Deregistrieren von {headsetId} (vielleicht abgesteckt?)", ex);
                }
            }
        }

        public void StartMonitoring()
        {
            Log.Info("AudioLogic: Starte Statuserfassung...");
            _lastActiveHeadset = null; // Zurücksetzen

            foreach (string headsetId in _settings.HeadsetMicIds)
            {
                try
                {
                    var device = _deviceEnumerator.GetDevice(headsetId);
                    // Sicherstellen, dass der Listener nicht doppelt hinzugefügt wird
                    device.AudioEndpointVolume.OnVolumeNotification -= OnAudioEndpointVolumeChanged;
                    device.AudioEndpointVolume.OnVolumeNotification += OnAudioEndpointVolumeChanged;
                    Log.Info($"AudioLogic: Listener für '{device.FriendlyName}' registriert.");
                    if (!device.AudioEndpointVolume.Mute)
                    {
                        _lastActiveHeadset = device.ID;
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"Fehler beim Registrieren von {headsetId}", ex);
                }
            }
            SetInitialDevice();
        }

        private void SetInitialDevice()
        {
            if (_lastActiveHeadset != null)
            {
                Log.Info($"Initial: Headset {_lastActiveHeadset} ist aktiv.");
                SwitchToDevice(_lastActiveHeadset);
            }
            else if (!string.IsNullOrEmpty(_settings.StandMicId))
            {
                Log.Info("Initial: Alle Headsets stumm. Wechsle zu Stand-Mikro.");
                SwitchToDevice(_settings.StandMicId);
            }
        }

        private void OnAudioEndpointVolumeChanged(AudioVolumeNotificationData data)
        {
            Log.Info(">>> MUTE/VOLUME EVENT ERKANNT! <<<");
            Log.Info($"    -> Mute-Status: {data.Muted}, MasterVolume: {data.MasterVolume}");

            if (_isCurrentlySwitching)
            {
                Log.Info("    -> Event ignoriert (isCurrentlySwitching = true).");
                return;
            }

            foreach (string headsetId in _settings.HeadsetMicIds)
            {
                try
                {
                    var device = _deviceEnumerator.GetDevice(headsetId);
                    if (device.AudioEndpointVolume.Mute) // Prüfe Mute-Status
                    {
                        if (_lastActiveHeadset == headsetId)
                        {
                            Log.Info($"Aktion: Headset {device.FriendlyName} wurde gemutet. Wechsle zu Stand-Mikro.");
                            _lastActiveHeadset = null;
                            SwitchToDevice(_settings.StandMicId);
                            break;
                        }
                    }
                    else // Nicht gemutet
                    {
                        if (_lastActiveHeadset != headsetId)
                        {
                            Log.Info($"Aktion: Headset {device.FriendlyName} wurde entmutet. Wechsle zu diesem Headset.");
                            _lastActiveHeadset = headsetId;
                            SwitchToDevice(headsetId);
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"Fehler in OnAudioEndpointVolumeChanged Schleife", ex);
                }
            }
        }

        public void SwitchToDevice(string? deviceId)
        {
            if (string.IsNullOrEmpty(deviceId))
            {
                Log.Error("SwitchToDevice: FEHLER, keine deviceId übergeben.", null);
                return;
            }

            _isCurrentlySwitching = true;
            string deviceName = "Unbekannt";
            try
            {
                // Erstelle den Client hier, auf dem AKTUELLEN Thread.
                PolicyConfigClient policyConfigClient = new PolicyConfigClient();

                deviceName = _deviceEnumerator.GetDevice(deviceId).FriendlyName;
                Log.Info($"SwitchToDevice: Versuche Umschalten auf: {deviceName}");

                policyConfigClient.SetDefaultEndpoint(deviceId, ERole.eConsole);
                policyConfigClient.SetDefaultEndpoint(deviceId, ERole.eCommunications);

                Log.Info($"SwitchToDevice: Umschalten auf {deviceName} ERFOLGREICH.");
                OnDeviceSwitched?.Invoke(deviceName);
            }
            catch (Exception ex)
            {
                Log.Error($"SwitchToDevice: FEHLER beim Umschalten auf {deviceName}", ex);
            }

            Task.Delay(100).ContinueWith(_ =>
            {
                _isCurrentlySwitching = false;
                Log.Info("SwitchToDevice: Switching-Sperre aufgehoben.");
            });
        }

        private void SetAutostart(bool enable)
        {
            try
            {
                string appPath = Process.GetCurrentProcess().MainModule?.FileName;
                if (string.IsNullOrEmpty(appPath)) return;

                RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                if (rk == null) return;

                if (enable)
                {
                    rk.SetValue(AppDomain.CurrentDomain.FriendlyName, appPath);
                    Log.Info("Autostart aktiviert.");
                }
                else
                {
                    rk.DeleteValue(AppDomain.CurrentDomain.FriendlyName, false);
                    Log.Info("Autostart deaktiviert.");
                }
            }
            catch (Exception ex)
            {
                Log.Error("Fehler bei Autostart-Registry", ex);
            }
        }

        public void Dispose()
        {
            Log.Info("AudioLogic: Dispose aufgerufen.");
            StopMonitoring(); // Stellt sicher, dass alle Listener entfernt werden
            _deviceEnumerator.UnregisterEndpointNotificationCallback(_notificationClient);
            _deviceEnumerator.Dispose();
        }
    }

    // --- COM-Klassen und Listener (Unverändert) ---

    internal class MuteNotificationClient : NAudio.CoreAudioApi.Interfaces.IMMNotificationClient
    {
        private MicSwitcherLogic _logic;
        public MuteNotificationClient(MicSwitcherLogic logic)
        {
            _logic = logic;
        }
        public void OnPropertyValueChanged(string pwstrDeviceId, PropertyKey key)
        {
            Log.Info($"IMMNotificationClient: OnPropertyValueChanged für Gerät {pwstrDeviceId}");
        }
        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string pwstrDefaultDeviceId)
        {
            Log.Info($"IMMNotificationClient: OnDefaultDeviceChanged");
        }
        public void OnDeviceAdded(string pwstrDeviceId)
        {
            Log.Info($"IMMNotificationClient: OnDeviceAdded");
        }
        public void OnDeviceRemoved(string pwstrDeviceId)
        {
            Log.Info($"IMMNotificationClient: OnDeviceRemoved");
        }
        public void OnDeviceStateChanged(string pwstrDeviceId, DeviceState dwNewState)
        {
            Log.Info($"IMMNotificationClient: OnDeviceStateChanged");
        }
        public void OnDeviceQueryRemove() { }
        public void OnDeviceQueryRemoveFailed() { }
        public void OnDeviceRemovePending() { }
    }

    [ComImport, Guid("F8679F50-850A-41CF-9C72-430F290290C8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPolicyConfig
    {
        [PreserveSig] int GetMixFormat(string pwszDeviceId, [Out] out IntPtr ppFormat);
        [PreserveSig] int GetDeviceFormat(string pwszDeviceId, bool bDefault, [Out] out IntPtr ppFormat);
        [PreserveSig] int ResetDeviceFormat(string pwszDeviceId);
        [PreserveSig] int SetDeviceFormat(string pwszDeviceId, IntPtr pEndpointFormat, IntPtr pMixFormat);
        [PreserveSig] int GetProcessingPeriod(string pwszDeviceId, bool bDefault, [Out] out long pmftDefaultPeriod, [Out] out long pmftMinimumPeriod);
        [PreserveSig] int SetProcessingPeriod(string pwszDeviceId, ref long pmftPeriod);
        [PreserveSig] int GetShareMode(string pwszDeviceId, [Out] out int pMode);
        [PreserveSig] int SetShareMode(string pwszDeviceId, int mode);
        [PreserveSig] int GetPropertyValue(string pwszDeviceId, ref PropertyKey pKey, [Out] out PropVariant pValue);
        [PreserveSig] int SetPropertyValue(string pwszDeviceId, ref PropertyKey pKey, PropVariant pValue);
        [PreserveSig] int SetDefaultEndpoint(string pwszDeviceId, ERole role);
        [PreserveSig] int SetEndpointVisibility(string pwszDeviceId, bool bVisible);
    }

    [ComImport, Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
    internal class CPolicyConfigClient { }

    internal class PolicyConfigClient
    {
        public void SetDefaultEndpoint(string deviceId, ERole role)
        {
            IPolicyConfig policyConfig = (IPolicyConfig)new CPolicyConfigClient();
            Marshal.ThrowExceptionForHR(policyConfig.SetDefaultEndpoint(deviceId, role));
        }
    }

    internal enum ERole { eConsole = 0, eMultimedia = 1, eCommunications = 2, eRole_Count = 3 }

    [StructLayout(LayoutKind.Explicit)]
    internal struct PropVariant { [FieldOffset(0)] public short vt; [FieldOffset(8)] public IntPtr pszVal; }


    // ####################################################################
    // ## DATEI 2: SettingsForm.cs (Mit Live-Update)
    // ####################################################################

    public partial class SettingsForm : Form
    {
        private MicSwitcherLogic _logic;
        private Dictionary<string, string> _allMics = new Dictionary<string, string>();
        private AppSettings _settings;
        private System.Windows.Forms.Label labelStandMic;
        private System.Windows.Forms.ComboBox comboStandMic;
        private System.Windows.Forms.Label labelHeadsets;
        private System.Windows.Forms.CheckedListBox listHeadsets;
        private System.Windows.Forms.CheckBox checkAutostart;
        private System.Windows.Forms.Button btnSave;
        private System.ComponentModel.IContainer components = null;

        public SettingsForm(MicSwitcherLogic logic)
        {
            InitializeComponent();
            _logic = logic;
            _settings = _logic.GetCurrentSettings();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            PopulateMicLists();
            LoadSettingsToUI();
        }

        private void PopulateMicLists()
        {
            _allMics = _logic.GetAllMicrophones();
            comboStandMic.DataSource = new BindingSource(_allMics, null);
            comboStandMic.DisplayMember = "Value";
            comboStandMic.ValueMember = "Key";
            listHeadsets.DataSource = new BindingSource(_allMics, null);
            listHeadsets.DisplayMember = "Value";
            listHeadsets.ValueMember = "Key";
        }

        private void LoadSettingsToUI()
        {
            if (!string.IsNullOrEmpty(_settings.StandMicId) && _allMics.ContainsKey(_settings.StandMicId))
            {
                comboStandMic.SelectedValue = _settings.StandMicId;
            }
            for (int i = 0; i < listHeadsets.Items.Count; i++)
            {
                if (listHeadsets.Items[i] is KeyValuePair<string, string> item)
                {
                    if (_settings.HeadsetMicIds.Contains(item.Key))
                    {
                        listHeadsets.SetItemChecked(i, true);
                    }
                }
            }
            checkAutostart.Checked = _settings.Autostart;
        }

        // KORREKTUR: Speichern-Button startet die Logik jetzt live neu
        private void btnSave_Click(object sender, EventArgs e)
        {
            Log.Info("Einstellungen: Speichern geklickt.");
            var newSettings = new AppSettings();
            if (comboStandMic.SelectedValue != null)
            {
                newSettings.StandMicId = comboStandMic.SelectedValue.ToString();
            }
            foreach (var item in listHeadsets.CheckedItems)
            {
                if (item is KeyValuePair<string, string> mic)
                {
                    newSettings.HeadsetMicIds.Add(mic.Key);
                }
            }
            newSettings.Autostart = checkAutostart.Checked;

            Log.Info("Einstellungen: Stoppe alte Statuserfassung...");
            _logic.StopMonitoring(); // Alte Listener entfernen

            _logic.SaveSettings(newSettings); // Neue Einstellungen speichern

            Log.Info("Einstellungen: Starte Statuserfassung mit neuen Einstellungen...");
            _logic.StartMonitoring(); // Neue Listener registrieren

            MessageBox.Show("Einstellungen wurden übernommen.", "Gespeichert", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Hide();
        }

        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.labelStandMic = new System.Windows.Forms.Label();
            this.comboStandMic = new System.Windows.Forms.ComboBox();
            this.labelHeadsets = new System.Windows.Forms.Label();
            this.listHeadsets = new System.Windows.Forms.CheckedListBox();
            this.checkAutostart = new System.Windows.Forms.CheckBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelStandMic
            // 
            this.labelStandMic.AutoSize = true;
            this.labelStandMic.Location = new System.Drawing.Point(12, 15);
            this.labelStandMic.Name = "labelStandMic";
            this.labelStandMic.Size = new System.Drawing.Size(117, 20);
            this.labelStandMic.Text = "Stand-Mikrofon:";
            // 
            // comboStandMic
            // 
            this.comboStandMic.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboStandMic.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboStandMic.FormattingEnabled = true;
            this.comboStandMic.Location = new System.Drawing.Point(16, 38);
            this.comboStandMic.Name = "comboStandMic";
            this.comboStandMic.Size = new System.Drawing.Size(455, 28);
            // 
            // labelHeadsets
            // 
            this.labelHeadsets.AutoSize = true;
            this.labelHeadsets.Location = new System.Drawing.Point(12, 80);
            this.labelHeadsets.Name = "labelHeadsets";
            this.labelHeadsets.Size = new System.Drawing.Size(269, 20);
            this.labelHeadsets.Text = "Headset-Mikrofone (für Mute-Erkennung):";
            // 
            // listHeadsets
            // 
            this.listHeadsets.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listHeadsets.CheckOnClick = true;
            this.listHeadsets.FormattingEnabled = true;
            this.listHeadsets.Location = new System.Drawing.Point(16, 103);
            this.listHeadsets.Name = "listHeadsets";
            this.listHeadsets.Size = new System.Drawing.Size(455, 202);
            // 
            // checkAutostart
            // 
            this.checkAutostart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkAutostart.AutoSize = true;
            this.checkAutostart.Location = new System.Drawing.Point(16, 320);
            this.checkAutostart.Name = "checkAutostart";
            this.checkAutostart.Size = new System.Drawing.Size(155, 24);
            this.checkAutostart.Text = "Mit Windows starten";
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnSave.Location = new System.Drawing.Point(317, 360);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(154, 35);
            this.btnSave.Text = "Speichern & Schließen";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(489, 407);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.checkAutostart);
            this.Controls.Add(this.listHeadsets);
            this.Controls.Add(this.labelHeadsets);
            this.Controls.Add(this.comboStandMic);
            this.Controls.Add(this.labelStandMic);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(400, 450);
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MicSwitcher - Einstellungen";
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SettingsForm_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }

    // ####################################################################
    // ## DATEI 3: Program.cs (Hauptstartpunkt) (Unverändert)
    // ####################################################################

    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Log.Info("========================================");
            Log.Info("MicSwitcherContext wird gestartet...");
            Application.Run(new MicSwitcherContext());
        }
    }

    public class MicSwitcherContext : ApplicationContext
    {
        private NotifyIcon _notifyIcon;
        private SettingsForm _settingsForm;
        private MicSwitcherLogic _logic;

        public MicSwitcherContext()
        {
            _logic = new MicSwitcherLogic();
            _settingsForm = new SettingsForm(_logic);

            _notifyIcon = new NotifyIcon()
            {
                Icon = SystemIcons.Asterisk,
                Text = "MicSwitcher",
                Visible = true
            };

            var contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("Einstellungen...", null, OnSettingsClicked);
            contextMenu.Items.Add("-");
            contextMenu.Items.Add("Beenden", null, OnExitClicked);
            _notifyIcon.ContextMenuStrip = contextMenu;

            _notifyIcon.DoubleClick += OnSettingsClicked;

            _logic.OnDeviceSwitched += ShowDeviceSwitchNotification;

            StartLogic();
        }

        private void ShowDeviceSwitchNotification(string deviceName)
        {
            Log.Info($"Zeige Benachrichtigung: Gerät gewechselt zu {deviceName}");
            _notifyIcon.ShowBalloonTip(
                2000,
                "Mikrofon gewechselt",
                $"Aktives Mikrofon: {deviceName}",
                ToolTipIcon.Info
            );
        }



        private void StartLogic()
        {
            _logic.LoadSettings();
            var settings = _logic.GetCurrentSettings();

            if (string.IsNullOrEmpty(settings.StandMicId) || settings.HeadsetMicIds.Count == 0)
            {
                Log.Info("StartLogic: Keine Konfig gefunden, zeige Einstellungen.");
                ShowSettingsForm();
            }
            else
            {
                Log.Info("StartLogic: Konfig geladen, starte Erfassung.");
                _logic.StartMonitoring();
            }
        }

        private void ShowSettingsForm()
        {
            if (_settingsForm.Visible)
            {
                _settingsForm.Activate();
            }
            else
            {
                _settingsForm.ShowDialog();
            }
        }

        private void OnSettingsClicked(object sender, EventArgs e)
        {
            Log.Info("Kontextmenü: Einstellungen geklickt.");
            ShowSettingsForm();
        }

        private void OnExitClicked(object sender, EventArgs e)
        {
            Log.Info("Kontextmenü: Beenden geklickt.");
            _notifyIcon.Visible = false;
            _logic.Dispose();
            Application.Exit();
        }
    }
}