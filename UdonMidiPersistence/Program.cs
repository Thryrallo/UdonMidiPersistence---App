using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.VisualBasic.Devices;
using Newtonsoft.Json;
using System.IO;
using System.Text;
using Commons.Music.Midi;
using System.Collections;
using Newtonsoft.Json.Linq;
using Microsoft.Win32;
using System.Security.Policy;
using System.Globalization;
using System.Reflection;

namespace UdonMidiPersistence
{

    // Required virtual midi port
    // https://www.tobias-erichsen.de/software/loopmidi.html
    internal static class Program
    {
        const string VERSION = "1.0.16";

        static string LOCAL_LOW_FOLDER = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData).Replace("Roaming", "LocalLow");
        static string VRCHAT_APP_DATA = LOCAL_LOW_FOLDER + "\\VRChat\\VRChat";
        static string SAVE_FILE_PATH = VRCHAT_APP_DATA + "\\UMP.json";
        static string LOG_FILE_PATH = VRCHAT_APP_DATA + "\\UMP_Log.txt";
        static bool _hasSpecificLogPath = false;
        static string _specificLogPath;
        static StreamReader _logStream;

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        static extern IntPtr GetShellWindow();

        [DllImport("user32.dll")]
        static extern IntPtr GetDesktopWindow();

        static NotifyIcon notifyIcon = new NotifyIcon();
        static bool Visible = false;
        static IntPtr processHandle;
        static IntPtr WinShell;
        static IntPtr WinDesktop;

        static ToolStripItem _toggleStartupItem;
        static bool _runOnStartup;


        static int _lastReadLogLine;
        static bool _isRunning = true;

        static Dictionary<string,Dictionary<string,object>> _savedValuesMap = new Dictionary<string,Dictionary<string,object>>();

        [STAThread]
        static void Main()
        {
            // Check if already running
            if (System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1)
            {
                Console.WriteLine("Udon Midi Persistence v" + VERSION);
                Console.WriteLine("Process is already running.");
                Thread.Sleep(3500);
                return;
            }
            
            // Clear Log file
            File.WriteAllText(LOG_FILE_PATH, "");
            using (new ConsoleCopy(LOG_FILE_PATH))
            {
                Console.WriteLine("Udon Midi Persistence v" + VERSION);
                string[] args = Environment.GetCommandLineArgs();
                if (args.Length > 1)
                {
                    _specificLogPath = args[1];
                    _hasSpecificLogPath = true;
                    Console.WriteLine($"Custom Log Path: {_specificLogPath}");
                    _logStream = new StreamReader(new FileStream(_specificLogPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), Encoding.UTF8);
                    ReadTillCurrentLine();
                }

                notifyIcon.DoubleClick += (s, e) =>
                {
                    Visible = !Visible;
                    SetConsoleWindowVisibility(Visible);
                };
                notifyIcon.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                notifyIcon.Visible = true;
                notifyIcon.Text = Application.ProductName;

                _runOnStartup = HasStartup();


                var contextMenu = new ContextMenuStrip();
                contextMenu.Items.Add("Show", null, (s, e) => { Visible = !Visible; SetConsoleWindowVisibility(Visible); });
                CreateStartupToggleItem(contextMenu);
                contextMenu.Items.Add("Exit", null, (s, e) => { _isRunning = false; Application.Exit(); });

                notifyIcon.ContextMenuStrip = contextMenu;

                processHandle = Process.GetCurrentProcess().MainWindowHandle;

                WinShell = GetShellWindow();

                WinDesktop = GetDesktopWindow();

                // Create folder
                // Load dictionary from file
                if (File.Exists(SAVE_FILE_PATH))
                {
                    _savedValuesMap = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, object>>>(File.ReadAllText(SAVE_FILE_PATH), new JsonSerializerSettings
                    {
                        TypeNameHandling = TypeNameHandling.Auto
                    });
                }

                Thread mainThread = new Thread(MainLoop);
                mainThread.Start();

                SetConsoleWindowVisibility(Visible);

                // Standard message loop to catch click-events on notify icon
                // Code after this method will be running only after Application.Exit()
                Application.Run();
                OnExit();
            }
        }

        static void OnExit()
        {
            Console.WriteLine("Clean Exit");
            _isRunning = false;
            notifyIcon.Visible = false;
            notifyIcon.Dispose();
        }

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        public static void SetConsoleWindowVisibility(bool visible)
        {
            IntPtr hWnd = FindWindow(null, Console.Title);
            if (hWnd != IntPtr.Zero)
            {
                if (visible) ShowWindow(hWnd, 1); //1 = SW_SHOWNORMAL           
                else ShowWindow(hWnd, 0); //0 = SW_HIDE               
            }
        }

        struct RequestData
        {
            public string id;
            public string dictionaryId;
        }

        struct SaveData
        {
            public string id;
            public string dictionaryId;
            public object value;
            public string type;
        }

        struct Vector3
        {
            public float x;
            public float y;
            public float z;

            public Vector3(float x, float y, float z)
            {
                this.x = x;
                this.y = y;
                this.z = z;
            }

            public override string ToString()
            {
                return $"({x},{y},{z})";
            }
        }

        static string _currentWorldId;
        static IMidiOutput _midiOutput;

        static void MainLoop()
        {
            var access = MidiAccessManager.Default;

            while(access.Outputs.Last().Name.Contains("loopMIDI") == false)
            {
                SetConsoleWindowVisibility(true);
                Console.WriteLine("");
                Console.WriteLine("You need to install & run loopMIDI before using MidiPersistence!");
                Console.WriteLine("Press 'Enter' to open loopMIDI donwload website.");
                Console.WriteLine("Press 'Spacebar' to try finding loopMIDI again.");
                ConsoleKeyInfo key;
                do
                {
                    key = Console.ReadKey();
                    if (key.Key == ConsoleKey.Enter)
                        OpenUrl("https://www.tobias-erichsen.de/software/loopmidi.html");
                } while (key.Key != ConsoleKey.Spacebar);
                Console.WriteLine("");
            }
            _midiOutput = access.OpenOutputAsync(access.Outputs.Last().Id).Result;

            string currentPath = null;
            long lastLogUpdate = 0;
            while (_isRunning) 
            {
                //Console.WriteLine("Main Loop");
                Thread.Sleep(100);
                // Check for most recent VRC Log file
                if (!_hasSpecificLogPath)
                {
                    while(currentPath == null || DateTimeOffset.Now.ToUnixTimeMilliseconds() - lastLogUpdate > 10000) // check for new log every 10 seconds
                    {
                        if (!Directory.Exists(VRCHAT_APP_DATA))
                        {
                            SetConsoleWindowVisibility(true);
                            Console.WriteLine("VRChat Logs folder cannot be found under '" + VRCHAT_APP_DATA + "'");
                            Thread.Sleep(10000);
                        }
                        else
                        {
                            lastLogUpdate = DateTimeOffset.Now.ToUnixTimeMilliseconds();
                            DirectoryInfo info = new DirectoryInfo(VRCHAT_APP_DATA);
                            FileInfo[] files = info.GetFiles().Where(f => f.Name.EndsWith(".txt") && f.Name.StartsWith("output_log_")).OrderByDescending(p => p.CreationTime).ToArray();
                            if(files.Length > 0)
                            {
                                if(currentPath != files[0].FullName)
                                {
                                    currentPath = files[0].FullName;
                                    Console.WriteLine("Opening Log: " + currentPath);
                                    if (_logStream != null)
                                        _logStream.Close();
                                    _logStream = new StreamReader(new FileStream(currentPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), Encoding.UTF8);
                                    ReadTillCurrentLine();
                                }
                            }
                            else
                            {
                                Console.WriteLine("No Logs can be found in directory '" + VRCHAT_APP_DATA + "'. Waiting...");
                                Thread.Sleep(20000);
                            }
                        }
                    }
                }

                // Parse file: Find all saves / Requests
                string line;
                while((line = _logStream.ReadLine()) != null)
                {
                    //Console.WriteLine(line);
                    if(line.Contains("[UdonMidiPersistence][Request]", StringComparison.Ordinal))
                    {
                        string json = line.Split("[UdonMidiPersistence][Request]")[1];
                        RequestData request = JsonConvert.DeserializeObject<RequestData>(json);
                        Console.WriteLine($"[Request] {request.id}");
                        HandleDataReqeust(request);
                    }

                    if (line.Contains("[UdonMidiPersistence][Save]", StringComparison.Ordinal))
                    {
                        string json = line.Split("[UdonMidiPersistence][Save]")[1];
                        SaveData save = JsonConvert.DeserializeObject<SaveData>(json);
                        Console.WriteLine($"[Save] {save.id} => {save.value}");
                        Save(save);
                    }
                }
            }
        }

        static void HandleDataReqeust(RequestData request)
        {
            if (_savedValuesMap.ContainsKey(request.dictionaryId) && _savedValuesMap[request.dictionaryId].ContainsKey(request.id))
            {
                object value = _savedValuesMap[request.dictionaryId][request.id];
                Console.WriteLine($"[Send] {request.id} => {value}");
                if (value.GetType() == typeof(int)) SendInt(request.id, (int)value);
                if (value.GetType() == typeof(float)) SendFloat(request.id, (float)value);
                if (value.GetType() == typeof(string)) SendString(request.id, (string)value);
                if (value.GetType() == typeof(Vector3)) SendVector3(request.id, (Vector3)value);
            }
            else
            {
                Console.WriteLine($"[Send] {request.id} => NONE");
                SendDataMessage(request.id, typeof(object), new byte[0]);
            }
        }

        static void SendInt(string id, int value)
        {
            SendDataMessage(id, typeof(int), BitConverter.GetBytes(value));
        }

        static void SendFloat(string id, float value)
        {
            SendDataMessage(id, typeof(float), BitConverter.GetBytes(value));
        }

        static void SendString(string id, string value)
        {
            SendDataMessage(id, typeof(string), Encoding.UTF8.GetBytes(value));
        }

        static void SendVector3(string id, Vector3 value)
        {
            byte[] data = new byte[12];
            Array.Copy(BitConverter.GetBytes(value.x), 0, data, 0, 4);
            Array.Copy(BitConverter.GetBytes(value.y), 0, data, 4, 4);
            Array.Copy(BitConverter.GetBytes(value.z), 0, data, 8, 4);
            SendDataMessage(id, typeof(Vector3), data);
        }

        static void SendDataMessage(string id, Type type, byte[] value)
        {
            int valueLength = value.Length + (type == typeof(string) ? 2 : 0);
            byte typeValue = 0;
            if (type == typeof(int)) typeValue = 1;
            if (type == typeof(float)) typeValue = 2;
            if (type == typeof(string)) typeValue = 3;
            if (type == typeof(Vector3)) typeValue = 4;
            byte[] data = new byte[2 + 2 + id.Length + valueLength];
            data[0] = (byte)(data.Length >> 8);
            data[1] = (byte)(data.Length & 0xFF);
            data[2] = (byte)(id.Length);
            data[3] = (byte)(typeValue);
            Array.Copy(Encoding.UTF8.GetBytes(id), 0, data, 4, id.Length);
            if (type == typeof(string))
            {
                data[4 + id.Length] = (byte)(value.Length >> 8);
                data[5 + id.Length] = (byte)(value.Length & 0xFF);
                Array.Copy(value, 0, data, 6 + id.Length, value.Length);
            }
            else
            {
                Array.Copy(value, 0, data, 4 + id.Length, value.Length);
            }
            SendBytes(data);
        }

        static byte BoolArrayToByte(bool[] ar, int offset, int length)
        {
            byte val = 0;
            for(int i = 0; (offset + i) < ar.Length && i < length; i++)
            {
                val += (byte)((ar[offset + i] ? 1 : 0) << (length - i - 1));
            }
            return val;
        }

        static void SendBytes(byte[] bytes)
        {
            // convert bytes to bits
            bool[] bits = new bool[bytes.Length * 8];
            for(int i = 0; i < bytes.Length; i++)
            {
                bits[i * 8 + 7] = (bytes[i] & 1) == 1;
                bits[i * 8 + 6] = (bytes[i] & 2) == 2;
                bits[i * 8 + 5] = (bytes[i] & 4) == 4;
                bits[i * 8 + 4] = (bytes[i] & 8) == 8;
                bits[i * 8 + 3] = (bytes[i] & 16) == 16;
                bits[i * 8 + 2] = (bytes[i] & 32) == 32;
                bits[i * 8 + 1] = (bytes[i] & 64) == 64;
                bits[i * 8 + 0] = (bytes[i] & 128) == 128;
                //Console.WriteLine(bytes[i]);
            }
            // iterate bits array with 14 bit jumps
            for(int i = 0; i < bits.Length; i += 14)
            {
                byte b1 = BoolArrayToByte(bits, i, 7);
                byte b2 = BoolArrayToByte(bits, i+7, 7);
                //Console.WriteLine(b1 + " , " + b2);
                SendMidi(MidiEvent.NoteOn, 0, b1, b2);
            }
        }

        static void SendMidi(byte type, int channel, int number, int value)
        {
            _midiOutput.Send(new byte[] { type, (byte)number, (byte)value }, 0, 3, 0);
        }

        static void Save(SaveData data)
        {
            if(data.dictionaryId == null)
            {
                Console.WriteLine("[Error][Save] dictionaryId is null! This should never happen!");
                return;
            }
            if (!_savedValuesMap.ContainsKey(data.dictionaryId))
                _savedValuesMap.Add(data.dictionaryId, new Dictionary<string, object>());
            object value = data.value;
            if (data.type == "System.Int32") value = Convert.ToInt32(value);
            else if(data.type == "System.Single") value = Convert.ToSingle(value);
            else if(data.type == "UnityEngine.Vector3") value = ConvertToVector3((string)value);
            else if (data.type != "System.String")
            {
                Console.WriteLine($"[Error][Save] Type {data.type} is not supported.");
                return;
            }
            _savedValuesMap[data.dictionaryId][data.id] = value;
            File.WriteAllText(SAVE_FILE_PATH, JsonConvert.SerializeObject(_savedValuesMap, Formatting.Indented, new JsonSerializerSettings
            {
                TypeNameHandling = TypeNameHandling.All
            }));
        }

        static void AnnounceNewWorld(string worldId)
        {
            if(worldId != _currentWorldId)
            {
                Console.WriteLine("New World ID");
                _currentWorldId = worldId;
            }
        }

        static void ReadTillCurrentLine()
        {
            while(_logStream.ReadLine() != null)
            {

            }
        }

        static void OpenUrl(string url)
        {
            Process pro = new Process();
            pro.StartInfo.FileName = url;
            pro.StartInfo.UseShellExecute = true;
            pro.Start();
        }

        static void CreateStartupToggleItem(ContextMenuStrip menu)
        {
            _toggleStartupItem = new ToolStripMenuItem((_runOnStartup ? "✓" : "✖") + "Autorun on startup", null, (s, e) =>
            {
                _runOnStartup = !_runOnStartup;
                SetStartup(_runOnStartup);
                _toggleStartupItem.Text = (_runOnStartup ? "✓" : "✖") + "Autorun on startup";
            });
            menu.Items.Add(_toggleStartupItem);
        }

        static void SetStartup(bool enabled)
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey
                ("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

            if (enabled)
                rk.SetValue("MidiPersistene", "\""+Application.ExecutablePath+"\"");
            else
                rk.DeleteValue("MidiPersistene", false);
        }

        static bool HasStartup()
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey
                ("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            return rk.GetValue("MidiPersistene") != null;
        }

        static Vector3 ConvertToVector3(string data)
        {
            data = data.Trim('(', ')');
            string[] split = data.Split(',');
            return new Vector3(Convert.ToSingle(split[0], CultureInfo.InvariantCulture), Convert.ToSingle(split[1], CultureInfo.InvariantCulture), Convert.ToSingle(split[2], CultureInfo.InvariantCulture));
        }
    }
}