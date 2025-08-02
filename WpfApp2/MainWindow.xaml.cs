using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Diagnostics;
using Syncfusion.UI.Xaml.Grid;
using Google.Protobuf.WellKnownTypes;
using Syncfusion.SfSkinManager;
using iText.Commons.Utils;
using Syncfusion.UI.Xaml.Grid.Helpers;
using Syncfusion.UI.Xaml.ScrollAxis;
using System.Reflection;
using Syncfusion.Data;
using System.ComponentModel;
using Org.BouncyCastle.Asn1.X9;
using System.Collections.ObjectModel;
using Syncfusion.Linq;
using Syncfusion.Windows.Data;
using Syncfusion.UI.Xaml.Charts;
using Syncfusion.CompoundFile.DocIO;
using System.Text.RegularExpressions;
using static Mysqlx.Datatypes.Scalar.Types;
using Syncfusion.UI.Xaml.CellGrid;
using Syncfusion.PMML;
using Mysqlx.Cursor;
using Syncfusion.Data.Extensions;
using iText.IO.Font.Otf;
using System.Timers;
using Syncfusion.Windows.Shared;
using System.Runtime.InteropServices;
using System.Printing;
using System.Media;
using System.IO;
using Syncfusion.XlsIO;

namespace WpfApp2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : ChromelessWindow
    {
        BitmapImage WaitingImage = new BitmapImage();
        BitmapImage RunningImage = new BitmapImage();
        BitmapImage StoppedImage = new BitmapImage();
        // Shared stop flag (volatile to ensure proper thread synchronization)
        private static volatile bool stop_Main_Function_thread = false;
        private static int audio_counter = 1;
        private static int audio_interate = 0;
        private static List<string> soundtracks = GetWavFiles();
        private static List<Tuple<string,string>> error_objects = new List<Tuple<string,string>>();
        Thread Main_Function_Thread;




        //All about capturing windows
        /**********************************************************************************************************/
        // Importing necessary functions from User32.dll                                                          
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool IsWindowVisible(IntPtr hWnd);
        // Delegate for enumerating windows                                                                       
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        /**********************************************************************************************************/




        //All about Input Simulation
        /**********************************************************************************************************/
        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public uint type;
            public InputUnion u;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        public const int INPUT_KEYBOARD = 1;
        public const int KEYEVENTF_KEYUP = 0x0002;

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern ushort VkKeyScan(char ch);

        public static void SimulateKeyPress(ushort keyCode)
        {
            INPUT[] inputs = new INPUT[]
            {
            // Press the key
            new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = keyCode,  // Virtual key code
                        dwFlags = 0 // Key press
                    }
                }
            },
            // Release the key
            new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = keyCode,  // Virtual key code
                        dwFlags = KEYEVENTF_KEYUP // Key release
                    }
                }
            }
            };

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        public static void HoldKey(ushort keyCode)
        {
            INPUT[] inputs = new INPUT[]
            {
            new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = keyCode,  // Virtual key code for the key
                        dwFlags = 0 // 0 for key press (hold)
                    }
                }
            }
            };

            // Send the key press event (but do not release it yet)
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        public static void ReleaseKey(ushort keyCode)
        {
            INPUT[] inputs = new INPUT[]
            {
            new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = keyCode,  // Virtual key code for the key
                        dwFlags = KEYEVENTF_KEYUP // KEYEVENTF_KEYUP for key release
                    }
                }
            }
            };

            // Send the key release event
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }
        /**********************************************************************************************************/




        static void PlayErrorSound()
        {
            try
            {
                if (audio_counter != 40)
                {
                    // Specify the path to the WAV file
                    string filePath1 = @"./soundtracks/portal2buzzer.wav";
                    string filePath2 = @"./soundtracks/transfer-order-failed-to-post-101soundboards.wav";

                    // Create a SoundPlayer instance
                    SoundPlayer player = new SoundPlayer(filePath1);
                    SoundPlayer player2 = new SoundPlayer(filePath2);

                    // Play the audio file
                    // player.PlaySync(); // Use this to play synchronously (blocks until the sound finishes playing)
                    player.PlaySync(); // Plays asynchronously
                    player2.PlaySync(); // Plays asynchronously

                    audio_counter++;
                }
                else
                {
                    string filePath3 = soundtracks[audio_interate];
                    // Create a SoundPlayer instance
                    SoundPlayer player3 = new SoundPlayer(filePath3);
                    player3.PlaySync(); // Plays asynchronously
                    audio_counter = 1;
                    if (audio_interate == soundtracks.Count - 1)
                    {
                        audio_interate = 0;
                        ShuffleList(soundtracks);
                        Debug.WriteLine("Shuffled List:");
                        foreach (string wav_file in soundtracks)
                        {
                            Debug.WriteLine(wav_file + "\n\n");
                        }
                    }
                    else
                    {
                        audio_interate++;
                    }
                }
                Debug.WriteLine("audio_counter: " + audio_counter);
                Debug.WriteLine("audio_interate: " + audio_interate);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error playing sound: {ex.Message}", "Exception Error", MessageBoxButton.OK);
            }
        }
        static void PlayPostedSound()
        {
            try
            {
                if (audio_counter != 40)
                {
                    // Specify the path to the WAV file
                    string filePath1 = @"./soundtracks/button_beep-in.wav";
                    string filePath2 = @"./soundtracks/transfer-order-posted-101soundboards.wav";

                    // Create a SoundPlayer instance
                    SoundPlayer player = new SoundPlayer(filePath1);
                    SoundPlayer player2 = new SoundPlayer(filePath2);

                    // Play the audio file
                    // player.PlaySync(); // Use this to play synchronously (blocks until the sound finishes playing)
                    player.PlaySync(); // Plays asynchronously
                    player2.PlaySync(); // Plays asynchronously
                    audio_counter++;
                }
                else
                {
                    string filePath3 = soundtracks[audio_interate];
                    // Create a SoundPlayer instance
                    SoundPlayer player3 = new SoundPlayer(filePath3);
                    player3.PlaySync(); // Plays asynchronously
                    audio_counter = 1;
                    if (audio_interate == soundtracks.Count - 1)
                    {
                        audio_interate = 0;
                        ShuffleList(soundtracks);
                        Debug.WriteLine("Shuffled List:");
                        foreach (string wav_file in soundtracks)
                        {
                            Debug.WriteLine(wav_file + "\n\n");
                        }
                    }
                    else
                    {
                        audio_interate++;
                    }
                }
                Debug.WriteLine("audio_counter: " + audio_counter);
                Debug.WriteLine("audio_interate: " + audio_interate);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error playing sound: {ex.Message}", "Exception Error", MessageBoxButton.OK);
            }
        }

        static void Randomized_Soundtrack()
        {
            try
            {
                // Shuffle the list
                ShuffleList(soundtracks);

                // Print the shuffled list
                Debug.WriteLine("Shuffled List:");
                foreach (string wav_file in soundtracks)
                {
                    Debug.WriteLine(wav_file + "\n\n");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error playing sound: {ex.Message}", "Exception Error", MessageBoxButton.OK);
            }
        }

        static void ShuffleList(List<string> soundtracks)
        {
            try
            {
                Random rng = new Random();
                int n = soundtracks.Count;

                for (int i = n - 1; i > 0; i--)
                {
                    // Pick a random index from 0 to i
                    int j = rng.Next(i + 1);

                    // Swap the elements at indices i and j
                    string temp = soundtracks[i];
                    soundtracks[i] = soundtracks[j];
                    soundtracks[j] = temp;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error playing sound: {ex.Message}", "Exception Error", MessageBoxButton.OK);
            }
        }

        static List<string> GetWavFiles()
        {
            List<string> wavFiles = new List<string>();
            try
            {
                // Get all .mp3 files from the directory and its subdirectories
                string[] files = Directory.GetFiles(Directory.GetCurrentDirectory() + "./soundtracks/randomized_soundtrack/", "*.wav", SearchOption.AllDirectories);

                // Add each .mp3 file to the list
                wavFiles.AddRange(files);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error: {ex.Message}");
            }

            return wavFiles;
        }


        public void The_Main_Function()
        {
            /*
            //window captor
            while (true)
            {
                IntPtr navWindow = FindWindow(null, "Microsoft Dynamics NAV");

                if (navWindow != IntPtr.Zero)
                {
                    EnumChildWindows(navWindow, new EnumWindowsProc((hWnd, lParam) =>
                                {
                                    // Step 3: Retrieve window properties for each child window
                                    StringBuilder windowText = new StringBuilder(256);
                                    StringBuilder className = new StringBuilder(256);

                                    // Get the window title (if any)
                                    GetWindowText(hWnd, windowText, windowText.Capacity);

                                    // Get the window class name
                                    GetClassName(hWnd, className, className.Capacity);

                                    // Only list visible windows
                                    if (IsWindowVisible(hWnd))
                                    {
                                        Console.WriteLine("Child Window:");
                                        Console.WriteLine("  - Title: " + windowText);
                                        Console.WriteLine("  - Class: " + className);
                                        Console.WriteLine("  - Handle: " + hWnd);
                                        Console.WriteLine("***************************************");

                                    }

                                    return true; // Continue enumeration
                                }), IntPtr.Zero);
                    Console.WriteLine("0000000000000000000000000000000000000000000000000");
                    Thread.Sleep(500);   //sleeps for 10 secs




                }
            }
            */


            /*
            // List all windows
            EnumWindows(new EnumWindowsProc((hWnd, lParam) =>
            {
                // Check if the window is visible
                if (IsWindowVisible(hWnd))
                {
                    // Get the window title (max 256 characters)
                    StringBuilder windowText = new StringBuilder(256);
                    GetWindowText(hWnd, windowText, windowText.Capacity);

                    // Only print windows with a title
                    if (!string.IsNullOrEmpty(windowText.ToString()))
                    {
                        Console.WriteLine("Window Title: " + windowText);
                    }
                }
                return true; // Continue enumeration
            }), IntPtr.Zero);
            */

            try
            {
                Randomized_Soundtrack();
                if (stop_Main_Function_thread == true)
                {
                    UpdateOutputTextBox2("\nMain_Function_Thread stopped1");
                    return;
                }
                UpdateOutputTextBox2("\nExecuting Main function.");
                Thread.Sleep(500);   //sleeps for 0.5 secs
                UpdateOutputTextBox2("\n***********************************************\n\n");
                bool ContinuesToLoop = true;


                while (stop_Main_Function_thread != true && ContinuesToLoop == true)
                {
                    if (stop_Main_Function_thread == true)
                    {
                        UpdateOutputTextBox2("\nMain_Function_Thread stopped2");
                        return;
                    }
                    Thread.Sleep(500);   //sleeps for 0.5 secs
                    UpdateOutputTextBox2("\nExecuting Loop function.");
                    UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                    UpdateOutputTextBox2("\n***********************************************\n\n");
                    if (Main_Window_Exist() == false)
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("Main_Function_Thread stopped3");
                            return;
                        }
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting Main_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        ContinuesToLoop = false;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }
                    else if (
                        Initialize_Posting_Window_Exist() == false &&
                        Posting_Item_Not_In_Inventory_Window_Exist() == false &&
                        Posting_Lines_Window_Exist() == false &&
                        Adjusting_Inventory_Level_Window_Exist() == false &&
                        Posted_Successfully_Window_Exist() == false &&
                        Just_Updated_This_Page_Window_Exist() == false &&
                        Insufficient_Quantity_Window_Exist() == false
                        )
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("\nMain_Function_Thread stopped4");
                            return;
                        }
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting Posting_Item_Not_In_Inventory_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x78);  //F9
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        ContinuesToLoop = true;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }
                    else if (Initialize_Posting_Window_Exist() == true)
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("\nMain_Function_Thread stopped5");
                            return;
                        }
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting Initialize_Posting_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        HoldKey(0xA4);  //Left Alt
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x73);  //F4
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        ReleaseKey(0xA4);  //Left Alt
                        ContinuesToLoop = true;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }
                    else if (Posting_Item_Not_In_Inventory_Window_Exist() == true)
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("\nMain_Function_Thread stopped6");
                            return;
                        }
                        PlayErrorSound();
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting Posting_Item_Not_In_Inventory_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x28);  //Down Arrow
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x78);  //F9
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        ContinuesToLoop = true;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }
                    else if (There_is_no_Inventory_Posting_Group_Setup_Window_Exist() == true)
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("\nMain_Function_Thread stopped7");
                            return;
                        }
                        PlayErrorSound();
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting There_is_no_Inventory_Posting_Group_Setup_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x28);  //Down Arrow
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x78);  //F9
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        ContinuesToLoop = true;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }
                    else if (Insufficient_Quantity_Window_Exist() == true)
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("\nMain_Function_Thread stopped8");
                            return;
                        }
                        PlayErrorSound();
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting Insufficient_Quantity_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x28);  //Down Arrow
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x78);  //F9
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        ContinuesToLoop = true;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }
                    else if (Posted_Successfully_Window_Exist() == true)
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("\nMain_Function_Thread stopped9");
                            return;
                        }
                        PlayPostedSound();
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting Posted_Successfully_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x28);  //Down Arrow
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x78);  //F9
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        ContinuesToLoop = true;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }
                    else if (Posting_Lines_Window_Exist() == true)
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("\nMain_Function_Thread stopped10");
                            return;
                        }
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting Posting_Lines_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        ContinuesToLoop = true;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }


                    else if (Adjusting_Inventory_Level_Window_Exist() == true)
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("\nMain_Function_Thread stopped11");
                            return;
                        }
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting Adjusting_Inventory_Level_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        ContinuesToLoop = true;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }
                    else if (Just_Updated_This_Page_Window_Exist() == true)
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("\nMain_Function_Thread stopped12");
                            return;
                        }
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting Just_Updated_This_Page_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress(0x74);  //F5
                        ContinuesToLoop = true;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }
                    else if (Version_Confliction_Window_Exist() == true)
                    {
                        if (stop_Main_Function_thread == true)
                        {
                            UpdateOutputTextBox2("\nMain_Function_Thread stopped13");
                            return;
                        }
                        Thread.Sleep(500);   //sleeps for 0.5 secs
                        UpdateOutputTextBox2("\nExecuting Version_Confliction_Window_Exist function.");
                        UpdateOutputTextBox2("\nCountdown from 0.5 seconds...");
                        Thread.Sleep(500);  //sleeps for 0.5 secs
                        SimulateKeyPress((char)ConsoleKey.Enter);
                        ContinuesToLoop = true;
                        UpdateOutputTextBox2("\n***********************************************\n\n");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception Error", MessageBoxButton.OK);
            }
        }

        // Step 1: Find the main window of Microsoft Dynamics NAV
        public static bool Main_Window_Exist()
        {
            IntPtr navWindow = FindWindow(null, "Transfer Orders - Microsoft Dynamics NAV");
            if (navWindow != IntPtr.Zero)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        // Step 1: Find the main window of Microsoft Dynamics NAV
        public static bool Initialize_Posting_Window_Exist()
        {
            IntPtr navWindow = FindWindow(null, "Microsoft Dynamics NAV");
            bool contains_sub_window_name = false;
            bool contains_ship_text = false;
            bool contains_receive_text = false;
            bool contains_OK_text = false;
            bool contains_Cancel_text = false;

            if (navWindow != IntPtr.Zero)
            {
                EnumChildWindows(navWindow, new EnumWindowsProc((hWnd, lParam) =>
                            {
                                // Step 3: Retrieve window properties for each child window
                                StringBuilder windowText = new StringBuilder(256);
                                StringBuilder className = new StringBuilder(256);

                                // Get the window title (if any)
                                GetWindowText(hWnd, windowText, windowText.Capacity);

                                // Get the window class name
                                GetClassName(hWnd, className, className.Capacity);

                                // Only list visible windows
                                if (IsWindowVisible(hWnd))
                                {
                                    if (windowText.ToString() == "Microsoft Dynamics NAV")
                                    {
                                        contains_sub_window_name = true;
                                    }
                                    if (windowText.ToString() == "&Ship")
                                    {
                                        contains_ship_text = true;
                                    }
                                    if (windowText.ToString() == "&Receive")
                                    {
                                        contains_receive_text = true;
                                    }
                                    if (windowText.ToString() == "Cancel")
                                    {
                                        contains_OK_text = true;
                                    }
                                    if (windowText.ToString() == "OK")
                                    {
                                        contains_Cancel_text = true;
                                    }
                                }

                                return true; // Continue enumeration
                            }), IntPtr.Zero);
                if (contains_sub_window_name == true &&
                        contains_ship_text == true &&
                        contains_receive_text == true &&
                        contains_OK_text == true &&
                        contains_Cancel_text == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }




        // Step 1: Find the main window of Microsoft Dynamics NAV
        public static bool Posting_Item_Not_In_Inventory_Window_Exist()
        {
            IntPtr navWindow = FindWindow(null, "Microsoft Dynamics NAV");
            bool contains_sub_window_name = false;
            bool contains_inventoryError_text = false;
            bool contains_OK_text = false;
            StringBuilder windowText = new StringBuilder(256);
            StringBuilder className = new StringBuilder(256);
            string error_text = "";
            //error_objects

            if (navWindow != IntPtr.Zero)
            {
                EnumChildWindows(navWindow, new EnumWindowsProc((hWnd, lParam) =>
                            {
                                // Step 3: Retrieve window properties for each child window

                                // Get the window title (if any)
                                GetWindowText(hWnd, windowText, windowText.Capacity);

                                // Get the window class name
                                GetClassName(hWnd, className, className.Capacity);

                                // Only list visible windows
                                if (IsWindowVisible(hWnd))
                                {
                                    if (windowText.ToString() == "Microsoft Dynamics NAV")
                                    {
                                        contains_sub_window_name = true;
                                    }
                                    if (windowText.ToString().Contains("is not in inventory."))
                                    {
                                        contains_inventoryError_text = true;
                                        error_text = windowText.ToString().Replace("Item","").Replace("is not in inventory.", "").Trim();
                                    }
                                    if (windowText.ToString() == "OK")
                                    {
                                        contains_OK_text = true;
                                    }
                                }

                                return true; // Continue enumeration
                            }), IntPtr.Zero);


                if (contains_sub_window_name == true &&
                        contains_inventoryError_text == true &&
                        contains_OK_text == true)
                {
                    //Debug.WriteLine(error_text);
                    error_objects.Add(Tuple.Create("Item is not in inventory: ", error_text));
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public static bool There_is_no_Inventory_Posting_Group_Setup_Window_Exist()
        {
            IntPtr navWindow = FindWindow(null, "Microsoft Dynamics NAV");
            bool contains_sub_window_name = false;
            bool contains_inventory_posting_group_Error_text = false;
            bool contains_OK_text = false;
            StringBuilder windowText = new StringBuilder(256);
            StringBuilder className = new StringBuilder(256);
            string error_text = "";

            if (navWindow != IntPtr.Zero)
            {
                EnumChildWindows(navWindow, new EnumWindowsProc((hWnd, lParam) =>
                            {
                                // Step 3: Retrieve window properties for each child window

                                // Get the window title (if any)
                                GetWindowText(hWnd, windowText, windowText.Capacity);

                                // Get the window class name
                                GetClassName(hWnd, className, className.Capacity);

                                // Only list visible windows
                                if (IsWindowVisible(hWnd))
                                {
                                    if (windowText.ToString() == "Microsoft Dynamics NAV")
                                    {
                                        contains_sub_window_name = true;
                                    }
                                    if (windowText.ToString().Contains("There is no Inventory Posting Setup within the filter."))
                                    {
                                        contains_inventory_posting_group_Error_text = true;
                                        error_text = windowText.ToString().Replace("There is no Inventory Posting Setup within the filter.","").Replace("Filters: Location Code:", "").Trim();
                                    }
                                    if (windowText.ToString() == "OK")
                                    {
                                        contains_OK_text = true;
                                    }
                                }

                                return true; // Continue enumeration
                            }), IntPtr.Zero);


                if (contains_sub_window_name == true &&
                        contains_inventory_posting_group_Error_text == true &&
                        contains_OK_text == true)
                {
                    error_objects.Add(Tuple.Create("no Inventory Posting Setup: ", error_text));
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public static bool Posting_Lines_Window_Exist()
        {
            IntPtr navWindow = FindWindow(null, "Microsoft Dynamics NAV");
            bool contains_sub_window_name = false;
            bool contains_posting_lines_text = false;
            bool contains_Cancel_text = false;

            if (navWindow != IntPtr.Zero)
            {
                EnumChildWindows(navWindow, new EnumWindowsProc((hWnd, lParam) =>
                            {
                                // Step 3: Retrieve window properties for each child window
                                StringBuilder windowText = new StringBuilder(256);
                                StringBuilder className = new StringBuilder(256);

                                // Get the window title (if any)
                                GetWindowText(hWnd, windowText, windowText.Capacity);

                                // Get the window class name
                                GetClassName(hWnd, className, className.Capacity);

                                // Only list visible windows
                                if (IsWindowVisible(hWnd))
                                {
                                    if (windowText.ToString() == "Microsoft Dynamics NAV")
                                    {
                                        contains_sub_window_name = true;
                                    }
                                    if (windowText.ToString() == "Posting transfer lines:")
                                    {
                                        contains_posting_lines_text = true;
                                    }
                                    if (windowText.ToString() == "Cancel")
                                    {
                                        contains_Cancel_text = true;
                                    }
                                }

                                return true; // Continue enumeration
                            }), IntPtr.Zero);


                if (contains_sub_window_name == true &&
                        contains_posting_lines_text == true &&
                        contains_Cancel_text == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }



        public static bool Adjusting_Inventory_Level_Window_Exist()
        {
            IntPtr navWindow = FindWindow(null, "Microsoft Dynamics NAV");
            bool contains_sub_window_name = false;
            bool contains_Adjusting_Value_text = false;
            bool contains_Adjmt_Level_text = false;
            bool contains_Adjust_text = false;
            bool contains_FW_Level_text = false;
            bool contains_Entry_No_text = false;
            bool contains_Remaining_Entries_text = false;
            bool contains_Cancel_text = false;

            if (navWindow != IntPtr.Zero)
            {
                EnumChildWindows(navWindow, new EnumWindowsProc((hWnd, lParam) =>
                            {
                                // Step 3: Retrieve window properties for each child window
                                StringBuilder windowText = new StringBuilder(256);
                                StringBuilder className = new StringBuilder(256);

                                // Get the window title (if any)
                                GetWindowText(hWnd, windowText, windowText.Capacity);

                                // Get the window class name
                                GetClassName(hWnd, className, className.Capacity);

                                // Only list visible windows
                                if (IsWindowVisible(hWnd))
                                {
                                    if (windowText.ToString() == "Microsoft Dynamics NAV")
                                    {
                                        contains_sub_window_name = true;
                                    }
                                    if (windowText.ToString() == "Adjusting value entries...")
                                    {
                                        contains_Adjusting_Value_text = true;
                                    }
                                    if (windowText.ToString() == "Adjmt. Level:")
                                    {
                                        contains_Adjmt_Level_text = true;
                                    }
                                    if (windowText.ToString() == "Adjust:")
                                    {
                                        contains_Adjust_text = true;
                                    }
                                    if (windowText.ToString() == "Cost FW. Level:")
                                    {
                                        contains_FW_Level_text = true;
                                    }
                                    if (windowText.ToString() == "Entry No.:")
                                    {
                                        contains_Entry_No_text = true;
                                    }
                                    if (windowText.ToString() == "Remaining Entries:")
                                    {
                                        contains_Remaining_Entries_text = true;
                                    }
                                    if (windowText.ToString() == "Cancel")
                                    {
                                        contains_Cancel_text = true;
                                    }
                                }

                                return true; // Continue enumeration
                            }), IntPtr.Zero);


                if (contains_sub_window_name == true &&
                    contains_Adjusting_Value_text == true &&
                    contains_Adjmt_Level_text == true &&
                    contains_Adjust_text == true &&
                    contains_FW_Level_text == true &&
                    contains_Entry_No_text == true &&
                    contains_Remaining_Entries_text == true &&
                    contains_Cancel_text == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }


        public static bool Posted_Successfully_Window_Exist()
        {
            IntPtr navWindow = FindWindow(null, "Microsoft Dynamics NAV");
            bool contains_sub_window_name = false;
            bool contains_inventoryError_text = false;
            bool contains_OK_text = false;

            if (navWindow != IntPtr.Zero)
            {
                EnumChildWindows(navWindow, new EnumWindowsProc((hWnd, lParam) =>
                            {
                                // Step 3: Retrieve window properties for each child window
                                StringBuilder windowText = new StringBuilder(256);
                                StringBuilder className = new StringBuilder(256);

                                // Get the window title (if any)
                                GetWindowText(hWnd, windowText, windowText.Capacity);

                                // Get the window class name
                                GetClassName(hWnd, className, className.Capacity);

                                // Only list visible windows
                                if (IsWindowVisible(hWnd))
                                {
                                    if (windowText.ToString() == "Microsoft Dynamics NAV")
                                    {
                                        contains_sub_window_name = true;
                                    }
                                    if (windowText.ToString().Contains("was successfully posted and is now deleted."))
                                    {
                                        contains_inventoryError_text = true;
                                    }
                                    if (windowText.ToString() == "OK")
                                    {
                                        contains_OK_text = true;
                                    }
                                }

                                return true; // Continue enumeration
                            }), IntPtr.Zero);


                if (contains_sub_window_name == true &&
                        contains_inventoryError_text == true &&
                        contains_OK_text == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }



        public static bool Just_Updated_This_Page_Window_Exist()
        {
            IntPtr navWindow = FindWindow(null, "Microsoft Dynamics NAV");
            bool contains_sub_window_name = false;
            bool contains_UpdatedThisPage_text = false;
            bool contains_OK_text = false;

            if (navWindow != IntPtr.Zero)
            {
                EnumChildWindows(navWindow, new EnumWindowsProc((hWnd, lParam) =>
                            {
                                // Step 3: Retrieve window properties for each child window
                                StringBuilder windowText = new StringBuilder(256);
                                StringBuilder className = new StringBuilder(256);

                                // Get the window title (if any)
                                GetWindowText(hWnd, windowText, windowText.Capacity);

                                // Get the window class name
                                GetClassName(hWnd, className, className.Capacity);

                                // Only list visible windows
                                if (IsWindowVisible(hWnd))
                                {
                                    if (windowText.ToString() == "Microsoft Dynamics NAV")
                                    {
                                        contains_sub_window_name = true;
                                    }
                                    if (windowText.ToString().Contains("just updated this page"))
                                    {
                                        contains_UpdatedThisPage_text = true;
                                    }
                                    if (windowText.ToString() == "OK")
                                    {
                                        contains_OK_text = true;
                                    }
                                }

                                return true; // Continue enumeration
                            }), IntPtr.Zero);


                if (contains_sub_window_name == true &&
                        contains_UpdatedThisPage_text == true &&
                        contains_OK_text == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }



        public static bool Insufficient_Quantity_Window_Exist()
        {
            IntPtr navWindow = FindWindow(null, "Microsoft Dynamics NAV");
            bool contains_sub_window_name = false;
            bool contains_Insufficient_Quantity_text = false;
            bool contains_OK_text = false;
            StringBuilder windowText = new StringBuilder(256);
            StringBuilder className = new StringBuilder(256);
            string error_text = "";

            if (navWindow != IntPtr.Zero)
            {
                EnumChildWindows(navWindow, new EnumWindowsProc((hWnd, lParam) =>
                            {
                                // Step 3: Retrieve window properties for each child window

                                // Get the window title (if any)
                                GetWindowText(hWnd, windowText, windowText.Capacity);

                                // Get the window class name
                                GetClassName(hWnd, className, className.Capacity);

                                // Only list visible windows
                                if (IsWindowVisible(hWnd))
                                {
                                    if (windowText.ToString() == "Microsoft Dynamics NAV")
                                    {
                                        contains_sub_window_name = true;
                                    }
                                    if (windowText.ToString().Contains("You have insufficient quantity of Item"))
                                    {
                                        contains_Insufficient_Quantity_text = true;
                                        error_text = windowText.ToString().Replace("You have insufficient quantity of Item","").Replace("on inventory.", "").Trim();
                                    }
                                    if (windowText.ToString() == "OK")
                                    {
                                        contains_OK_text = true;
                                    }
                                }

                                return true; // Continue enumeration
                            }), IntPtr.Zero);


                if (contains_sub_window_name == true &&
                        contains_Insufficient_Quantity_text == true &&
                        contains_OK_text == true)
                {
                    error_objects.Add(Tuple.Create("Insufficient Stock: ", error_text));
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }



        public static bool Version_Confliction_Window_Exist()
        {
            IntPtr navWindow = FindWindow(null, "Microsoft Dynamics NAV");
            bool contains_sub_window_name = false;
            bool contains_Version_Confliction_text = false;
            bool contains_OK_text = false;

            if (navWindow != IntPtr.Zero)
            {
                EnumChildWindows(navWindow, new EnumWindowsProc((hWnd, lParam) =>
                            {
                                // Step 3: Retrieve window properties for each child window
                                StringBuilder windowText = new StringBuilder(256);
                                StringBuilder className = new StringBuilder(256);

                                // Get the window title (if any)
                                GetWindowText(hWnd, windowText, windowText.Capacity);

                                // Get the window class name
                                GetClassName(hWnd, className, className.Capacity);

                                // Only list visible windows
                                if (IsWindowVisible(hWnd))
                                {
                                    if (windowText.ToString() == "Microsoft Dynamics NAV")
                                    {
                                        contains_sub_window_name = true;
                                    }
                                    if (windowText.ToString().Contains("An attempt was made to change an old version of a Transfer Header record."))
                                    {
                                        contains_Version_Confliction_text = true;
                                    }
                                    if (windowText.ToString() == "OK")
                                    {
                                        contains_OK_text = true;
                                    }
                                }

                                return true; // Continue enumeration
                            }), IntPtr.Zero);


                if (contains_sub_window_name == true &&
                        contains_Version_Confliction_text == true &&
                        contains_OK_text == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }




        public MainWindow()
        {
            try
            {
                InitializeComponent();
                ButtonAdv.IsEnabled = true;
                ButtonAdv1.IsEnabled = false;
                // Create source

                // BitmapImage.UriSource must be in a BeginInit/EndInit block
                WaitingImage.BeginInit();
                WaitingImage.UriSource = new Uri(@"Waiting.png", UriKind.Relative);

                // To save significant application memory, set the DecodePixelWidth or
                // DecodePixelHeight of the BitmapImage value of the image source to the desired
                // height or width of the rendered image. If you don't do this, the application will
                // cache the image as though it were rendered as its normal size rather than just
                // the size that is displayed.
                // Note: In order to preserve aspect ratio, set DecodePixelWidth
                // or DecodePixelHeight but not both.
                WaitingImage.DecodePixelWidth = 384;
                WaitingImage.EndInit();
                TextBox1.Text = "Waiting...";
                TextBox1.Foreground = new SolidColorBrush(Colors.Black);
                Image1.Source = WaitingImage;


                RunningImage.BeginInit();
                RunningImage.UriSource = new Uri(@"Running.png", UriKind.Relative);
                RunningImage.DecodePixelWidth = 384;
                RunningImage.EndInit();



                StoppedImage.BeginInit();
                StoppedImage.UriSource = new Uri(@"Stopped.png", UriKind.Relative);
                StoppedImage.DecodePixelWidth = 384;
                StoppedImage.EndInit();

                OutputTextBox2.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception Error", MessageBoxButton.OK);
            }
        }



        private async void ButtonAdv_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //disable start button first then start the countdown
                ButtonAdv.IsEnabled = false;
                //once the countdown is finished re-enable the stop button
                await StartCountDown();
                ButtonAdv1.IsEnabled = true;
                TextBox1.Text = "Running...";
                TextBox1.Foreground = new SolidColorBrush(Colors.Black);
                Image1.Source = RunningImage;
                Main_Function_Thread = new Thread(The_Main_Function);
                stop_Main_Function_thread = false;
                Main_Function_Thread.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception Error", MessageBoxButton.OK);
            }
        }

        private async void ButtonAdv_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                ButtonAdv1.IsEnabled = false;
                stop_Main_Function_thread = true;
                await Task.Run(() => Main_Function_Thread.Join());
                ButtonAdv.IsEnabled = true;
                TextBox1.Text = "Stopped...";
                TextBox1.Foreground = new SolidColorBrush(Colors.Black);
                Image1.Source = StoppedImage;
                if (error_objects.Count > 0)
                {
                    Export_Error();
                    error_objects.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception Error", MessageBoxButton.OK);
            }
        }



        // Specify what you want to happen when the Elapsed event is raised.
        private async Task StartCountDown()
        {
            try
            {
                int timer = 10;
                while (timer != 0)
                {
                    await Task.Delay(1000); // Asynchronously wait for 1 second
                    timer -= 1;
                    // Update the UI on the main thread
                    TextBox1.Text = $"{timer}";
                    TextBox1.Foreground = new SolidColorBrush(Colors.Yellow);
                    TextBox1.TextAlignment = TextAlignment.Center;

                    Image1.Source = WaitingImage;
                    OutputTextBox2.AppendText($"\n\nTimer: {timer}\n");
                    OutputTextBox2.ScrollToEnd();
                }
                TextBox1.TextAlignment = TextAlignment.Center;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception Error", MessageBoxButton.OK);
            }
        }
        public void UpdateOutputTextBox2(string Text)
        {
            try
            {
                OutputTextBox2.Dispatcher.Invoke(new System.Action(() =>
                {
                    OutputTextBox2.AppendText(Text);
                    OutputTextBox2.ScrollToEnd();
                }));
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception Error", MessageBoxButton.OK);
            }
        }

        private void ChromelessWindow_Closed(object sender, EventArgs e)
        {
            if (error_objects.Count > 0)
            {
                Export_Error();
                error_objects.Clear();
            }
        }


        private void Export_Error()
        {// Specify the file path
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/Transfer_Error.xlsx"; 
            string newfilePath = filePath;
            int fileIndex = 1;

            // Check if the file exists
            while (File.Exists(newfilePath) == true)
            {
                newfilePath =Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + $"/Transfer_Error_{fileIndex}" + ".xlsx";
                fileIndex++;
            }

            // Create an Excel engine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                // Instantiate the application object
                IApplication application = excelEngine.Excel;

                // Create a workbook with a worksheet
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                worksheet.Range[1,1].Text = "Transfer Error";
                worksheet.Range[1,1].CellStyle.ColorIndex = ExcelKnownColors.Teal;
                worksheet.Range[1,1].CellStyle.Font.Size = 12;
                worksheet.Range[1,1].AutofitColumns();
                worksheet.Range[1,1].AutofitRows();


                // Fill the row with column names
                for (int i = 0; i < error_objects.Count; i++)
                {
                    worksheet.Range[i + 4, 1].Text = error_objects[i].Item1;
                    worksheet.Range[i + 4, 1].CellStyle.ColorIndex = ExcelKnownColors.Pale_blue;
                    worksheet.Range[i + 4, 1].CellStyle.Font.Size = 12;
                    worksheet.Range[i + 4, 1].AutofitColumns();
                    worksheet.Range[i + 4, 1].AutofitRows();
                    worksheet.Range[i + 4, 2].Text = error_objects[i].Item2;
                    worksheet.Range[i + 4, 2].CellStyle.ColorIndex = ExcelKnownColors.Pale_blue;
                    worksheet.Range[i + 4, 2].CellStyle.Font.Size = 12;
                    worksheet.Range[i + 4, 2].AutofitColumns();
                    worksheet.Range[i + 4, 2].AutofitRows();
                }

                workbook.SaveAs(newfilePath);
                MessageBox.Show($"Excel file saved successfully at {newfilePath}", "Template Exported", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}