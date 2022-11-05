using EarTrumpet.DataModel.WindowsAudio;
using EarTrumpet.Diagnosis;
using EarTrumpet.Extensibility;
using EarTrumpet.Extensibility.Hosting;
using EarTrumpet.Extensions;
using EarTrumpet.Interop.Helpers;
using EarTrumpet.UI.Helpers;
using EarTrumpet.UI.ViewModels;
using EarTrumpet.UI.Views;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using LgTv.Clients;
using LgTv.Clients.Mouse;
using LgTv.Connections;
using LgTv.Stores;
using System.Windows.Forms;
using System.IO;
using MessageBox = System.Windows.MessageBox;
using Bugsnag;
using LgTv.Networking;
using System.Net.NetworkInformation;
using Quartz;
using Quartz.Impl;

namespace EarTrumpet
{
    public class MorningRoutine
    {
        private const bool SecureConnection = false;
        private const string TvHost = "192.168.254.173";
        private const int TvPort = 3000;
        private const string ClientKeyStoreFileName = "client-keys.json";
        private static readonly string ClientKeyStoreFilePath = Path.Combine(Path.GetDirectoryName(typeof(App).Assembly.Location), ClientKeyStoreFileName);

        public async Task GraduallyIncreaseVolume(LgTvClient client, TimeSpan timeSpan, int from, int to)
        {
            int volume = from;
            TimeSpan slice = TimeSpan.FromMilliseconds(timeSpan.TotalMilliseconds / (to - from));

            while (volume <= to)
            {
                Console.WriteLine($"Setting volume to {volume}");
                await client.Audio.SetVolume(volume);
                await Task.Delay(slice);

                volume++;
            }
        }

        public async Task Execute()
        {
            await Console.Out.WriteLineAsync("Running MorningRoutine");

            // Initialization
            var client = new LgTvClient(
                () => new LgTvConnection(),
                new JsonFileClientKeyStore(ClientKeyStoreFilePath),
                SecureConnection, TvHost, TvPort);

            await WakeOnLan.SendMagicPacket(App.MAC_ADDRESS);

            await Task.Delay(1000);

            await client.Connect();
            await client.MakeHandShake();

            await Task.Delay(1000);

            dynamic payload = new
            {
                id = "youtube.leanback.v4",
                contentId = "list=UU6isuGFtrmhgWNgFOP6h_fA"
            };

            await client.SendCommand(new LgTv.RequestMessage("ssap://system.launcher/launch", payload));

            await Task.Delay(1000);

            dynamic outputPayload = new
            {
                output = "tv_speaker"
            };

            await client.SendCommand(new LgTv.RequestMessage("ssap://com.webos.service.apiadapter/audio/changeSoundOutput", outputPayload));

            await Task.Delay(1000);

            await GraduallyIncreaseVolume(client, new TimeSpan(0, 0, 15), 0, 15);
        }
    }

    public class MorningJob : IJob
    {
        public async Task Execute(IJobExecutionContext context)
        {
            await Console.Out.WriteLineAsync("Greetings from HelloJob!");

            var morningRoutine = new MorningRoutine();

            await morningRoutine.Execute();
        }
    }

    public partial class App
    {
        // 	ac:5a:f0:9a:50:d5
        public static string MAC_ADDRESS = "80-5B-65-8A-6D-5E";

        public static bool IsShuttingDown { get; private set; }
        public static bool HasIdentity { get; private set; }
        public static bool HasDevIdentity { get; private set; }
        public static string PackageName { get; private set; }
        public static Version PackageVersion { get; private set; }
        public static TimeSpan Duration => s_appTimer.Elapsed;

        public FlyoutWindow FlyoutWindow { get; private set; }
        public DeviceCollectionViewModel CollectionViewModel { get; private set; }
        public DeviceCollectionViewModel RecordingCollectionViewModel { get; private set; }

        private static readonly Stopwatch s_appTimer = Stopwatch.StartNew();
        private FlyoutViewModel _flyoutViewModel;

        private ShellNotifyIcon _trayIcon;
        private WindowHolder _mixerWindow;
        private WindowHolder _settingsWindow;
        private ErrorReporter _errorReporter;
        private AppSettings _settings;

        private const string ClientKeyStoreFileName = "client-keys.json";

        private static readonly string ClientKeyStoreFilePath = Path.Combine(Path.GetDirectoryName(typeof(App).Assembly.Location), ClientKeyStoreFileName);

        private const bool SecureConnection = false;
        private const string TvHost = "192.168.254.173";
        private const int TvPort = 3000;

        private LgTvClient lgTvClient;

        IScheduler scheduler;

        private void OnAppStartup(object sender, StartupEventArgs e)
        {
            Exit += (_, __) => IsShuttingDown = true;
            Exit += async (_, __) =>
            {
                // and last shut down the scheduler when you are ready to close your program
                await scheduler.Shutdown();
            };
            HasIdentity = PackageHelper.CheckHasIdentity();
            HasDevIdentity = PackageHelper.HasDevIdentity();
            PackageVersion = PackageHelper.GetVersion(HasIdentity);
            PackageName = PackageHelper.GetFamilyName(HasIdentity);

            _settings = new AppSettings();
            _errorReporter = new ErrorReporter(_settings);

            var schedTask = InitializeScheduler();
            schedTask.Wait();

            if (SingleInstanceAppMutex.TakeExclusivity())
            {
                Exit += (_, __) => SingleInstanceAppMutex.ReleaseExclusivity();

                try
                {
                    ContinueStartup();
                }
                catch (Exception ex) when (IsCriticalFontLoadFailure(ex))
                {
                    ErrorReporter.LogWarning(ex);
                    OnCriticalFontLoadFailure();
                }
            }
            else
            {
                Shutdown();
            }
        }

        private async Task InitializeScheduler()
        {
            if (scheduler == null)
            {
                StdSchedulerFactory factory = new StdSchedulerFactory();
                scheduler = await factory.GetScheduler();
            }

            // and start it off
            await scheduler.Start();

            // define the job and tie it to our HelloJob class
            IJobDetail job = JobBuilder.Create<MorningJob>()
                .Build();

            // Trigger the job to run now, and then repeat every 10 seconds
            ITrigger trigger = TriggerBuilder.Create()
                .WithCronSchedule("0 15 10 * * ?")
                .Build();

            // Tell quartz to schedule the job using our trigger
            await scheduler.ScheduleJob(job, trigger);

            Console.WriteLine("Scheduled job");
        }

        private async Task<LgTvClient> InitializeTv()
        {
            if (lgTvClient != null)
            {
                return lgTvClient;
            }

            // Initialization
            var client = new LgTvClient(
                () => new LgTvConnection(),
                new JsonFileClientKeyStore(ClientKeyStoreFilePath),
                SecureConnection, TvHost, TvPort);

            //await client.Power.TurnOn();

            await client.Connect();
            await client.MakeHandShake();

            lgTvClient = client;

            return lgTvClient;

            /*var systemInfo = await client.Info.GetSystemInfo();
            var softwareInfo = await client.Info.GetSoftwareInfo();
            var connectionInfo = await client.Info.GetConnectionInfo();


            // Volume control
            await client.Audio.VolumeDown();
            await client.Audio.VolumeUp();


            // Playback control
            await client.Playback.Pause();
            await client.Playback.Play();


            var channels = await client.Channels.GetChannels();
            var inputs = await client.Inputs.GetInputs();
            var apps = await client.Apps.GetApps();

            using (var mouse = await client.GetMouse())
            {
                await mouse.SendButton(ButtonType.Up);
                await mouse.SendButton(ButtonType.Left);
                await mouse.SendButton(ButtonType.Right);
                await mouse.SendButton(ButtonType.Down);
            }


            await client.Power.TurnOff();*/
        }

        private void ContinueStartup()
        {
            ((UI.Themes.Manager)Resources["ThemeManager"]).Load();

            var recordingDeviceManager = WindowsAudioFactory.Create(AudioDeviceKind.Recording);
            RecordingCollectionViewModel = new DeviceCollectionViewModel(recordingDeviceManager, _settings);

            var deviceManager = WindowsAudioFactory.Create(AudioDeviceKind.Playback);
            deviceManager.Loaded += (_, __) => CompleteStartup();
            CollectionViewModel = new DeviceCollectionViewModel(deviceManager, _settings);

            _trayIcon = new ShellNotifyIcon(new TaskbarIconSource(RecordingCollectionViewModel, _settings));
            Exit += (_, __) => _trayIcon.IsVisible = false;

            CollectionViewModel.TrayPropertyChanged += () => { };
            RecordingCollectionViewModel.TrayPropertyChanged += () => _trayIcon.SetTooltip(RecordingCollectionViewModel.GetTrayToolTip());

            _flyoutViewModel = new FlyoutViewModel(RecordingCollectionViewModel, () => _trayIcon.SetFocus(), _settings);
            FlyoutWindow = new FlyoutWindow(_flyoutViewModel);
            // Initialize the FlyoutWindow last because its Show/Hide cycle will pump messages, causing UI frames
            // to be executed, breaking the assumption that startup is complete.
            FlyoutWindow.Initialize();
        }

        private void CompleteStartup()
        {
            AddonManager.Load(shouldLoadInternalAddons: HasDevIdentity);
            Exit += (_, __) => AddonManager.Shutdown();
#if DEBUG
            DebugHelpers.Add();
#endif
            _mixerWindow = new WindowHolder(CreateMixerExperience);
            _settingsWindow = new WindowHolder(CreateSettingsExperience);

            _settings.FlyoutHotkeyTyped += () => _flyoutViewModel.OpenFlyout(InputType.Keyboard);
            _settings.MixerHotkeyTyped += () => _mixerWindow.OpenOrClose();
            _settings.SettingsHotkeyTyped += () => _settingsWindow.OpenOrBringToFront();
            _settings.AbsoluteVolumeUpHotkeyTyped += AbsoluteVolumeIncrement;
            _settings.AbsoluteVolumeDownHotkeyTyped += AbsoluteVolumeDecrement;
            _settings.RegisterHotkeys();

            _trayIcon.PrimaryInvoke += (_, type) => _flyoutViewModel.OpenFlyout(type);
            _trayIcon.SecondaryInvoke += (_, __) => _trayIcon.ShowContextMenu(GetTrayContextMenuItems());
            _trayIcon.TertiaryInvoke += (_, __) =>
            {
                var device = RecordingCollectionViewModel.AllDevices.FirstOrDefault((d) => d.DeviceDescription == "Digital-In");

                if (device != null)
                {
                    device.IsMuted = !device.IsMuted;
                }
            };
            _trayIcon.Scrolled += (_, wheelDelta) =>
            {
                var device = RecordingCollectionViewModel.AllDevices.FirstOrDefault((d) => d.DeviceDescription == "Digital-In");

                if (device != null)
                {
                    device.IncrementVolume(Math.Sign(wheelDelta) * 2);
                }
            };
            _trayIcon.SetTooltip(RecordingCollectionViewModel.GetTrayToolTip());
            _trayIcon.IsVisible = true;

            DisplayFirstRunExperience();
        }

        private void DisplayFirstRunExperience()
        {
            if (!_settings.HasShownFirstRun
#if DEBUG
                || Keyboard.IsKeyDown(Key.LeftCtrl)
#endif
                )
            {
                Trace.WriteLine($"App DisplayFirstRunExperience Showing welcome dialog");
                _settings.HasShownFirstRun = true;

                var dialog = new DialogWindow { DataContext = new WelcomeViewModel(_settings) };
                dialog.Show();
                dialog.RaiseWindow();
            }
        }

        private bool IsCriticalFontLoadFailure(Exception ex)
        {
            return ex.StackTrace.Contains("MS.Internal.Text.TextInterface.FontFamily.GetFirstMatchingFont") ||
                   ex.StackTrace.Contains("MS.Internal.Text.Line.Format");
        }

        private void OnCriticalFontLoadFailure()
        {
            Trace.WriteLine($"App OnCriticalFontLoadFailure");

            new Thread(() =>
            {
                if (MessageBox.Show(
                    EarTrumpet.Properties.Resources.CriticalFailureFontLookupHelpText,
                    EarTrumpet.Properties.Resources.CriticalFailureDialogHeaderText,
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Error,
                    MessageBoxResult.OK) == MessageBoxResult.OK)
                {
                    Trace.WriteLine($"App OnCriticalFontLoadFailure OK");
                    ProcessHelper.StartNoThrow("https://eartrumpet.app/jmp/fixfonts");
                }
                Environment.Exit(0);
            }).Start();

            // Stop execution because callbacks to the UI thread will likely cause another cascading font error.
            new AutoResetEvent(false).WaitOne();
        }

        private IEnumerable<ContextMenuItem> GetTrayContextMenuItems()
        {
            var ret = new List<ContextMenuItem>();

            ret.AddRange(new List<ContextMenuItem>
                {
                    new ContextMenuSeparator(),
                    new ContextMenuItem
                    {
                        DisplayName = "Submenu",
                        Children = new List<ContextMenuItem>
                        {
                            new ContextMenuItem { DisplayName = "Nested", Command =  new RelayCommand(() =>
                            {
                                MessageBox.Show("Nested");
                            }) },
                        },
                    },
                    new ContextMenuSeparator(),
                    new ContextMenuItem { DisplayName = "Power on", Command =  new RelayCommand(async () =>
                            {
                                // Run on a loop until awakened.
                                //var mac = "80-5B-65-8A-6D-5E";
                                //var parsed = PhysicalAddress.Parse(mac);
                                //WakeOnLan.SendWakeOnLan(parsed);
                                //Console.WriteLine("Sent WoL packet");
                                //var client = await InitializeTv();

                                //await client.Power.TurnOn();

                                //var ipAddress = IPAddressResolver.GetIPAddress(TvHost);
                                //var macAddress = MacAddressResolver.GetMacAddress(ipAddress);

                                await WakeOnLan.SendMagicPacket(MAC_ADDRESS);
                            }) },
                    new ContextMenuItem { DisplayName = "Morning Routine", Command =  new RelayCommand(async () =>
                            {
                                var job = new MorningRoutine();

                                await job.Execute();
                            }) },
                    new ContextMenuItem { DisplayName = "Power off", Command =  new RelayCommand(async () =>
                            {
                                var client = await InitializeTv();

                                await client.Power.TurnOff();
                            }) },
                    new ContextMenuItem { DisplayName = "Play", Command =  new RelayCommand(async () =>
                            {
                                var client = await InitializeTv();

                                await client.Playback.Play();
                            }) },
                    new ContextMenuItem { DisplayName = "Pause", Command =  new RelayCommand(async () =>
                            {
                                var client = await InitializeTv();

                                await client.Playback.Pause();
                            }) },
                    new ContextMenuItem { DisplayName = "YouTube", Command =  new RelayCommand(async () =>
                            {
                                var client = await InitializeTv();

                                dynamic payload = new {
                                    id = "youtube.leanback.v4",
                                    contentId = "v=CzH5gHaiYg4"
                                };

                                await client.SendCommand(new LgTv.RequestMessage("ssap://system.launcher/launch", payload));
                            }) },
                    new ContextMenuItem { DisplayName = "Plex", Command =  new RelayCommand(async () =>
                            {
                                var client = await InitializeTv();

                                dynamic payload = new {
                                    id = "cdp-30",
                                    contentId = "v=CzH5gHaiYg4"
                                };

                                await client.SendCommand(new LgTv.RequestMessage("ssap://system.launcher/launch", payload));
                            }) },
                    new ContextMenuItem { DisplayName = EarTrumpet.Properties.Resources.ContextMenuExitTitle, Command = new RelayCommand(Shutdown) },
                });

            return ret;
        }

        private Window CreateSettingsExperience()
        {
            var defaultCategory = new SettingsCategoryViewModel(
                EarTrumpet.Properties.Resources.SettingsCategoryTitle,
                "\xE71D",
                EarTrumpet.Properties.Resources.SettingsDescriptionText,
                null,
                new SettingsPageViewModel[]
                    {
                        new EarTrumpetShortcutsPageViewModel(_settings),
                        new EarTrumpetLegacySettingsPageViewModel(_settings),
                        new EarTrumpetAboutPageViewModel(() => _errorReporter.DisplayDiagnosticData(), _settings)
                    });

            var allCategories = new List<SettingsCategoryViewModel>();
            allCategories.Add(defaultCategory);

            if (AddonManager.Host.SettingsItems != null)
            {
                allCategories.AddRange(AddonManager.Host.SettingsItems.Select(a => CreateAddonSettingsPage(a)));
            }

            var viewModel = new SettingsViewModel(EarTrumpet.Properties.Resources.SettingsWindowText, allCategories);
            return new SettingsWindow { DataContext = viewModel };
        }

        private SettingsCategoryViewModel CreateAddonSettingsPage(IEarTrumpetAddonSettingsPage addonSettingsPage)
        {
            var addon = (EarTrumpetAddon)addonSettingsPage;
            var category = addonSettingsPage.GetSettingsCategory();

            if (!addon.IsInternal())
            {
                category.Pages.Add(new AddonAboutPageViewModel(addon));
            }
            return category;
        }

        private Window CreateMixerExperience() => new FullWindow { DataContext = new FullWindowViewModel(RecordingCollectionViewModel) };

        private void AbsoluteVolumeIncrement()
        {
            foreach (var device in RecordingCollectionViewModel.AllDevices.Where(d => !d.IsMuted || d.IsAbsMuted))
            {
                // in any case this device is not abs muted anymore
                device.IsAbsMuted = false;
                device.IncrementVolume(2);
            }
        }

        private void AbsoluteVolumeDecrement()
        {
            foreach (var device in RecordingCollectionViewModel.AllDevices.Where(d => !d.IsMuted))
            {
                // if device is not muted but will be muted by 
                bool wasMuted = device.IsMuted;
                // device.IncrementVolume(-2);
                device.Volume -= 2;
                // if device is muted by this absolute down
                // .IsMuted is not already updated
                if (!wasMuted == (device.Volume <= 0))
                {
                    device.IsAbsMuted = true;
                }
            }
        }
    }
}
