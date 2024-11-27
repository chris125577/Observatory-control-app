using ASCOM.DeviceInterface;
using ASCOM.DriverAccess;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;


// Observatory control panel 
// controls mount, roof shutter, KMTronics power toggle and monitors environment and safety sensors
// added additional non-ASCOM standard controls to override safety systems
// 6.4 - modified park status so that moving mount at park overrides sensor
// 6.5 - removed windows messages on connection issues - causes hang up
// 6.6 - added some connected status updates for devices that do not connect
// 6.6.1 - added tracking off before homing, as a precaution
// 6.6.2 - added further precaution of turning off relays if abort mount does not work.
// 6.6.3 - further tests to abort slew or roof move in unsafe condition
// 7.0 - refinements to dome timing
// 8.0 - added graphs for environmental
// 8.1 - moved AAG powerup to safety from weather, chart improvements
// 8.2 - trying to find fix for TSX64 issues (turns out to be an ASCOM issue)
// 8.3 - trying to find way to get AAG to work.
// 8.4 - changing the AAG type and making it dynamic worked
// 8.5 - added further set of mount toggle disables if mount is not connected to this utility but is elsewhere
// 8.6 - increased rooftime window to 40 seconds
// 8.7 - changed switch logic over to normal polarity, using Pegasus rather than USB-powered switch.
// 8.8 - enhanced mount toggle disconnect - made method and included scan of "mount" in switch description too
// 8.9 - problems with Unity driver - direct relay1 disconnect for mount, and longer timer interval to 2000 ms
// 9.0 - improved clarity of toggle power control (when Unity server is fixed, poll will return to 1000ms)
// 9.2 - after Unity server update, put back quick toggle status refresh
// 10.0 - added beeper enable and disable - use with ARDUINO driver 3.0 and ASCOM driver  3.0
// 10.1 - added safety Shutdown 
// 10.2 - modified to use a range of safety monitors but still fires up AAG CoudWatcher if it detect file output
// 10.5 - delayed AAG graph output for 1 minute to prevent exception, conditional 4-way switch control
// 10.6 - added CloudCalculator cover (sky temperature)
// 11.0 - overhaul, including better comments
// 11.1 - changed sky temperature to CloudCalculator cover
// 11.2 - changed 1s form timer to 2s and 0.5s for non- and critical conditions
// 11.3 - the rain sensor status is updated to a file (safefile.txt) when conditions change 
// 11.4 - introduced new tab to do some of the mount settings (location, guide rate and time)
// 11.6 - several changes to accommodate new WD20 mount, which works best homing and parking from a stable tracking state
// 11.7 - several further refinements to accommodate wd20 mount which needs to be tracking before parking from home position
// 11.8 - mount status will show tracking or slewing, even if it is at sensor position, so it can react quicker 
//        also put in wait for AtHome to be confirmed after homing for mounts implementing async operations

namespace Observatory
{
    public partial class Obsyform : Form
    {
        private Telescope mount; // ASCOM telescope 
        private SafetyMonitor safe; // ASCOM safety monitor
        private ObservingConditions weather; // ASCOM weather
        private Dome dome;               // ASCOM dome
        private Switch toggle;           // ASCOM swtich
        //        private AAG_CloudWatcher.CloudWatcher oCW;  // for remote control of AAG cloudwatcher
        private string domeId, mountId, weatherId, safetyId, switchID;  // holds ASCOM device IDs
        private bool busy;  // flag to stop conflicting commands to Arduino
        private bool mountConnected, roofConnected, weatherConnected, safetyConnected, switchconnected;  // local variable to note current connected state
        private string mountSafe, roofSafe;      // used for local storage
        private string settingsfile = "\\obsy.txt";
        private string dryfile = "\\safety.txt";
        private bool priorDry = true;  // old rain sensor status, to check for change 
        private bool priority = false;   // used to toggle priority actions in RefreshAll()
        private ShutterState roofStatus;             // used to store shutter status
        String path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ASCOM\\Obsy";
        String safetyPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\AAG";
        private bool queOpen, queClose, dry, goodConditions, clearAir, clearSky; // boolean flags for deciding if roof is on "auto" mode operation
        private double maxHumidity; // threshold for fog/mist
        private int mountTimeout;  // period in multiples of 2 seconds for error condition to apply to movements
        private int roofTimeout;  // period in multiples of 2 seconds for roof to move
        private bool aborted;  // set if abort command is issued, reset if mount is tracking, homed or parked
        private bool[] power = new bool[4]; // local values of toggle states
        // arrays to hold last three hours of data for charting
        private double[] chartValues = new double[120];
        private double[] tempValues = new double[120];
        private double[] humidValues = new double[120];
        private double[] dewValues = new double[120];
        private double[] SQMValues = new double[120];
        private double[] rainValues = new double[120];
        private double[] cloudValues = new double[120];
        private int sampleCount = 0;  // sampling value increments every 2 seconds
        private int chartType = 0;  //selection variable
        // changes the true and false around on the following two lines to change the toggle logic sense
        private const bool switchOn = true;
        private const bool switchOff = false;
        private double SkyUL = 10; // 100% CloudCalculator cover (initial value)
        private double SkyLL = -10; // 0% CloudCalculator cover (iniitial value)
        private const double skyTmax = 25;
        private const double skyTmin = -25;//private short relayinterval = 0; // used for original Pegasus toggle box that needed slow polled switch reads
        private double longitude = 0.6162;  // default value for SWF
        private double latitude = 51.6387;   // default value for SWF
        private double elevation = 10.0;   // default value for SWF
        private double sidereal = 0.00416666666666667;   // sidereal rate deg/sec

        public Obsyform()
        {
            InitializeComponent();
            // initialise form status boxes
            drytext.Text = "not connected";
            rooftext.Text = "not connected";
            mountText.Text = "not connected";
            humidText.Text = "not connected";
            tempText.Text = "not connected";
            pressureText.Text = "not connected";
            sqmText.Text = "not connected";
            imagingText.Text = "not connected";
            cloudText.Text = "not connected";
            statusbox.Text = "";
            LongitudeIP.Text = longitude.ToString();
            LatitudeIP.Text = latitude.ToString();
            ElevationIP.Text = elevation.ToString();
            // initialise connection box colors
            btnConnDome.ForeColor = Color.White;
            btnDiscDome.ForeColor = Color.Gray;
            btnConnMount.ForeColor = Color.White;
            btnDiscMount.ForeColor = Color.Gray;
            btnConnWeather.ForeColor = Color.White;
            btnDiscWeather.ForeColor = Color.Gray;
            btnConnSafety.ForeColor = Color.White;
            btnDiscSafety.ForeColor = Color.Gray;
            btnConnSwitch.ForeColor = Color.White;
            btnDiscSwitch.ForeColor = Color.Gray;

            // initialise flags 
            // ASCOM connection status
            roofConnected = false;
            mountConnected = false;
            weatherConnected = false;
            safetyConnected = false;
            switchconnected = false;
            for (int i = 0; i < 4; i++) power[i] = false;
            busy = false; // flag used to prevent multiple serial commands overlapping
            queOpen = false;  // flag used to tell roof to semi-automate opening
            queClose = false;  // flag used to tell roof to semi-automate closure
            // initilise variables hold ID's of ASCOM drivers - which are stored and recalled from file
            domeId = null;
            mountId = null;
            weatherId = null;
            safetyId = null;
            // initialise weather and safety variables (Per's weather stick)
            maxHumidity = 97.0; // default but overidden by file read
            mountTimeout = 15; // response time period for mount to move (x2)
            roofTimeout = 20;  //  40 seconds for roof to complete movement
            // initialise flags  used to build up a composite safety flag
            dry = false;
            clearAir = false;
            clearSky = false;
            goodConditions = false;
            aborted = false;
            roofStatus = ShutterState.shutterError; // assume an error to start with
            // read in, if it exists, the connection settings file           
            if (File.Exists(path + settingsfile)) ReadFile();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (weatherId != null)
            {
                //set up chart
                // need to move y-axis setting dependendent on series
                chart1.Series.Clear(); //remove default series
                chart1.Series.Add("CloudW"); //add series called CloudW
                chart1.Series.FindByName("CloudW").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line; //Change chart type to line
                chart1.Series.FindByName("CloudW").Color = Color.White; //Change series color to red
                chart1.Series["CloudW"].BorderWidth = 2;
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                chart1.ChartAreas[0].AxisX.Maximum = 120; //(2 hours) - y-axis determined by source
            }
        }

        //  Each of the setxxx methods call up the ASCOM chooser and
        //  updates the configuration / ASCOM profile file.
        //  dome device (roll off roof)
        private void SetRoof(object sender, EventArgs e)
        {
            try
            {
                domeId = Dome.Choose(domeId);
                this.WriteFile();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("roof selection failed");
            }

        }
        // telescope (mount) device
        private void SetMount(object sender, EventArgs e)
        {
            try
            {
                mountId = Telescope.Choose(mountId);
                this.WriteFile();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("mount selection failed");
            }
        }

        // observing conditions device
        private void SetWeather(object sender, EventArgs e)
        {
            try
            {
                weatherId = ObservingConditions.Choose(weatherId);
                this.WriteFile();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("weather selection failed");
            }
        }
        // safety monitor device ( either separate or using ASCOM device that in turn links to weather monitors
        private void SetSafety(object sender, EventArgs e)
        {
            try
            {
                safetyId = SafetyMonitor.Choose(safetyId);
                this.WriteFile();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("safety selection failed");
            }
        }
        // switch devices (binary)
        private void SetSwitch(object sender, EventArgs e)
        {
            try
            {
                switchID = Switch.Choose(switchID);
                this.WriteFile();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("switch selection failed");
            }
        }

        //  Each of the connectxxx methods create an object instance, connect to object
        //  and set up initial states for status boxes.
        private void ConnectDome(object sender, EventArgs e)
        {
            try
            {
                if (domeId != null)
                {
                    dome = new Dome(domeId);
                    dome.Connected = true;
                    roofConnected = dome.Connected;
                    if (roofConnected)
                    {
                        rooftext.Text = "connected";
                        drytext.Text = "connected";
                        btnConnDome.ForeColor = Color.Gray;
                        btnDiscDome.ForeColor = Color.White;
                        DisplaySensor(); // update sensor backgrounds to indicate sensors
                        // initial Hydreon state
                        if (dome.CommandBool("RAIN", false))
                        {
                            File.WriteAllText(path + dryfile, "unsafe");
                            priorDry = false; // note reverse sense RAINING vs DRY
                        }
                        else
                        {
                            File.WriteAllText(path + dryfile, "safe");
                            priorDry = true;
                        }
                    }
                    else System.Windows.Forms.MessageBox.Show("Dome did not connect");
                }
                else
                {
                    roofConnected = false;
                    rooftext.Text = "not connected";
                    drytext.Text = "not connected";
                    btnConnDome.ForeColor = Color.White;
                    btnDiscDome.ForeColor = Color.Gray;
                }
            }
            catch (Exception)
            {
                roofConnected = false;
                dome.Connected = false;
                System.Windows.Forms.MessageBox.Show("Dome did not connect");
            }
        }
        private void ConnectMount(object sender, EventArgs e)
        {
            try
            {
                if (mountId != null)
                {
                    mount = new Telescope(mountId);
                    mount.Connected = true;
                    mountConnected = true;
                    mountText.Text = "connected";
                    btnConnMount.ForeColor = Color.Gray;
                    btnDiscMount.ForeColor = Color.White;
                    GuideRate.Value = (decimal)Math.Round(mount.GuideRateDeclination / sidereal, 2);
                    LongitudeIP.Text = Math.Round(mount.SiteLongitude, 6).ToString();
                    LatitudeIP.Text = Math.Round(mount.SiteLatitude, 6).ToString();
                    ElevationIP.Text = Math.Round(mount.SiteElevation, 2).ToString();
                    mount.UTCDate = (System.DateTime.UtcNow); // see if this works here
                    if (!mount.CanFindHome || !mount.CanPark || !mount.CanPulseGuide)
                        System.Windows.Forms.MessageBox.Show("check mount park/home/guide settings");
                }
                else
                {
                    mountConnected = false;
                    mountText.Text = "not connected";
                    btnConnMount.ForeColor = Color.White;
                    btnDiscMount.ForeColor = Color.Gray;
                }
            }
            catch (Exception)  // added some changes to status, to prevent new connections
            {
                mountConnected = false;
                mountText.Text = "not connected!";
                System.Windows.Forms.MessageBox.Show("Mount did not connect");
            }
        }
        private void ConnectSwitch(object sender, EventArgs e)
        {
            try
            {
                if (switchID != null)
                {
                    toggle = new Switch(switchID);
                    toggle.Connected = true;
                    switchconnected = true;
                    btnConnSwitch.ForeColor = Color.Gray;
                    btnDiscSwitch.ForeColor = Color.White;
                    // populate text boxes with toggle names from driver/profile
                    if(!switchID.Contains( "Pegasus"))
                    {
                        btnSwitch1.Text = toggle.GetSwitchName(0);
                        btnSwitch2.Text = toggle.GetSwitchName(1); // 
                        btnSwitch3.Text = toggle.GetSwitchName(2); //
                        btnSwitch4.Text = toggle.GetSwitchName(3);
                    }
                    else
                    {
                        btnSwitch1.Text = toggle.GetSwitchDescription(0);
                        btnSwitch2.Text = toggle.GetSwitchDescription(1); // 
                        btnSwitch3.Text = toggle.GetSwitchDescription(2); //
                        btnSwitch4.Text = toggle.GetSwitchDescription(3);
                    }
                        GetSwitchState();  // all at once
                }
                else
                {
                    switchconnected = false;
                    btnConnSwitch.ForeColor = Color.White;
                    btnDiscSwitch.ForeColor = Color.Gray;
                }
            }
            catch (Exception)
            {
                toggle.Connected = false;
                switchconnected = false;
                System.Windows.Forms.MessageBox.Show("Relay did not connect");
            }
        }

        //  weather devices are complicated by the fact that weather and safety devices
        //  are inextricably linked. Here, the AAG cloudwatcher is checked for as a safety device but also
        //  there are useful ASCOM utilities that create a safe/unsafe condition from multiple weather
        //  sources
         
        private void ConnectWeather(object sender, EventArgs e)
        {
            try
            {
                if (weatherId != null)
                {
                    // as safety monitor also has environmental sensors;
                    if (!safetyConnected) ConnectSafety(sender, e);
                    System.Threading.Thread.Sleep(5000);
                    // connect to weatherstick or other device
                    weather = new ObservingConditions(weatherId);
                    weather.Connected = true;
                    weatherConnected = true;
                    humidText.Text = "connected";
                    pressureText.Text = "connected";
                    tempText.Text = "connected";
                    sqmText.Text = "connected";
                    cloudText.Text = "connected";
                    btnConnWeather.ForeColor = Color.Gray;
                    btnDiscWeather.ForeColor = Color.White;
                    humidText.BackColor = Color.LightGreen;
                    pressureText.BackColor = Color.LightGreen;
                    sqmText.BackColor = Color.LightGreen;
                    tempText.BackColor = Color.LightGreen;
                    cloudText.BackColor = Color.LightGreen;
                    // default chart is temp C
                    chartType = 0;
                    chart1.ChartAreas[0].AxisY.Minimum = -5;
                    chart1.ChartAreas[0].AxisY.Maximum = 25;
                    btngraphsel.Text = "Temp C";
                    System.Threading.Thread.Sleep(5000); // time for sensors to read
                    RefreshWeather();

                }
                else // weather device not set
                {
                    weatherConnected = false;
                    humidText.Text = "not connected";
                    pressureText.Text = "not connected";
                    sqmText.Text = "not connected";
                    tempText.Text = "not connected";
                    cloudText.Text = "not connected";
                    btnConnWeather.ForeColor = Color.White;
                    btnDiscWeather.ForeColor = Color.Gray;
                    humidText.BackColor = Color.Silver;
                    pressureText.BackColor = Color.Silver;
                    sqmText.BackColor = Color.Silver;
                    tempText.BackColor = Color.Silver;
                    cloudText.BackColor = Color.Silver;
                }
            }
            catch (Exception)
            {
                weather.Connected = false;
                weatherConnected = false;
                System.Windows.Forms.MessageBox.Show("AAG/WS did not connect");
            }
        }
        // special case if it detects AAG cloudwatcher log file
        private void ConnectSafety(object sender, EventArgs e)
        {
            try
            {
                if (safetyId != null)
                {
                    if (File.Exists(safetyPath + "\\AAG_SLD.dat")) // detect presence of AAG Cloudwatcher
                    {
                        Type foo = Type.GetTypeFromProgID("AAG_CloudWatcher.CloudWatcher");
                        dynamic oCW;
                        oCW = Activator.CreateInstance(foo);
                        //oCW.WindowLeft = 0;
                        //oCW.WindowTop = 0;
                        oCW.Device_Start();
                        oCW.RecordStart(true);
                        oCW = null;
                    }
                        safe = new SafetyMonitor(safetyId);
                        safe.Connected = true;
                        safetyConnected = true;
                        imagingText.Text = "connected";
                        btnConnSafety.ForeColor = Color.Gray;
                        btnDiscSafety.ForeColor = Color.White;
                }
                else
                {
                    safetyConnected = false;
                    imagingText.Text = "not connected";
                    btnConnSafety.ForeColor = Color.White;
                    btnDiscSafety.ForeColor = Color.Gray;
                }
            }
            catch (Exception)
            {
                safe.Connected = false;
                safetyConnected = false;
                System.Windows.Forms.MessageBox.Show("Safety Monitor did not connect");
            }
        }

        // catchall shortcut does connect to all equipment
        // starts with SWITCH device only, to allow toggle setup before further connections
        // if SWITCH device connected, it connects the remaining devices
        public void ConnectAll(object sender, EventArgs e)
        {
            statusbox.Clear();
            if (!switchconnected)
            {
                ConnectSwitch(sender, e);
                System.Windows.Forms.MessageBox.Show("check equipment power");
            }

            else
            {
                ConnectSafety(sender, e);
                System.Threading.Thread.Sleep(5000);
                ConnectDome(sender, e);
                ConnectMount(sender, e);
                ConnectWeather(sender, e);
            }

        }
        // functions to do the device disconnections (called by several button presses)
        private void DisconnectTelescope()
        {
            try
            {
                if (mountConnected)
                {
                    mount.Connected = false;  // disconnect mount
                    mountConnected = false;
                    mountText.Text = "not connected";
                    btnConnMount.ForeColor = Color.White;
                    btnDiscMount.ForeColor = Color.Gray;
                    mountText.BackColor = Color.Silver;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect mount error");
            }
        }
        private void DisconnectDome()
        {
            try
            {
                if (roofConnected)
                {
                    dome.Connected = false;   //disconnect roof
                    roofConnected = false;
                    rooftext.Text = "not connected";
                    drytext.Text = "not connected";
                    btnConnDome.ForeColor = Color.White;
                    btnDiscDome.ForeColor = Color.Gray;
                    rooftext.BackColor = Color.Silver;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect roof error");
            }
        }
        private void DisconnectSwitch()
        {
            try
            {
                if (switchconnected)
                {
                    toggle.Connected = false;   //disconnect roof
                    switchconnected = false;
                    btnConnSwitch.ForeColor = Color.White;
                    btnDiscSwitch.ForeColor = Color.Gray;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect relay error");
            }
        }
        private void DisconnectWeather()
        {
            try
            {
                if (weatherConnected)
                {
                    weather.Connected = false;   //disconnect weather sensors
                    weatherConnected = false;
                    humidText.Text = "not connected";
                    pressureText.Text = "not connected";
                    tempText.Text = "not connected";
                    sqmText.Text = "not connected";
                    cloudText.Text = "not connected";
                    btnConnWeather.ForeColor = Color.White;
                    btnDiscWeather.ForeColor = Color.Gray;
                    humidText.BackColor = Color.Silver;
                    pressureText.BackColor = Color.Silver;
                    tempText.BackColor = Color.Silver;
                    sqmText.BackColor = Color.Silver;
                    cloudText.BackColor = Color.Silver;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect weather error");
            }
        }
        private void DisconnectSafety()
        {
            try
            {
                if (safetyConnected)
                {
                    safe.Connected = false;   //diconnect roof
                    safetyConnected = false;
                    imagingText.Text = "not connected";
                    btnConnSafety.ForeColor = Color.White;
                    btnDiscSafety.ForeColor = Color.Gray;
                    imagingText.BackColor = Color.Silver;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect safe error");
            }
        }

        //  Each of the disconnectxxx methods disconnects
        //  and changes state for status boxes.
        private void DisconnectMount(object sender, EventArgs e)
        {
            try
            {
                this.DisconnectTelescope();
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect mount error");
            }
        }

        private void DisconnectRoof(object sender, EventArgs e)
        {
            try
            {
                this.DisconnectDome();
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect roof error");
            }
        }

        private void DisconnectEnvironment(object sender, EventArgs e)
        {
            try
            {
                this.DisconnectWeather();
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect weather error");
            }
        }

        private void DisconnectSafetyMonitor(object sender, EventArgs e)
        {
            try
            {
                this.DisconnectSafety();
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect safety error");
            }
        }

        private void DisconnectRelay(object sender, EventArgs e)
        {
            try
            {
                this.DisconnectSwitch();
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect relay error");
            }
        }

        // tries to disconnect all, after confirmation from user
        private void DisconnectAll(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Do you want to continue?", "Disconnecting", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                if (queClose) this.CueAutoRoofClose(sender, e);
                if (queOpen) this.CueAutoRoofOpen(sender, e);
                this.DisconnectRoof(sender, e);
                this.DisconnectMount(sender, e);
                this.DisconnectEnvironment(sender, e);
                this.DisconnectSafetyMonitor(sender, e);
                mountText.BackColor = Color.Silver;
                rooftext.BackColor = Color.Silver;
                drytext.BackColor = Color.Silver;
                imagingText.BackColor = Color.Silver;
                humidText.BackColor = Color.Silver;
                pressureText.BackColor = Color.Silver;
                tempText.BackColor = Color.Silver;
                sqmText.BackColor = Color.Silver; 
                cloudText.BackColor = Color.Silver;
                statusbox.Clear();
                result = MessageBox.Show("And relays?", "Disconnecting", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes) this.DisconnectRelay(sender, e);
            }
        }

        // parks or unparks mount
        private void TogglePark(object sender, EventArgs e)
        {
            if (mountSafe == "Parked") UnparkMount();
            else if (mountSafe != "Slewing") ParkMount();  // added extra test to avoid trying to park while slewing
        }

        // parks telescope mount, only if roof is open
        private void ParkMount()
        {
            try
            {
                if (mountConnected && roofStatus == ShutterState.shutterOpen)
                {
                    if (!mount.AtPark)
                    {
                        if (mount.CanPark)
                        {
                            mountText.Text = "Parking"; 
                            if (mountId.Contains("OnStep")&& mountSafe == "Homed")
                            { mount.Tracking = true;
                              System.Threading.Thread.Sleep(1000);
                            }
                            mount.Park();
                            btnTrackTog.Text = "Tracking --";
                            btnPark.Text = "Parking";
                            if (roofConnected)  // toggle auto cue enabler, or it will unpark again
                            {
                            queOpen = false;
                            btnAutoOpen.BackColor = Color.DarkOrange;
                            btnAutoOpen.ForeColor = Color.White;
                            }
                        }
                        else statusbox.AppendText(Environment.NewLine + "parking disabled");                        
                    }
                }
                else if (roofStatus == ShutterState.shutterClosed) System.Windows.Forms.MessageBox.Show("Roof is Closed!");
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "mount park error ");
            }
        }

        // unparks telescope mount, only if roof is open
        private void UnparkMount()
        {
            try
            {
                if (mountConnected && roofStatus == ShutterState.shutterOpen)
                {
                    if (mount.AtPark)
                    {
                        if (mount.CanUnpark)
                        {
                            mount.Unpark();
                            mountText.Text = "Unparked";
                        }
                        else statusbox.AppendText(Environment.NewLine + "mount cannot unpark");
                    }
                    else System.Windows.Forms.MessageBox.Show("mount not parked!");
                }
                else if (roofStatus == ShutterState.shutterClosed) System.Windows.Forms.MessageBox.Show("Roof is closed!");
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "mount unpark error");
            }
        }

        // open, home mount, park mount and close roof (only if safe to do so)
        private void Hibernate(object sender, EventArgs e)
        {
            int i;
            try
            {
                if (queOpen) this.CueAutoRoofOpen(sender, e);  // disable autoque for Hibernate command
                if (queClose) this.CueAutoRoofClose(sender, e);
                btnHibernate.BackColor = Color.Yellow;
                btnHibernate.ForeColor = Color.Black;
                if (mountConnected && roofConnected && mount.CanPark)
                {
                    if (roofStatus == ShutterState.shutterClosed)
                    {
                        if (mountSafe == "Parked" && !busy)
                        {
                            mountText.Text = "hibernating";
                            mountText.BackColor = Color.DarkOrange;
                            statusbox.AppendText(Environment.NewLine + "opening");
                            this.OpenRoof(sender, e);
                            this.RefreshShutterState();
                        }

                        for (i = 0; i < roofTimeout && roofStatus != ShutterState.shutterOpen; ++i)
                        {
                            System.Threading.Thread.Sleep(2000);
                            Console.Beep(1000, 200);
                            this.RefreshShutterState();
                        }
                    }
                    if (roofStatus == ShutterState.shutterOpen)
                    {                       
                        if (mount.AtPark)
                        {
                            mountText.Text = "unparking"; 
                            mount.Unpark();
                            if (mountId.Contains("OnStep"))  // WD20 mount doesn't like parking from home
                            {
                                mount.Tracking = true;
                                System.Threading.Thread.Sleep(4000);
                            }
                        }

                        if (!mount.AtHome)
                        {
                            mountText.Text = "homing";
                            mount.FindHome();
                            i = 0;
                            while (!mount.AtHome && i<60)
                            {
                                System.Threading.Thread.Sleep(1000); // for V3/V4 ambiguities
                                i++;
                            }
                        }
;
                        if (mount.CanPark)
                        {
                            mountText.Text = "parking";
                            if (mountId.Contains("OnStep"))  // WD20 mount doesn't like parking from home
                            {
                                mount.Tracking = true;
                                System.Threading.Thread.Sleep(1000);
                            }
                            mount.Park(); // ensure mount is parked before closing roof
                        }
                        for (i = 0; i < mountTimeout && !mount.AtPark; ++i)
                        {
                            System.Threading.Thread.Sleep(2000);
                            Console.Beep(1000, 200);
                        }
                        System.Threading.Thread.Sleep(3000);
                        statusbox.AppendText(Environment.NewLine + "ready to close");
                        this.RefreshMount();

                        if (mountSafe == "Parked" && !busy)  // sensor agreed position
                        {
                            busy = true;
                            dome.CloseShutter();  // if mount parked, close roof
                            busy = false;
                            statusbox.AppendText(Environment.NewLine + "closing roof");
                        }
                        else System.Windows.Forms.MessageBox.Show("Error - try manual close");
                        mountText.BackColor = Color.LightGreen;
                        btnHibernate.BackColor = Color.Transparent;
                        btnHibernate.ForeColor = Color.White;
                    }
                }
                else statusbox.AppendText(Environment.NewLine + "mount cannot park");
            }

            catch (Exception)
            {
                // DisconnectMount();
                statusbox.AppendText(Environment.NewLine + "Hibernate error");
                btnHibernate.BackColor = Color.DarkOrange;
                btnHibernate.ForeColor = Color.White;
            }
        }

        // simple command to stop tracking - does not need to know roof position
        private void toggletrack(object sender, EventArgs e)
        {
            try
            {
                if (mountConnected && mount.CanSetTracking)
                {
                    if (mountSafe == "Tracking")
                    {
                        if (mount.CanSetTracking) mount.Tracking = false;
                        btnTrackTog.Text = "Tracking On";
                        mountText.Text = "Tracking Off";
                    }
                    else if (mountSafe != "Tracking")
                    {
                        if (mount.CanSetTracking) mount.Tracking = true;
                        btnTrackTog.Text = "Tracking Off";
                        mountText.Text = "Tracking On";
                    }
                }
                else statusbox.AppendText(Environment.NewLine + "no tracking control");
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "tracking off error");
            }
        }

        // sends ASCOM home command, but only if the roof is open 
        private void homemount(object sender, EventArgs e)
        {
            try
            {
                int i;
                if (mountConnected && roofConnected && roofStatus == ShutterState.shutterOpen)
                {
                    if (mount.CanUnpark)
                    {
                        mountText.Text = "Unparking";
                        if (mount.AtPark) mount.Unpark();
                    }
                    else statusbox.AppendText(Environment.NewLine + "mount cannot unpark");

                    if (mount.CanSetTracking)
                    {
                        mountText.Text = " Tracking ";
                        mount.Tracking = true;
                        // delay 4s for WD20
                        System.Threading.Thread.Sleep(4000);
                    }
                    if (mount.CanFindHome)
                    {
                        mountText.Text = "homing";
                        mount.FindHome();
                        i = 0;
                        while (!mount.AtHome && i<60)
                        {
                            System.Threading.Thread.Sleep(500); // for V3/V4 ambiguities
                            i++;
                        }

                    }
                    else
                    {
                        statusbox.AppendText(Environment.NewLine + "mount cannot home");
                        if (mount.CanSetTracking) mount.Tracking = false;
                        mountText.Text = " Stopped ";
                    }
                    }
                else if (roofStatus != ShutterState.shutterOpen) System.Windows.Forms.MessageBox.Show("No mount movements when roof not open!");
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "home error ");
            }
        }

        // force home and park ignore roof status - unique to LifeRoof motor system
        private void forcehome(object sender, EventArgs e)
        {
            try
            {
                int i;
                if (mountConnected)
                {
                    if (mount.CanUnpark)
                    {
                        mountText.Text = "Unparking";
                        if (mount.AtPark) mount.Unpark();
                    }
                    if (mount.CanSetTracking && mountId.Contains("OnStep"))  // WD20 mount doesn't like parking from home
                    {
                            mount.Tracking = true;
                            mountText.Text = " Tracking on";
                            System.Threading.Thread.Sleep(2000);
                    }
                    if (mount.CanFindHome)
                    {
                        mount.FindHome();
                        mountText.Text = "homing";
                        i = 0;
                        while (!mount.AtHome && i<60)
                        {
                            System.Threading.Thread.Sleep(1000); // for V3/V4 ambiguities
                            i++;
                        }
                    }
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "force home error");
            }
        }
        private void forcepark(object sender, EventArgs e)
        {
            try
            {
                if (mountConnected && mount.CanPark)
                {
                    if (mountId.Contains("OnStep"))  // WD20 mount doesn't like parking from home
                    {
                        mount.Tracking = true;
                        System.Threading.Thread.Sleep(1000);
                    }
                    mount.Park();
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "force park error");
            }
        }

        // force roof open and close, ignoring sensors
        private void forceopen(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("FORCEOPEN", false);
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "force open error");
            }
        }

        private void forceclose(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("FORCECLOSE", false);
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "force close error");
            }
        }

        // abort command tells Arduino to stop all movement and mount too, turns mount toggle off
        private void abort(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.AbortSlew();
                    busy = false;
                }
                if (mountConnected)
                {
                    mount.AbortSlew();
                    aborted = true;
                    mount.Tracking = false;
                    btnTrackTog.Text = "Tracking On";
                    statusbox.AppendText(Environment.NewLine + "All Abort");
                }
            }
            catch (Exception)  // turn relays off if mount is not answering
            {
                statusbox.AppendText(Environment.NewLine + "power down");
                if (switchconnected)
                {
                    TurnOffMountSwitch();
                }
            }
        }

        private void TurnOffMountSwitch()
        {
            try
            {
                for (short i=0; i<4; i++)
                {
                    if (toggle.GetSwitchName(i).Contains("ount") || toggle.GetSwitchDescription(i).Contains("ount"))  // (mount or Mount)
                    {
                        toggle.SetSwitch(i, switchOff);  // turn toggle off
                        power[i] = false;
                    }
                }
            }
            catch
            {
                statusbox.AppendText(Environment.NewLine + "mount off error");                
            }
        }

        // initiates Arduino controller - resetting rain sensor
        private void resetdome(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    var result = MessageBox.Show("Do you want to continue?", "Initialize Roof", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes) dome.CommandBlind("INIT", false);
                    busy = false;
                    statusbox.Clear();
                }
            }

            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "reset dome error");
            }
        }

        // general roof moves, with safety considerations
        private void roofclose(object sender, EventArgs e)
        {
            try
            {
                if (!mountConnected)
                {
                    const string message = "Mount not connected, do you want to continue?";
                    var result = MessageBox.Show(message, "Closing Roof", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        if (roofConnected && !busy)
                        {
                            busy = true;
                            dome.CloseShutter();
                            busy = false;
                        }
                    }
                }
                else
                {
                    if (roofConnected && mountSafe == "Parked" && !busy)
                    {
                        busy = true;
                        dome.CloseShutter();
                        busy = false;
                    }
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "roof close error");
            }

        }

        // general roof move, with safety considerations
        private void OpenRoof(object sender, EventArgs e)
        {
            try
            {
                if (!mountConnected)
                {
                    const string message = "Mount not connected, do you want to continue?";
                    var result = MessageBox.Show(message, "Opening Roof", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        if (roofConnected && !busy)
                        {
                            busy = true;
                            dome.OpenShutter();
                            busy = false;
                        }
                    }
                }
                else
                {
                    if (roofConnected && mountSafe == "Parked" && !busy)
                    {
                        busy = true;
                        dome.OpenShutter();
                        busy = false;
                    }
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "roof open error");
            }
        }

        // update sensor enable flags
        void DisplaySensor()
        {
            if (dome.CommandBool("PARKSENSOR", false))
            {
                btnEnParkSense.BackColor = Color.LightGreen;
                btnDisParkSense.BackColor = Color.Moccasin;
            }
            else
            {
                btnEnParkSense.BackColor = Color.Honeydew;
                btnDisParkSense.BackColor = Color.DarkOrange;
            }
            if (dome.CommandBool("RAINSENSOR", false))
            {
                btnEnRainSense.BackColor = Color.LightGreen;
                btnDisRainSense.BackColor = Color.Moccasin;
            }
            else
            {
                btnEnRainSense.BackColor = Color.Honeydew;
                btnDisRainSense.BackColor = Color.DarkOrange;
            }
            if (dome.CommandBool("BEEPSTATUS", false))
            {
                btnEnBeep.BackColor = Color.LightGreen;
                btnDisBeep.BackColor = Color.Moccasin;
            }
            else
            {
                btnEnBeep.BackColor = Color.Honeydew;
                btnDisBeep.BackColor = Color.DarkOrange;
            }
        }

        // special command outside standard ASCOM to modify sensor usage in Arduino
        // check this type of command goes through the hub
        private void NoRainSensor(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("NORAINSENSE", false);
                    System.Threading.Thread.Sleep(2000);
                    DisplaySensor();
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "rain disable error");
            }
        }

        // special command outside standard ASCOM to modify sensor usage in Arduino
        // when disabled, Arduino acts as if sensor always reads safe
        private void RainSensor(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("RAINSENSE", false);
                    System.Threading.Thread.Sleep(2000);
                    DisplaySensor();
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "rain enable error");
            }
        }

        // special command outside standard ASCOM to modify sensor usage in Arduino
        // when disabled, Arduino acts as if sensor always reads safe
        private void NoParkSensor(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("NOPARKSENSE", false);
                    System.Threading.Thread.Sleep(2000);
                    DisplaySensor();
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "park disable error");
            }
        }

        // special command outside standard ASCOM to modify sensor usage in Arduino
        private void ParkSensor(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("PARKSENSE", false);
                    System.Threading.Thread.Sleep(2000);
                    DisplaySensor();
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "park enable error");
            }
        }

        private void EnableBeep(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("BEEPON", false);
                    System.Threading.Thread.Sleep(2000);
                    DisplaySensor();
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "park enable error");
            }
        }

        private void DisableBeep(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("BEEPOFF", false);
                    System.Threading.Thread.Sleep(2000);
                    DisplaySensor();
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "park enable error");
            }
        }

        // toggles que open trigger flag if roof is connected for auto open when safe
        private void CueAutoRoofOpen(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected)  // toggle auto cue enabler
                {
                    if (!queOpen)
                    {
                        queOpen = true;
                        btnAutoOpen.BackColor = Color.Yellow;
                        btnAutoOpen.ForeColor = Color.Black;
                    }
                    else
                    {
                        queOpen = false;
                        btnAutoOpen.BackColor = Color.Transparent;
                        btnAutoOpen.ForeColor = Color.White;
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Roof not connected");
                    queClose = false;
                    queOpen = false;
                    btnAutoOpen.BackColor = Color.DarkOrange;
                    btnAutoOpen.ForeColor = Color.White;
                    btnAutoClose.BackColor = Color.DarkOrange;
                    btnAutoClose.ForeColor = Color.White;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "que open error");
            }
        }

        // toggles que close trigger flag if roof is connected for when auto close when not safe
        private void CueAutoRoofClose(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected)   // toggle auto cue enabler
                {
                    if (!queClose)
                    {
                        queClose = true;
                        btnAutoClose.BackColor = Color.Yellow;
                        btnAutoClose.ForeColor = Color.Black;
                    }
                    else
                    {
                        queClose = false;
                        btnAutoClose.BackColor = Color.Transparent;
                        btnAutoClose.BackColor = Color.White;
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Roof not connected");
                    queClose = false;
                    queOpen = false;
                    btnAutoOpen.BackColor = Color.DarkOrange;
                    btnAutoClose.BackColor = Color.DarkOrange;
                    btnAutoClose.BackColor = Color.White;
                    btnAutoOpen.ForeColor = Color.White;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "que open error");
            }
        }

        // If one or both of the autocues is active, AutoRoof works out whether to open or 
        // close the roof and park the mount depending on weather conditions.
        // It uses an amalgam of multiple sensors to determine the action.
        private void AutoRoof()
        {
            // initial states, weathersafe is conglomerate of multiple rain, humidity and safety sensors
            int i;
            try
            {
                if (mountConnected && roofConnected && !busy)  // only if mount and roof are connected
                {
                    // check for auto close conditions
                    if (roofStatus == ShutterState.shutterOpen && !goodConditions && queClose)
                    {
                        mountText.Text = "Parking";
                        if (!mount.AtPark && mount.CanPark)
                        {
                            mount.Tracking = true;
                            System.Threading.Thread.Sleep(1000);
                            mount.Park(); // ensure mount is parked before closing roof
                        }
                        for (i = 0; i < mountTimeout && !mount.AtPark; ++i)
                        {
                            System.Threading.Thread.Sleep(2000);
                            Console.Beep(1000, 200);
                        }
                        System.Threading.Thread.Sleep(3000);
                        this.RefreshMount();
                        if (mountSafe == "Parked" && !busy)  // sensor agreed position
                        {
                            busy = true;
                            dome.CloseShutter();  // if mount parked, close roof
                            busy = false;
                            if (!dry) TurnOffMountSwitch(); // for extreme poor conditions, shut down
                            queOpen = false;
                            btnAutoOpen.BackColor = Color.DarkOrange;
                            btnAutoOpen.ForeColor = Color.White;
                            statusbox.AppendText(Environment.NewLine + "auto close");
                        }
                        else System.Windows.Forms.MessageBox.Show("busy, try manual close");
                        this.RefreshShutterState();
                    }

                    // check for auto open conditions
                    else if (roofStatus == ShutterState.shutterClosed && goodConditions && queOpen)
                    {
                        if (mountSafe == "Parked" && !busy) // use sensor rather than mount status
                        {
                            busy = true;
                            dome.OpenShutter();
                            busy = false;
                            statusbox.AppendText(Environment.NewLine + "auto open");
                        }
                        else System.Windows.Forms.MessageBox.Show("busy, try manual open");
                        this.RefreshShutterState();
                    }
                    // experiment to trap roof closing on unparked mount
                    else if (roofStatus == ShutterState.shutterClosing && mountSafe != "Parked")
                        dome.AbortSlew();
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "auto roof error");
            }
        }

        //  checks if mount is parked and everything otherwise is good and finds home - uses sensor for mount position rather than parked status 
        //  due to unkonwn mount condition at power up, even if in park position,
        //  double checks that mount is not moving while the roof is closed and stops it dead if it is 
        private void AutoMount()
        {
            try
            {
                int i;
                if (mountConnected && roofConnected)  // only if mount and roof are connected
                {
                    //auto open
                    if (roofStatus == ShutterState.shutterOpen && mountSafe == "Parked" && goodConditions && queOpen)  // if it opened, move mount
                    {
                        mountText.Text = "unparking";
                        if (mount.AtPark)
                        {
                            mount.Unpark();
                            mount.Tracking = true;
                            // delay 4s for WD20
                            System.Threading.Thread.Sleep(4000);
                        }
                        mountText.Text = "homing";
                        mount.FindHome();
                        i = 0;
                        while (!mount.AtHome && i<60)
                        {
                            System.Threading.Thread.Sleep(1000); // for V3/V4 ambiguities
                            i++;
                        }
                        // disable que open
                        queOpen = false;
                        btnAutoOpen.BackColor = Color.DarkOrange;
                        btnAutoOpen.ForeColor = Color.White;
                        RefreshMount();  // update mount status
                    }

                    // with new roof status, if mount is doing anything strange, need to stop it dead - either with roof closed or in error state
                    if (roofStatus != ShutterState.shutterOpen)
                    {
                        if (mountSafe != "Parked")
                        {
                            if (mountSafe == "Slewing")
                            {
                                mount.AbortSlew();
                                aborted = true;
                                mountText.Text = "aborting";
                                statusbox.AppendText(Environment.NewLine + "aborting slew");
                            }
                            mount.Tracking = false; // stop tracking
                            btnTrackTog.Text = "Tracking --";
                            statusbox.AppendText(Environment.NewLine + "aborted tracking");
                            statusbox.AppendText(Environment.NewLine + "make safe, home mount");
                            if (switchconnected)
                            {
                                if (mountConnected) DisconnectTelescope();
                                TurnOffMountSwitch(); // look for switch labelled 'Mount' or 'mount' and disconnect mount power
                                aborted = true;
                                statusbox.AppendText(Environment.NewLine + "mount power off");
                            }
                            // disable que open
                            queOpen = false;
                            btnAutoOpen.BackColor = Color.DarkOrange;
                            btnAutoOpen.ForeColor = Color.White;
                        }
                    }
                }
                // added further catch in case ASCOM mount is disconnected
                else if (roofConnected)
                {
                    if (roofStatus != ShutterState.shutterOpen && mountSafe != "Parked") // in case mount is disconnected
                    {
                        if (switchconnected)
                        {
                            // conventional power control 
                            TurnOffMountSwitch();  // look for switch labelled 'Mount' or 'mount' and disconnect mount power
                            aborted = true;
                            statusbox.AppendText(Environment.NewLine + "mount power off");
                        }
                    }
                }
            }
            
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "mount ASCOM error");
            }
        }

        // status refresh routines - for roof, weather, safety monitor and mount, called by form timer
        // creates overall good conditions status from various readings
        // checks on mount and roof actions every cycle weather stuff on every other
        private void RefreshAll(object sender, EventArgs e)
        {
            try
            {
                RefreshShutterState();
                RefreshMount();
                AutoMount(); // checks mount logic
                AutoRoof();  // checks roof logic
                if (!priority)
                {
                    RefreshSwitch();
                    RefreshWeather();
                    RefreshSafetyMonitor();
                    goodConditions = dry && clearAir && clearSky;  // amalgamation of air, sky and weather                          
                }
                priority = !priority;                
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "data refresh error");
            }
        }

        // small routine to update switches, in case there is a parallel hub command - every 4 cycles
        private void RefreshSwitch()
        {
            try  // all at once
            {
                if (switchconnected)
                {
                    GetSwitchState();
                    ShowSwitchState();
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "switch refresh error");
            }
        }

        // uses safetymonitor to update status and display
        private void RefreshSafetyMonitor()
        {
            try
            {
                if (safetyConnected)
                {
                    if (safe.IsSafe)
                    {
                        imagingText.BackColor = Color.LightGreen;
                        imagingText.Text = "OK to Open";
                        clearSky = true;  // update weatherstatus
                    }
                    else
                    {
                        imagingText.BackColor = Color.DarkOrange;
                        imagingText.Text = "NOK to Open";
                        clearSky = false;  // update weatherstatus
                    }
                }
                else clearSky = true;  // if no safetydevice connected
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "safety monitor error");
            }
        }

        // checks shutter status from latest Arduino broadcast
        private void RefreshShutterState()
        {
            try
            {
                if (roofConnected && !busy)
                {
                    //  get shutter status
                    busy = true;
                    roofStatus = dome.ShutterStatus;  // get shutter status   
                    busy = false;
                    switch (roofStatus)
                    {
                        case ShutterState.shutterOpen:
                            rooftext.Text = "Open";
                            rooftext.BackColor = Color.LightGreen;
                            break;
                        case ShutterState.shutterClosed:
                            rooftext.Text = "Closed";
                            rooftext.BackColor = Color.DarkOrange;
                            break;
                        case ShutterState.shutterOpening:
                            rooftext.Text = "Opening";
                            rooftext.BackColor = Color.Yellow;
                            break;
                        case ShutterState.shutterClosing:
                            rooftext.Text = "Closing";
                            rooftext.BackColor = Color.Yellow;
                            break;
                        case ShutterState.shutterError:
                            rooftext.Text = "Error";
                            rooftext.BackColor = Color.DarkOrange;
                            break;
                        default:
                            rooftext.Text = "Comms error";
                            rooftext.BackColor = Color.DarkOrange;
                            break;
                    }
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "roofStatus error");
            }
        }

        /* refreshes status of rain sensor connected to Arduino and other sensors. Note that
         the rain is detected with a non-standard ASCOM command through its CommandBool method. If
         you are using a hub, you need to make sure that it passes all commands and not just the
         standard ones */
        private void RefreshWeather()
        {
            try
            {
                if (roofConnected) // assumes  LifeRoof system with rain sensor
                {
                    if (!busy)
                    {
                        busy = true;
                        dry = !dome.CommandBool("RAIN", false);
                        busy = false;
                        if (!dry)
                        {
                            roofSafe = "Rain";
                            drytext.BackColor = Color.DarkOrange;
                        }
                        else
                        {
                            roofSafe = "Dry";
                            drytext.BackColor = Color.LightGreen;
                        }
                        drytext.Text = roofSafe;
                        // update filesafe file
                        if (priorDry != dry)
                        {
                            priorDry = dry;
                            WriteSafeFile(dry); // update status
                        }
                    }
                }
                else dry = true;  // default if no rain detector
                if (weatherConnected)
                {
                    pressureText.Text = Math.Round(weather.Pressure, 2).ToString() + " hPa";
                    tempText.Text = Math.Round(weather.Temperature, 2).ToString() + " °C";
                    humidText.Text = Math.Round(weather.Humidity, 2).ToString() + " %";
                    sqmText.Text = Math.Round(weather.SkyQuality, 1).ToString() + " M/asec2";
                    cloudText.Text = Math.Round(CloudCalculator(weather.SkyTemperature), 1).ToString() + " %";
                    if (weather.Humidity > maxHumidity)
                    {
                        humidText.BackColor = Color.DarkOrange;
                        clearAir = false;
                    }
                    else
                    {
                        humidText.BackColor = Color.LightGreen;
                        clearAir = true;
                    }
                    ChartDataUpdate();
                    ChartUpdate();
                }
                else clearAir = true;  // default for no weather connection
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "weather error");
            }
        }
        //  CloudCalculator takes sky temperature and scales to % CloudCalculator cover - auto adjusts limits
        private double CloudCalculator(double adjtemp)
        {
            // automatic trim of limits SkyUL and SkyLL to sky temp extremes - but less than error state limits
            if (adjtemp < skyTmin)
            {
                SkyLL = skyTmin;// very clear, less than limit
                adjtemp = skyTmin;
            }
            if (adjtemp > skyTmax)
            {
                SkyUL = skyTmax; // very cloudy
                adjtemp = skyTmax;
            }
            // otherwise modifiy SkyLL and SkyUL between outer limts
            if (adjtemp < SkyLL && adjtemp > skyTmin) SkyLL = adjtemp;// very clear, less than limit
            if (adjtemp > SkyUL && adjtemp < skyTmax) SkyUL = adjtemp; // very cloudy

            // CloudCalculator cover calculation - ratio of sky temp between limits
            return(100 * (adjtemp - SkyLL) / (SkyUL - SkyLL));
        }
        // routine to change over the graph data source and y axis and update
        private void graphselect(object sender, EventArgs e)
        {
            if ((string)btngraphsel.SelectedItem == "temp C")
            {
                chartType = 0;
                chart1.ChartAreas[0].AxisY.Minimum = -10;
                chart1.ChartAreas[0].AxisY.Maximum = +30;
            }
            if ((string)btngraphsel.SelectedItem == "humidity %")
            {
                chartType = 1;
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                chart1.ChartAreas[0].AxisY.Maximum = 100;
            }
            if ((string)btngraphsel.SelectedItem == "dewpoint C")
            {
                chartType = 2;
                chart1.ChartAreas[0].AxisY.Minimum = -10;
                chart1.ChartAreas[0].AxisY.Maximum = +20;
            }
            if ((string)btngraphsel.SelectedItem == "sky quality SQM")
            {
                chartType = 3;
                chart1.ChartAreas[0].AxisY.Minimum = 10;
                chart1.ChartAreas[0].AxisY.Maximum = 20;
            }
            if ((string)btngraphsel.SelectedItem == "CloudCalculator %")
            {
                chartType = 4;
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                chart1.ChartAreas[0].AxisY.Maximum = 100;
            }
            // update chart values
            ChartDataUpdate(); // update latest values and transpose
            // either copy non zero over, or use all values
            if (sampleCount < 7200) //(every minute)
            {
                var arrayNZ = chartValues.Select(x => x).Where(x => x != 0).ToArray(); // only do non-zero values at the beginning
                chart1.Series[0].Points.DataBindY(arrayNZ);
            }
            else chart1.Series[0].Points.DataBindY(chartValues);
        }

        private void InitMountParams(object sender, EventArgs e)
        {
            try
            {
                if (mount.Connected)
                {
                    double guiderate;
                    mount.UTCDate = (System.DateTime.UtcNow);
                    mount.SiteElevation = Convert.ToDouble(ElevationIP.Text);
                    guiderate = sidereal * (double)(GuideRate.Value);  // calculate guide rate
                    if (mount.CanSetGuideRates) mount.GuideRateRightAscension = guiderate;
                    mount.SiteLongitude = Convert.ToDouble(LongitudeIP.Text);
                    mount.SiteLatitude = Convert.ToDouble(LatitudeIP.Text);
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("communication failure");
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void ElevationIP_TextChanged(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void altitude_TextChanged(object sender, EventArgs e)
        {

        }

        // ChartData updates several arrays of weather data, covering last 3 hours and 
        // copies one set over to the chart array, for display, according to chartType's value
        // waits for a minute upon startup before displaying anything
        private void ChartDataUpdate()
        {
            int index;
            sampleCount += 1;
            if(sampleCount > 59)
            {
                if ((sampleCount < 7200) && (sampleCount % 60 == 0)) //(every minute)
                {
                    index = ((int)((sampleCount / 60.0) - (sampleCount % 60)));
                    tempValues[index] = Math.Round(weather.Temperature, 2);
                    humidValues[index] = Math.Round(weather.Humidity, 2);
                    dewValues[index] = Math.Round(weather.DewPoint, 2);
                    SQMValues[index] = Math.Round(weather.SkyQuality, 2);
                    cloudValues[index] = Math.Round(CloudCalculator(weather.SkyTemperature), 1);
                }
                if ((sampleCount >= 7200) && (sampleCount % 60 == 0))
                {
                    for (int i = 0; i < 119; i++)  // shift left
                    {
                        tempValues[i] = tempValues[i + 1];
                        humidValues[i] = humidValues[i + 1];
                        dewValues[i] = dewValues[i + 1];
                        SQMValues[i] = SQMValues[i + 1];
                        cloudValues[i] = cloudValues[i + 1];
                    }
                    // rhs value is current value
                    tempValues[119] = Math.Round(weather.Temperature, 2);
                    humidValues[119] = Math.Round(weather.Humidity, 2);
                    dewValues[119] = Math.Round(weather.DewPoint, 2);
                    SQMValues[119] = Math.Round(weather.SkyQuality, 2);
                    cloudValues[119] = Math.Round(CloudCalculator(weather.SkyTemperature), 1);
                }
            }
            
            switch (chartType) // copy applicable data into chart array
            {
                case 0:
                    chartValues = tempValues;
                    //for (int i = 0; i < 120; i++) chartValues[i] = tempValues[i];
                    break;
                case 1:
                    chartValues = humidValues;
                    //for (int i = 0; i < 120; i++) chartValues[i] = humidValues[i];
                    break;
                case 2:
                    chartValues = dewValues;
                    //for (int i = 0; i < 120; i++) chartValues[i] = dewValues[i];
                    break;
                case 3:
                    chartValues = SQMValues;
                    //for (int i = 0; i < 120; i++) chartValues[i] = SQMValues[i];
                    break;
                default:
                    chartValues = cloudValues;
                    //for (int i = 0; i < 120; i++) chartValues[i] = cloudValues[i];
                    break;
            }
        }
        // updates chart plot, according to the selected data source, does not display zero values.
        private void ChartUpdate()
        {
            if (sampleCount > 59)
            {
                if ((sampleCount < 7200) && (sampleCount % 60 == 0)) //(every minute)
                {
                    var arrayNZ = chartValues.Select(x => x).Where(x => x != 0).ToArray(); // only do non-zero values at the beginning
                    chart1.Series[0].Points.DataBindY(arrayNZ);
                }
                if ((sampleCount >= 7200) && (sampleCount % 60 == 0)) chart1.Series[0].Points.DataBindY(chartValues);
            }
        }
        /* updates mount from status
        note that unlike most systems, this uses an amalgamation of the mount position and a separate
        sensor which is part of the roof system. This is due to the fact that most mounts have
        programmable park positions and if set incorrectly, the mount may not clear the roofline. The sensor uses
        a non-standard ASCOM command through its CommandBool method. If you are using a hub for the roof, you
        need to make sure it passes all commands through and not just the standard ones 
        mountSafe has JUST these outcomes:
        "Parked"
        "Not at Park"
        "Parked" is sensor(no mount) or sensor/mount confirm
        "Tracking"
        "Homed"
        "Slewing"
        "Stopped"
        any movement (tracking, slewing, homing) from mount invalidates Park status
         */
        private void RefreshMount()
        {
            bool parkconfirm;  // park sensor status
            try
            {
                string trackingtext = "Tracking On"; // assume it is not tracking by default
                string parktext = "Park"; // assume not at park
                // roof not connected or busy//default
                mountSafe = "Stopped";
                if (mountConnected)
                {
                    // get coordinates
                    // update coordinates
                    declination.Text = Math.Round(mount.Declination, 4).ToString();
                    RAcension.Text = Math.Round(mount.RightAscension, 4).ToString();
                    altitude.Text = Math.Round(mount.Altitude, 4).ToString();
                    azimuth.Text = Math.Round(mount.Azimuth, 4).ToString();
                    if (mount.AtPark)
                    {
                        mountSafe = "Parked";
                        parktext = "Unpark";
                    }
                }
                if (roofConnected && !busy) // if roof connected, override park position using sensors rather than mount
                {
                    busy = true;
                    parkconfirm = dome.CommandBool("PARK", false);  // from sensor
                    busy = false;
                    if (parkconfirm)
                    {
                        mountSafe = "Parked";
                        parktext = "Unpark";
                    }
                    else  // sensor does not say it is at park
                    {
                        mountSafe = "Not at Park";
                        parktext = "Park";
                    }
                }

                if (mountConnected)  // update mount status from mount and potentially override Park status
                {
                    if (mount.Tracking)
                    {
                        mountSafe = "Tracking";
                        trackingtext = "Tracking Off";
                        aborted = false;
                        parktext = "Park";
                    }
                    else if (mount.AtHome)
                    {
                        mountSafe = "Homed";
                        aborted = false;
                        trackingtext = "Tracking On";
                        parktext = "Park";
                    }
                    if (!aborted)
                    {
                        if (mount.Slewing)  // mount can report slewing and tracking at same time, so worse case is slewing for abort actionss
                        {
                            mountSafe = "Slewing";
                            trackingtext = "Tracking --";
                            parktext = "Park --";
                        }
                    }

                }
                btnTrackTog.Text = trackingtext;  // doing this here with a string variable stops flicker on display
                btnPark.Text = parktext;
                mountText.Text = mountSafe;  // update display
                if (mountSafe == "Parked") mountText.BackColor = Color.LightGreen;
                else mountText.BackColor = Color.DarkOrange;
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "mount refresh error");
            }
        }

        // power switch functions  ( x4)
        // toggleswitch(x) is customized to select different outputs depending on the connected hardware
        // or if multiple devices are being combined using darkskygeek's hub
        private void ToggleSwitch0(object sender, EventArgs e)
        {
            try
            {
                if (switchconnected)
                {
                    if (power[0]) toggle.SetSwitch(0, switchOff);  // turn toggle off                   
                    else toggle.SetSwitch(0, switchOn);
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "relay 0 error");
            }
        }
        private void ToggleSwitch1(object sender, EventArgs e)
        {
            try
            {
                if (switchconnected)
                {
                    /*if (switchID == "ASCOM.DarkSkyGeek.SwitchHub")
                    {
                        if (power[1]) toggle.SetSwitch(10, switchOff);  // turn toggle off                   
                        else toggle.SetSwitch(10, switchOn);
                    }*/
                    if (power[1]) toggle.SetSwitch(1, switchOff);  // turn toggle off                   
                        else toggle.SetSwitch(1, switchOn);
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "switch 1 error");
            }
        }
        
        private void ToggleSwitch2(object sender, EventArgs e)
        {
            try
            {
                if (switchconnected)
                {
                    /*if (switchID == "ASCOM.DarkSkyGeek.SwitchHub")
                    {
                        if (power[1]) toggle.SetSwitch(10, switchOff);  // turn toggle off                   
                        else toggle.SetSwitch(10, switchOn);
                    }*/
                    if (power[2]) toggle.SetSwitch(2, switchOff);  // turn toggle off                   
                    else toggle.SetSwitch(2, switchOn);
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "switch 2 error");
            }
        }
        private void ToggleSwitch3(object sender, EventArgs e)
        {
            try
            {
                if (switchconnected)
                {
                    /*if (switchID == "ASCOM.DarkSkyGeek.SwitchHub")
                    {
                        if (power[1]) toggle.SetSwitch(10, switchOff);  // turn toggle off                   
                        else toggle.SetSwitch(10, switchOn);
                    }*/
                    if (power[3]) toggle.SetSwitch(3, switchOff);  // turn toggle off                   
                    else toggle.SetSwitch(3, switchOn);
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "switch 3 error");
            }
        }

        // reads switch position - normal logic
        private void GetSwitchState()
        {
            try
            {
                if (switchconnected)
                {
                    power[0] = toggle.GetSwitch(0);
                    /*if (switchID == "ASCOM.DarkSkyGeek.SwitchHub")
                    {
                        power[1] = toggle.GetSwitch(10);
                        power[2] = toggle.GetSwitch(12);
                        power[3] = toggle.GetSwitch(13); 
                    }*/
                    for (short i = 1; i < 4; i++) power[i] = toggle.GetSwitch(i); //all at once                   
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "Relay status error");
            }
        }

        // display toggle state
        private void ShowSwitchState()
        {
            try
            {
                if (power[0]) btnSwitch1.BackColor = Color.LightGreen;
                else btnSwitch1.BackColor = Color.DarkOrange;
                if (power[1]) btnSwitch2.BackColor = Color.LightGreen;
                else btnSwitch2.BackColor = Color.DarkOrange;
                if (power[2]) btnSwitch3.BackColor = Color.LightGreen;
                else btnSwitch3.BackColor = Color.DarkOrange;
                if (power[3]) btnSwitch4.BackColor = Color.LightGreen;
                else btnSwitch4.BackColor = Color.DarkOrange;
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "Relay status error");
            }
        }

        // utility functions

        // read humidity value from form and store
        private void SetHumidityLimit(object sender, EventArgs e)
        {
            try
            {
                maxHumidity = Convert.ToDouble(humidlimit.Value);
                this.WriteFile();
            }
            catch
            {
                statusbox.AppendText(Environment.NewLine + "humidity set error");
            }
        }

        // experimental overall shuttdown method - disconnect devices and then power through switches
        private void Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (mountConnected) DisconnectTelescope();  // disconnect mount driver
                if (roofConnected) DisconnectDome();  // disconnect roof driver
                if (safetyConnected) DisconnectSafety();  // disconnect safety monitor
                if (weatherConnected) DisconnectWeather();  // disconnect weather monitor
                if (switchconnected)
                {
                    for (short i = 0; i < 4; i++)
                    {
                        toggle.SetSwitch(i, switchOff); // turn off all switches
                        power[i] = false;                                 
                    }
                    DisconnectSwitch();  // now disconnect ASCOM driver
                }
                Application.Exit();
            }
            catch
            {
                statusbox.AppendText(Environment.NewLine + "aborting error");
            }
        }

        // WriteFile() saves device choices and variables to MyDocuments/ASCOM/Obsy/obsy.txt
        private void WriteFile()
        {
            try
            {
                string[] configure = new string[7];
                configure[0] = domeId;
                configure[1] = mountId;
                configure[2] = weatherId;
                configure[3] = safetyId;
                configure[4] = maxHumidity.ToString();
                configure[5] = switchID;
              
                if (!System.IO.Directory.Exists(path)) System.IO.Directory.CreateDirectory(path);
                File.WriteAllLines(path + settingsfile, configure);
            }
            catch (System.UnauthorizedAccessException e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }
        // WriteSafeFile writes 'safe' or 'unsafe' to safety.txt file
        private void WriteSafeFile(bool safe)
        {
            try
            {
                string drystring;
                if (!System.IO.Directory.Exists(path)) System.IO.Directory.CreateDirectory(path);
                if (safe) drystring = "safe";
                else drystring = "unsafe";
                File.WriteAllText(path + dryfile, drystring);
            }
            catch (System.UnauthorizedAccessException e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }
        // ReadFile() reads device choices and variables from MyDocuments/ASCOM/Obsy/obsy.txt
        private void ReadFile()
        {
            try
            {
                string[] configure = new string[7];
                configure = File.ReadAllLines(path + settingsfile);
                domeId = configure[0];
                mountId = configure[1];
                weatherId = configure[2];
                safetyId = configure[3];
                maxHumidity = Convert.ToDouble(configure[4]);
                humidlimit.Value = Convert.ToDecimal(maxHumidity);
                switchID = configure[5];
            }
            catch (System.UnauthorizedAccessException e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }
    }
}
