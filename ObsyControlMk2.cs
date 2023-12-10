using ASCOM.DeviceInterface;
using ASCOM.DriverAccess;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
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
// 10.1 - added safety shutdown 
// 10.2 - modified to use a range of safety monitors but still fires up AAG CoudWatcher if it detect file output
// 10.5 - delayed AAG graph output for 1 minute to prevent exception, conditional 4-way switch control
// 10.6 - added cloud cover (sky temperature)
// 11.0 - overhaul, including better comments
// 11.1 - changed sky temperature to cloud cover
// 11.2 - changed 1s form timer to 2s and 0.5s for non- and critical conditions

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
        private bool mountConnected, roofConnected, weatherconnected, safetyconnected, switchconnected;  // local variable to note current connected state
        private string mountsafe, roofsafe;      // used for local storage
        private ShutterState roofstatus;             // used to store shutter status
        String path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ASCOM\\Obsy";
        String safetypath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\AAG";
        private bool queOpen, queClose, noRain, goodConditions, clearAir, clearSky; // boolean flags for deciding if roof is on "auto" mode operation
        private double maxhumidity; // threshold for fog/mist
        private int mountimeout;  // period in multiples of 2 seconds for error condition to apply to movements
        private int rooftimeout;  // period in multiples of 2 seconds for roof to move
        private bool aborted;  // set if abort command is issued, reset if mount is tracking, homed or parked
        private bool[] power = new bool[4]; // local values of toggle states
        // arrays to hold last three hours of data for charting
        private double[] chartvalues = new double[120];
        private double[] tempvalues = new double[120];
        private double[] humidvalues = new double[120];
        private double[] dewvalues = new double[120];
        private double[] SQMvalues = new double[120];
        private double[] rainvalues = new double[120];
        private double[] cloudvalues = new double[120];
        private int samplecount = 0;  // sampling value increments every 2 seconds
        private int charttype = 0;  //selection variable
        // changes the true and false around on the following two lines to change the toggle logic sense
        private const bool switchon = true;
        private const bool switchoff = false;
        private double SkyUL = 10; // 100% cloud cover (initial value)
        private double SkyLL = -10; // 0% cloud cover (iniitial value)
        private const double skyTmax = 25;
        private const double skyTmin = -25;//private short relayinterval = 0; // used for original Pegasus toggle box that needed slow polled switch reads

        public Obsyform()
        {
            InitializeComponent();
            // initialise form status boxes
            drytext.Text = "not connected";
            rooftext.Text = "not connected";
            mountext.Text = "not connected";
            humidtext.Text = "not connected";
            temptext.Text = "not connected";
            pressuretext.Text = "not connected";
            sqmtext.Text = "not connected";
            imagingtext.Text = "not connected";
            cloudtext.Text = "not connected";
            statusbox.Text = "";
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
            weatherconnected = false;
            safetyconnected = false;
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
            maxhumidity = 97.0; // default but overidden by file read
            mountimeout = 15; // response time period for mount to move (x2)
            rooftimeout = 20;  //  40 seconds for roof to complete movement
            // initialise flags  used to build up a composite safety flag
            noRain = false;
            clearAir = false;
            clearSky = false;
            goodConditions = false;
            aborted = false;
            roofstatus = ShutterState.shutterError; // assume an error to start with
            // read in, if it exists, the connection settings file           
            if (File.Exists(path + "\\obsy.txt")) fileread();
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
        private void setroof(object sender, EventArgs e)
        {
            try
            {
                domeId = Dome.Choose(domeId);
                this.filewrite();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("roof selection failed");
            }

        }
        // telescope (mount) device
        private void setmount(object sender, EventArgs e)
        {
            try
            {
                mountId = Telescope.Choose(mountId);
                this.filewrite();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("mount selection failed");
            }
        }

        // observing conditions device
        private void setweather(object sender, EventArgs e)
        {
            try
            {
                weatherId = ObservingConditions.Choose(weatherId);
                this.filewrite();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("weather selection failed");
            }
        }
        // safety monitor device ( either separate or using ASCOM device that in turn links to weather monitors
        private void setsafety(object sender, EventArgs e)
        {
            try
            {
                safetyId = SafetyMonitor.Choose(safetyId);
                this.filewrite();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("safety selection failed");
            }
        }
        // switch devices (binary)
        private void setswitch(object sender, EventArgs e)
        {
            try
            {
                switchID = Switch.Choose(switchID);
                this.filewrite();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("switch selection failed");
            }
        }

        //  Each of the connectxxx methods create an object instance, connect to object
        //  and set up initial states for status boxes.
        private void connectroof(object sender, EventArgs e)
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
                        sensorDisplay(); // update sensor backgrounds to indicate sensors
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
        private void connecttelescope(object sender, EventArgs e)
        {
            try
            {
                if (mountId != null)
                {
                    mount = new Telescope(mountId);
                    mount.Connected = true;
                    mountConnected = true;
                    mountext.Text = "connected";
                    btnConnMount.ForeColor = Color.Gray;
                    btnDiscMount.ForeColor = Color.White;
                    if(!mount.CanFindHome || !mount.CanPark || !mount.CanPulseGuide)
                        System.Windows.Forms.MessageBox.Show("check mount park/home/guide settings");
                }
                else
                {
                    mountConnected = false;
                    mountext.Text = "not connected";
                    btnConnMount.ForeColor = Color.White;
                    btnDiscMount.ForeColor = Color.Gray;
                }
            }
            catch (Exception)  // added some changes to status, to prevent new connections
            {
                mountConnected = false;
                mountext.Text = "not connected!";
                System.Windows.Forms.MessageBox.Show("Mount did not connect");
            }
        }
        private void connectswitch(object sender, EventArgs e)
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
                    btnSwitch1.Text = toggle.GetSwitchDescription(0);
                    if(switchID == "ASCOM.DarkSkyGeek.SwitchHub")
                    {        
                        btnSwitch2.Text = toggle.GetSwitchDescription(12); // Camera 1 usb
                        btnSwitch3.Text = toggle.GetSwitchDescription(13); // Camera 2 usb
                        btnSwitch4.Text = toggle.GetSwitchDescription(14); // guider usb
                    }
                    if (switchID == "ASCOM.PegasusAstroPPBAdvance.Switch")
                    {
                        btnSwitch2.Text = toggle.GetSwitchDescription(1); // Camera 1 usb
                        btnSwitch3.Text = toggle.GetSwitchDescription(2); // aps-c
                        btnSwitch4.Text = toggle.GetSwitchDescription(3); // dew
                    }
                    if (switchID == "ASCOM.PegasusAstroUPBv2.Switch")
                    {
                        btnSwitch2.Text = toggle.GetSwitchDescription(4); // cam 1 usb
                        btnSwitch3.Text = toggle.GetSwitchDescription(5); // cam 2 usb
                        btnSwitch4.Text = toggle.GetSwitchDescription(6); // guider
                    }
                        getswitchstate();  // all at once
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
         
        private void connectweather(object sender, EventArgs e)
        {
            try
            {
                if (weatherId != null)
                {
                    // as safety monitor also has environmental sensors;
                    if (!safetyconnected) connectsafety(sender, e);
                    // connect to weatherstick or other device
                    weather = new ObservingConditions(weatherId);
                    weather.Connected = true;
                    weatherconnected = true;
                    humidtext.Text = "connected";
                    pressuretext.Text = "connected";
                    temptext.Text = "connected";
                    sqmtext.Text = "connected";
                    cloudtext.Text = "connected";
                    btnConnWeather.ForeColor = Color.Gray;
                    btnDiscWeather.ForeColor = Color.White;
                    humidtext.BackColor = Color.LightGreen;
                    pressuretext.BackColor = Color.LightGreen;
                    sqmtext.BackColor = Color.LightGreen;
                    temptext.BackColor = Color.LightGreen;
                    cloudtext.BackColor = Color.LightGreen;
                    // default chart is temp C
                    charttype = 0;
                    chart1.ChartAreas[0].AxisY.Minimum = -5;
                    chart1.ChartAreas[0].AxisY.Maximum = 25;
                    btngraphsel.Text = "Temp C";
                    refreshweather();

                }
                else // weather device not set
                {
                    weatherconnected = false;
                    humidtext.Text = "not connected";
                    pressuretext.Text = "not connected";
                    sqmtext.Text = "not connected";
                    temptext.Text = "not connected";
                    cloudtext.Text = "not connected";
                    btnConnWeather.ForeColor = Color.White;
                    btnDiscWeather.ForeColor = Color.Gray;
                    humidtext.BackColor = Color.Silver;
                    pressuretext.BackColor = Color.Silver;
                    sqmtext.BackColor = Color.Silver;
                    temptext.BackColor = Color.Silver;
                    cloudtext.BackColor = Color.Silver;
                }
            }
            catch (Exception)
            {
                weather.Connected = false;
                weatherconnected = false;
                System.Windows.Forms.MessageBox.Show("AAG/WS did not connect");
            }
        }
        // special case if it detects AAG cloudwatcher log file
        private void connectsafety(object sender, EventArgs e)
        {
            try
            {
                if (safetyId != null)
                {
                    if (File.Exists(safetypath + "\\AAG_SLD.dat")) // detect presence of AAG Cloudwatcher
                    {
                        Type foo = Type.GetTypeFromProgID("AAG_CloudWatcher.CloudWatcher");
                        dynamic oCW;
                        oCW = Activator.CreateInstance(foo);
                        //oCW.WindowLeft = 0;
                        //oCW.WindowTop = 0;
                        oCW.Device_Start();
                        oCW.RecordStart(true);
                        oCW = null;                       
                        safe = new SafetyMonitor(safetyId);
                        safe.Connected = true;
                        safetyconnected = true;
                        imagingtext.Text = "connected";
                        btnConnSafety.ForeColor = Color.Gray;
                        btnDiscSafety.ForeColor = Color.White;
                    }
                }
                else
                {
                    safetyconnected = false;
                    imagingtext.Text = "not connected";
                    btnConnSafety.ForeColor = Color.White;
                    btnDiscSafety.ForeColor = Color.Gray;
                }
            }
            catch (Exception)
            {
                safe.Connected = false;
                safetyconnected = false;
                System.Windows.Forms.MessageBox.Show("Safety Monitor did not connect");
            }
        }

        // catchall shortcut does connect to all equipment
        // starts with SWITCH device only, to allow toggle setup before further connections
        // if SWITCH device connected, it connects the remaining devices
        public void connectall(object sender, EventArgs e)
        {
            statusbox.Clear();
            if (!switchconnected)
            {
                connectswitch(sender, e);
                System.Windows.Forms.MessageBox.Show("check equipment power");
            }

            else
            {
                connectsafety(sender, e);
                connectroof(sender, e);
                connecttelescope(sender, e);
                connectweather(sender, e);
            }

        }
        // functions to do the device disconnections (called by several button presses)
        private void disctelescope()
        {
            try
            {
                if (mountConnected)
                {
                    mount.Connected = false;  // disconnect mount
                    mountConnected = false;
                    mountext.Text = "not connected";
                    btnConnMount.ForeColor = Color.White;
                    btnDiscMount.ForeColor = Color.Gray;
                    mountext.BackColor = Color.Silver;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect mount error");
            }
        }
        private void discroof()
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
        private void discswitch()
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
        private void discweather()
        {
            try
            {
                if (weatherconnected)
                {
                    weather.Connected = false;   //disconnect weather sensors
                    weatherconnected = false;
                    humidtext.Text = "not connected";
                    pressuretext.Text = "not connected";
                    temptext.Text = "not connected";
                    sqmtext.Text = "not connected";
                    cloudtext.Text = "not connected";
                    btnConnWeather.ForeColor = Color.White;
                    btnDiscWeather.ForeColor = Color.Gray;
                    humidtext.BackColor = Color.Silver;
                    pressuretext.BackColor = Color.Silver;
                    temptext.BackColor = Color.Silver;
                    sqmtext.BackColor = Color.Silver;
                    cloudtext.BackColor = Color.Silver;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect weather error");
            }
        }
        private void discsafe()
        {
            try
            {
                if (safetyconnected)
                {
                    safe.Connected = false;   //diconnect roof
                    safetyconnected = false;
                    imagingtext.Text = "not connected";
                    btnConnSafety.ForeColor = Color.White;
                    btnDiscSafety.ForeColor = Color.Gray;
                    imagingtext.BackColor = Color.Silver;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect safe error");
            }
        }

        //  Each of the disconnectxxx methods disconnects
        //  and changes state for status boxes.
        private void disconnecttelescope(object sender, EventArgs e)
        {
            try
            {
                this.disctelescope();
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect mount error");
            }
        }

        private void disconnectroof(object sender, EventArgs e)
        {
            try
            {
                this.discroof();
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect roof error");
            }
        }

        private void disconnectweather(object sender, EventArgs e)
        {
            try
            {
                this.discweather();
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect weather error");
            }
        }

        private void disconnectsafety(object sender, EventArgs e)
        {
            try
            {
                this.discsafe();
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect safety error");
            }
        }

        private void disconnectrelay(object sender, EventArgs e)
        {
            try
            {
                this.discswitch();
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "disconnect relay error");
            }
        }

        // tries to disconnect all, after confirmation from user
        private void disconnectall(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Do you want to continue?", "Disconnecting", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                if (queClose) this.queautoroofclose(sender, e);
                if (queOpen) this.queautoroofopen(sender, e);
                this.disconnectroof(sender, e);
                this.disconnecttelescope(sender, e);
                this.disconnectweather(sender, e);
                this.disconnectsafety(sender, e);
                mountext.BackColor = Color.Silver;
                rooftext.BackColor = Color.Silver;
                drytext.BackColor = Color.Silver;
                imagingtext.BackColor = Color.Silver;
                humidtext.BackColor = Color.Silver;
                pressuretext.BackColor = Color.Silver;
                temptext.BackColor = Color.Silver;
                sqmtext.BackColor = Color.Silver; 
                cloudtext.BackColor = Color.Silver;
                statusbox.Clear();
                result = MessageBox.Show("And relays?", "Disconnecting", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes) this.disconnectrelay(sender, e);
            }
        }

        // parks or unparks mount
        private void parktoggle(object sender, EventArgs e)
        {
            if (mountsafe == "Parked") mountunpark();
            else if (mountsafe != "Slewing") mountpark();  // added extra test to avoid trying to park while slewing
        }

        // parks telescope mount, only if roof is open
        private void mountpark()
        {
            try
            {
                if (mountConnected && roofstatus == ShutterState.shutterOpen)
                {
                    if (!mount.AtPark)
                    {
                        mountext.Text = "Parking";
                        if (mount.CanPark)
                        {
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
                else if (roofstatus == ShutterState.shutterClosed) System.Windows.Forms.MessageBox.Show("Roof is Closed!");
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "mount park error ");
            }
        }

        // unparks telescope mount, only if roof is open
        private void mountunpark()
        {
            try
            {
                if (mountConnected && roofstatus == ShutterState.shutterOpen)
                {
                    if (mount.AtPark)
                    {
                        if (mount.CanUnpark)
                        {
                            mount.Unpark();
                            mountext.Text = "Unparked";
                        }
                        else statusbox.AppendText(Environment.NewLine + "mount cannot unpark");
                    }
                    else System.Windows.Forms.MessageBox.Show("mount not parked!");
                }
                else if (roofstatus == ShutterState.shutterClosed) System.Windows.Forms.MessageBox.Show("Roof is closed!");
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "mount unpark error");
            }
        }

        // open, home mount, park mount and close roof (only if safe to do so)
        private void hibernate(object sender, EventArgs e)
        {
            int i;
            try
            {
                if (queOpen) this.queautoroofopen(sender, e);  // disable autoque for hibernate command
                if (queClose) this.queautoroofclose(sender, e);
                btnHibernate.BackColor = Color.Yellow;
                btnHibernate.ForeColor = Color.Black;
                if (mountConnected && roofConnected && mount.CanPark)
                {
                    if (roofstatus == ShutterState.shutterClosed)
                    {
                        if (mountsafe == "Parked" && !busy)
                        {
                            mountext.Text = "hibernating";
                            mountext.BackColor = Color.DarkOrange;
                            statusbox.AppendText(Environment.NewLine + "opening");
                            this.roofopen(sender, e);
                            this.refreshshutterstate();
                        }

                        for (i = 0; i < rooftimeout && roofstatus != ShutterState.shutterOpen; ++i)
                        {
                            System.Threading.Thread.Sleep(2000);
                            Console.Beep(1000, 200);
                            this.refreshshutterstate();
                        }
                    }
                    if (roofstatus == ShutterState.shutterOpen)
                    {
                        mountext.Text = "unparking";
                        if (mount.AtPark) mount.Unpark();
                        mount.Tracking = false;
                        mountext.Text = "homing";
                        if (!mount.AtHome) mount.FindHome();
                        mountext.Text = "parking";
                        if (mount.CanPark) mount.Park(); // ensure mount is parked before closing roof
                        for (i = 0; i < mountimeout && !mount.AtPark; ++i)
                        {
                            System.Threading.Thread.Sleep(2000);
                            Console.Beep(1000, 200);
                        }
                        System.Threading.Thread.Sleep(3000);
                        statusbox.AppendText(Environment.NewLine + "ready to close");
                        this.refreshmount();
                        if (mountsafe == "Parked" && !busy)  // sensor agreed position
                        {
                            busy = true;
                            dome.CloseShutter();  // if mount parked, close roof
                            busy = false;
                            statusbox.AppendText(Environment.NewLine + "closing roof");
                        }
                        else System.Windows.Forms.MessageBox.Show("Error - try manual close");
                        mountext.BackColor = Color.LightGreen;
                        btnHibernate.BackColor = Color.Transparent;
                        btnHibernate.ForeColor = Color.White;
                    }
                }
                else statusbox.AppendText(Environment.NewLine + "mount cannot park");
            }

            catch (Exception)
            {
                // disctelescope();
                statusbox.AppendText(Environment.NewLine + "hibernate error");
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
                    if (mountsafe == "Tracking")
                    {
                        if (mount.CanSetTracking) mount.Tracking = false;
                        btnTrackTog.Text = "Tracking On";
                        mountext.Text = "Tracking Off";
                    }
                    else if (mountsafe != "Tracking")
                    {
                        if (mount.CanSetTracking) mount.Tracking = true;
                        btnTrackTog.Text = "Tracking Off";
                        mountext.Text = "Tracking On";
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
                if (mountConnected && roofConnected && roofstatus == ShutterState.shutterOpen)
                {
                    if (mount.CanUnpark)
                    {
                        mountext.Text = "Unparking";
                        if (mount.AtPark) mount.Unpark();
                    }
                    else statusbox.AppendText(Environment.NewLine + "mount cannot unpark");

                    if (mount.CanSetTracking) mount.Tracking = false;
                    mountext.Text = " Tracking off";
                    if (mount.CanFindHome)
                    {
                        mountext.Text = "homing";
                        mount.FindHome();
                    }
                    else statusbox.AppendText(Environment.NewLine + "mount cannot home");
                }
                else if (roofstatus != ShutterState.shutterOpen) System.Windows.Forms.MessageBox.Show("No mount movements when roof not open!");
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
                if (mountConnected)
                {
                    if (mount.CanUnpark)
                    {
                        mountext.Text = "Unparking";
                        if (mount.AtPark) mount.Unpark();
                    }
                    if (mount.CanSetTracking) mount.Tracking = false;
                    mountext.Text = " Tracking off";
                    if (mount.CanFindHome)
                    {
                        mount.FindHome();
                        mountext.Text = "homing";
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
                if (mountConnected && mount.CanPark) mount.Park();
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
                        toggle.SetSwitch(i, switchoff);  // turn toggle off
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
                    if (roofConnected && mountsafe == "Parked" && !busy)
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
        private void roofopen(object sender, EventArgs e)
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
                    if (roofConnected && mountsafe == "Parked" && !busy)
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
        void sensorDisplay()
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
        private void norainsense(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("NORAINSENSE", false);
                    System.Threading.Thread.Sleep(2000);
                    sensorDisplay();
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
        private void rainsense(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("RAINSENSE", false);
                    System.Threading.Thread.Sleep(2000);
                    sensorDisplay();
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
        private void noparksense(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("NOPARKSENSE", false);
                    System.Threading.Thread.Sleep(2000);
                    sensorDisplay();
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "park disable error");
            }
        }

        // special command outside standard ASCOM to modify sensor usage in Arduino
        private void parksense(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("PARKSENSE", false);
                    System.Threading.Thread.Sleep(2000);
                    sensorDisplay();
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "park enable error");
            }
        }

        private void enablebeep(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("BEEPON", false);
                    System.Threading.Thread.Sleep(2000);
                    sensorDisplay();
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "park enable error");
            }
        }

        private void disablebeep(object sender, EventArgs e)
        {
            try
            {
                if (roofConnected && !busy)
                {
                    busy = true;
                    dome.CommandBlind("BEEPOFF", false);
                    System.Threading.Thread.Sleep(2000);
                    sensorDisplay();
                    busy = false;
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "park enable error");
            }
        }

        // toggles que open trigger flag if roof is connected for auto open when safe
        private void queautoroofopen(object sender, EventArgs e)
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
        private void queautoroofclose(object sender, EventArgs e)
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

        // If one or both of the autocues is active, autoroof works out whether to open or 
        // close the roof and park the mount depending on weather conditions.
        // It uses an amalgam of multiple sensors to determine the action.
        private void autoroof()
        {
            // initial states, weathersafe is conglomerate of multiple rain, humidity and safety sensors
            int i;
            try
            {
                if (mountConnected && roofConnected && !busy)  // only if mount and roof are connected
                {
                    // check for auto close conditions
                    if (roofstatus == ShutterState.shutterOpen && !goodConditions && queClose)
                    {
                        mountext.Text = "Parking";
                        if (!mount.AtPark) mount.Park(); // ensure mount is parked before closing roof
                        for (i = 0; i < mountimeout && !mount.AtPark; ++i)
                        {
                            System.Threading.Thread.Sleep(2000);
                            Console.Beep(1000, 200);
                        }
                        System.Threading.Thread.Sleep(3000);
                        this.refreshmount();
                        if (mountsafe == "Parked" && !busy)  // sensor agreed position
                        {
                            busy = true;
                            dome.CloseShutter();  // if mount parked, close roof
                            busy = false;
                            if (!noRain) TurnOffMountSwitch(); // for extreme poor conditions, shut down
                            queOpen = false;
                            btnAutoOpen.BackColor = Color.DarkOrange;
                            btnAutoOpen.ForeColor = Color.White;
                            statusbox.AppendText(Environment.NewLine + "auto close");
                        }
                        else System.Windows.Forms.MessageBox.Show("busy, try manual close");
                        this.refreshshutterstate();
                    }

                    // check for auto open conditions
                    else if (roofstatus == ShutterState.shutterClosed && goodConditions && queOpen)
                    {
                        if (mountsafe == "Parked" && !busy) // use sensor rather than mount status
                        {
                            busy = true;
                            dome.OpenShutter();
                            busy = false;
                            statusbox.AppendText(Environment.NewLine + "auto open");
                        }
                        else System.Windows.Forms.MessageBox.Show("busy, try manual open");
                        this.refreshshutterstate();
                    }
                    // experiment to trap roof closing on unparked mount
                    else if (roofstatus == ShutterState.shutterClosing && mountsafe != "Parked")
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
        private void automount()
        {
            try
            {
                if (mountConnected && roofConnected)  // only if mount and roof are connected
                {
                    //auto open
                    if (roofstatus == ShutterState.shutterOpen && mountsafe == "Parked" && goodConditions && queOpen)  // if it opened, move mount
                    {
                        mountext.Text = "unparking";
                        if (mount.AtPark) mount.Unpark();
                        mount.Tracking = false;
                        mountext.Text = "homing";
                        mount.FindHome();
                        // disable que open
                        queOpen = false;
                        btnAutoOpen.BackColor = Color.DarkOrange;
                        btnAutoOpen.ForeColor = Color.White;
                        refreshmount();  // update mount status
                    }

                    // with new roof status, if mount is doing anything strange, need to stop it dead - either with roof closed or in error state
                    if (roofstatus != ShutterState.shutterOpen)
                    {
                        if (mountsafe != "Parked")
                        {
                            if (mountsafe == "Slewing")
                            {
                                mount.AbortSlew();
                                aborted = true;
                                mountext.Text = "aborting";
                                statusbox.AppendText(Environment.NewLine + "aborting slew");
                            }
                            mount.Tracking = false; // stop tracking
                            btnTrackTog.Text = "Tracking --";
                            statusbox.AppendText(Environment.NewLine + "aborted tracking");
                            statusbox.AppendText(Environment.NewLine + "make safe, home mount");
                            if (switchconnected)
                            {
                                if (mountConnected) disctelescope();
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
                    if (roofstatus != ShutterState.shutterOpen && mountsafe != "Parked") // in case mount is disconnected
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
        // checks on mount and roof actions
        private void refreshall(object sender, EventArgs e)
        {
            try
            {
                // refreshshutterstate();
                // refreshmount();
                refreshswitch(); 
                refreshweather();
                refreshsafetymonitor();
                goodConditions = noRain && clearAir && clearSky;  // amalgamation of air, sky and weather                          
                automount(); // checks mount logic
                autoroof();  // checks roof logic
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "data refresh error");
            }
        }

        //  test of two form timers, this faster one just checks for collisions
        private void collision(object sender, EventArgs e)
        {
            try
            {
                refreshshutterstate();
                refreshmount();
                automount(); // checks mount logic
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "data refresh error");
            }
        }

        // small routine to update switches, in case there is a parallel hub command - every 4 cycles
        private void refreshswitch()
        {
            try  // all at once
            {
                if (switchconnected)
                {
                    getswitchstate();
                    showswitchstate();
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "switch refresh error");
            }
        }

        // uses safetymonitor to update status and display
        private void refreshsafetymonitor()
        {
            try
            {
                if (safetyconnected)
                {
                    if (safe.IsSafe)
                    {
                        imagingtext.BackColor = Color.LightGreen;
                        imagingtext.Text = "OK to Open";
                        clearSky = true;  // update weatherstatus
                    }
                    else
                    {
                        imagingtext.BackColor = Color.DarkOrange;
                        imagingtext.Text = "NOK to Open";
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
        private void refreshshutterstate()
        {
            try
            {
                if (roofConnected && !busy)
                {
                    //  get shutter status
                    busy = true;
                    roofstatus = dome.ShutterStatus;  // get shutter status   
                    busy = false;
                    switch (roofstatus)
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
                statusbox.AppendText(Environment.NewLine + "roofstatus error");
            }
        }

        /* refreshes status of rain sensor connected to Arduino and other sensors. Note that
         the rain is detected with a non-standard ASCOM command through its CommandBool method. If
         you are using a hub, you need to make sure that it passes all commands and not just the
         standard ones */
        private void refreshweather()
        {
            try
            {
                if (roofConnected) // assumes  LifeRoof system with rain sensor
                {
                    if (!busy)
                    {
                        busy = true;
                        noRain = !dome.CommandBool("RAIN", false);
                        busy = false;
                        if (!noRain)
                        {
                            roofsafe = "Rain";
                            drytext.BackColor = Color.DarkOrange;
                        }
                        else
                        {
                            roofsafe = "Dry";
                            drytext.BackColor = Color.LightGreen;
                        }
                        drytext.Text = roofsafe;
                    }
                }
                else noRain = true;  // default if no rain detector
                if (weatherconnected)
                {
                    pressuretext.Text = Math.Round(weather.Pressure, 2).ToString() + " hPa";
                    temptext.Text = Math.Round(weather.Temperature, 2).ToString() + " °C";
                    humidtext.Text = Math.Round(weather.Humidity, 2).ToString() + " %";
                    sqmtext.Text = Math.Round(weather.SkyQuality, 1).ToString() + " M/asec2";
                    cloudtext.Text = Math.Round(cloud(weather.SkyTemperature), 1).ToString() + " %";
                    if (weather.Humidity > maxhumidity)
                    {
                        humidtext.BackColor = Color.DarkOrange;
                        clearAir = false;
                    }
                    else
                    {
                        humidtext.BackColor = Color.LightGreen;
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
        //  cloud takes sky temperature and scales to % cloud cover - auto adjusts limits
        private double cloud(double adjtemp)
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

            // cloud cover calculation - ratio of sky temp between limits
            return(100 * (adjtemp - SkyLL) / (SkyUL - SkyLL));
        }
        // routine to change over the graph data source and y axis and update
        private void graphselect(object sender, EventArgs e)
        {
            if ((string)btngraphsel.SelectedItem == "temp C")
            {
                charttype = 0;
                chart1.ChartAreas[0].AxisY.Minimum = -10;
                chart1.ChartAreas[0].AxisY.Maximum = +30;
            }
            if ((string)btngraphsel.SelectedItem == "humidity %")
            {
                charttype = 1;
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                chart1.ChartAreas[0].AxisY.Maximum = 100;
            }
            if ((string)btngraphsel.SelectedItem == "dewpoint C")
            {
                charttype = 2;
                chart1.ChartAreas[0].AxisY.Minimum = -10;
                chart1.ChartAreas[0].AxisY.Maximum = +20;
            }
            if ((string)btngraphsel.SelectedItem == "sky quality SQM")
            {
                charttype = 3;
                chart1.ChartAreas[0].AxisY.Minimum = 10;
                chart1.ChartAreas[0].AxisY.Maximum = 20;
            }
            if ((string)btngraphsel.SelectedItem == "cloud %")
            {
                charttype = 4;
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                chart1.ChartAreas[0].AxisY.Maximum = 100;
            }
            // update chart values
            ChartDataUpdate(); // update latest values and transpose
            // either copy non zero over, or use all values
            if (samplecount < 7200) //(every minute)
            {
                var arrayNZ = chartvalues.Select(x => x).Where(x => x != 0).ToArray(); // only do non-zero values at the beginning
                chart1.Series[0].Points.DataBindY(arrayNZ);
            }
            else chart1.Series[0].Points.DataBindY(chartvalues);
        }
        // ChartData updates several arrays of weather data, covering last 3 hours and 
        // copies one set over to the chart array, for display, according to charttype's value
        // waits for a minute upon startup before displaying anything
        private void ChartDataUpdate()
        {
            int index;
            samplecount += 1;
            if(samplecount > 59)
            {
                if ((samplecount < 7200) && (samplecount % 60 == 0)) //(every minute)
                {
                    index = ((int)((samplecount / 60.0) - (samplecount % 60)));
                    tempvalues[index] = Math.Round(weather.Temperature, 2);
                    humidvalues[index] = Math.Round(weather.Humidity, 2);
                    dewvalues[index] = Math.Round(weather.DewPoint, 2);
                    SQMvalues[index] = Math.Round(weather.SkyQuality, 2);
                    cloudvalues[index] = Math.Round(cloud(weather.SkyTemperature), 1);
                }
                if ((samplecount >= 7200) && (samplecount % 60 == 0))
                {
                    for (int i = 0; i < 119; i++)  // shift left
                    {
                        tempvalues[i] = tempvalues[i + 1];
                        humidvalues[i] = humidvalues[i + 1];
                        dewvalues[i] = dewvalues[i + 1];
                        SQMvalues[i] = SQMvalues[i + 1];
                        cloudvalues[i] = cloudvalues[i + 1];
                    }
                    // rhs value is current value
                    tempvalues[119] = Math.Round(weather.Temperature, 2);
                    humidvalues[119] = Math.Round(weather.Humidity, 2);
                    dewvalues[119] = Math.Round(weather.DewPoint, 2);
                    SQMvalues[119] = Math.Round(weather.SkyQuality, 2);
                    cloudvalues[119] = Math.Round(cloud(weather.SkyTemperature), 1);
                }
            }
            
            switch (charttype) // copy applicable data into chart array
            {
                case 0:
                    for (int i = 0; i < 120; i++) chartvalues[i] = tempvalues[i];
                    break;
                case 1:
                    for (int i = 0; i < 120; i++) chartvalues[i] = humidvalues[i];
                    break;
                case 2:
                    for (int i = 0; i < 120; i++) chartvalues[i] = dewvalues[i];
                    break;
                case 3:
                    for (int i = 0; i < 120; i++) chartvalues[i] = SQMvalues[i];
                    break;
                default:
                    for (int i = 0; i < 120; i++) chartvalues[i] = cloudvalues[i];
                    break;
            }
        }
        // updates chart plot, according to the selected data source, does not display zero values.
        private void ChartUpdate()
        {
            if (samplecount > 59)
            {
                if ((samplecount < 7200) && (samplecount % 60 == 0)) //(every minute)
                {
                    var arrayNZ = chartvalues.Select(x => x).Where(x => x != 0).ToArray(); // only do non-zero values at the beginning
                    chart1.Series[0].Points.DataBindY(arrayNZ);
                }
                if ((samplecount >= 7200) && (samplecount % 60 == 0)) chart1.Series[0].Points.DataBindY(chartvalues);
            }
        }
        /* updates mount from status
        note that unlike most systems, this uses an amalgamation of the mount position and a separate
        sensor which is part of the roof system. This is due to the fact that most mounts have
        programmable park positions and if set incorrectly, the mount may not clear the roofline. The sensor uses
        a non-standard ASCOM command through its CommandBool method. If you are using a hub for the roof, you
        need to make sure it passes all commands through and not just the standard ones 
        mountsafe has these outcomes:
        "Parked"
        "Not at Park"
        "Parked" is sensor(no mount) or sensor/mount confirm
        "Tracking"
        "Homed"
        "Slewing"
        any movement (tracking, slewing, homing) from mount invalidates Park status
         */
        private void refreshmount()
        {
            bool parkconfirm;  // park sensor status
            try
            {
                string trackingtext = "Tracking On"; // assume it is not tracking by default
                string parktext = "Park"; // assume not at park
                if (roofConnected && !busy) // if roof connected, confirm park position using sensors rather than mount
                {
                    busy = true;
                    parkconfirm = dome.CommandBool("PARK", false);  // from sensor
                    busy = false;
                    if (parkconfirm)
                    {
                        mountsafe = "Parked";
                        mountext.BackColor = Color.LightGreen;
                        parktext = "Unpark";
                    }
                    else  // sensor does not say it is at park
                    {
                        mountsafe = "Not at Park";
                        mountext.BackColor = Color.DarkOrange;
                        parktext = "Park";
                    }
                }
                else  // roof not connected or busy
                {
                    mountsafe = "Unknown";
                    mountext.BackColor = Color.DarkOrange;
                }
                if (mountConnected && mountsafe != "Parked")  // update mount status if not in park position-
                {
                    if (mount.Tracking)
                    {
                        mountsafe = "Tracking";
                        trackingtext = "Tracking Off";
                        mountext.BackColor = Color.DarkOrange;
                        aborted = false;
                        parktext = "Park";
                    }
                    else if (mount.AtHome)
                    {
                        mountsafe = "Homed";
                        mountext.BackColor = Color.DarkOrange;
                        aborted = false;
                        trackingtext = "Tracking --";
                        parktext = "Park";
                    }
                    if (!aborted)
                    {
                        if (mount.Slewing)  // mount can report slewing and tracking at same time, so worse case is slewing for abort actionss
                        {
                            mountsafe = "Slewing";
                            mountext.BackColor = Color.DarkOrange;
                            trackingtext = "Tracking --";
                            parktext = "Park --";
                        }
                    }
                }
                btnTrackTog.Text = trackingtext;  // doing this here with a string variable stops flicker on display
                btnPark.Text = parktext;
                mountext.Text = mountsafe;  // update display
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "mount refresh error");
            }
        }

        // power switch functions  ( x4)
        // toggleswitch(x) is customized to select different outputs depending on the connected hardware
        // or if multiple devices are being combined using darkskygeek's hub
        private void toggleswitch0(object sender, EventArgs e)
        {
            try
            {
                if (switchconnected)
                {
                    if (power[0]) toggle.SetSwitch(0, switchoff);  // turn toggle off                   
                    else toggle.SetSwitch(0, switchon);
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "relay 0 error");
            }
        }
        private void toggleswitch1(object sender, EventArgs e)
        {
            try
            {
                if (switchconnected)
                {

                    if (switchID == "ASCOM.DarkSkyGeek.SwitchHub")
                    {
                        if (power[1]) toggle.SetSwitch(12, switchoff);  // turn toggle off                   
                        else toggle.SetSwitch(12, switchon);
                    }
                    if (switchID == "ASCOM.PegasusAstroPPBAdvance.Switch")
                    {
                        if (power[1]) toggle.SetSwitch(1, switchoff);  // turn toggle off                   
                        else toggle.SetSwitch(1, switchon);
                    }
                    if (switchID == "ASCOM.PegasusAstroUPBv2.Switch")
                    {
                        if (power[1]) toggle.SetSwitch(4, switchoff);  // turn toggle off                   
                        else toggle.SetSwitch(4, switchon);
                    }
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "switch 1 error");
            }
        }
        
        private void toggleswitch2(object sender, EventArgs e)
        {
            try
            {
                if (switchconnected)
                {

                    if (switchID == "ASCOM.DarkSkyGeek.SwitchHub")
                    {
                        if (power[2]) toggle.SetSwitch(13, switchoff);  // turn switch off                   
                        else toggle.SetSwitch(13, switchon);
                    }
                    if (switchID == "ASCOM.PegasusAstroPPBAdvance.Switch")
                    {
                        if (power[2]) toggle.SetSwitch(2, switchoff);  // turn switch off                   
                        else toggle.SetSwitch(2, switchon);
                    }
                    if (switchID == "ASCOM.PegasusAstroUPBv2.Switch")
                    {
                        if (power[2]) toggle.SetSwitch(5, switchoff);  // turn switch off                   
                        else toggle.SetSwitch(5, switchon);
                    }
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "switch 2 error");
            }
        }
        private void toggleswitch3(object sender, EventArgs e)
        {
            try
            {
                if (switchconnected)
                {

                    if (switchID == "ASCOM.DarkSkyGeek.SwitchHub")
                    {
                        if (power[3]) toggle.SetSwitch(14, switchoff);  // turn toggle off                   
                        else toggle.SetSwitch(14, switchon);
                    }
                    if (switchID == "ASCOM.PegasusAstroPPBAdvance.Switch")
                    {
                        if (power[3]) toggle.SetSwitch(3, switchoff);  // turn toggle off                   
                        else toggle.SetSwitch(3, switchon);
                    }
                    if (switchID == "ASCOM.PegasusAstroUPBv2.Switch")
                    {
                        if (power[3]) toggle.SetSwitch(6, switchoff);  // turn toggle off                   
                        else toggle.SetSwitch(6, switchon);
                    }
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "switch 3 error");
            }
        }

        // reads switch position - normal logic
        private void getswitchstate()
        {
            try
            {
                if (switchconnected)
                {
                    power[0] = toggle.GetSwitch(0);
                    if (switchID == "ASCOM.DarkSkyGeek.SwitchHub")
                    {
                        power[1] = toggle.GetSwitch(12);
                        power[2] = toggle.GetSwitch(13);
                        power[3] = toggle.GetSwitch(14); 
                    }
                    if (switchID == "ASCOM.PegasusAstroPPBAdvance.Switch")
                    {
                        for (short i = 1; i < 4; i++) power[i] = toggle.GetSwitch(i); //all at once
                    }
                    if (switchID == "ASCOM.PegasusAstroUPBv2.Switch")
                    {
                        power[1] = toggle.GetSwitch(4);
                        power[2] = toggle.GetSwitch(5);
                        power[3] = toggle.GetSwitch(6);
                    }                    
                }
            }
            catch (Exception)
            {
                statusbox.AppendText(Environment.NewLine + "Relay status error");
            }
        }

        // display toggle state
        private void showswitchstate()
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
        private void sethumidity(object sender, EventArgs e)
        {
            try
            {
                maxhumidity = Convert.ToDouble(humidlimit.Value);
                this.filewrite();
            }
            catch
            {
                statusbox.AppendText(Environment.NewLine + "humidity set error");
            }
        }

        // experimental overall shuttdown method - disconnect devices and then power through switches
        private void shutdown(object sender, EventArgs e)
        {
            try
            {
                if (mountConnected) disctelescope();  // disconnect mount driver
                if (roofConnected) discroof();  // disconnect roof driver
                if (safetyconnected) discsafe();  // disconnect safety monitor
                if (weatherconnected) discweather();  // disconnect weather monitor
                if (switchconnected)
                {
                    for (short i = 0; i < 4; i++)
                    {
                        toggle.SetSwitch(i, switchoff); // turn off all switches
                        power[i] = false;                                 
                    }
                    discswitch();  // now disconnect ASCOM driver
                }
                Application.Exit();
            }
            catch
            {
                statusbox.AppendText(Environment.NewLine + "aborting error");
            }
        }

        // filewrite() saves device choices and variables to MyDocuments/ASCOM/Obsy/obsy.txt
        private void filewrite()
        {
            try
            {
                string[] configure = new string[7];
                configure[0] = domeId;
                configure[1] = mountId;
                configure[2] = weatherId;
                configure[3] = safetyId;
                configure[4] = maxhumidity.ToString();
                configure[5] = switchID;
                if (!System.IO.Directory.Exists(path)) System.IO.Directory.CreateDirectory(path);
                File.WriteAllLines(path + "\\obsy.txt", configure);
            }
            catch (System.UnauthorizedAccessException e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }

        // fileread() reads device choices and variables from MyDocuments/ASCOM/Obsy/obsy.txt
        private void fileread()
        {
            try
            {
                string[] configure = new string[7];
                configure = File.ReadAllLines(path + "\\obsy.txt");
                domeId = configure[0];
                mountId = configure[1];
                weatherId = configure[2];
                safetyId = configure[3];
                maxhumidity = Convert.ToDouble(configure[4]);
                humidlimit.Value = Convert.ToDecimal(maxhumidity);
                switchID = configure[5];
            }
            catch (System.UnauthorizedAccessException e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }
    }
}
