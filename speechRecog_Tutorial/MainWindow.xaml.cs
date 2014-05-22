/////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////
/*
 * 
 * 
 * 
 * 
 * Change it, play with it, do with it what you will and enjoy, 
 * please just keep this intact. 
 * 
 * Kinect media controller designed & developed by M Palmer 2013
 * Copyright 2013, Michael Palmer
 * michaelpalmer.mp@gmail.com
 * www.michaelpalmerwebdesign.com
 * Downloaded from www.michaelpalmerwebdesign.com/downloads.php
 * 
 * 
 * 
 * 
 * 
 */
/////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////


//http://www.codeproject.com/Articles/6042/Interoperating-with-Windows-Media-Player-using-P-I
//minimze to tray: http://blogs.msdn.com/b/delay/archive/2009/08/31/get-out-of-the-way-with-the-tray-minimize-to-tray-sample-implementation-for-wpf.aspx

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Kinect;

using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Recognition;
using System.IO;
//lets us open programs
using System.Diagnostics;
//lets us run commands on other processes
using System.Runtime.InteropServices;
//for the timer
using System.Timers;
using System.Windows.Threading;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using Microsoft.VisualBasic.Devices;
using Interceptor;
namespace speechRecog_Tutorial
{

    public partial class MainWindow : Window
    {      
        String[] speechCommands = new string [17];

        public bool isListening;

        String myKinectName, openMyMediaPlayer, closeMyMediaPlayer, play, pause, stop, rewind, fastforward, next, previous, volUp, volDwn, mute, fullscreen, exitFullscreen, browse, hide;

        public System.Int32 iHandle;
        Process mediaPlayerOpen = new Process();

        public KinectSensor CurrentSensor;
        private SpeechRecognitionEngine speechRecognizer;
        private static RecognizerInfo GetKinectRecognizer()
        {
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);
                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) && "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };
            return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault();
        }

        public MainWindow()
        {
            InitializeComponent();
            Window window = Window.GetWindow(this);

            Loaded += (o, e) =>
            {
                var workingArea = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea;
                var transform = PresentationSource.FromVisual(this).CompositionTarget.TransformFromDevice;
                var corner = transform.Transform(new Point(workingArea.Right, workingArea.Top));

                this.Left = corner.X - this.ActualWidth;
                this.Top = corner.Y;
            };

            // Minimize
            window.Height = 183;
            MinimizeToTray.Enable(this);
            kinectIdle();
        }

        //timer for SR initilaiztion
        int countDown1 = 4;
        public DispatcherTimer sr_ReadyTimer;
        public void sr_ReadyCountDown()
        {
            this.sr_ReadyTimer = new DispatcherTimer();
            this.sr_ReadyTimer.Tick += new EventHandler(this.srInitializing);

            this.sr_ReadyTimer.Interval = new TimeSpan(0, 0, 0, 1, 0); // 100 Milliseconds 
            this.sr_ReadyTimer.Start();
        }
        public void srInitializing(object o, EventArgs sender)
        {
            BlurEffect myBlurEffect = new BlurEffect();
            myBlurEffect.Radius = 40;
            speechRectangle.Effect = myBlurEffect;

            --countDown1;
            if(countDown1 == 3){
                myKinectReply.Text = "preparing media controller.";
            }
            if (countDown1 == 2)
            {
                myKinectReply.Text = "preparing media controller..";
            }
            if (countDown1 == 1)
            {
                myKinectReply.Text = "preparing media controller...";
            }
            if (countDown1 == 0)
            {
                myKinectReply.Text = "Ready";
            }
            if (countDown1 == -1)
            {
                myKinectReply.Text = "";

                this.sr_ReadyTimer.Stop();
                this.sr_ReadyTimer = null;

                Window window = Window.GetWindow(this);
                //Uncomment this if you want it to minimise automatically after 3 seconds
                //window.WindowState = WindowState.Minimized;

                countDown1 = 4;
            }

        }

        //Called kinect - initiate speech timer: time out in 8secs
        int countDown2 = 20;
        public DispatcherTimer kinectListningTimer;
        public void kinectListening_ReadyCountDown()
        {
            if (this.kinectListningTimer != null)
            {
                this.kinectListningTimer.Stop();
                this.kinectListningTimer = null;
                countDown2 = 20;
            }

            this.kinectListningTimer = new DispatcherTimer();
            this.kinectListningTimer.Tick += new EventHandler(this.kinectListeningTimerInitialized);

            this.kinectListningTimer.Interval = new TimeSpan(0, 0, 0, 1, 0); // 100 Milliseconds 
            this.kinectListningTimer.Start();
        }
        public void kinectListeningTimerInitialized(object o, EventArgs sender)
        {          
            --countDown2;

            if (countDown2 == 16)
            {
                myKinectReply.Text = "";
            }

            if (countDown2 == 15)
            {
                myKinectReply.Text = "I am listening.";
            }

            if (countDown2 == 14)
            {
                myKinectReply.Text = "I am listening..";
            }
            if (countDown2 == 13)
            {
                myKinectReply.Text = "I am listening...";
            }

            if (countDown2 == 11)
            {
                myKinectReply.Text = "OK.";
            }
            if (countDown2 == 10)
            {
                myKinectReply.Text = "OK..";
            }
            if (countDown2 == 9)
            {
                myKinectReply.Text = "OK...";
            }

            if (countDown2 == 8)
            {
                myKinectReply.Text = "Entering standby.";
            }
            if (countDown2 == 7)
            {
                myKinectReply.Text = "Entering standby..";
            }
            if (countDown2 == 6)
            {
                myKinectReply.Text = "Entering standby...";
            }

            if (countDown2 <= 5)
            {
                myKinectReply.Text = "";
                kinectNotListening();
            }

            if (countDown2 <= 0)
            {
                this.kinectListningTimer.Stop();
                this.kinectListningTimer = null;

                // Find the window that contains the control
                Window window = Window.GetWindow(this);

                // Minimize
                //window.WindowState = WindowState.Minimized;

                countDown2 = 20;
            }
        }

        private KinectSensor InitializeKinect()
        {
            //get the first available sensor and set it to the current sensor variable
            CurrentSensor = KinectSensor.KinectSensors
                                  .FirstOrDefault(s => s.Status == KinectStatus.Connected);
            speechRecognizer = CreateSpeechRecognizer();
            //Start ther sensor
            CurrentSensor.Start();
            //then run the start method to start streaming audio
            Start();
            return CurrentSensor;
        }

        private void Start()
        {
            //set sensor audio source to variable
            var audioSource = CurrentSensor.AudioSource;
            //Set the beam angle mode - the direction the audio beam is pointing
            //we want it to be set to adaptive
            audioSource.BeamAngleMode = BeamAngleMode.Manual;
            audioSource.ManualBeamAngle = Math.PI / 180.0 * 10.0; //angle in radians

            
            //start the audiosource 
            var kinectStream = audioSource.Start();
            //configure incoming audio stream
            speechRecognizer.SetInputToAudioStream(
                kinectStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            //make sure the recognizer does not stop after completing 	
            speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);
            //reduce background and ambient noise for better accuracy
            CurrentSensor.AudioSource.EchoCancellationMode = EchoCancellationMode.None;
            CurrentSensor.AudioSource.AutomaticGainControlEnabled = false;
        }

        public void startMediaController()
        {
            myKinectName = kinectName.Text;
            openMyMediaPlayer = openMediaPlayer.Text;
            closeMyMediaPlayer = closeMediaPlayer.Text;
            play = wrdPlay.Text;
            pause = wrdPause.Text;
            stop = wrdStop.Text;
            rewind = wrdRwd.Text;
            fastforward = wrdFF.Text;
            next = wrdNext.Text;
            previous = wrdPrevious.Text;
            volUp = wrdVolUp.Text;
            volDwn = wrdVolDwn.Text;
            mute = wrdMute.Text;
            fullscreen = wrdFullScreen.Text;
            exitFullscreen = wrdExitFullScreen.Text;
            browse = wrdBrowse.Text;
            hide = wrdHide.Text;
            InitializeKinect();
            initializeSpeechCommandsArray();
            sr_ReadyCountDown();
        }

        public void initializeSpeechCommandsArray()
        {
            try
            {
                speechCommands[0] = myKinectName;
                speechCommands[1] = openMyMediaPlayer;
                speechCommands[2] = closeMyMediaPlayer;
                speechCommands[3] = play;
                speechCommands[4] = pause;
                speechCommands[5] = stop;
                speechCommands[6] = rewind;
                speechCommands[7] = fastforward;
                speechCommands[8] = next;
                speechCommands[9] = previous;
                speechCommands[10] = volUp;
                speechCommands[11] = volDwn;
                speechCommands[12] = mute;
                speechCommands[13] = fullscreen;
                speechCommands[14] = exitFullscreen;
                speechCommands[15] = browse;
                speechCommands[16] = hide;
            }
            catch (NullReferenceException)
            {
                //do nothing
            }
            writeToFile();
        }

        private void writeToFile()
        {
            System.IO.File.WriteAllLines("kinectMediaController_speechCommands.txt", speechCommands);
        }

        public void loadDefaultCommands()
        {
            speechCommands = System.IO.File.ReadAllLines("kinectMediaController_speechCommands.txt");

            kinectName.Text = speechCommands[0];
            openMediaPlayer.Text = speechCommands[1];
            closeMediaPlayer.Text = speechCommands[2];
            wrdPlay.Text = speechCommands[3];
            wrdPause.Text = speechCommands[4];
            wrdStop.Text = speechCommands[5];
            wrdRwd.Text = speechCommands[6];
            wrdFF.Text = speechCommands[7];
            wrdNext.Text = speechCommands[8];
            wrdPrevious.Text = speechCommands[9];
            wrdVolUp.Text = speechCommands[10];
            wrdVolDwn.Text = speechCommands[11];
            wrdMute.Text = speechCommands[12];
            wrdFullScreen.Text = speechCommands[13];
            wrdExitFullScreen.Text = speechCommands[14];
            wrdBrowse.Text = speechCommands[15];
            wrdHide.Text = speechCommands[16];
        }

        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            myKinectName = kinectName.Text;
            openMyMediaPlayer = openMediaPlayer.Text;
            closeMyMediaPlayer = closeMediaPlayer.Text;
            play = wrdPlay.Text;
            pause = wrdPause.Text;
            stop = wrdStop.Text;
            rewind = wrdRwd.Text;
            fastforward = wrdFF.Text;
            next = wrdNext.Text;
            previous = wrdPrevious.Text;
            volUp = wrdVolUp.Text;
            volDwn = wrdVolDwn.Text;
            mute = wrdMute.Text;
            fullscreen = wrdFullScreen.Text;
            exitFullscreen = wrdExitFullScreen.Text;
            browse = wrdBrowse.Text;
            hide = wrdHide.Text;
            initializeSpeechCommandsArray();
            loadDefaultCommands();
        }

        private void defaults_Click(object sender, RoutedEventArgs e)
        {
            loadDefaultCommands();
        }

        private SpeechRecognitionEngine CreateSpeechRecognizer()
        {
            //set recognizer info
            RecognizerInfo ri = GetKinectRecognizer();
            //create instance of SRE
            SpeechRecognitionEngine sre;
            sre = new SpeechRecognitionEngine(ri.Id);

            //Add the words we want our program to recognise
            var grammar = new Choices();
            grammar.Add(myKinectName);
            grammar.Add(openMyMediaPlayer);
            grammar.Add(closeMyMediaPlayer);
            grammar.Add(play);
            grammar.Add(pause);
            grammar.Add(stop);
            grammar.Add(rewind);
            grammar.Add(fastforward);
            grammar.Add(next);
            grammar.Add(previous);
            grammar.Add(volUp);
            grammar.Add(volDwn);
            grammar.Add(mute);
            grammar.Add(fullscreen);
            grammar.Add(exitFullscreen);
            grammar.Add(browse);
            grammar.Add(hide);

            //set culture - language, country/region
            var gb = new GrammarBuilder { Culture = ri.Culture };
            gb.Append(grammar);

            //set up the grammar builder
            var g = new Grammar(gb);
            sre.LoadGrammar(g);

            //Set events for recognizing, hypothesising and rejecting speech
            sre.SpeechRecognized += SreSpeechRecognized;
            sre.SpeechHypothesized += SreSpeechHypothesized;
            sre.SpeechRecognitionRejected += SreSpeechRecognitionRejected;
            return sre;
        }

        private void SreSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Confidence < .5)
            {
                RejectSpeech(e.Result);
            }
            else
            {
                if (e.Result.Text == myKinectName) 
                { 
                    myKinectReply.Text = "How can I help?";
                    kinectListening();
                    isListening = true;
                    Window window = Window.GetWindow(this);
                    window.WindowState = WindowState.Normal;
                    kinectListening_ReadyCountDown();
                }
                if (isListening)
                {
                    if (e.Result.Text == openMyMediaPlayer)
                    {
                        this.mediaPlayerOpen.StartInfo.FileName = "notepad.exe";
                        this.mediaPlayerOpen.Start();
                    }
                    if (e.Result.Text == closeMyMediaPlayer)
                    {
                        try
                        {
                            this.mediaPlayerOpen.StartInfo.FileName = "notepad.exe";
                            this.mediaPlayerOpen.CloseMainWindow();
                            this.mediaPlayerOpen.WaitForExit();
                        }
                        catch (InvalidOperationException)
                        {
                            //MessageBox.Show("Something went wrong.");
                        }
                    }
                    if (e.Result.Text == play) { myKinectReply.Text = "Okey dokey, playing."; commandOne(); }
                    if (e.Result.Text == pause) { myKinectReply.Text = "Sure thing boss, paused."; commandTwo(); }
                    if (e.Result.Text == stop) { myKinectReply.Text = "Alrighty, stopped."; commandThree(); }
                    if (e.Result.Text == rewind) { myKinectReply.Text = "Aye Aye, rewinding."; commandFour(); }
                    if (e.Result.Text == fastforward) { myKinectReply.Text = "You got it, fastforwarding."; commandFive(); }
                    if (e.Result.Text == next) { myKinectReply.Text = "True dat, next."; commandSix(); }
                    if (e.Result.Text == previous) { myKinectReply.Text = "No problemo, previous."; commandSeven(); }
                    if (e.Result.Text == volUp) { myKinectReply.Text = "Pump up the volume!"; commandEight(); }
                    if (e.Result.Text == volDwn) { myKinectReply.Text = "Volume down."; commandNine(); }
                    if (e.Result.Text == mute) { myKinectReply.Text = "Shhhhhhh..., muted."; commandTen(); }
                    if (e.Result.Text == fullscreen) { myKinectReply.Text = "Tada! full screen."; commandEleven(); }
                    if (e.Result.Text == browse) { myKinectReply.Text = "OK! Browsing."; commandTwelve(); }
                    if (e.Result.Text == exitFullscreen) { myKinectReply.Text = "Your wish is my command, exit full screen."; commandThirteen(); }
                    if (e.Result.Text == hide)
                    {
                        // Find the window that contains the control
                        Window window = Window.GetWindow(this);

                        // Minimize
                        window.WindowState = WindowState.Minimized;
                        isListening = false;
                    }
                    
                }
                else
                {
                    //do nothing for now
                }
            }
        }

        private void RejectSpeech(RecognitionResult result)
        {
            if(isListening){
                myKinectReply.Text = "Pardon Moi?";
            }
        }

        private void SreSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            RejectSpeech(e.Result);
        }

        private void SreSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            if (e.Result.Confidence > 0)
            {
               //feedback
            }

        }
        
        //play/pause
        static void commandOne() 
        {

           
            Input input = new Input();
            
            // Be sure to set your keyboard filter to be able to capture key presses and simulate key presses
            // KeyboardFilterMode.All captures all events; 'Down' only captures presses for non-special keys; 'Up' only captures releases for non-special keys; 'E0' and 'E1' capture presses/releases for special keys
            input.KeyboardFilterMode = KeyboardFilterMode.All;
            // You can set a MouseFilterMode as well, but you don't need to set a MouseFilterMode to simulate mouse clicks

            // Finally, load the driver
            input.Load();
            Microsoft.VisualBasic.Interaction.AppActivate("Untitled - Notepad");

            input.MoveMouseTo(5, 5);
            input.MoveMouseBy(25, 25);
            input.SendLeftClick();

            input.KeyDelay = 1; // See below for explanation; not necessary in non-game apps
            input.SendKeys(Keys.Enter);  // Presses the ENTER key down and then up (this constitutes a key press)

            // Or you can do the same thing above using these two lines of code
            input.SendKeys(Keys.Enter, KeyState.Down);
            System.Threading.Thread.Sleep(1); // For use in games, be sure to sleep the thread so the game can capture all events. A lagging game cannot process input quickly, and you so you may have to adjust this to as much as 40 millisecond delay. Outside of a game, a delay of even 0 milliseconds can work (instant key presses).
            input.SendKeys(Keys.Enter, KeyState.Up);

            input.SendText("hello, I am typing!");

            /* All these following characters / numbers / symbols work */
            input.SendText("abcdefghijklmnopqrstuvwxyz");
            input.SendText("1234567890");
            input.SendText("!@#$%^&*()");
            input.SendText("[]\\;',./");
            input.SendText("{}|:\"<>?");


            // And finally
            input.Unload();
        }
        //stop
        private void commandTwo()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        //rewind
        private void commandThree()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        //fastforward
        private void commandFour()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        //next
        private void commandFive()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        //previous
        private void commandSix()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        //Volume up
        private void commandSeven()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        //Volume down
        private void commandEight()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        //Mute
        private void commandNine()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        //fullscreen
        private void commandTen()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        //browse
        private void commandEleven()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        private void commandTwelve()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }
        private void commandThirteen()
        {
            Microsoft.VisualBasic.Interaction.AppActivate("Notepad");

        }

        private void btnConfigure_Click(object sender, RoutedEventArgs e)
        {
            btnConfigure.Visibility = Visibility.Hidden;
            lblConfig.Visibility = Visibility.Hidden;

            btnClose.Visibility = Visibility.Visible;
            lblClose.Visibility = Visibility.Visible;

            btnStart.Visibility = Visibility.Hidden;
            lblStart.Visibility = Visibility.Hidden;

            DoubleAnimation growImage = new DoubleAnimation();
            growImage.To = 658;
            growImage.Duration = new Duration(TimeSpan.FromSeconds(1));
            Window window = Window.GetWindow(this);
            window.BeginAnimation(Image.HeightProperty, growImage);
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            lblStart.Visibility = Visibility.Hidden;
            btnStart.Visibility = Visibility.Hidden;
            loadDefaultCommands();
            startMediaController();
            Window window = Window.GetWindow(this);
            DoubleAnimation growImage = new DoubleAnimation();
            growImage.To = 183;
            growImage.Duration = new Duration(TimeSpan.FromSeconds(1));
            window.BeginAnimation(Image.HeightProperty, growImage);
            myKinectReply.Text = "preparing media controller";
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            btnClose.Visibility = Visibility.Hidden;
            lblClose.Visibility = Visibility.Hidden;

            btnConfigure.Visibility = Visibility.Visible;
            lblConfig.Visibility = Visibility.Visible;

            btnStart.Visibility = Visibility.Hidden;
            lblStart.Visibility = Visibility.Hidden;

            if (kinectName.Text == "" || openMediaPlayer.Text == "" || closeMediaPlayer.Text == "" ||
    wrdPlay.Text == "" || wrdPause.Text == "" || wrdStop.Text == "" || wrdRwd.Text == "" || wrdFF.Text
    == "" || wrdNext.Text == "" || wrdPrevious.Text == "" || wrdVolUp.Text == "" || wrdVolDwn.Text
    == "" || wrdMute.Text == "" || wrdFullScreen.Text == "" || wrdExitFullScreen.Text == "" ||
    wrdBrowse.Text == "" || wrdHide.Text == "")
            {
                MessageBox.Show("Make sure you have filled everything out!");
            }
            else
            {
                myKinectName = kinectName.Text;
                openMyMediaPlayer = openMediaPlayer.Text;
                closeMyMediaPlayer = closeMediaPlayer.Text;
                play = wrdPlay.Text;
                pause = wrdPause.Text;
                stop = wrdStop.Text;
                rewind = wrdRwd.Text;
                fastforward = wrdFF.Text;
                next = wrdNext.Text;
                previous = wrdPrevious.Text;
                volUp = wrdVolUp.Text;
                volDwn = wrdVolDwn.Text;
                mute = wrdMute.Text;
                fullscreen = wrdFullScreen.Text;
                exitFullscreen = wrdExitFullScreen.Text;
                browse = wrdBrowse.Text;
                hide = wrdHide.Text;
                InitializeKinect();
                initializeSpeechCommandsArray();
                sr_ReadyCountDown();
            }

            Window window = Window.GetWindow(this);
            DoubleAnimation growImage = new DoubleAnimation();
            growImage.To = 183;
            growImage.Duration = new Duration(TimeSpan.FromSeconds(1));
            window.BeginAnimation(Image.HeightProperty, growImage);
        }

        public void kinectListening()
        {
            BlurEffect myBlurEffect = new BlurEffect();
            myBlurEffect.Radius = 40;
            speechRectangle.Effect = myBlurEffect;

            SolidColorBrush mySolidColorBrush = new SolidColorBrush();
           mySolidColorBrush = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FF00F500"));

           speechRectangle.Fill = mySolidColorBrush;
           speechRectangle.Stroke = mySolidColorBrush;
        }

        public void kinectNotListening()
        {
            isListening = false;
            BlurEffect myBlurEffect = new BlurEffect();
            myBlurEffect.Radius = 40;
            speechRectangle.Effect = myBlurEffect;

            speechRectangle.Fill = new SolidColorBrush(Colors.Red);
            speechRectangle.Stroke = new SolidColorBrush(Colors.Red);
        }

        public void kinectIdle()
        {
            BlurEffect myBlurEffect = new BlurEffect();
            myBlurEffect.Radius = 40;
            speechRectangle.Effect = myBlurEffect;

            SolidColorBrush mySolidColorBrush = new SolidColorBrush();
            mySolidColorBrush = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FF40BCEC"));

            speechRectangle.Fill = mySolidColorBrush;
            speechRectangle.Stroke = mySolidColorBrush;
        }
    }
}
