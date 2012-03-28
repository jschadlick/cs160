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
using System.Windows.Threading;
using System.Threading;
using System.IO;



namespace SpeechToText
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private KinectSensor kinect;
        private DispatcherTimer readyTimer;
        private SpeechRecognitionEngine speechRecognizer;
        private String wordsForGrammar;
        private String[] wordsArray;
        private String comment;

        private bool pre_recording = false;
        private bool post_recording = true;
        private bool annotating = false;

        public MainWindow()
        {
            InitializeComponent();

            SensorChooser.KinectSensorChanged += this.SensorChooserKinectSensorChanged;
        }

        #region Kinect setup and takedown

        private void InitializeKinect()
        {
            var sensor = this.kinect;
            this.speechRecognizer = this.CreateSpeechRecognizer();
            try
            {
                sensor.Start();
            }
            catch (Exception)
            {
                SensorChooser.AppConflictOccurred();
                return;
            }

            if (this.speechRecognizer != null && sensor != null)
            {
                // NOTE: Need to wait 4 seconds for device to be ready to stream audio right after initialization
                this.readyTimer = new DispatcherTimer();
                this.readyTimer.Tick += this.ReadyTimerTick;
                this.readyTimer.Interval = new TimeSpan(0, 0, 4);
                this.readyTimer.Start();

                this.Closing += this.MainWindowClosing;
            }
        }

        private void ReadyTimerTick(object sender, EventArgs e)
        {
            this.Start();
            this.readyTimer.Stop();
            this.readyTimer = null;
        }

        private void Start()
        {
            var audioSource = this.kinect.AudioSource;
            audioSource.BeamAngleMode = BeamAngleMode.Adaptive;

            // This should be off by default, but just to be explicit, this MUST be set to false.
            audioSource.AutomaticGainControlEnabled = false;

            var kinectStream = audioSource.Start();
            this.speechRecognizer.SetInputToAudioStream(
                kinectStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            // Keep recognizing speech until window closes
            this.speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);
        }

        private void SensorChooserKinectSensorChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            KinectSensor oldSensor = e.OldValue as KinectSensor;
            if (oldSensor != null)
            {
                this.UninitializeKinect();
            }

            KinectSensor newSensor = e.NewValue as KinectSensor;
            this.kinect = newSensor;

            if (newSensor != null)
            {
                this.InitializeKinect();
            }
        }

        private void UninitializeKinect()
        {
            var sensor = this.kinect;
            if (this.speechRecognizer != null && sensor != null)
            {
                sensor.AudioSource.Stop();
                sensor.Stop();
                this.speechRecognizer.RecognizeAsyncCancel();
                this.speechRecognizer.RecognizeAsyncStop();
            }

            if (this.readyTimer != null)
            {
                this.readyTimer.Stop();
                this.readyTimer = null;
            }
        }

        private void MainWindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.UninitializeKinect();
        }
        #endregion

        #region Speech recognizer setup

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

        //takes vocabulary of words from text file, puts it into a string array
        private void LoadWords()
        {
            var path = System.IO.Path.GetFullPath("english_words.txt");
         
            wordsForGrammar = File.ReadAllText(path);

            wordsArray = wordsForGrammar.Split('\n');

            for (int i = 0; i < wordsArray.Length; i++)
            {
                wordsArray[i] = wordsArray[i].Trim();
            }

        }

        private SpeechRecognitionEngine CreateSpeechRecognizer()
        {
            #region Initialization
            RecognizerInfo ri = GetKinectRecognizer();
            if (ri == null)
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
                    Ensure you have the Microsoft Speech SDK installed.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                this.Close();
                return null;
            }

            SpeechRecognitionEngine sre;
            try
            {
                sre = new SpeechRecognitionEngine(ri.Id);
            }
            catch
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
                    Ensure you have the Microsoft Speech SDK installed and configured.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                this.Close();
                return null;
            }
            #endregion

            #region Build grammar

            //takes vocabulary of words from text file, puts it into a string array
            LoadWords();

            var wordChoices = new Choices(wordsArray);

            var preRecordingChoices = new Choices(new string[] { "start" });
            var gb_preR = new GrammarBuilder { Culture = ri.Culture };

            var postRecordingChoices = new Choices(new string[] { "keep","cancel","redo","play" });
            var gb_postR = new GrammarBuilder { Culture = ri.Culture };

            var gb_1 = new GrammarBuilder { Culture = ri.Culture };
            gb_1.Append(wordChoices);

            var gb_2 = new GrammarBuilder { Culture = ri.Culture };
            gb_2.Append(wordChoices);

            var gb_3 = new GrammarBuilder { Culture = ri.Culture };
            gb_3.Append(wordChoices);

            var gb_4 = new GrammarBuilder { Culture = ri.Culture };
            gb_4.Append(wordChoices);

            var gb = new GrammarBuilder { Culture = ri.Culture };

            gb.Append(gb_preR, 0, 1);
            gb.Append(gb_postR, 0, 1);

            //gb.Append(new SemanticResultKey("Words0", wordChoices));
            gb.Append(gb_1, 0, 1);
            //gb.Append(gb_2, 0, 1);
            //gb.Append(gb_3, 0, 1);
            //gb.Append(gb_4, 0, 1);
           
          

            // Create the actual Grammar instance, and then load it into the speech recognizer.
            var g = new Grammar(gb);

  
            sre.LoadGrammar(g);
            
            #endregion

            #region Hook up events
            sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
            sre.SpeechRecognitionRejected += new EventHandler<SpeechRecognitionRejectedEventArgs>(sre_SpeechRecognitionRejected);
            /*
            sre.SpeechHypothesized += this.SreSpeechHypothesized;
            sre.SpeechRecognitionRejected += this.SreSpeechRecognitionRejected;
            */
            #endregion

            return sre;
        }
        #endregion

        #region Speech recognition events

        void sre_SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            this.RejectSpeech(e.Result);
        }

        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
           if(annotating)
           {
            if (e.Result.Confidence < 0.5)
            {
                this.RejectSpeech(e.Result);
                return;
            }
            else
            {
                this.RecognizeSpeech(e.Result);
                return;
            }
           }
           else if (pre_recording)
           {
               switch (e.Result.Text.ToString().ToUpperInvariant())
               {
                   case "START":
                       start_label.Visibility = Visibility.Visible;
                       return;
                   default:
                       return;
               }
           }
           else if (post_recording)
           {
               
               switch (e.Result.Text.ToString().ToUpperInvariant())
               {
                   case "KEEP":
                       keep_label.Visibility = Visibility.Visible;
                       return;
                   case "CANCEL":
                       cancel_label.Visibility = Visibility.Visible;
                       return;
                   case "REDO":
                       redo_label.Visibility = Visibility.Visible;
                       return;
                   case "PLAY":
                       play_label.Visibility = Visibility.Visible;
                       return;
                   default:
                       return;
               }

           }
           else
           {
               return;
           }
           return;
        }

        private void RecognizeSpeech(String resultText)
        {
            string status = "Recognzied: " + resultText;
            this.ReportStatus(status);
            string newText = resultText.ToString();
            this.UpdateText(newText);

        }

        private void RecognizeSpeech(RecognitionResult result)
        {
            string status = "Recognzied: " + (result == null ? string.Empty : result.Text + " " + result.Confidence);
            this.ReportStatus(status);
            string newText = result.Text.ToString();
            this.UpdateText(newText);
            this.comment = this.comment + result.Text.ToString();
        }

        private void RejectSpeech(RecognitionResult result)
        {
            string status = "Rejected: " + (result == null ? string.Empty : result.Text + " " + result.Confidence);
            this.ReportStatus(status);
        }

        String getComment()
        {
            if (this.comment != null)
            {
                return this.comment;
            }
            return "";
        }
      
        #endregion

        #region UI update functions
        private void ReportStatus(string status)
        {
            Dispatcher.BeginInvoke(new Action(() => { statusLabel.Content = status; }), DispatcherPriority.Normal);
        }
        private void UpdateText(string newText)
        {
            Dispatcher.BeginInvoke(new Action(() => { mainTextBlock.Text = mainTextBlock.Text + " " + newText; }), DispatcherPriority.Normal);
        }
        #endregion
    }
}
