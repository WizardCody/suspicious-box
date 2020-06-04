

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;

namespace SharedResources
{
	public class FormProgressManager
	{
		public FormProgressManager()
		{
            Worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
        }
        
        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FormProgress.Close();
        }

        public BackgroundWorker Worker { get; private set; } = new BackgroundWorker
        {
            WorkerSupportsCancellation = true
        };

        public FormProgress FormProgress { get; set; } = null;

        public void Run(object argument = null)
        {
            if (!Worker.IsBusy)
            {
                if (FormProgress == null || FormProgress.IsDisposed)
                {
                    FormProgress = SetupFormProgress();
                }

                if (!FormProgress.Visible)
                {
                    FormProgress.Show();
                }

                FormProgress.Activate();

                if (argument != null)
                {
                    Worker.RunWorkerAsync(argument);
                }   
                else
                {
                    Worker.RunWorkerAsync();
                }
            }
        }

        private void StopWorker()
        {
            if (Worker.IsBusy)
            {
                CurrentStage = ProcessStage.Cancel;

                Worker.CancelAsync();
            }
        }

        private ProcessStage _currentStage = ProcessStage.Prepare;

        public ProcessStage CurrentStage
        {
            get
            {
                return _currentStage;
            }
            set
            {
                _currentStage = value;

                FormProgress.Invoke((MethodInvoker)
                    delegate ()
                    {
                        switch (value)
                        {
                            case ProcessStage.Prepare:
                                FormProgress.Text = "Preparing...";
                                break;

                            case ProcessStage.Process:
                                FormProgress.Text = "Processing...";
                                break;

                            case ProcessStage.Cancel:
                                FormProgress.Text = "Cancelling...";
                                FormProgress.ButtonCancel.Enabled = false;
                                break;
                        }
                    });
            }
        }

        public void SetStatus(int current, int end)
        {
            Debug.WriteLine(string.Format("{0}/{1}", current, end));

            FormProgress.Invoke((MethodInvoker)
            delegate ()
            {
                FormProgress.LabelStatus.Text = string.Format("{0}/{1}", current, end);
                FormProgress.ProgressBar.Value = Convert.ToInt32(current * (100.0 / end));
            });
        }

        public enum ProcessStage
        {
            Prepare,
            Process,
            Cancel
        };

        private FormProgress SetupFormProgress()
        {
            var result = new FormProgress();

            result.ButtonCancel.Click += ButtonCancel_Click;
            result.FormClosing += FormProgress_FormClosing;

            return result;
        }

        private void FormProgress_FormClosing(object sender, FormClosingEventArgs e)
        {
            StopWorker();
        }

        private void ButtonCancel_Click(object sender, EventArgs e)
        {
            StopWorker();
        }
    }

}

