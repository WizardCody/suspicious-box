using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn
{
    public partial class FormProgress : Form
    {
        public FormProgress()
        {
            InitializeComponent();
        }

        public Label LabelStatus
        {
            get
            {
                return labelStatus;
            }
        }

        public Button ButtonCancel
        {
            get
            {
                return buttonCancel;
            }
        }

        public ProgressBar ProgressBar
        {
            get
            {
                return progressBar;
            }
        }
    }
}
