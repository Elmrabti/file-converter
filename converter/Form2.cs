using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
/*-----------------------------------------------
MY C# Project
Name:        File Converter
Author :     EL MRABTI KHALID
License:     1.0
----------------------------------------------- */
namespace converter
{
    // this form for display waiting Message Dialogue
    public partial class Form2 : Form
    {
        
        public Action Worker { get; set; }
        public Form2(Action worker)
        {
            InitializeComponent();
            if(worker == null)
                throw new ArgumentNullException();
                Worker = worker;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Task.Factory.StartNew(Worker).ContinueWith(t=> { this.Close(); }, TaskScheduler.FromCurrentSynchronizationContext());
        }
    }
    }

