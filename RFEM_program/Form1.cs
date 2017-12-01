using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Dlubal.RFEM5;
using System.Runtime.InteropServices;

namespace RFEM_program
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            KarambaToRFEM.OpenConnection();
            ErrorInfo[] err = KarambaToRFEM.CalcModel();
            KarambaToRFEM.CloseConnection();
            if (err != null)
            {
                foreach (ErrorInfo er in err)
                {
                    richTextBox1.Text = richTextBox1.Text + " " + er.Description;
                }
            }
            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            IApplication app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
            app.UnlockLicense();
            app = null;   
        }

        private void button3_Click(object sender, EventArgs e)
        {
            chart1.Series["deflection"].Points.Clear();

            IApplication app = null;
            IModel rModel = null;
            ILoads rLoads = null;
            try
            {
                //Get active RFEM5 application
                app = Marshal.GetActiveObject("RFEM5.Application") as IApplication;
                app.LockLicense();
                rModel = app.GetActiveModel();
                rLoads = rModel.GetLoads();
                /*
                ILoadCase lcBase = rLoads.GetLoadCase(1, ItemAt.AtNo);
                AnalysisParameters param = lcBase.GetAnalysisParameters();
                param.ModifyLoadingByFactor = true;
                param.LoadingFactor = param.LoadingFactor + 0.1;
                LoadCase lc = lcBase.GetData();
                */
                ICalculation calc = rModel.GetCalculation();
                ErrorInfo[] err = { };
                int i = 2;
                double mult = 1.1;
                double addor = 1;


                while (Math.Abs(addor) > 0.1)
                {
                    duStuff(rLoads, mult, i, ref err, ref calc);
                    if (err.Length != 0)
                    {
                        addor = (Math.Abs(addor) - 0.1) * -1;                 
                    }
                    else { addor = Math.Abs(addor); }
                    mult += addor;
                    i++;


                }
                mult -= addor+0.05;
                i++;
                duStuff(rLoads, mult, i, ref err, ref calc);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Release COM object
            }
            finally
            {
                rLoads = null;
                rModel = null;
                app.UnlockLicense();
                app = null;
                
                
                //Cleans Garbage collector for releasing all COM interfaces and objects
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
            }

            
            
        }


        private void duStuff(ILoads rLoads, double mult, int i, ref ErrorInfo[] err, ref ICalculation calc)
        {
            ILoadCase lCase = rLoads.GetLoadCase(1, ItemAt.AtNo);
            AnalysisParameters param = new AnalysisParameters();

            param.ModifyLoadingByFactor = true;
            param.LoadingFactor = mult;
            param.Method = AnalysisMethodType.SecondOrder;

            //Loads
            //lc.PrepareModification();
            rLoads.PrepareModification();
            try
            {
                lCase.SetAnalysisParameters(ref param);
            }
            finally
            {
                rLoads.FinishModification();
                //lc.FinishModification();
                lCase = null;
            }

            err = calc.Calculate(LoadingType.LoadCaseType, 1);
            if (err.Length == 0)
            {
                IResults res = calc.GetResultsInFeNodes(LoadingType.LoadCaseType, 1);
                MaximumResults max = res.GetMaximum();
                Point3D point = max.Displacement;
                double value = Math.Sqrt(Math.Pow(point.X, 2) + Math.Pow(point.Y, 2) + Math.Pow(point.Z, 2));
                chart1.Series["deflection"].Points.AddXY(i, value);
            }

        }
    }

    
}
