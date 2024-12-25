using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UnilityCommand.SubControl;
using UnilityCommand.Plc.Mitsubishi;
using UILAlignProject.PLC;

namespace SeojinMeasurement.View.PLC
{
    //UCMonitorPLCStatus --Object
    public partial class UCMonitorPlc : UserControl
    {
        private Color color = Color.Black;

        public UCMonitorPlc()
        {
          
            InitializeComponent();
        }

        private RegisterPlc2PC _plcStatus;

        public RegisterPlc2PC PlcStatus
        {
            get { return _plcStatus; }
            set
            {
                _plcStatus = value;
                if (value == null)
                {
                    return;
                }
                PropertyInfo[] propertyInfos = value.GetType().GetProperties();
                flowControl1.Controls.Clear();
                foreach (var item in propertyInfos)
                {
                    CustomLabel customLabel = new CustomLabel();
                    //toolTip1.SetToolTip(customLabel, ThemCachNeuCoVietHoa(item.Name));
                    if (item.PropertyType.Name == "Boolean")
                    {
                        customLabel.color = color;
                        //customLabel.Name = ThemCachNeuCoVietHoa(item.Name);
                       
                        customLabel.DataBindings.Add("IsConnect", _plcStatus, item.Name);
                        if (InvokeRequired)
                        {
                            this.Invoke(new Action(() =>
                            {
                                flowControl1.Controls.Add(customLabel);
                            }));
                        }
                        else
                        {
                            flowControl1.Controls.Add(customLabel);
                        }
                    }
                    else if (item.PropertyType.Name == "Int32")
                    {
                        customLabel.color = color;
                        customLabel.IsConnect = true;
                        //customLabel.Name = ThemCachNeuCoVietHoa(item.Name);
                        customLabel.DataBindings.Add("Name", _plcStatus, item.Name, true);
                        if (InvokeRequired)
                        {
                            this.Invoke(new Action(() =>
                            {
                                flowControl1.Controls.Add(customLabel);
                            }));
                        }
                        else
                        {
                            flowControl1.Controls.Add(customLabel);
                        }
                    }



                    
                }
            }
        }

        
    }
}