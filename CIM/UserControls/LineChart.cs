using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;

namespace CIM
{
    public partial class LineChart : UserControl
    {
        private LineSeries mc1Series;
        private LineSeries mc2Series;
        public LineChart()
        {

            InitializeComponent();

            mc1Series = new LineSeries
            {
                Title = "Data",
                StrokeThickness = 3,
                StrokeDashArray = new DoubleCollection(1),
                Stroke = new SolidColorBrush(Color.FromRgb(0, 0, 255)),
                Fill = Brushes.Transparent,
                LineSmoothness = 0,
                PointGeometry = DefaultGeometries.Circle,
                PointGeometrySize = 10,
                PointForeground = new SolidColorBrush(Color.FromRgb(0, 0, 255)),
                Values = new ChartValues<double>()
            };

            

            cartesianChart1.Series.Add(mc1Series);
            cartesianChart1.Series.Add(mc2Series);

            cartesianChart1.AxisX.Add(new Axis
            {
                Title = "Days",
                Separator = new Separator
                {
                    StrokeThickness = 1,
                    StrokeDashArray = new DoubleCollection(2),
                    Stroke = new SolidColorBrush(Color.FromRgb(64, 79, 86)),
                    Step = 1
                },
                DisableAnimations = true
            });

            cartesianChart1.AxisY.Add(new Axis
            {
                Title = "EA",
                LabelFormatter = value => value.ToString("N0") + "%",
                Separator = new Separator
                {
                    StrokeThickness = 1,
                    StrokeDashArray = new DoubleCollection(4),
                    Stroke = new SolidColorBrush(Color.FromRgb(64, 79, 86))
                }
            });

            cartesianChart1.LegendLocation = LegendLocation.Bottom;
        }
        public List<double> Mc1data { get; set; }
        public List<double> Mc2data { get; set; }
        public List<string> Days { get; set; }

        public void UpdateChart()
        {
            // Cập nhật giá trị của LineSeries
            mc1Series.Values.Clear();
            mc2Series.Values.Clear();

            foreach (var value in Mc1data)
            {
                mc1Series.Values.Add(value);
            }

            foreach (var value in Mc2data)
            {
                mc2Series.Values.Add(value);
            }

            // Cập nhật Labels của trục X
            cartesianChart1.AxisX[0].Labels = Days;
        }
    }

}
