using LiveCharts.Wpf;
using LiveCharts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace CIM
{
    public partial class PieChart : UserControl
    {
        private LiveCharts.Wpf.PieChart pieChart;
        private PieSeries okSeries;
        private PieSeries ngSeries;


        public PieChart()
        {
            InitializeComponent();
            Load += Piechart_Load;


            okSeries = new PieSeries
            {
                Title = "OK",
                Values = new ChartValues<int> { 0 },
                DataLabels = true,
                Fill = System.Windows.Media.Brushes.Green,
                Foreground = System.Windows.Media.Brushes.White,
                FontSize = 15
            };

            ngSeries = new PieSeries
            {
                Title = "NG",
                Values = new ChartValues<int> { 0 },
                DataLabels = true,
                Fill = System.Windows.Media.Brushes.OrangeRed,
                Foreground = System.Windows.Media.Brushes.White,
                FontSize = 15
            };

            series.Add(okSeries);
            series.Add(ngSeries);

            pieChart1.Series = series;
            pieChart1.DisableAnimations = true;
            pieChart1.LegendLocation = LegendLocation.Right;
            pieChart1.ForeColor = System.Drawing.Color.White;
        }

        LiveCharts.SeriesCollection series = new LiveCharts.SeriesCollection();

        public void UpdateChartData(int okCount, int ngCount)
        {
            okSeries.Values[0] = okCount;
            ngSeries.Values[0] = ngCount;
        }

        private void Piechart_Load(object sender, EventArgs e)
        {
            pieChart1.Width = this.Width;
            pieChart1.Height = this.Height;

            ElementHost host = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = pieChart
            };
            Controls.Add(host);

            pieChart1.DisableAnimations = true;
            pieChart1.LegendLocation = LegendLocation.Right;
            pieChart1.ForeColor = System.Drawing.Color.White;
        }
    }
}
