using Sunny.UI;
using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace CIM
{
    public partial class SearchForm : UIForm
    {
        string lastBarcode = string.Empty;

        public SearchForm()
        {
            InitializeComponent();

            txtQR.Focus();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            var currentBarcode = txtQR.Text;

            if (string.IsNullOrWhiteSpace(currentBarcode))
            {
                MessageBox.Show("QR_Code can not empty!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtQR.Focus();
                return;
            }

            if (lastBarcode == currentBarcode)
            {
                return;
            }

            lastBarcode = currentBarcode;

            DataSet ds = SqlLite.Instance.SearchData(currentBarcode);
            DataTable dt = ds.Tables[0];

            if (dt.Rows.Count > 0)
            {
                dataGridView1.Rows.Clear();

                foreach (DataRow row in dt.Rows)
                {
                    int rowIndex = dataGridView1.Rows.Add(
                         row["TOPHOUSING"],
                         row["BOX1_GLUE_AMOUNT"],
                         row["BOX1_GLUE_DISCHARGE_VOLUME_VISION"],
                         row["INSULATOR_BAR_CODE"],
                         row["BOX1_GLUE_OVERFLOW_VISION"],
                         row["BOX1_HEATED_AIR_CURING"],
                         row["BOX2_GLUE_AMOUNT"],
                         row["BOX2_GLUE_DISCHARGE_VOLUME_VISION"],
                         row["FPCB_BAR_CODE"],
                         row["BOX2_GLUE_OVERFLOW_VISION"],
                         row["BOX2_HEATED_AIR_CURING"],
                         row["BOX3_DISTANCE"],
                         row["BOX3_GLUE_AMOUNT"],
                         row["BOX3_GLUE_DISCHARGE_VOLUME_VISION"],
                         row["BOX3_GLUE_OVERFLOW_VISION"],
                         row["BOX3_HEATED_AIR_CURING"],
                         row["BOX4_TIGHTNESS_AND_LOCATION_VISION"],
                         row["BOX4_HEIGHT_PARALLELISM"],
                         row["BOX4_RESISTANCE"],
                         row["BOX4_AIR_LEAKAGE_TEST_DETAIL"],
                         row["BOX4_AIR_LEAKAGE_TEST_RESULT"],
                         row["BOX4_TestTime"],//add QR code
                         row["Remark"]
                    );

                    if (row["Remark"].ToString() == "Doublicate")
                    {
                        dataGridView1.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.Red;
                        dataGridView1.Rows[rowIndex].Cells["Remark"].Style.ForeColor = Color.Red; // Đổi màu chữ của cột "Remark" thành màu trắng
                    }
                }
                dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Descending);
            }
            else
            {
                MessageBox.Show("QR_Code not found!", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if ((e.ColumnIndex == 2 || e.ColumnIndex == 4 || e.ColumnIndex == 7 || e.ColumnIndex == 9 || e.ColumnIndex == 13 || e.ColumnIndex == 15 || e.ColumnIndex == 16 || e.ColumnIndex == 20) && e.Value != null)
            {
                string cellValue = e.Value.ToString().Trim();

                if (cellValue.ToUpper() == "NG")
                {
                    e.CellStyle.ForeColor = System.Drawing.Color.Red;
                }
                else if (cellValue.ToUpper() == "OK")
                {
                    e.CellStyle.ForeColor = System.Drawing.Color.Green;
                }
            }
        }
    }
}
