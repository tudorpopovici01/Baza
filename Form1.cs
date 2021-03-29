using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Baza
{
    public partial class Form1 : Form
    {
        private SqlConnection sqlConnction = null;
        private SqlCommandBuilder sqlBuilder = null;
        private SqlDataAdapter sqlDataAdapter = null;
        private DataSet dataSet = null;
        private bool newRowAdding = false;


        public Form1()
        {
            InitializeComponent();
        }

        private void LoadData()

        {
            try
            {
                sqlDataAdapter = new SqlDataAdapter("SELECT *, 'Delete'AS [Command] FROM Users ", sqlConnction);

                sqlBuilder = new SqlCommandBuilder(sqlDataAdapter);

                sqlBuilder.GetInsertCommand();
                sqlBuilder.GetUpdateCommand();
                sqlBuilder.GetDeleteCommand();

                dataSet = new DataSet();

                sqlDataAdapter.Fill(dataSet,"Users");

                dataGridView1.DataSource = dataSet. Tables["Users"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[6, i] = linkCell;
                }
            }


          
           catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

           
        }

        private void ReloadData()

        {
            try
            {
                dataSet.Tables["Users"].Clear();
                sqlDataAdapter.Fill(dataSet, "Users");

                dataGridView1.DataSource = dataSet.Tables["Users"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[6, i] = linkCell;
                }
            }



            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }




        private void Form1_Load(object sender, EventArgs e)
        {
            sqlConnction = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\HP\source\repos\Baza\Database1.mdf;Integrated Security=True");
            sqlConnction.Open();

            LoadData();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            ReloadData();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 6)
                {
                    string task = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();

                    if (task == "Delete")
                    {
                       if ( MessageBox.Show("Ștergeți acest rînd?", "Ștergere", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                            == DialogResult.Yes)
                       {
                            int rowIndex = e.RowIndex;

                            dataGridView1.Rows.RemoveAt(rowIndex);
                            dataSet.Tables["Users"].Rows[rowIndex].Delete();
                            sqlDataAdapter.Update(dataSet, "Users");

                       }    
                    }

                    else if (task == "Insert")
                    {
                        int rowIndex = dataGridView1.Rows.Count - 2;
                        DataRow row = dataSet.Tables["Users"].NewRow();

                        row["IDNP"] = dataGridView1.Rows[rowIndex].Cells["IDNP"].Value;
                        row["Nume"] = dataGridView1.Rows[rowIndex].Cells["Nume"].Value;
                        row["Prenume"] = dataGridView1.Rows[rowIndex].Cells["Prenume"].Value;
                        row["Ziua.Data.Anul Nașterii"] = dataGridView1.Rows[rowIndex].Cells["Ziua.Data.Anul Nașterii"].Value;
                        row["Statut ocupațional"] = dataGridView1.Rows[rowIndex].Cells["Statut ocupațional"].Value;

                        dataSet.Tables["Users"].Rows.Add(row);
                        dataSet.Tables["Users"].Rows.RemoveAt(dataSet.Tables["Users"].Rows.Count - 1);
                        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 2);
                        dataGridView1.Rows[e.RowIndex].Cells[6].Value = "Delete";
                        sqlDataAdapter.Update(dataSet, "Users");
                        newRowAdding = false;
                         
                    }
                    else if (task == "Update")
                    {
                        int r = e.RowIndex;

                        dataSet.Tables["Users"].Rows[r]["IDNP"] = dataGridView1.Rows[r].Cells["IDNP"].Value;
                        dataSet.Tables["Users"].Rows[r]["Nume"] = dataGridView1.Rows[r].Cells["Nume"].Value;
                        dataSet.Tables["Users"].Rows[r]["Prenume"] = dataGridView1.Rows[r].Cells["Prenume"].Value;
                        dataSet.Tables["Users"].Rows[r]["Ziua.Data.Anul Nașterii"] = dataGridView1.Rows[r].Cells["Ziua.Data.Anul Nașterii"].Value;
                        dataSet.Tables["Users"].Rows[r]["Statut ocupațional"] = dataGridView1.Rows[r].Cells["Statut ocupațional"].Value;

                        sqlDataAdapter.Update(dataSet, "Users");
                        dataGridView1.Rows[e.RowIndex].Cells[6].Value = "Delete";



                    }

                    ReloadData();
                }
            }



            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if(newRowAdding == false)
                {
                    newRowAdding = true;

                    int lastRow = dataGridView1.Rows.Count - 2;
                    DataGridViewRow row = dataGridView1.Rows[lastRow];
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[6, lastRow] = linkCell;
                    row.Cells["Command"].Value = "Insert";


                }
            }



            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
               if (newRowAdding == false)
               {
                    int rowIndex = dataGridView1.SelectedCells[0].RowIndex;
                    DataGridViewRow editingRow = dataGridView1.Rows[rowIndex];
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[6, rowIndex] = linkCell;
                    editingRow.Cells["Command"].Value = "Update";
                }
            }



            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);

            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                TextBox textBox = e.Control as TextBox;

                if (textBox != null)
                {
                    textBox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
                }
            }
        }

        private void Column_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void salveazăCaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            int i, j;
            for (i=0; i<= dataGridView1.RowCount -2;i++)
            {
                for (j=0; j <= dataGridView1.ColumnCount -1;j++)
                {
                    wsh.Cells[i + 1, j + 1] = dataGridView1[j, i].Value.ToString();
                }
            }
            exApp.Visible = true;
        }
    }
}
