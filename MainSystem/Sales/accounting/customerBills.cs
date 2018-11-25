﻿using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MainSystem
{
    public partial class customerBills : Form
    {
        MySqlConnection dbconnection;
        MySqlConnection dbconnection1;
        private string Customer_Type;
        private bool loaded = false;
        CultureInfo c;
        public customerBills()
        {
            InitializeComponent();
            dbconnection = new MySqlConnection(connection.connectionString);
            dbconnection1 = new MySqlConnection(connection.connectionString);
            labelClient.Visible = false;
            labelEng.Visible = false;//label of مهندس/مقاول
            comClient.Visible = false;
            comEngCon.Visible = false;
            txtCustomerID.Visible = false;
            txtClientID.Visible = false;

        }
        
        private void customerBills_Load(object sender, EventArgs e)
        {
            try
            {
                this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;

                this.dataGridView1.ColumnHeadersHeight = this.dataGridView1.ColumnHeadersHeight * 2;

                this.dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                this.dataGridView1.CellPainting += new DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);

                this.dataGridView1.Paint += new PaintEventHandler(dataGridView1_Paint);

                this.dataGridView1.Scroll += new ScrollEventHandler(dataGridView1_Scroll);

                this.dataGridView1.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void dataGridView1_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            try
            {
                Rectangle rtHeader = this.dataGridView1.DisplayRectangle;

                rtHeader.Height = this.dataGridView1.ColumnHeadersHeight / 2;

                this.dataGridView1.Invalidate(rtHeader);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {

            Rectangle rtHeader = this.dataGridView1.DisplayRectangle;

            rtHeader.Height = this.dataGridView1.ColumnHeadersHeight / 2;

            this.dataGridView1.Invalidate(rtHeader);

        }

        void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            try
            {
                if (e.RowIndex == -1 && e.ColumnIndex > -1)
                {

                    Rectangle r2 = e.CellBounds;

                    r2.Y += e.CellBounds.Height / 2;

                    r2.Height = e.CellBounds.Height / 2;

                    e.PaintBackground(r2, true);
                
                    e.PaintContent(r2);

                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        void dataGridView1_Paint(object sender, PaintEventArgs e)
        {
            try
            {
                string[] monthes = { " الفواتير", "السدادات ", "الاسم", "الكود"};
                Rectangle r2= this.dataGridView1.GetCellDisplayRectangle(0, -1, true);
                for (int j = 0; j < 6;)
                {
                    Rectangle r1 = this.dataGridView1.GetCellDisplayRectangle(j, -1, true);

                    int w2 = this.dataGridView1.GetCellDisplayRectangle(j + 1, -1, true).Width;

                    r1.X += -dataGridView1.Columns[0].Width;

                    r1.Y += 1;

                    r1.Width = r1.Width + w2 - 2;

                    r1.Height = r1.Height / 2 - 2;

                    e.Graphics.FillRectangle(new SolidBrush(this.dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r1);

                    StringFormat format = new StringFormat();

                    format.Alignment = StringAlignment.Center;

                    format.LineAlignment = StringAlignment.Center;

                    e.Graphics.DrawString(monthes[j / 2],

                        this.dataGridView1.ColumnHeadersDefaultCellStyle.Font,

                        new SolidBrush(this.dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor),

                        r1,

                        format);
                    r2 = r1;
                    j += 2;

                }
                r2.X += -(100);

                r2.Y += 1;
                r2.Width = 100;

                r2.Height = 37;

                e.Graphics.FillRectangle(new SolidBrush(this.dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r2);

                StringFormat format1 = new StringFormat();

                format1.Alignment = StringAlignment.Center;

                format1.LineAlignment = StringAlignment.Center;

                e.Graphics.DrawString(monthes[3],

                    this.dataGridView1.ColumnHeadersDefaultCellStyle.Font,

                    new SolidBrush(this.dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor),

                    r2,

                    format1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        //check type of customer if engineer,client or contract 
        private void radiotype_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radio = (RadioButton)sender;
            Customer_Type = radio.Text;

            loaded = false; //this is flag to prevent action of SelectedValueChanged event until datasource fill combobox
            try
            {
                if (Customer_Type == "عميل")
                {
                    labelEng.Visible = false;
                    comEngCon.Visible = false;
                    txtCustomerID.Visible = false;
                    labelClient.Visible = true;
                    comClient.Visible = true;
                    txtClientID.Visible = true;

                    string query = "select * from customer where Customer_Type='" + Customer_Type + "'";
                    MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    comClient.DataSource = dt;
                    comClient.DisplayMember = dt.Columns["Customer_Name"].ToString();
                    comClient.ValueMember = dt.Columns["Customer_ID"].ToString();
                    comClient.Text = "";
                    txtClientID.Text = "";
                }
                else
                {
                    labelEng.Visible = true;
                    comEngCon.Visible = true;
                    txtCustomerID.Visible = true;
                    labelClient.Visible = false;
                    comClient.Visible = false;
                    txtClientID.Visible = false;

                    string query = "select * from customer where Customer_Type='" + Customer_Type + "'";
                    MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    comEngCon.DataSource = dt;
                    comEngCon.DisplayMember = dt.Columns["Customer_Name"].ToString();
                    comEngCon.ValueMember = dt.Columns["Customer_ID"].ToString();
                    comEngCon.Text = "";
                    txtCustomerID.Text = "";
                }

                loaded = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void comClient_SelectedValueChanged(object sender, EventArgs e)
        {
            if (loaded)
            {
                try
                {
                    txtClientID.Text = comClient.SelectedValue.ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        //when select customer(مهندس,مقاول)display in comCustomer the all clients of th customer 
        private void comEngCon_SelectedValueChanged(object sender, EventArgs e)
        {
            if (loaded)
            {
                try
                {
                    txtCustomerID.Text = comEngCon.SelectedValue.ToString();
                    labelClient.Visible = true;
                    comClient.Visible = true;
                    txtClientID.Visible = true;

                    string query = "select * from customer where Customer_ID in(select Client_ID from custmer_client where Customer_ID=" + comEngCon.SelectedValue + ")";
                    MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    comClient.DataSource = dt;
                    comClient.DisplayMember = dt.Columns["Customer_Name"].ToString();
                    comClient.ValueMember = dt.Columns["Customer_ID"].ToString();
                    comClient.Text = "";
                    txtClientID.Text = "";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void txtBox_KeyDown(object sender, KeyEventArgs e)
        {

            TextBox txtBox = (TextBox)sender;
            string query;
            MySqlCommand com;
            string Name;
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    if (txtBox.Text != "")
                    {
                        dbconnection.Open();
                        switch (txtBox.Name)
                        {
                            case "txtCustomerID":
                                query = "select Customer_Name from customer where Customer_ID=" + txtCustomerID.Text + "";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comEngCon.Text = Name;
                                    comEngCon.SelectedValue = txtCustomerID.Text;

                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtClientID":
                                query = "select Customer_Name from customer where Customer_ID=" + txtClientID.Text + "";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comClient.Text = Name;
                                    comClient.SelectedValue = txtClientID.Text;
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                        }

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                dbconnection.Close();
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                dbconnection1.Open();
                dataGridView1.Rows.Clear();

                displayBill();
                //  displayReturnBill();
                //  displayPaidBill();
                double totalBill = 0, TotalReturn = 0, totalPaidBill = 0, TotalPaidReturn = 0;

                foreach (DataGridViewRow row1 in dataGridView1.Rows)
                {
                    totalBill += Convert.ToDouble(row1.Cells[0].Value);
                    TotalReturn += Convert.ToDouble(row1.Cells[1].Value);
                    totalPaidBill += Convert.ToDouble(row1.Cells[2].Value);
                    TotalPaidReturn += Convert.ToDouble(row1.Cells[3].Value);

                }
           
                labBills.Text = (totalBill - TotalReturn) + "";
                labpaid.Text = (totalPaidBill - TotalPaidReturn).ToString("000,000.00");
                labRest.Text = ((totalBill - TotalReturn)- (totalPaidBill - TotalPaidReturn))+"";

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
            dbconnection1.Close();
        }
    
        //function
        // display Customer bills
        public void displayBill()
        {
            DateTime date = dateTimeFrom.Value.Date;
            string d = date.ToString("yyyy-MM-dd");
            DateTime date2 = dateTimeTo.Value.Date;
            string d2 = date2.ToString("yyyy-MM-dd");
            string query = "";
            if (txtClientID.Text != "" && txtCustomerID.Text != "")
            {
                query = "select sum(TotalCostAD),sum(Total_CostAD),customer_bill.Customer_ID,customer_bill.Client_ID ,c1.Customer_Name,c2.Customer_Name from customer_return_bill, customer_bill left join customer as c1 on c1.Customer_ID=customer_bill.Customer_ID left join customer as c2 on c2.Customer_ID=customer_bill.Client_ID where customer_bill.Client_ID='" + txtClientID.Text + "' and customer_bill.Customer_ID='" + txtCustomerID.Text + "' and Bill_Date between '" + d + "' and '" + d2 + "' group by customer_bill.Client_ID,customer_bill.Customer_ID";
            }
            else if (txtClientID.Text == "" && txtCustomerID.Text != "")
            {
                query = "select sum(TotalCostAD),sum(Total_CostAD),customer_bill.Customer_ID,customer_bill.Client_ID ,c1.Customer_Name,c2.Customer_Name from customer_return_bill, customer_bill left join customer as c1 on c1.Customer_ID=customer_bill.Customer_ID left join customer as c2 on c2.Customer_ID=customer_bill.Client_ID where  customer_bill.Customer_ID='" + txtCustomerID.Text + "' and Bill_Date between '" + d + "' and '" + d2 + "' group by customer_bill.Customer_ID";
            }
            else if (txtClientID.Text != "" && txtCustomerID.Text == "")
            {
                query = "select sum(TotalCostAD),sum(Total_CostAD),customer_bill.Customer_ID,customer_bill.Client_ID ,c1.Customer_Name,c2.Customer_Name from customer_return_bill, customer_bill left join customer as c1 on c1.Customer_ID=customer_bill.Customer_ID left join customer as c2 on c2.Customer_ID=customer_bill.Client_ID where customer_bill.Client_ID='" + txtClientID.Text + "' and customer_bill.Customer_ID is null and Bill_Date between '" + d + "' and '" + d2 + "' group by customer_bill.Client_ID";
            }
            else
            {
                query = "select sum(TotalCostAD),sum(Total_CostAD), sum(tt1.Amount) as t1,sum(tt2.Amount)as t2,customer_bill.Customer_ID,customer_bill.Client_ID,c1.Customer_Name,c2.Customer_Name from transitions as tt1,transitions as tt2,customer_return_bill, customer_bill left join customer as c1 on c1.Customer_ID=customer_bill.Customer_ID left join customer as c2 on c2.Customer_ID=customer_bill.Client_ID where Bill_Date between '" + d + "' and '" + d2 + "' and tt1.Transition='ايداع' and tt2.Transition='سحب'  group by customer_bill.Client_ID";
            }
            MySqlCommand com = new MySqlCommand(query, dbconnection1);
            MySqlDataReader dr = com.ExecuteReader();

            while (dr.Read())
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = dr["sum(Total_CostAD)"].ToString();
                dataGridView1.Rows[n].Cells[1].Value = dr["sum(TotalCostAD)"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = dr["t1"].ToString();
                dataGridView1.Rows[n].Cells[3].Value = dr["t2"].ToString();

                if (!dr.IsDBNull(6))
                    dataGridView1.Rows[n].Cells[4].Value = dr["Customer_Name"].ToString();
                else
                    dataGridView1.Rows[n].Cells[4].Value = "";

                dataGridView1.Rows[n].Cells[5].Value = dr[7].ToString();

                if (!dr.IsDBNull(4) )
                    dataGridView1.Rows[n].Cells[6].Value = dr["Customer_ID"].ToString();
                else
                    dataGridView1.Rows[n].Cells[6].Value = "";

                dataGridView1.Rows[n].Cells[7].Value = dr["Client_ID"].ToString();
            }
            dr.Close();
        }

        // display Customer return bills
        public void displayReturnBill()
        {
            string query = "";
            if (txtClientID.Text != "" && txtCustomerID.Text != "")
            {
                query = "select sum(TotalCostAD) from customer_return_bill where Client_ID='" + txtClientID.Text + "' and Customer_ID='" + txtCustomerID.Text + "'  group by Client_ID,Customer_ID";
            }
            else if (txtClientID.Text == "" && txtCustomerID.Text != "")
            {
                query = "select sum(TotalCostAD) from customer_return_bill where  Customer_ID='" + txtCustomerID.Text + "'  group by Customer_ID";

            }
            else if (txtClientID.Text != "" && txtCustomerID.Text == "")
            {
                query = "select sum(TotalCostAD) from customer_return_bill where Client_ID='" + txtClientID.Text + "' and Customer_ID is null  group by Client_ID";
            }
            else
            {
                query = "select sum(TotalCostAD),Customer_ID,Client_ID from customer_return_bill  group by Client_ID";
            }
            MySqlCommand com = new MySqlCommand(query, dbconnection1);
            MySqlDataReader dr = com.ExecuteReader();

            while (dr.Read())
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = "0.00";
                dataGridView1.Rows[n].Cells[1].Value = dr["sum(TotalCostAD)"].ToString();
                if (dr["Customer_ID"].ToString() != "")
                {
                    String q = "select Customer_Name from customer where Customer_ID=" + dr["Customer_ID"].ToString();
                    MySqlCommand com1 = new MySqlCommand(q, dbconnection);
                    string Customer_Name = com1.ExecuteScalar().ToString();
                    dataGridView1.Rows[n].Cells[2].Value = Customer_Name;
                }
                else
                {
                    dataGridView1.Rows[n].Cells[2].Value = "";
                }
                if (dr["Client_ID"].ToString() != "")
                {
                    String q = "select Customer_Name from customer where Customer_ID=" + dr["Client_ID"].ToString();
                    MySqlCommand com1 = new MySqlCommand(q, dbconnection);
                    string Customer_Name = com1.ExecuteScalar().ToString();
                    dataGridView1.Rows[n].Cells[3].Value = Customer_Name;
                }
                else
                {
                    dataGridView1.Rows[n].Cells[3].Value = "";
                }

                
            }
            dr.Close();
        }

        // display Customer Paid bills
        public void displayPaidBill()
        {
            string query = "";
            if (txtClientID.Text != "" && txtCustomerID.Text != "")
            {
                query = "select * from transitions where Client_ID=" + txtClientID.Text + " and Customer_ID='" + txtCustomerID.Text + "'";
            }
            else if (txtClientID.Text == "" && txtCustomerID.Text != "")
            {
                query = "select * from transitions where Customer_ID='" + txtCustomerID.Text + "' ";
            }
            else
            {
                query = "select * from transitions where Client_ID='" + txtClientID.Text + "'";

            }
            MySqlCommand com = new MySqlCommand(query, dbconnection1);
            MySqlDataReader dr = com.ExecuteReader();

            while (dr.Read())
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = "0.00";
                dataGridView1.Rows[n].Cells[1].Value = dr["Transition_Amount"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = dr["Transition_ID"].ToString();
                dataGridView1.Rows[n].Cells[3].Value = dr["Beneficiary_Name"].ToString();
                dataGridView1.Rows[n].Cells[4].Value = "";
                dataGridView1.Rows[n].Cells[5].Value = dr["Transition_Date"].ToString();
            }
            dr.Close();
        }

    }
}
