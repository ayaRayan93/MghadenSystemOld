using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraTab;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MainSystem
{
    public partial class UpdateSellPrice : Form
    {
        MySqlConnection dbconnection;
        ProductsSellPriceForm productsSellPriceForm = null;
        XtraTabControl xtraTabControlSalesContent = null;
        List<DataRowView> rows = null;
        customPanel customPanel;
        bool load = false;
        string updateType = "NormalCase";
        int id = 0;
        String query = "";
        double Price = -1, ProfitRatio = -1, Sell_Discount = -1, Normal_Increase = -1, Categorical_Increase = -1, PercentageDelegate = -1;

        public UpdateSellPrice(List<DataRowView> rows, ProductsSellPriceForm productsSellPriceForm, string query, XtraTabControl xtraTabControlSalesContent)
        {
            try
            {
                InitializeComponent();
                this.xtraTabControlSalesContent = xtraTabControlSalesContent;
                this.rows = rows;
                this.productsSellPriceForm = productsSellPriceForm;
                this.query = query;
                dbconnection = new MySqlConnection(connection.connectionString);

                customPanel = new customPanel();
                panContent.Controls.Add(customPanel.Controls["panContent"]);
                panContent.Controls["panContent"].Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void UpdateSellPrice_Load(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string ids = "";
                for (int i = 0; i < rows.Count - 1; i++)
                {
                    ids += rows[i][0] + ",";
                }
                ids += rows[rows.Count - 1][0];
                query = "SELECT SellPrice.SellPrice_ID,data.Data_ID, data.Code as 'الكود',concat( product.Product_Name,' ',type.Type_Name,' ',factory.Factory_Name,' ',groupo.Group_Name,' ' ,COALESCE(color.Color_Name,''),' ' ,COALESCE(size.Size_Value,'') )as 'البند',sort.Sort_Value as 'الفرز',data.Classification as 'التصنيف',data.Description as 'الوصف',sellprice.Price as 'السعر',sellprice.Price_Type as 'نوع السعر',sellprice.Sell_Discount as 'خصم البيع',sellprice.Normal_Increase as 'الزيادة العادية',sellprice.Categorical_Increase as 'الزيادة القطعية',sellprice.ProfitRatio as 'نسبة الاضافة',sellprice.Sell_Price as 'سعر البيع',sellprice.PercentageDelegate as 'نسبة المندوب' ,sellprice.OfferFlag from data INNER JOIN sellprice on sellprice.Data_ID=data.Data_ID  INNER JOIN type ON type.Type_ID = data.Type_ID INNER JOIN product ON product.Product_ID = data.Product_ID INNER JOIN factory ON data.Factory_ID = factory.Factory_ID INNER JOIN groupo ON data.Group_ID = groupo.Group_ID LEFT outer JOIN color ON data.Color_ID = color.Color_ID LEFT outer  JOIN size ON data.Size_ID = size.Size_ID LEFT outer  JOIN sort ON data.Sort_ID = sort.Sort_ID where   SellPrice.SellPrice_ID in(" + ids + ")";

                MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                DataTable dt = new DataTable();
                da.Fill(dt);

                gridControl1.DataSource = dt;
            
                gridView1.Columns[0].Visible = false;
                gridView1.Columns[1].Visible = false;
                gridView1.Columns["OfferFlag"].Visible = false;
                gridView1.Columns[2].Width = 200;
                gridView1.Columns[3].Width = 300;
                gridView1.BestFitColumns();
                setData((DataRowView)gridView1.GetRow(0));
                dbconnection.Close();
                gridView1.SelectAll();
                chBoxDelegatePersentage.Checked = true;
                checkBoxSell_Discount.Checked = true;
                checkBoxPrice.Checked = true;
                checkBoxNormal_Increase.Checked = true;
                checkBoxCategorical_Increase.Checked = true;
                load = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            dbconnection.Close();
        }
       
        //deign event
        private void chBoxAdditionalIncrease_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chBoxAdditionalIncrease.Checked)
                {
                    tLPanCpntent.RowStyles[1].Height = 360;                    
                }
                else
                {
                    tLPanCpntent.RowStyles[1].Height = 200;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void radioList_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                label14.Text = "خصم البيع";
                txtNormal.Visible = true;
                txtUnNormal.Visible = true;
                label15.Visible = true;
                label16.Visible = true;
                checkBoxNormal_Increase.Visible = true;
                checkBoxCategorical_Increase.Visible = true;
       
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void radioQata3y_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                label14.Text = "نسبة الاضافة";
                txtNormal.Visible = false;
                txtUnNormal.Visible = false;
                label15.Visible = false;
                label16.Visible = false;
                checkBoxNormal_Increase.Visible = false;
                checkBoxCategorical_Increase.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void chBoxSpecialIncrease_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chBoxSpecialIncrease.Checked)
                {
                    gridView1.SelectAll();
                    tLPanCpntent.RowStyles[1].Height = 200;
                   
                    foreach (Control item in panContent.Controls)
                    {
                        item.Visible = false;
                    }
                    panContent.Controls["panContent"].Visible = true;
                    label19.Visible = false;
                    labSellPrice.Visible = false;
                }
                else
                {
                    if(chBoxAdditionalIncrease.Checked)
                        tLPanCpntent.RowStyles[1].Height = 360;
                    else
                        tLPanCpntent.RowStyles[1].Height = 200;
                    
                    foreach (Control item in panContent.Controls)
                    {
                        item.Visible = true;
                    }
                    panContent.Controls["panContent"].Visible = false;
                    label19.Visible = true;
                    labSellPrice.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
       
        //main events
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
                            case "txtCodePart1":
                                query = "select Type_Name from type where Type_ID='" + txtCodePart1.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                 
                                    makeCode(txtCodePart1);
                                    txtCodePart2.Focus();
                                    dbconnection.Close();
                                }
                                else
                                {
                                    MessageBox.Show("هذا الكود غير مسجل");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtCodePart2":
                                query = "select Factory_Name from factory where Factory_ID='" + txtCodePart2.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                   
                                    makeCode(txtCodePart2);
                                    txtCodePart3.Focus();
                                    dbconnection.Close();
                                }
                                else
                                {
                                    MessageBox.Show("هذا الكود غير مسجل");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtCodePart3":
                                query = "select Group_Name from Groupo where Group_ID='" + txtCodePart3.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                   
                                    makeCode(txtCodePart3);
                                    txtCodePart4.Focus();
                                    dbconnection.Close();
                                }
                                else
                                {
                                    MessageBox.Show("هذا الكود غير مسجل");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtCodePart4":
                                query = "select Product_Name from Product where Product_ID='" + txtCodePart4.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                   
                                    makeCode(txtCodePart4);
                                    txtCodePart5.Focus();
                                    dbconnection.Close();
                                }
                                else
                                {
                                    MessageBox.Show("هذا الكود غير مسجل");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtCodePart5":
                                makeCode(txtCodePart5);
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
        private void gridControl1_EditorKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)(((GridView)gridControl1.MainView).GetRow(((GridView)gridControl1.MainView).GetSelectedRows()[0]));
                if (row != null)
                {
                    txtCode.Text = row[1].ToString();
                    id = Convert.ToInt32(row[0].ToString());
                    String code = txtCode.Text;
                    displayCode(code);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                if (!chBoxSpecialIncrease.Checked)
                {
                    chackUpdateType();
                    if (updateType == "NormalCase")
                    {
                        #region ideal case of update 
                        double price = double.Parse(txtPrice.Text);
                        double SellPercent = double.Parse(txtSell.Text);

                        if (radioQata3y.Checked == true)
                        {
                            #region set qata3yPrice for list item

                            double NormalPercent = double.Parse(txtNormal.Text);
                            double UnNormalPercent = double.Parse(txtUnNormal.Text);

                            int[] listOfSelectedRows = gridView1.GetSelectedRows();
                            for (int i = 0; i < gridView1.SelectedRowsCount; i++)
                            {
                                DataRowView row = (DataRowView)(((GridView)gridControl1.MainView).GetRow(((GridView)gridControl1.MainView).GetSelectedRows()[i]));
                               gridControl1.MainView.GetRow(((GridView)gridControl1.MainView).GetSelectedRows()[i]);

                                additionalIncreaseSellPrice(Convert.ToInt32(row[0].ToString()));
                                String query = "update sellprice set Sell_Discount=@Sell_Discount,Normal_Increase=@Normal_Increase,Categorical_Increase=@Categorical_Increase,Price_Type=@Price_Type,Sell_Price=@Sell_Price,ProfitRatio=@ProfitRatio,Price=@Price,PercentageDelegate=@PercentageDelegate,Date=@Date,OfferFlag=@OfferFlag where SellPrice_ID=" + row[0].ToString();
                                MySqlCommand command = new MySqlCommand(query, dbconnection);
                                command.Parameters.AddWithValue("@Price_Type", "قطعى");
                                command.Parameters.AddWithValue("@Sell_Price", price + (price * SellPercent / 100.0));
                                command.Parameters.AddWithValue("@ProfitRatio", SellPercent);
                                command.Parameters.AddWithValue("@Sell_Discount", 0.00);
                                command.Parameters.AddWithValue("@Price", price);
                                command.Parameters.AddWithValue("@Normal_Increase", 0.00);
                                command.Parameters.AddWithValue("@Categorical_Increase", 0.00);
                                command.Parameters.AddWithValue("@PercentageDelegate", double.Parse(txtPercentageDelegate.Text));
                                command.Parameters.Add("?Date", MySqlDbType.Date);
                                command.Parameters["?Date"].Value = DateTime.Now.Date;
                                command.Parameters.Add("?OfferFlag", MySqlDbType.Int16);
                                if (checkBoxOfferFlag.Checked)
                                {
                                    command.Parameters["?OfferFlag"].Value = 1;
                                }
                                else
                                {
                                    command.Parameters["?OfferFlag"].Value = 0;
                                }
                                command.ExecuteNonQuery();
               
                                UserControl.ItemRecord("sellprice", "تعديل", Convert.ToInt32(row[0].ToString()), DateTime.Now, "", dbconnection);

                                //insert into Archif Table
                                query = "INSERT INTO oldsellprice (Sell_Discount,Price_Type, Sell_Price, ProfitRatio, Data_ID, Price,Last_Price, PercentageDelegate,Date) VALUES(?Sell_Discount,?Price_Type,?Sell_Price,?ProfitRatio,?Data_ID,?Price,?Last_Price,?PercentageDelegate,?Date)";
                                command = new MySqlCommand(query, dbconnection);
                                command.Parameters.AddWithValue("?Price_Type", "قطعى");
                                command.Parameters.AddWithValue("?Sell_Price", calSellPrice());
                                command.Parameters.AddWithValue("?Data_ID", row[1].ToString());
                                command.Parameters.AddWithValue("?ProfitRatio", double.Parse(txtSell.Text));
                                command.Parameters.AddWithValue("?Price", price);
                                command.Parameters.AddWithValue("?Last_Price", calSellPrice());
                                command.Parameters.AddWithValue("?Sell_Discount", 0.0);
                                command.Parameters.AddWithValue("?PercentageDelegate", double.Parse(txtPercentageDelegate.Text));
                                command.Parameters.Add("?Date", MySqlDbType.Date);
                                command.Parameters["?Date"].Value = DateTime.Now.Date;

                                command.ExecuteNonQuery();
                                additionalIncreaseOldSellPrice();

                            }

                            #endregion
                        }
                        else
                        {
                            #region set priceList for collection of items

                            double NormalPercent = double.Parse(txtNormal.Text);
                            double unNormalPercent = double.Parse(txtUnNormal.Text);

                            double sellPrice = (price + NormalPercent) - ((price + NormalPercent) * SellPercent / 100.0);

                            sellPrice = sellPrice + unNormalPercent;

                            int[] listOfSelectedRows = gridView1.GetSelectedRows();
                            for (int i = 0; i < gridView1.SelectedRowsCount; i++)
                            {
                                DataRowView row = (DataRowView)(((GridView)gridControl1.MainView).GetRow(((GridView)gridControl1.MainView).GetSelectedRows()[i]));

                                additionalIncreaseSellPrice(Convert.ToInt32(row[0].ToString()));

                                string query = "update sellprice set ProfitRatio=@ProfitRatio, Price_Type=@Price_Type,Sell_Price=@Sell_Price,Sell_Discount=@Sell_Discount,Price=@Price,Normal_Increase=@Normal_Increase,Categorical_Increase=@Categorical_Increase,PercentageDelegate=@PercentageDelegate ,Last_Price=@Last_Price,Date=@Date,OfferFlag=@OfferFlag where SellPrice_ID =" + row[0].ToString();

                                MySqlCommand command = new MySqlCommand(query, dbconnection);
                                command.Parameters.AddWithValue("@Price_Type", "لستة");
                                command.Parameters.AddWithValue("@Sell_Price", sellPrice);
                                command.Parameters.AddWithValue("@ProfitRatio", 0.00);
                                command.Parameters.AddWithValue("@Sell_Discount", double.Parse(txtSell.Text));
                                command.Parameters.AddWithValue("@Price", price);
                                command.Parameters.AddWithValue("@Last_Price", lastPrice());
                                command.Parameters.AddWithValue("@Normal_Increase", double.Parse(txtNormal.Text));
                                command.Parameters.AddWithValue("@Categorical_Increase", double.Parse(txtUnNormal.Text));
                                command.Parameters.AddWithValue("@PercentageDelegate", double.Parse(txtPercentageDelegate.Text));
                                command.Parameters.Add("?Date", MySqlDbType.Date);
                                command.Parameters["?Date"].Value = DateTime.Now.Date;
                                command.Parameters.Add("?OfferFlag", MySqlDbType.Int16);
                                if (checkBoxOfferFlag.Checked)
                                {
                                    command.Parameters["?OfferFlag"].Value = 1;
                                }
                                else
                                {
                                    command.Parameters["?OfferFlag"].Value = 0;
                                }
                                command.ExecuteNonQuery();

                                UserControl.ItemRecord("sellprice", "تعديل", Convert.ToInt32(row[0].ToString()), DateTime.Now, "", dbconnection);

                                //insert into Archif table
                                query = "INSERT INTO oldsellprice (Last_Price,Price_Type,Sell_Price,Data_ID,Sell_Discount,Price,Normal_Increase,Categorical_Increase,PercentageDelegate,Date) VALUES (?Last_Price,?Price_Type,?Sell_Price,?Data_ID,?Sell_Discount,?Price,?Normal_Increase,?Categorical_Increase,?PercentageDelegate,?Date)";
                                command = new MySqlCommand(query, dbconnection);
                                command.Parameters.AddWithValue("@Price_Type", "لستة");
                                command.Parameters.AddWithValue("@Sell_Price", calSellPrice());
                                command.Parameters.AddWithValue("?Data_ID", row[1].ToString());
                                command.Parameters.AddWithValue("@Sell_Discount", double.Parse(txtSell.Text));
                                command.Parameters.AddWithValue("@Price", price);
                                command.Parameters.AddWithValue("@Last_Price", lastPrice());
                                command.Parameters.AddWithValue("@Normal_Increase", double.Parse(txtNormal.Text));
                                command.Parameters.AddWithValue("@Categorical_Increase", double.Parse(txtUnNormal.Text));
                                command.Parameters.AddWithValue("@PercentageDelegate", double.Parse(txtPercentageDelegate.Text));
                                command.Parameters.Add("?Date", MySqlDbType.Date);
                                command.Parameters["?Date"].Value = DateTime.Now.Date;
                                command.Parameters.Add("?OfferFlag", MySqlDbType.Int16);
                                if (checkBoxOfferFlag.Checked)
                                {
                                    command.Parameters["?OfferFlag"].Value = 1;
                                }
                                else
                                {
                                    command.Parameters["?OfferFlag"].Value = 0;
                                }
                                command.ExecuteNonQuery();
                                additionalIncreaseOldSellPrice();

                            }

                            #endregion
                        }

                        #endregion
                    }
                    else
                    {
                        #region unIdeal case of update
                        resetValues();
                        setValues();
                        if (radioQata3y.Checked == true)
                        {
                            #region set qata3yPrice for list item

                            //double NormalPercent = double.Parse(txtNormal.Text);
                            //double UnNormalPercent = double.Parse(txtUnNormal.Text);

                            int[] listOfSelectedRows = gridView1.GetSelectedRows();
                            for (int i = 0; i < gridView1.SelectedRowsCount; i++)
                            {
                                DataRowView row = (DataRowView)(((GridView)gridControl1.MainView).GetRow(((GridView)gridControl1.MainView).GetSelectedRows()[i]));

                                additionalIncreaseSellPrice(Convert.ToInt32(row[0].ToString()));

                                String query = "update sellprice set Sell_Discount=@Sell_Discount,Normal_Increase=@Normal_Increase,Categorical_Increase=@Categorical_Increase,Price_Type=@Price_Type,Sell_Price=@Sell_Price,ProfitRatio=@ProfitRatio,Last_Price=@Last_Price,Price=@Price,PercentageDelegate=@PercentageDelegate,Date=@Date,OfferFlag=@OfferFlag where SellPrice_ID=" + row[0].ToString();

                                MySqlCommand command = new MySqlCommand(query, dbconnection);
                                command.Parameters.AddWithValue("@Price_Type", "قطعى");
                                if (Price != -1 && ProfitRatio != -1)
                                {
                                    command.Parameters.AddWithValue("@Sell_Price", Price + (Price * ProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@Last_Price", Price + (Price * ProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@ProfitRatio", ProfitRatio);
                                    command.Parameters.AddWithValue("@Price", Price);
                                }
                                else if (Price == -1 && ProfitRatio != -1)
                                {
                                    double xPrice = Convert.ToDouble(row[7]);
                                    command.Parameters.AddWithValue("@Sell_Price", xPrice + (xPrice * ProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@Last_Price", Price + (Price * ProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@ProfitRatio", ProfitRatio);
                                    command.Parameters.AddWithValue("@Price", xPrice);
                                }
                                else if (Price != -1 && ProfitRatio == -1)
                                {
                                    double xProfitRatio = Convert.ToDouble(row[12]);
                                    command.Parameters.AddWithValue("@Sell_Price", Price + (Price * xProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@Last_Price", Price + (Price * ProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@ProfitRatio", xProfitRatio);
                                    command.Parameters.AddWithValue("@Price", Price);
                                }
                                else if (Price == -1 && ProfitRatio == -1)
                                {
                                    double xPrice = Convert.ToDouble(row[7]);
                                    double xProfitRatio = Convert.ToDouble(row[12]);
                                    command.Parameters.AddWithValue("@Sell_Price", xPrice + (xPrice * xProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@Last_Price", Price + (Price * ProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@ProfitRatio", xProfitRatio);
                                    command.Parameters.AddWithValue("@Price", xPrice);
                                }
                                command.Parameters.AddWithValue("@Sell_Discount", 0.00);
                                command.Parameters.AddWithValue("@Normal_Increase", 0.00);
                                command.Parameters.AddWithValue("@Categorical_Increase", 0.00);
                                if (PercentageDelegate != -1)
                                    command.Parameters.AddWithValue("@PercentageDelegate", PercentageDelegate);
                                else
                                    command.Parameters.AddWithValue("@PercentageDelegate", Convert.ToDouble(row["نسبة المندوب"]));

                                command.Parameters.Add("?Date", MySqlDbType.Date);
                                command.Parameters["?Date"].Value = DateTime.Now.Date;
                                command.Parameters.Add("?OfferFlag", MySqlDbType.Int16);
                                if (checkBoxOfferFlag.Checked)
                                {
                                    command.Parameters["?OfferFlag"].Value = 1;
                                }
                                else
                                {
                                    command.Parameters["?OfferFlag"].Value = 0;
                                }
                                command.ExecuteNonQuery();

                                UserControl.ItemRecord("sellprice", "تعديل", Convert.ToInt32(row[0].ToString()), DateTime.Now, "", dbconnection);


                                //insert into Archif Table
                                query = "INSERT INTO oldsellprice (Sell_Discount,Price_Type, Sell_Price, ProfitRatio, Data_ID, Price,PercentageDelegate,Date) VALUES(?Sell_Discount,?Price_Type,?Sell_Price,?ProfitRatio,?Data_ID,?Price,?PercentageDelegate,?Date)";
                                command = new MySqlCommand(query, dbconnection);
                                command.Parameters.AddWithValue("@Price_Type", "قطعى");
                                command.Parameters.AddWithValue("@Data_ID", row[1]);
                                if (Price != -1 && ProfitRatio != -1)
                                {
                                    command.Parameters.AddWithValue("@Sell_Price", Price + (Price * ProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@ProfitRatio", ProfitRatio);
                                    command.Parameters.AddWithValue("@Price", Price);
                                }
                                else if (Price == -1 && ProfitRatio != -1)
                                {
                                    double xPrice = Convert.ToDouble(row[7]);
                                    command.Parameters.AddWithValue("@Sell_Price", xPrice + (xPrice * ProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@ProfitRatio", ProfitRatio);
                                    command.Parameters.AddWithValue("@Price", xPrice);
                                }
                                else if (Price != -1 && ProfitRatio == -1)
                                {
                                    double xProfitRatio = Convert.ToDouble(row[12]);
                                    command.Parameters.AddWithValue("@Sell_Price", Price + (Price * xProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@ProfitRatio", xProfitRatio);
                                    command.Parameters.AddWithValue("@Price", Price);
                                }
                                else if (Price == -1 && ProfitRatio == -1)
                                {
                                    double xPrice = Convert.ToDouble(row[7]);
                                    double xProfitRatio = Convert.ToDouble(row[12]);
                                    command.Parameters.AddWithValue("@Sell_Price", xPrice + (xPrice * xProfitRatio / 100.0));
                                    command.Parameters.AddWithValue("@ProfitRatio", xProfitRatio);
                                    command.Parameters.AddWithValue("@Price", xPrice);
                                }
                                command.Parameters.AddWithValue("@Sell_Discount", 0.00);
                                command.Parameters.AddWithValue("@Normal_Increase", 0.00);
                                command.Parameters.AddWithValue("@Categorical_Increase", 0.00);
                                if (PercentageDelegate != -1)
                                    command.Parameters.AddWithValue("@PercentageDelegate", PercentageDelegate);
                                else
                                    command.Parameters.AddWithValue("@PercentageDelegate", Convert.ToDouble(row["نسبة المندوب"]));

                                command.Parameters.Add("?Date", MySqlDbType.Date);
                                command.Parameters["?Date"].Value = DateTime.Now.Date;
                             
                                command.ExecuteNonQuery();
                                additionalIncreaseOldSellPrice();

                            }

                            #endregion
                        }
                        else
                        {
                            #region set priceList for collection of items

                            int[] listOfSelectedRows = gridView1.GetSelectedRows();
                            for (int i = 0; i < gridView1.SelectedRowsCount; i++)
                            {
                                DataRowView row = (DataRowView)(((GridView)gridControl1.MainView).GetRow(((GridView)gridControl1.MainView).GetSelectedRows()[listOfSelectedRows[i]]));

                                additionalIncreaseSellPrice(Convert.ToInt32(row[0].ToString()));

                                string query = "update sellprice set ProfitRatio=@ProfitRatio, Price_Type=@Price_Type," +/*Sell_Price=@Sell_Price*/"Sell_Discount=@Sell_Discount,Price=@Price,Normal_Increase=@Normal_Increase,Categorical_Increase=@Categorical_Increase,PercentageDelegate=@PercentageDelegate ,Last_Price=@Last_Price,Date=@Date,OfferFlag=@OfferFlag where SellPrice_ID =" + row[0].ToString();

                                MySqlCommand command = new MySqlCommand(query, dbconnection);
                                command.Parameters.AddWithValue("@Price_Type", "لستة");
                                command.Parameters.AddWithValue("@ProfitRatio", 0.00);
                                if (Sell_Discount != -1)
                                {
                                    command.Parameters.AddWithValue("@Sell_Discount", Sell_Discount);
                                }
                                else
                                {
                                    command.Parameters.AddWithValue("@Sell_Discount", Convert.ToDouble(row["خصم البيع"]));
                                }
                                if (Price != -1)
                                    command.Parameters.AddWithValue("@Price", Price);
                                else
                                    command.Parameters.AddWithValue("@Price", Convert.ToDouble(row["السعر"]));

                                command.Parameters.AddWithValue("@Last_Price", lastPrice());

                                if (Normal_Increase != -1)
                                    command.Parameters.AddWithValue("@Normal_Increase", Normal_Increase);
                                else
                                    command.Parameters.AddWithValue("@Normal_Increase", Convert.ToDouble(row["الزيادة العادية"]));
                                if (Categorical_Increase != -1)
                                    command.Parameters.AddWithValue("@Categorical_Increase", Categorical_Increase);
                                else
                                    command.Parameters.AddWithValue("@Categorical_Increase", Convert.ToDouble(row["الزيادة القطعية"]));

                                if (PercentageDelegate != -1)
                                    command.Parameters.AddWithValue("@PercentageDelegate", PercentageDelegate);
                                else
                                    command.Parameters.AddWithValue("@PercentageDelegate", Convert.ToDouble(row["نسبة المندوب"]));

                                command.Parameters.Add("?Date", MySqlDbType.Date);
                                command.Parameters["?Date"].Value = DateTime.Now.Date;
                                command.Parameters.Add("?OfferFlag", MySqlDbType.Int16);
                                if (checkBoxOfferFlag.Checked)
                                {
                                    command.Parameters["?OfferFlag"].Value = 1;
                                }
                                else
                                {
                                    command.Parameters["?OfferFlag"].Value = 0;
                                }
                                command.ExecuteNonQuery();

                                
                                UserControl.ItemRecord("sellprice", "تعديل", Convert.ToInt32(row[0].ToString()), DateTime.Now, "", dbconnection);

                                //insert into Archif table
                                query = "INSERT INTO oldsellprice (Price_Type,Data_ID,Sell_Discount,Price,Normal_Increase,Categorical_Increase,PercentageDelegate,Date) VALUES (?Price_Type,?Data_ID,?Sell_Discount,?Price,?Normal_Increase,?Categorical_Increase,?PercentageDelegate,?Date)";
                                command = new MySqlCommand(query, dbconnection);
                                command.Parameters.AddWithValue("@Price_Type", "لستة");
                                command.Parameters.AddWithValue("@Data_ID", row[1].ToString());

                                command.Parameters.AddWithValue("@ProfitRatio", 0.00);
                                if (Sell_Discount != -1)
                                {
                                    command.Parameters.AddWithValue("@Sell_Discount", Sell_Discount);
                                }
                                else
                                {
                                    command.Parameters.AddWithValue("@Sell_Discount", Convert.ToDouble(row["خصم البيع"]));
                                }
                                if (Price != -1)
                                    command.Parameters.AddWithValue("@Price", Price);
                                else
                                    command.Parameters.AddWithValue("@Price", Convert.ToDouble(row["السعر"]));

                                if (Normal_Increase != -1)
                                    command.Parameters.AddWithValue("@Normal_Increase", Normal_Increase);
                                else
                                    command.Parameters.AddWithValue("@Normal_Increase", Convert.ToDouble(row["الزيادة العادية"]));
                                if (Categorical_Increase != -1)
                                    command.Parameters.AddWithValue("@Categorical_Increase", Categorical_Increase);
                                else
                                    command.Parameters.AddWithValue("@Categorical_Increase", Convert.ToDouble(row["الزيادة القطعية"]));

                                if (PercentageDelegate != -1)
                                    command.Parameters.AddWithValue("@PercentageDelegate", PercentageDelegate);
                                else
                                    command.Parameters.AddWithValue("@PercentageDelegate", Convert.ToDouble(row["نسبة المندوب"]));

                                command.Parameters.Add("?Date", MySqlDbType.Date);
                                command.Parameters["?Date"].Value = DateTime.Now.Date;
                                command.ExecuteNonQuery();
                                additionalIncreaseOldSellPrice();
                            }

                            #endregion
                        } 
                        #endregion
                    }
               
                    
                }
                else
                {
                    double addationalValue = 0.0;
                    DataGridView DataGridView = (DataGridView)panContent.Controls["panContent"].Controls["dataGridView1"];
                    foreach (DataGridViewRow item1 in DataGridView.Rows)
                    {
                        addationalValue += Convert.ToDouble(item1.Cells[0].Value);
                    }
                    int[] rows = (((GridView)gridControl1.MainView).GetSelectedRows());
                    List<DataRowView> recordList = new List<DataRowView>();
                    for (int i = 0; i < rows.Length; i++)
                    {
                        DataRowView a = (DataRowView)(((GridView)gridControl1.MainView).GetRow(rows[i]));
                        recordList.Add(a);
                    }
                    foreach (DataRowView item in recordList)
                    {
                        String query = "update sellprice set Sell_Price=@Sell_Price where SellPrice_ID=" + item[0].ToString();

                        MySqlCommand command = new MySqlCommand(query, dbconnection);
                        command.Parameters.AddWithValue("@Sell_Price",Convert.ToDouble(item["سعر البيع"])+addationalValue);           
                        command.ExecuteNonQuery();

                        foreach (DataGridViewRow item1 in DataGridView.Rows)
                        {
                            double addational = Convert.ToDouble(item1.Cells[0].Value);
                            string  queryx = "insert into special_increase (SellPrice_ID,Value,Description,Date) values (@SellPrice_ID,@Value,@Description,@Date)";
                            MySqlCommand com = new MySqlCommand(queryx, dbconnection);
                            com.Parameters.AddWithValue("@SellPrice_ID", item[0]);
                            com.Parameters.AddWithValue("@Value", item1.Cells[0].Value);
                            com.Parameters.AddWithValue("@Description", item1.Cells[1].Value);
                            com.Parameters.Add("@Date", MySqlDbType.Date);
                            com.Parameters["@Date"].Value = DateTime.Now.Date;

                            com.ExecuteNonQuery();

                        }
                    }
            
                }
                displayData();
               // productsSellPriceForm.displayProducts();
                MessageBox.Show("تم");
                gridView1.SelectAll();
                XtraTabPage xtraTabPage = getTabPage("تعديل اسعار البنود");
                xtraTabPage.ImageOptions.Image = null;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void chBoxSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                //gridView1.OptionsSelection.MultiSelect = true;
                //if (chBoxSelectAll.Checked)
                //{
                //    gridView1.SelectRows(0, gridView1.RowCount - 1);
                //    txtCode.Text = "";
                //    txtCode.Visible = false;
                //    panCodeParts.Visible = false;
                //    label6.Visible = false;
                //    gridView1.OptionsSelection.MultiSelect = true;
                //    gridView1.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect;
                //}
                //else
                //{
                //    int selectedRowsCount = gridView1.SelectedRowsCount;
                //    for (int i = 0; i < selectedRowsCount; i++)
                //    {
                //        gridView1.UnselectRow(i);
                //    }
                //    txtCode.Visible = true;
                //    panCodeParts.Visible = true;
                //    label6.Visible = true;
                //    gridView1.OptionsSelection.MultiSelect = false;
                //    gridView1.OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect;
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void txtBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (load)
                {
                    XtraTabPage xtraTabPage = getTabPage("تعديل اسعار البنود");
                    if (!IsClear())
                        xtraTabPage.ImageOptions.Image = Properties.Resources.unsave;
                    else
                        xtraTabPage.ImageOptions.Image = null;

                    labSellPrice.Text = calSellPrice() + "";
                }
            }
            catch (Exception)
            {

            }

        }
        private void btnNewPlus_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridView1.SelectedRowsCount>0)
                {
                    if (txtDes.Text != "" && txtPlus.Text != "")
                    {
                        if (radioNormal.Checked)
                        {
                            int n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = txtPlus.Text;
                            dataGridView1.Rows[n].Cells[1].Value = "عادية";
                            dataGridView1.Rows[n].Cells[2].Value = txtDes.Text;
                            labSellPrice.Text = "" + calSellPrice();
                        }
                        else if (radioQata3a.Checked)
                        {
                            int n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = txtPlus.Text;
                            dataGridView1.Rows[n].Cells[1].Value = "قطعية";
                            dataGridView1.Rows[n].Cells[2].Value = txtDes.Text;
                            labSellPrice.Text = "" + calSellPrice();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridViewRow row1 = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];
                if (row1 != null)
                {
                    dataGridView1.Rows.Remove(row1);
                    labSellPrice.Text = calSellPrice()+"";
                }
                else
                {
                    MessageBox.Show("you must select an item");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        private void gridView1_SelectionChanged(object sender, DevExpress.Data.SelectionChangedEventArgs e)
        {
            try
            {
                dbconnection.Close();
                dbconnection.Open();
                if (gridView1.SelectedRowsCount == 1)
                {
                    txtCode.Visible = true;
                    panCodeParts.Visible = true;
                    label6.Visible = true;
                    label19.Visible = true;
                    labSellPrice.Visible = true;
                    DataRowView row = (DataRowView)(((GridView)gridControl1.MainView).GetRow(((GridView)gridControl1.MainView).GetSelectedRows()[0]));
                    if (row != null)
                    {
                        txtCode.Text = row[2].ToString();
                        id = Convert.ToInt32(row[0].ToString());
                        String code = txtCode.Text;
                        displayCode(code);
                        setData(row);
                    }
                }
                else if (gridView1.SelectedRowsCount > 1)
                {
                    DataRowView row = (DataRowView)(((GridView)gridControl1.MainView).GetRow(((GridView)gridControl1.MainView).GetSelectedRows()[0]));

                    txtCode.Visible = false;
                    panCodeParts.Visible = false;
                    label6.Visible = false;
                    label19.Visible = false;
                    labSellPrice.Visible = false;
                    setData(row);
                }
                else
                {
                    txtCode.Visible = true;
                    panCodeParts.Visible = true;
                    label6.Visible = true;
                    label19.Visible = true;
                    labSellPrice.Visible = true;
                    txtCode.Text = "";
                    txtCodePart1.Text = txtCodePart2.Text = txtCodePart3.Text = txtCodePart4.Text = txtCodePart5.Text = "";
                    id = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                CheckBox checkBox = (CheckBox)sender;
                switch (checkBox.Name)
                {
                    case "chBoxDelegatePersentage":
                        if (chBoxDelegatePersentage.Checked)
                        {
                            try
                            {
                                txtPercentageDelegate.BackColor = Color.White;
                                PercentageDelegate = Convert.ToDouble(txtPercentageDelegate.Text);

                            }
                            catch
                            {
                                txtPercentageDelegate.Focus();
                                txtPercentageDelegate.BackColor = Color.RosyBrown;
                            }
                        }
                        else
                        {
                            PercentageDelegate = -1;
                        }
                        break;
                    case "checkBoxSell_Discount":
                        if (checkBoxSell_Discount.Checked)
                        {
                            try
                            {
                                txtSell.BackColor = Color.White;
                                if (radioList.Checked)
                                    Sell_Discount = Convert.ToDouble(txtSell.Text);
                                else
                                    ProfitRatio = Convert.ToDouble(txtSell.Text);
                            }
                            catch
                            {
                                txtSell.Focus();
                                txtSell.BackColor = Color.RosyBrown;
                            }
                        }
                        else
                        {
                            Sell_Discount = -1;
                            ProfitRatio = -1;
                        }
                        break;
                    case "checkBoxPrice":
                        if (checkBoxPrice.Checked)
                        {
                            try
                            {
                                txtPrice.BackColor = Color.White;
                                Price = Convert.ToDouble(txtPrice.Text);
                            }
                            catch
                            {
                                txtPrice.Focus();
                                txtPrice.BackColor = Color.RosyBrown;
                            }
                        }
                        else
                        {
                            Price = -1;
                        }
                        break;
                    case "checkBoxNormal_Increase":
                        if (checkBoxNormal_Increase.Checked)
                        {
                            try
                            {
                                txtNormal.BackColor = Color.White;
                                Normal_Increase = Convert.ToDouble(txtNormal.Text);
                            }
                            catch
                            {
                                txtNormal.Focus();
                                txtNormal.BackColor = Color.RosyBrown;
                            }
                        }
                        else
                        {
                            Normal_Increase = -1;
                        }
                        break;
                    case "checkBoxCategorical_Increase":
                        if (checkBoxCategorical_Increase.Checked)
                        {
                            try
                            {
                                txtUnNormal.BackColor = Color.White;
                                Categorical_Increase = Convert.ToDouble(txtUnNormal.Text);

                            }
                            catch
                            {
                                txtUnNormal.Focus();
                                txtUnNormal.BackColor = Color.RosyBrown;
                            }
                        }
                        else
                        {
                            Categorical_Increase = -1;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //function 
        public void displayData()
        {
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            gridControl1.DataSource = dt;
          
            gridView1.Columns[0].Visible = false;
            gridView1.Columns[1].Visible = false;
            gridView1.Columns[2].Width = 200;
            gridView1.Columns[3].Width = 300;
            gridView1.BestFitColumns();
        }
        public void setData(DataRowView row1)
        {
            //txtCode.Text = row1["الكود"].ToString();
            int offerFlag = Convert.ToInt16(row1["OfferFlag"].ToString());
            if (offerFlag == 1)
            {
                checkBoxOfferFlag.Checked = true;
            }
            else
            {
                checkBoxOfferFlag.Checked = false;
            }
            txtPrice.Text = row1["السعر"].ToString();
            labSellPrice.Text = row1["سعر البيع"].ToString();
            txtPercentageDelegate.Text = row1["نسبة المندوب"].ToString();
            string str = row1["نوع السعر"].ToString();
            if (str == "لستة")
            {
                txtSell.Text = row1["خصم البيع"].ToString();
                txtNormal.Text = row1["الزيادة العادية"].ToString();
                txtUnNormal.Text = row1["الزيادة القطعية"].ToString();
                radioList.Checked = true;
            }
            else
            {
                txtSell.Text = row1["نسبة الاضافة"].ToString();
                radioQata3y.Checked = true;
            }
            string query = "select AdditionalValue,Type,Description from additional_increase_sellprice where SellPrice_ID="+ row1[0].ToString(); ;
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            MySqlDataReader dr = com.ExecuteReader();
            dataGridView1.Rows.Clear();
            while (dr.Read())
            {
                chBoxAdditionalIncrease.Checked = true;
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = dr[0].ToString();
                dataGridView1.Rows[n].Cells[1].Value = dr[1].ToString();
                dataGridView1.Rows[n].Cells[2].Value = dr[2].ToString();

            }
            dr.Close();
            calSellPrice();

        }
        public XtraTabPage getTabPage(string text)
        {
            for (int i = 0; i < xtraTabControlSalesContent.TabPages.Count; i++)
                if (xtraTabControlSalesContent.TabPages[i].Text == text)
                {
                    return xtraTabControlSalesContent.TabPages[i];
                }
            return null;
        }
        public bool IsClear()
        {
            if (txtCode.Text == "" &&
                txtPrice.Text == "0" &&
                txtSell.Text == "0" &&
                txtNormal.Text == "0" &&
                txtUnNormal.Text == "0" &&
                radioList.Checked == true)
                return true;
            else
                return false;
        }
        public void Clear()
        {
            txtCode.Text = "";
            txtPrice.Text = "0";
            txtSell.Text = "0";
            txtNormal.Text = "0";
            txtUnNormal.Text = "0";
            txtDes.Text = "";
            txtPlus.Text = "";
            dataGridView1.Rows.Clear();
            txtCodePart1.Text = txtCodePart2.Text = txtCodePart3.Text = txtCodePart4.Text = txtCodePart5.Text = "";
            radioList.Checked = true;
        }
        public double calSellPrice()
        {
            double addational = 0.0;
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                addational += Convert.ToDouble(item.Cells[0].Value);
            }
            double price = double.Parse(txtPrice.Text);
            double SellPercent = double.Parse(txtSell.Text);
            if (radioQata3y.Checked == true)
            {
                return price + (price * SellPercent / 100.0) + addational;
            }
            else
            {
                double NormalPercent = double.Parse(txtNormal.Text);
                double unNormalPercent = double.Parse(txtUnNormal.Text);
                double sellPrice = (price + NormalPercent) - ((price + NormalPercent) * SellPercent / 100.0);
                sellPrice = sellPrice + unNormalPercent;
                return sellPrice + addational;
            }
        }
        public double lastPrice()
        {
            double price = double.Parse(txtPrice.Text);
            if (dataGridView1.Rows.Count > 0)
            {
                foreach (DataGridViewRow item in dataGridView1.Rows)
                {
                    if (item.Cells[1].Value.ToString() == "عادية")
                        price += Convert.ToDouble(item.Cells[0].Value);
                }
            }
            price += double.Parse(txtNormal.Text);
            return price;
        }
        public void makeCode(TextBox txtBox)
        {
            string code = txtCode.Text;
            int j = 4 - txtBox.TextLength;
            for (int i = 0; i < j; i++)
            {
                code += "0";
            }
            code += txtBox.Text;
            txtCode.Text = code;

        } 
        public void displayCode(string code)
        {
            char[] arrCode = code.ToCharArray();
            txtCodePart1.Text = Convert.ToInt32(arrCode[0].ToString() + arrCode[1].ToString() + arrCode[2].ToString() + arrCode[3].ToString()) + "";
            txtCodePart2.Text = Convert.ToInt32(arrCode[4].ToString() + arrCode[5].ToString() + arrCode[6].ToString() + arrCode[7].ToString()) + "";
            txtCodePart3.Text = Convert.ToInt32(arrCode[8].ToString() + arrCode[9].ToString() + arrCode[10].ToString() + arrCode[11].ToString()) + "";
            txtCodePart4.Text = Convert.ToInt32(arrCode[12].ToString() + arrCode[13].ToString() + arrCode[14].ToString() + arrCode[15].ToString()) + "";
            txtCodePart5.Text = "" + Convert.ToInt32(arrCode[16].ToString() + arrCode[17].ToString() + arrCode[18].ToString() + arrCode[19].ToString());
        }

        //for update
        public void chackUpdateType()
        {
            if (!chBoxDelegatePersentage.Checked||!checkBoxSell_Discount.Checked || !checkBoxPrice.Checked || !checkBoxNormal_Increase.Checked || !checkBoxCategorical_Increase.Checked)
            {
                updateType = "UnNormalCase";
            }
        }
        public void resetValues()
        {
            Price = -1;
            ProfitRatio = -1;
            Sell_Discount = -1;
            Normal_Increase = -1;
            Categorical_Increase = -1;
            PercentageDelegate = -1;
        }
        public void setValues()
        {
            if (chBoxDelegatePersentage.Checked)
            {
                try
                {
                    txtPercentageDelegate.BackColor = Color.White;
                    PercentageDelegate = Convert.ToDouble(txtPercentageDelegate.Text);
                }
                catch
                {
                    txtPercentageDelegate.Focus();
                    txtPercentageDelegate.BackColor = Color.RosyBrown;
                }
            }
            else
            {
                PercentageDelegate = -1;
            }
            if (checkBoxSell_Discount.Checked)
            {
                try
                {
                    txtSell.BackColor = Color.White;
                    if (radioList.Checked)
                        Sell_Discount = Convert.ToDouble(txtSell.Text);
                    else
                        ProfitRatio = Convert.ToDouble(txtSell.Text);
                }
                catch
                {
                    txtSell.Focus();
                    txtSell.BackColor = Color.RosyBrown;
                }
            }
            else
            {
                Sell_Discount = -1;
                ProfitRatio = -1;
            }
            if (checkBoxPrice.Checked)
            {
                try
                {
                    txtPrice.BackColor = Color.White;
                    Price = Convert.ToDouble(txtPrice.Text);
                }
                catch
                {
                    txtPrice.Focus();
                    txtPrice.BackColor = Color.RosyBrown;
                }
            }
            else
            {
                Price = -1;
            }
            if (checkBoxNormal_Increase.Checked)
            {
                try
                {
                    txtNormal.BackColor = Color.White;
                    Normal_Increase = Convert.ToDouble(txtNormal.Text);
                }
                catch
                {
                    txtNormal.Focus();
                    txtNormal.BackColor = Color.RosyBrown;
                }
            }
            else
            {
                Normal_Increase = -1;
            }
            if (checkBoxCategorical_Increase.Checked)
            {
                try
                {
                    txtUnNormal.BackColor = Color.White;
                    Categorical_Increase = Convert.ToDouble(txtUnNormal.Text);

                }
                catch
                {
                    txtUnNormal.Focus();
                    txtUnNormal.BackColor = Color.RosyBrown;
                }
            }
            else
            {
                Categorical_Increase = -1;
            }
        }
        public void additionalIncreaseSellPrice(int sellPrice_ID)
        {
            if (chBoxAdditionalIncrease.Checked)
            {
                string queryx = "delete from additional_increase_sellprice where SellPrice_ID=" + sellPrice_ID;
                MySqlCommand com = new MySqlCommand(queryx, dbconnection);
                com.ExecuteNonQuery();

                foreach (DataGridViewRow item in dataGridView1.Rows)
                {
                    #region addational increase
                    double addational = Convert.ToDouble(item.Cells[0].Value);
                    queryx = "insert into additional_increase_sellprice (SellPrice_ID,AdditionalValue,Type,Description) values (@SellPrice_ID,@AdditionalValue,@Type,@Description)";
                    com = new MySqlCommand(queryx, dbconnection);
                    com.Parameters.AddWithValue("@SellPrice_ID", sellPrice_ID);
                    com.Parameters.AddWithValue("@Type", item.Cells[1].Value);
                    com.Parameters.AddWithValue("@AdditionalValue", item.Cells[0].Value);
                    com.Parameters.AddWithValue("@Description", item.Cells[2].Value);
                    com.ExecuteNonQuery();

                    #endregion
                }
            }
        }
        public void additionalIncreaseOldSellPrice()
        {
            if (chBoxAdditionalIncrease.Checked)
            {
                int oldSellPrice_ID = 0;
                string queryx = "select OldSellPrice_ID from oldsellprice order by OldSellPrice_ID desc limit 1";
                MySqlCommand com = new MySqlCommand(queryx, dbconnection);
                if (com.ExecuteScalar() != null)
                {
                    oldSellPrice_ID = Convert.ToInt32(com.ExecuteScalar());
                }

                foreach (DataGridViewRow item in dataGridView1.Rows)
                {
                    #region addational increase
                    double addational = Convert.ToDouble(item.Cells[0].Value);

                    //insert into archive table
                    queryx = "insert into old_additional_increase_sellprice (OldSellPrice_ID,AdditionalValue,Type,Description) values (@OldSellPrice_ID,@AdditionalValue,@Type,@Description)";
                    com = new MySqlCommand(queryx, dbconnection);
                    com.Parameters.AddWithValue("@OldSellPrice_ID", oldSellPrice_ID);
                    com.Parameters.AddWithValue("@Type", item.Cells[1].Value);
                    com.Parameters.AddWithValue("@AdditionalValue", item.Cells[0].Value);
                    com.Parameters.AddWithValue("@Description", item.Cells[2].Value);
                    com.ExecuteNonQuery();
                    #endregion
                }
            }
        }
    }
}
