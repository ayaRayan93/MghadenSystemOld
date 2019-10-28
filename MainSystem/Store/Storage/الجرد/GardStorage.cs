﻿using DevExpress.XtraGrid.Views.Grid;
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
using Microsoft.VisualBasic;

namespace MainSystem
{
    public partial class GardStorage : Form
    {
        MySqlConnection dbconnection;
        bool load = false;
        bool factoryFlage = false;
        bool groupFlage = false;
        bool flagProduct = false;
        bool flag = false;
        //double noMeter = 0;
        MainForm MainForm;
        XtraTabControl xtraTabControlStoresContent;
        DataTable mdt;
        //DataTable mdtSaved;
        int Data_ID;
        string code = "";
        DataRowView mRow = null;
        List<int> ListOfDataIDs;
        List<int> ListOfSavedDataIDs;
        List<int> ListOfEditDataIDs;

        public GardStorage(/*XtraTabControl xtraTabControlStoresContent*/)
        {
            try
            {
                InitializeComponent();
                dbconnection = new MySqlConnection(connection.connectionString);
                this.MainForm = MainForm;
                ListOfDataIDs = new List<int>();
                ListOfSavedDataIDs = new List<int>();
                ListOfEditDataIDs = new List<int>();
                //this.xtraTabControlStoresContent = xtraTabControlStoresContent;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void GardStorage_Load(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string query = "select * from type";
                MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                comType.DataSource = dt;
                comType.DisplayMember = dt.Columns["Type_Name"].ToString();
                comType.ValueMember = dt.Columns["Type_ID"].ToString();
                comType.Text = "";
                txtType.Text = "";

                query = "select * from factory";
                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                comFactory.DataSource = dt;
                comFactory.DisplayMember = dt.Columns["Factory_Name"].ToString();
                comFactory.ValueMember = dt.Columns["Factory_ID"].ToString();
                comFactory.Text = "";
                txtFactory.Text = "";

                query = "select * from groupo";
                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                comGroup.DataSource = dt;
                comGroup.DisplayMember = dt.Columns["Group_Name"].ToString();
                comGroup.ValueMember = dt.Columns["Group_ID"].ToString();
                comGroup.Text = "";
                txtGroup.Text = "";

                query = "select * from product";
                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                comProduct.DataSource = dt;
                comProduct.DisplayMember = dt.Columns["Product_Name"].ToString();
                comProduct.ValueMember = dt.Columns["Product_ID"].ToString();
                comProduct.Text = "";
                txtProduct.Text = "";

                query = "select * from size";
                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                comSize.DataSource = dt;
                comSize.DisplayMember = dt.Columns["Size_Value"].ToString();
                comSize.ValueMember = dt.Columns["Size_ID"].ToString();
                comSize.Text = "";
                txtSize.Text = "";

                query = "select * from color";
                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                comColor.DataSource = dt;
                comColor.DisplayMember = dt.Columns["Color_Name"].ToString();
                comColor.ValueMember = dt.Columns["Color_ID"].ToString();
                comColor.Text = "";
                txtColor.Text = "";

                query = "select * from sort";
                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                comSort.DataSource = dt;
                comSort.DisplayMember = dt.Columns["Sort_Value"].ToString();
                comSort.ValueMember = dt.Columns["Sort_ID"].ToString();
                comSort.Text = "";
                txtSort.Text = "";

                query = "select distinct Classification from data";
                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                comClassfication.DataSource = dt;
                comClassfication.DisplayMember = dt.Columns["Classification"].ToString();
                comClassfication.Text = "";

                query = "select * from store";
                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                comStore.DataSource = dt;
                comStore.DisplayMember = dt.Columns["Store_Name"].ToString();
                comStore.ValueMember = dt.Columns["Store_ID"].ToString();
                comStore.Text = "";
                load = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            dbconnection.Close();
        }
        private void comBox_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (load)
                {
                    ComboBox comBox = (ComboBox)sender;

                    switch (comBox.Name)
                    {
                        case "comType":
                            if (load)
                            {
                                txtType.Text = comType.SelectedValue.ToString();
                                string query = "select * from factory inner join type_factory on factory.Factory_ID=type_factory.Factory_ID inner join type on type_factory.Type_ID=type.Type_ID where type_factory.Type_ID=" + txtType.Text;
                                MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                                DataTable dt = new DataTable();
                                da.Fill(dt);
                                comFactory.DataSource = dt;
                                comFactory.DisplayMember = dt.Columns["Factory_Name"].ToString();
                                comFactory.ValueMember = dt.Columns["Factory_ID"].ToString();
                                comFactory.Text = "";
                                txtFactory.Text = "";
                                dbconnection.Close();
                                dbconnection.Open();
                                query = "select TypeCoding_Method from type where Type_ID=" + txtType.Text;
                                MySqlCommand com = new MySqlCommand(query, dbconnection);
                                int TypeCoding_Method = (int)com.ExecuteScalar();
                                dbconnection.Close();
                                if (TypeCoding_Method == 1)
                                {
                                    string query2 = "";
                                    if (txtType.Text == "2" || txtType.Text == "1")
                                    {
                                        query2 = "select * from groupo where Factory_ID=-1";
                                    }
                                    else
                                    {
                                        query2 = "select * from groupo where Factory_ID=" + -Convert.ToInt32(txtType.Text) + " and Type_ID=" + txtType.Text;
                                    }

                                    MySqlDataAdapter da2 = new MySqlDataAdapter(query2, dbconnection);
                                    DataTable dt2 = new DataTable();
                                    da2.Fill(dt2);
                                    comGroup.DataSource = dt2;
                                    comGroup.DisplayMember = dt2.Columns["Group_Name"].ToString();
                                    comGroup.ValueMember = dt2.Columns["Group_ID"].ToString();
                                    comGroup.Text = "";
                                    txtGroup.Text = "";
                                    groupFlage = true;
                                }
                                factoryFlage = true;

                                query = "select * from color where Type_ID=" + txtType.Text;
                                da = new MySqlDataAdapter(query, dbconnection);
                                dt = new DataTable();
                                da.Fill(dt);
                                comColor.DataSource = dt;
                                comColor.DisplayMember = dt.Columns["Color_Name"].ToString();
                                comColor.ValueMember = dt.Columns["Color_ID"].ToString();
                                comColor.Text = "";
                                txtColor.Text = "";
                                comFactory.Focus();
                            }
                            break;
                        case "comFactory":
                            if (factoryFlage)
                            {
                                txtFactory.Text = comFactory.SelectedValue.ToString();
                                dbconnection.Close();
                                dbconnection.Open();
                                string query = "select TypeCoding_Method from type where Type_ID=" + txtType.Text;
                                MySqlCommand com = new MySqlCommand(query, dbconnection);
                                int TypeCoding_Method = (int)com.ExecuteScalar();
                                dbconnection.Close();
                                if (TypeCoding_Method == 2)
                                {
                                    string query2f = "select * from groupo where Type_ID=" + txtType.Text + " and Factory_ID=" + txtFactory.Text;
                                    MySqlDataAdapter da2f = new MySqlDataAdapter(query2f, dbconnection);
                                    DataTable dt2f = new DataTable();
                                    da2f.Fill(dt2f);
                                    comGroup.DataSource = dt2f;
                                    comGroup.DisplayMember = dt2f.Columns["Group_Name"].ToString();
                                    comGroup.ValueMember = dt2f.Columns["Group_ID"].ToString();
                                    comGroup.Text = "";
                                    txtGroup.Text = "";
                                }
                                else
                                {
                                    filterProduct();
                                }

                                groupFlage = true;

                                string query2 = "select * from size where Factory_ID=" + txtFactory.Text;
                                MySqlDataAdapter da2 = new MySqlDataAdapter(query2, dbconnection);
                                DataTable dt2 = new DataTable();
                                da2.Fill(dt2);
                                comSize.DataSource = dt2;
                                comSize.DisplayMember = dt2.Columns["Size_Value"].ToString();
                                comSize.ValueMember = dt2.Columns["Size_ID"].ToString();
                                comSize.Text = "";
                                txtSize.Text = "";
                                comGroup.Focus();

                                query = "select distinct Classification from data where Factory_ID=" + txtFactory.Text;
                                da2 = new MySqlDataAdapter(query, dbconnection);
                                dt2 = new DataTable();
                                da2.Fill(dt2);
                                comClassfication.DataSource = dt2;
                                comClassfication.DisplayMember = dt2.Columns["Classification"].ToString();
                                comClassfication.Text = "";
                            }
                            break;
                        case "comGroup":
                            if (groupFlage)
                            {
                                txtGroup.Text = comGroup.SelectedValue.ToString();
                                string supQuery = "", subQuery1 = "";
                                if (txtType.Text != "")
                                {
                                    supQuery += " and product.Type_ID=" + txtType.Text;
                                }
                                if (txtFactory.Text != "")
                                {
                                    supQuery += " and product_factory_group.Factory_ID=" + txtFactory.Text;
                                    subQuery1 += " and Factory_ID=" + txtFactory.Text;
                                }
                                string query3 = "select distinct  product.Product_ID  ,Product_Name  from product inner join product_factory_group on product.Product_ID=product_factory_group.Product_ID  where product_factory_group.Group_ID=" + txtGroup.Text + supQuery + "  order by product.Product_ID";
                                MySqlDataAdapter da3 = new MySqlDataAdapter(query3, dbconnection);
                                DataTable dt3 = new DataTable();
                                da3.Fill(dt3);
                                comProduct.DataSource = dt3;
                                comProduct.DisplayMember = dt3.Columns["Product_Name"].ToString();
                                comProduct.ValueMember = dt3.Columns["Product_ID"].ToString();
                                comProduct.Text = "";
                                txtProduct.Text = "";

                                string query2 = "select * from size where Group_ID=" + txtGroup.Text + subQuery1;
                                MySqlDataAdapter da2 = new MySqlDataAdapter(query2, dbconnection);
                                DataTable dt2 = new DataTable();
                                da2.Fill(dt2);
                                comSize.DataSource = dt2;
                                comSize.DisplayMember = dt2.Columns["Size_Value"].ToString();
                                comSize.ValueMember = dt2.Columns["Size_ID"].ToString();
                                comSize.Text = "";
                                txtSize.Text = "";

                                comProduct.Focus();
                                flagProduct = true;
                            }
                            break;

                        case "comProduct":

                            txtProduct.Text = comProduct.SelectedValue.ToString();
                            comColor.Focus();

                            break;

                        case "comColor":
                            txtColor.Text = comColor.SelectedValue.ToString();
                            comSize.Focus();
                            break;

                        case "comSize":
                            txtSize.Text = comSize.SelectedValue.ToString();
                            break;

                        case "comSort":
                            txtSort.Text = comSort.SelectedValue.ToString();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
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
                            case "txtType":
                                query = "select Type_Name from type where Type_ID='" + txtType.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comType.Text = Name;
                                    txtFactory.Focus();
                                    dbconnection.Close();

                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtFactory":
                                query = "select Factory_Name from factory where Factory_ID='" + txtFactory.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comFactory.Text = Name;
                                    txtGroup.Focus();
                                    dbconnection.Close();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtGroup":
                                query = "select Group_Name from groupo where Group_ID='" + txtGroup.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comGroup.Text = Name;
                                    txtProduct.Focus();
                                    dbconnection.Close();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtProduct":
                                query = "select Product_Name from product where Product_ID='" + txtProduct.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comProduct.Text = Name;
                                    txtType.Focus();
                                    dbconnection.Close();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;

                            case "txtColor":
                                query = "select Color_Name from color where Color_ID='" + txtColor.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comColor.Text = Name;
                                    txtSize.Focus();
                                    dbconnection.Close();
                                    displayProducts();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtSize":
                                query = "select Size_Value from size where Size_ID='" + txtSize.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comSize.Text = Name;
                                    txtSort.Focus();
                                    dbconnection.Close();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtSort":
                                query = "select Sort_Value from sort where Sort_ID='" + txtSort.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comProduct.Text = Name;

                                    dbconnection.Close();
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
        private void comStore_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (load)
                {
                    dbconnection.Close();
                    dbconnection.Open();
                    txtStoreID.Text = comStore.SelectedValue.ToString();
                    string query = "select Inventory_Num,Date from inventory where Store_ID="+txtStoreID.Text+ " order by Inventory_Num desc limit 1 ";
                    MySqlCommand com = new MySqlCommand(query,dbconnection);
                    MySqlDataReader dr = com.ExecuteReader();
                    while (dr.Read())
                    {
                        labGardPermission.Text = dr[0].ToString();
                        dateTimePicker1.Text = dr[1].ToString();
                        ListOfEditDataIDs.Clear();
                        //ListOfSavedDataIDs.Clear();
                        //ListOfDataIDs.Clear();
                        //displayProducts();
                    }
                    dr.Close();
                  
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                displayProducts();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnNewChooes_Click(object sender, EventArgs e)
        {
            try
            {
                comStore.Text = "";

                comType.Text = "";
                comFactory.Text = "";
                comGroup.Text = "";
                comProduct.Text = "";

                txtType.Text = "";
                txtFactory.Text = "";
                txtGroup.Text = "";
                txtProduct.Text = "";
                
                gridControl2.DataSource = null;
                mdt.Rows.Clear();
                flag = false;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void gridView2_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            try
            {
                GridView View = sender as GridView;
                if (e.RowHandle >= 0)
                {
                    int Data_ID = Convert.ToInt32(View.GetRowCellDisplayText(e.RowHandle, View.Columns[0]));

                    for (int i = 0; i < ListOfSavedDataIDs.Count; i++)
                    {
                        if (Data_ID == ListOfSavedDataIDs[i])
                        {
                            e.Appearance.BackColor = Color.MediumSeaGreen;
                            e.Appearance.BackColor2 = Color.MediumSeaGreen;
                            e.Appearance.ForeColor = Color.White;
                        }
                    }
                    for (int i = 0; i < ListOfEditDataIDs.Count; i++)
                    {
                        if (Data_ID == ListOfEditDataIDs[i])
                        {
                            e.Appearance.BackColor = Color.LightBlue;
                            e.Appearance.BackColor2 = Color.LightBlue;
                            e.Appearance.ForeColor = Color.Black;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void gridView2_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            try
            {
                GridView view = (GridView)sender;
                DataRow dataRow = view.GetFocusedDataRow();
                int Data_ID = Convert.ToInt32(dataRow["Data_ID"].ToString());
                addGardQuantity(Data_ID, dataRow);
                ListOfEditDataIDs.Add(Data_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }





        private void btnReport_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridView2.RowCount > 0)
                {
                    if (flag)//to ensure data save first
                    {
                        MainForm.bindReportStorageForm(gridControl2, "تقرير كميات البنود");
                    }
                    else
                    {
                        MessageBox.Show("يجب حفظ البيانات اولا");
                    }
                }
                else
                {
                    MessageBox.Show("لا يوجد بيانات للطباعة");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
           
        }
    
        //functions
        public void clear()
        {

        }
        public XtraTabPage getTabPage(string text)
        {
            for (int i = 0; i < xtraTabControlStoresContent.TabPages.Count; i++)
                if (xtraTabControlStoresContent.TabPages[i].Text == text)
                {
                    return xtraTabControlStoresContent.TabPages[i];
                }
            return null;
        }
        public bool IsClear()
        {
            foreach (Control item in this.Controls["tableLayoutPanel1"].Controls)
            {
                if (item is TextBox)
                {
                    if (item.Text != "")
                        return false;
                }
                else if (item is ComboBox)
                {
                    if (item.Text != "")
                        return false;
                }
                else if (item is DateTimePicker)
                {
                    DateTimePicker item1 = (DateTimePicker)item;
                    if (item1.Value.Date != DateTime.Now.Date)
                        return false;
                }
            }
            return true;
        }
        public void displayData(string code)
        {
            string q1, q2, q3, q4,q5;
            if (txtType.Text == "")
            {
                q1 = "select Type_ID from data";
            }
            else
            {
                q1 = txtType.Text;
            }
            if (txtFactory.Text == "")
            {
                q2 = "select Factory_ID from data";
            }
            else
            {
                q2 = txtFactory.Text;
            }
            if (txtProduct.Text == "")
            {
                q3 = "select Product_ID from data";
            }
            else
            {
                q3 = txtProduct.Text;
            }
            if (txtGroup.Text == "")
            {
                q4 = "select Group_ID from data";
            }
            else
            {
                q4 = txtGroup.Text;
            }
            if (txtProduct.Text == "")
            {
                q5 = "";
            }
            else
            {
                q5 = "and  data.Size_ID=" + txtProduct.Text;
            }
            string q = "";
            if (code != "")
            {
                q = " and data.Code='" + code + "'";
            }
            else
            {
                q = "";
            }
            string itemName = "concat(type.Type_Name,' ',factory.Factory_Name,' ',groupo.Group_Name,' ' ,COALESCE(color.Color_Name,''),' ',COALESCE(size.Size_Value,''),' ',COALESCE(sort.Sort_Value,''),' ',COALESCE(data.Classification,''),' ',COALESCE(data.Description,''))as 'البند'";
            string query = "";
            if (txtType.Text == "1" || txtType.Text == "2" || txtType.Text == "9")
            {
               query = "SELECT data.Data_ID, data.Code as 'الكود', product.Product_Name as 'الصنف'," + itemName + ",data.Carton as 'الكرتنة' from data INNER JOIN type ON type.Type_ID = data.Type_ID INNER JOIN product ON product.Product_ID = data.Product_ID INNER JOIN factory ON data.Factory_ID = factory.Factory_ID INNER JOIN groupo ON data.Group_ID = groupo.Group_ID LEFT outer JOIN color ON data.Color_ID = color.Color_ID LEFT outer  JOIN size ON data.Size_ID = size.Size_ID LEFT outer  JOIN sort ON data.Sort_ID = sort.Sort_ID where  data.Type_ID IN(" + q1 + ") and  data.Factory_ID  IN(" + q2 + ") "+q5+" and data.Group_ID IN (" + q4 + ") " + q+ " order by SUBSTR(data.Code,1,16),color.Color_Name,data.Description,data.Sort_ID";
            }
            else
            {
               query = "SELECT data.Data_ID, data.Code as 'الكود', product.Product_Name as 'الصنف'," + itemName + ",data.Carton as 'الكرتنة' from data INNER JOIN type ON type.Type_ID = data.Type_ID INNER JOIN product ON product.Product_ID = data.Product_ID INNER JOIN factory ON data.Factory_ID = factory.Factory_ID INNER JOIN groupo ON data.Group_ID = groupo.Group_ID LEFT outer JOIN color ON data.Color_ID = color.Color_ID LEFT outer  JOIN size ON data.Size_ID = size.Size_ID LEFT outer  JOIN sort ON data.Sort_ID = sort.Sort_ID where  data.Type_ID IN(" + q1 + ") and  data.Factory_ID  IN(" + q2 + ") and  data.Product_ID  IN(" + q3 + ") and data.Group_ID IN (" + q4 + ") " + q+ " order by SUBSTR(data.Code,1,16),color.Color_Name,data.Description,data.Sort_ID";
            }
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            gridControl2.DataSource = dt;
            gridView2.Columns[0].Visible = false;
            gridView2.Columns[1].Width = 200;
            gridView2.BestFitColumns();
        }
        public void displayProducts()
        {
            string q1, q2, q3, q4, fQuery = "";
            if (txtType.Text == "")
            {
                q1 = "select Type_ID from data";
            }
            else
            {
                q1 = txtType.Text;
            }
            if (txtFactory.Text == "")
            {
                q2 = "select Factory_ID from data";
            }
            else
            {
                q2 = txtFactory.Text;
            }
            if (txtProduct.Text == "")
            {
                q3 = "select Product_ID from data";
            }
            else
            {
                q3 = txtProduct.Text;
            }
            if (txtGroup.Text == "")
            {
                q4 = "select Group_ID from data";
            }
            else
            {
                q4 = txtGroup.Text;
            }

            if (comSize.Text != "")
            {
                fQuery += "and size.Size_ID='" + comSize.SelectedValue + "'";
            }

            if (comColor.Text != "")
            {
                fQuery += "and color.Color_ID='" + comColor.SelectedValue + "'";
            }
            if (comSort.Text != "")
            {
                fQuery += "and Sort.Sort_ID='" + comSort.SelectedValue + "'";
            }
            if (comClassfication.Text != "")
            {
                fQuery += "and data.Classification='" + comClassfication.Text + "'";
            }

            string query = "SELECT data.Data_ID,data.Code as 'الكود',concat( product.Product_Name,' ',type.Type_Name,' ',factory.Factory_Name,' ',groupo.Group_Name,' ' ,COALESCE(size.Size_Value,''),COALESCE(data.Classification,''),COALESCE(data.Description,'') )as 'البند',sort.Sort_Value as 'الفرز',color.Color_Name as 'اللون',Old_Quantity as 'الكمية قبل الجرد' ,Current_Quantity as 'الكمية المجردة' from inventory inner join inventory_details on inventory.Inventory_ID=inventory_details.Inventory_ID INNER JOIN data on inventory_details.Data_ID=data.Data_ID INNER JOIN type ON type.Type_ID = data.Type_ID INNER JOIN product ON product.Product_ID = data.Product_ID INNER JOIN factory ON data.Factory_ID = factory.Factory_ID INNER JOIN groupo ON data.Group_ID = groupo.Group_ID LEFT outer JOIN color ON data.Color_ID = color.Color_ID LEFT outer  JOIN size ON data.Size_ID = size.Size_ID LEFT outer  JOIN sort ON data.Sort_ID = sort.Sort_ID where  data.Type_ID IN(" + q1 + ") and  data.Factory_ID  IN(" + q2 + ") and  data.Product_ID  IN(" + q3 + ") and data.Group_ID IN (" + q4 + ")  " + fQuery + " and Inventory_Num=" + labGardPermission.Text + " and Store_ID=" + txtStoreID.Text+" order by SUBSTR(data.Code,1,16),color.Color_Name,data.Description,data.Sort_ID ";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            gridControl2.DataSource = dt;
            gridView2.Columns[0].Visible = false;
            gridView2.Columns[1].Width = 140;
            fQuery = "";
        }

        public string getDataIDsWhichHaveQuantity()
        {
            dbconnection.Close();
            dbconnection.Open();
            string query = "select Data_ID from storage where Data_ID is not null and Store_ID="+comStore.SelectedValue+" and Store_Place_ID=";
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            MySqlDataReader dr = com.ExecuteReader();
            string DataIDs = "";
            while (dr.Read())
            {
                DataIDs += dr[0].ToString() + ",";
            }
            dr.Close();
            DataIDs += "0";
            dbconnection.Close();
            return DataIDs;
        }
        public bool validation(int storeID,int StorePlaceID)
        {
            string query = "select Storage_ID from storage where Data_ID="+Data_ID+" and Store_ID="+ storeID + " and Store_Place_ID="+ StorePlaceID;
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            if (com.ExecuteScalar() != null)
            {
                return false;
            }
            else
            {
                return true;
            }

        }
        public void displayCode(string code)
        {
            char[] arrCode = code.ToCharArray();
     
        }
        public void makeCode(TextBox txtBox)
        {
            if (code.Length < 20)
            {
                int j = 4 - txtBox.TextLength;
                for (int i = 0; i < j; i++)
                {
                    code += "0";
                }
                code += txtBox.Text;
            }
        }
        //public void add2Store()
        //{
        //    if ( code.Length==20)
        //    {
        //        string query = "select Data_ID from data where Code='" + code + "'";
        //        MySqlCommand com = new MySqlCommand(query, dbconnection);
        //        Data_ID = Convert.ToInt32(com.ExecuteScalar());
        //        if (validation((int)comStore.SelectedValue, (int)comStorePlace.SelectedValue))
        //        {
        //            query = "insert into open_storage_account (Data_ID,Quantity,Store_ID,Store_Place_ID,Date,Note) values (@Data_ID,@Quantity,@Store_ID,@Store_Place_ID,@Date,@Note)";
        //            com = new MySqlCommand(query, dbconnection);
        //            com.Parameters.Add("@Data_ID", MySqlDbType.Int16);
        //            com.Parameters["@Data_ID"].Value = Data_ID;
        //            com.Parameters.Add("@Quantity", MySqlDbType.Decimal);
        //            com.Parameters["@Quantity"].Value = txtTotalMeter.Text;
        //            com.Parameters.Add("@Store_ID", MySqlDbType.Int16);
        //            com.Parameters["@Store_ID"].Value = comStore.SelectedValue;
        //            com.Parameters.Add("@Store_Place_ID", MySqlDbType.Int16);
        //            com.Parameters["@Store_Place_ID"].Value = comStorePlace.SelectedValue;
        //            com.Parameters.Add("@Date", MySqlDbType.Date, 0);
        //            com.Parameters["@Date"].Value = dateTimePicker1.Value;
        //            com.Parameters.Add("@Note", MySqlDbType.VarChar);
        //            com.Parameters["@Note"].Value = txtNote.Text;
        //            com.ExecuteNonQuery();

        //            query = "insert into storage (Store_ID,Type,Storage_Date,Data_ID,Store_Place_ID,Total_Meters,Note) values (@Store_ID,@Type,@Date,@Data_ID,@PlaceOfStore,@TotalOfMeters,@Note)";
        //            com = new MySqlCommand(query, dbconnection);
        //            com.Parameters.Add("@Store_ID", MySqlDbType.Int16);
        //            com.Parameters["@Store_ID"].Value = comStore.SelectedValue;
        //            com.Parameters.Add("@Type", MySqlDbType.VarChar);
        //            com.Parameters["@Type"].Value = "بند";
        //            com.Parameters.Add("@Date", MySqlDbType.Date, 0);
        //            com.Parameters["@Date"].Value = dateTimePicker1.Value;
        //            com.Parameters.Add("@Data_ID", MySqlDbType.Int16);
        //            com.Parameters["@Data_ID"].Value = Data_ID;
        //            com.Parameters.Add("@PlaceOfStore", MySqlDbType.Int16);
        //            com.Parameters["@PlaceOfStore"].Value = comStorePlace.SelectedValue;
        //            com.Parameters.Add("@TotalOfMeters", MySqlDbType.Decimal);
        //            com.Parameters["@TotalOfMeters"].Value = txtTotalMeter.Text;
        //            com.Parameters.Add("@Note", MySqlDbType.VarChar);
        //            com.Parameters["@Note"].Value = txtNote.Text;
        //            com.ExecuteNonQuery();
        //            MessageBox.Show("Add success");
        //        }
        //        else
        //        {
        //            MessageBox.Show("هذا البند مضاف فعلا");
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("you must fill all fields please");
        //        dbconnection.Close();
        //        return;
        //    }

        //}
        public DataTable createDataTable()
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("Data_ID", typeof(int));
            dt.Columns.Add("Store_ID", typeof(int));
            dt.Columns.Add("Store_Place_ID", typeof(int));
            dt.Columns.Add("كود", typeof(string));
            dt.Columns.Add("البند", typeof(string));
            dt.Columns.Add("رصيد البند", typeof(double));
            dt.Columns.Add("ملاحظة", typeof(string));
            dt.Columns.Add("التاريخ", typeof(DateTime));
           
            return dt;
        }
        public void add2GridView(DataTable dt, DataRowView row, string [] re)
        {
            double quantity;
            DateTime date = new DateTime();
            if (re[0] == null)
            {
               // quantity = Convert.ToDouble(txtTotalMeter.Text);
            }
            else
            {
                quantity = Convert.ToDouble(re[0]);
            }
            if (re[1] == null)
            {
                date = dateTimePicker1.Value;
            }
            else
            {
                date = Convert.ToDateTime(re[1]);
            }
            //dt.Rows.Add(new object[] {
            //    Convert.ToInt32(row[0].ToString()),
            //    comStore.SelectedValue,
            //     getStore_Place_ID((int)comStore.SelectedValue),
            //    row[1].ToString(),
            //      row[2].ToString(),
            //    row[3].ToString(),
            //    quantity,
            //    txtNote.Text,
            //    date
            //});
         
            gridControl2.DataSource = dt;
            gridView2.Columns[0].Visible = false;
            gridView2.Columns[1].Visible = false;
            gridView2.Columns[2].Visible = false;
            gridView2.Columns[3].OptionsColumn.AllowEdit = false;
            gridView2.Columns[4].OptionsColumn.AllowEdit = false;
            gridView2.Columns[5].OptionsColumn.AllowEdit = false;
            gridView2.Columns[6].OptionsColumn.AllowEdit = true;
            gridView2.Columns[7].OptionsColumn.AllowEdit = true;
            gridView2.Columns[8].OptionsColumn.AllowEdit = true;
            gridView2.BestFitColumns();
        }
        public int getStore_Place_ID(int Store_ID)
        {
            dbconnection.Close();
            dbconnection.Open();
            string query = "select Store_Place_ID from store_places inner join store on store_places.Store_ID=store.Store_ID where store_places.Store_ID=" + Store_ID + " limit 1";
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            int Store_Place_ID = Convert.ToInt32(com.ExecuteScalar());

            return Store_Place_ID;
        }
        public string [] getOldValueIfExist(int Data_ID)
        {
            string [] re =new string[2];
            bool f = false;
            string query = "select Quantity,Date from open_storage_account where Data_ID=" + Data_ID+" and Store_ID="+comStore.SelectedValue;
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            MySqlDataReader dr = com.ExecuteReader();
           while(dr.Read())
            { 
                re[0] =dr[0].ToString();
                re[1] = dr[1].ToString();
                f = true;
                ListOfDataIDs.Add(Data_ID);
            }
           if(f)
            MessageBox.Show("البند مضاف من قبل");           
            return re;
        }
        public bool IsEdited(int DataID)
        {
            for (int i = 0; i < ListOfEditDataIDs.Count; i++)
            {
                if (DataID == ListOfEditDataIDs[i])
                    return true;
            }
            return false;
        }
        public bool IsSaved(int DataID)
        {
            for (int i = 0; i < ListOfSavedDataIDs.Count; i++)
            {
                if (DataID == ListOfSavedDataIDs[i])
                    return true;
            }
            return false;
        }
        public bool IsExistInGridView(string c)
        {
            for (int i = 0; i < gridView2.DataRowCount; i++)
            {
                DataRowView r = (DataRowView)gridView2.GetRow(i);
                if (r[3].ToString()==c)
                {
                    // gridView2.MoveBy(i);
                    int x = gridView2.GetSelectedRows()[0];
                    if(x!=0)
                        gridView2.FocusedRowHandle = 0;
                    else
                        gridView2.FocusedRowHandle = gridView2.DataRowCount-1;
                    gridView2.FocusedRowHandle = i;
                    ListOfDataIDs.Add((int)r[0]);
                    return true;
                }
            }
            return false;
        }
        //public bool validation()
        //{
        //    if (code.Length != 20)
        //        labcode.Visible = true;
        //    else
        //        labcode.Visible = false;
        //    if (txtTotalMeter.Text == "")
        //        labQuantity.Visible = true;
        //    else
        //        labQuantity.Visible = false;
        //    if (comStore.Text == "")
        //        labStore.Visible = true;
        //    else
        //        labStore.Visible = false;
        //    if (comStorePlace.Text == "")
        //        labStorePlace.Visible = true;
        //    else
        //        labStorePlace.Visible = false;
        //    if (code.Length == 20 && txtTotalMeter.Text != "" && comStore.Text != "" && comStorePlace.Text != "")
        //    {
        //        return true;
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //}
        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }
        public void filterFactory()
        {
            if (comType.Text != "")
            {
                string query = "select * from factory inner join type_factory on factory.Factory_ID=type_factory.Factory_ID inner join type on type.Type_ID=type_factory.Type_ID  where type_factory.Type_ID=" + comType.SelectedValue;
                MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                comFactory.DataSource = dt;
                comFactory.DisplayMember = dt.Columns["Factory_Name"].ToString();
                comFactory.ValueMember = dt.Columns["Factory_ID"].ToString();
                comFactory.Text = "";
                txtFactory.Text = "";
            }
        }
        public void filterGroup()
        {
            string query = "select TypeCoding_Method from type where Type_ID=" + txtType.Text;
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            int TypeCoding_Method = (int)com.ExecuteScalar();
            if (TypeCoding_Method == 1)
            {
                string query2 = "";
                if (txtType.Text == "2" || txtType.Text == "1")
                    query2 = "select * from groupo where Factory_ID=" + -1;
                else
                    query2 = "select * from groupo where Factory_ID=" + -Convert.ToInt32(txtType.Text) + " and Type_ID=" + txtType.Text;

                MySqlDataAdapter da2 = new MySqlDataAdapter(query2, dbconnection);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                comGroup.DataSource = dt2;
                comGroup.DisplayMember = dt2.Columns["Group_Name"].ToString();
                comGroup.ValueMember = dt2.Columns["Group_ID"].ToString();
                comGroup.Text = "";
                txtGroup.Text = "";
            }
            else
            {
                string q = "";
                if (txtFactory.Text != "")
                {
                    q = " and Factory_ID = " + txtFactory.Text;
                }
                query = "select * from groupo where Type_ID=" + txtType.Text + q;
                MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                comGroup.DataSource = dt;
                comGroup.DisplayMember = dt.Columns["Group_Name"].ToString();
                comGroup.ValueMember = dt.Columns["Group_ID"].ToString();
                comGroup.Text = "";
                txtGroup.Text = "";
            }
        }
        public void filterProduct()
        {
            if (comType.Text != "")
            {
                if (comGroup.Text != "" || comFactory.Text != "" || comType.Text != "")
                {
                    if (txtType.Text != "1" && txtType.Text != "2" && txtType.Text != "9")
                    {
                        string supQuery = "";

                        supQuery = " product.Type_ID=" + txtType.Text + "";
                        if (comFactory.Text != "")
                        {
                            supQuery += " and product_factory_group.Factory_ID=" + txtFactory.Text + "";
                        }
                        if (comGroup.Text != "")
                        {
                            supQuery += " and product_factory_group.Group_ID=" + txtGroup.Text + "";
                        }
                        string query = "select distinct  product.Product_ID  ,Product_Name  from product inner join product_factory_group on product.Product_ID=product_factory_group.Product_ID  where  " + supQuery + "   order by product.Product_ID";
                        MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        comProduct.DataSource = dt;
                        comProduct.DisplayMember = dt.Columns["Product_Name"].ToString();
                        comProduct.ValueMember = dt.Columns["Product_ID"].ToString();
                        comProduct.Text = "";
                        txtProduct.Text = "";
                    }
                    else
                    {
                        string supQuery = "";
                        if (comFactory.Text != "")
                        {
                            supQuery += " and type_factory.Factory_ID=" + txtFactory.Text + "";
                        }
                        if (comGroup.Text != "")
                        {
                            supQuery += " and Group_ID=" + txtGroup.Text + "";
                        }
                        string query = "select * from size inner join type_factory on size.Factory_ID=type_factory.Factory_ID where type_factory.Type_ID="+txtType.Text + supQuery;
                        MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        comProduct.DataSource = dt;
                        comProduct.DisplayMember = dt.Columns["Size_Value"].ToString();
                        comProduct.ValueMember = dt.Columns["Size_ID"].ToString();
                        comProduct.Text = "";
                        txtProduct.Text = "";
                    }
                }
            }

        }
        public void filterSize()
        {
            if (comFactory.Text != "")
            {
                string query = "select * from size size inner join type_factory on size.Factory_ID=type_factory.Factory_ID where Type_ID=" + txtType.Text+" Factory_ID=" + comFactory.SelectedValue;
                MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                comProduct.DataSource = dt;
                comProduct.DisplayMember = dt.Columns["Size_Value"].ToString();
                comProduct.ValueMember = dt.Columns["Size_ID"].ToString();
                comProduct.Text = "";
                txtProduct.Text = "";
            }
        }
        public bool AllDataIsSaved()
        {
            if ((ListOfSavedDataIDs.Count - ListOfEditDataIDs.Count) != mdt.Rows.Count)
            {
                for (int i = 0; i < mdt.Rows.Count; i++)
                {
                    if (IsEdited((int)mdt.Rows[i][0]) || !IsSaved((int)mdt.Rows[i][0]))
                    {
                        return false;
                    }
                }
            }
            return true;
        }




        public void addGardQuantity(int Data_ID, DataRow row)
        {
            dbconnection.Close();
            dbconnection.Open();

            string query = "select Inventory_ID from inventory where Inventory_Num=" + labGardPermission.Text;
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            int Inventory_ID = Convert.ToInt16(com.ExecuteScalar());
            if (haveOpenStorageAccount(Data_ID))
            {
                query = "update  inventory_details set Current_Quantity=@Current_Quantity,Updated=@Updated where Data_ID=" + Data_ID + " and Inventory_ID=" + Inventory_ID;
                com = new MySqlCommand(query, dbconnection);
                com.Parameters.Add("@Current_Quantity", MySqlDbType.Decimal);
                com.Parameters["@Current_Quantity"].Value = row[6];
                com.Parameters.Add("@Updated", MySqlDbType.Int16);
                com.Parameters["@Updated"].Value = 1;
                com.ExecuteNonQuery();

                query = "update storage set Type=@Type,Storage_Date=@Storage_Date,Total_Meters=@Total_Meters,Note=@Note where Data_ID=" + Data_ID + " and Store_ID=" +txtStoreID.Text;
                com = new MySqlCommand(query, dbconnection);
                com.Parameters.Add("@Type", MySqlDbType.VarChar);
                com.Parameters["@Type"].Value = "بند";
                com.Parameters.Add("@Storage_Date", MySqlDbType.Date, 0);
                DateTime date = Convert.ToDateTime(dateTimePicker1.Text);
                string d = date.ToString("yyyy-MM-dd");
                com.Parameters["@Storage_Date"].Value = d;
                com.Parameters.Add("@Total_Meters", MySqlDbType.Decimal);
                com.Parameters["@Total_Meters"].Value = row[6];
                com.Parameters.Add("@Note", MySqlDbType.VarChar);
                com.Parameters["@Note"].Value = "تعديل جرد";
                com.ExecuteNonQuery();

                string value = "";
                UserControl.ItemRecord("inventory_details", "تعديل", (int)row[0], DateTime.Now, value, dbconnection);

            }
            else
            {
                //save to open storage account with inital value 0
                query = "insert into open_storage_account (Data_ID,Quantity,Store_ID,Store_Place_ID,Date,Note) values (@Data_ID,@Quantity,@Store_ID,@Store_Place_ID,@Date,@Note)";
                com = new MySqlCommand(query, dbconnection);
                com.Parameters.Add("@Data_ID", MySqlDbType.Int16);
                com.Parameters["@Data_ID"].Value = row[0];
                com.Parameters.Add("@Quantity", MySqlDbType.Decimal);
                com.Parameters["@Quantity"].Value = row[6];
                com.Parameters.Add("@Store_ID", MySqlDbType.Int16);
                com.Parameters["@Store_ID"].Value =Convert.ToInt16(txtStoreID.Text);
                com.Parameters.Add("@Store_Place_ID", MySqlDbType.Int16);
                com.Parameters["@Store_Place_ID"].Value = getStore_Place_ID(Convert.ToInt16(txtStoreID.Text));
                com.Parameters.Add("@Date", MySqlDbType.Date, 0);
                DateTime date = Convert.ToDateTime(dateTimePicker1.Text);
                string d = date.ToString("yyyy-MM-dd");
                com.Parameters["@Date"].Value = d;
                com.Parameters.Add("@Note", MySqlDbType.VarChar);
                com.Parameters["@Note"].Value ="جرد";
                com.ExecuteNonQuery();

                //save to storage with gard value
                query = "insert into storage (Store_ID,Type,Storage_Date,Data_ID,Store_Place_ID,Total_Meters,Note) values (@Store_ID,@Type,@Date,@Data_ID,@PlaceOfStore,@TotalOfMeters,@Note)";
                com = new MySqlCommand(query, dbconnection);
                com.Parameters.Add("@Store_ID", MySqlDbType.Int16);
                com.Parameters["@Store_ID"].Value = txtStoreID.Text;
                com.Parameters.Add("@Type", MySqlDbType.VarChar);
                com.Parameters["@Type"].Value = "بند";
                com.Parameters.Add("@Date", MySqlDbType.Date, 0);
                date = Convert.ToDateTime(dateTimePicker1.Text);
                d = date.ToString("yyyy-MM-dd");
                com.Parameters["@Date"].Value = d;
                com.Parameters.Add("@Data_ID", MySqlDbType.Int16);
                com.Parameters["@Data_ID"].Value = row[0];
                com.Parameters.Add("@PlaceOfStore", MySqlDbType.Int16);
                com.Parameters["@PlaceOfStore"].Value = getStore_Place_ID(Convert.ToInt16(txtStoreID.Text));
                com.Parameters.Add("@TotalOfMeters", MySqlDbType.Decimal);
                com.Parameters["@TotalOfMeters"].Value = row[6];
                com.Parameters.Add("@Note", MySqlDbType.VarChar);
                com.Parameters["@Note"].Value = "تعديل جرد";
                com.ExecuteNonQuery();

                //update inventory
                query = "update  inventory_details set Current_Quantity=@Current_Quantity,Updated=@Updated where Data_ID=" + Data_ID + " and Inventory_ID=" + Inventory_ID;
                com = new MySqlCommand(query, dbconnection);
                com.Parameters.Add("@Current_Quantity", MySqlDbType.Decimal);
                com.Parameters["@Current_Quantity"].Value = row[6];
                com.Parameters.Add("@Updated", MySqlDbType.Int16);
                com.Parameters["@Updated"].Value = 1;
                com.ExecuteNonQuery();

                UserControl.ItemRecord("inventory_details", "تعديل", (int)row[0], DateTime.Now, "", dbconnection);

            }
            dbconnection.Close();
        }
        public bool haveOpenStorageAccount(int Data_ID)
        {
            string query = "select OpenStorageAccount_ID from open_storage_account where Data_ID="+Data_ID;
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            if (com.ExecuteScalar() != null)
            {
                return true;
            }
            return false;
        }
    }

}
