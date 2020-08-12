using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
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
    public partial class ProductItems : Form
    {
        MySqlConnection dbconnection;
        DataGridViewRow row1 = null;
        int id = -1;
        int typeUpdateId = -1;
        int groupUpdateId = -1;
        int productUpdateId = -1;
        int colorUpdateId = -1;
        int SortUpdateId = -1;
        int SizeUpdateId = -1;
        bool load = false;
        bool flagFactory = false;//in group tap
        bool flagFactoryP = false;//in product tap
        bool flagGroup = false;//in product tap
        bool flagItemCheckGroup = true;

        List<factory_Group> listFactory_Group = new List<factory_Group>();
        AutoCompleteStringCollection collection = new AutoCompleteStringCollection();
        public ProductItems()
        {
            try
            {         
                InitializeComponent();
                dbconnection = new MySqlConnection(connection.connectionString);
                mTC_Content.SelectedIndex = 6;
                displayType();
                txtType.Focus();
                DeletePermission();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        //header tabs
        private void btnSwitchTab_Click(object sender, EventArgs e)
        {
            try
            {
                Button btn = (Button)sender;
                switch (btn.Name)
                {
                    case "btnType":
                        mTC_Content.SelectedIndex = 6;
                        displayType();
                        txtType.Focus();
                        break;

                    case "btnFactory":
                        mTC_Content.SelectedIndex = 5;                       
                        txtFactory.Focus();
                        dbconnection.Open();
                        string query = "select * from type";
                        MySqlCommand com = new MySqlCommand(query, dbconnection);
                        MySqlDataReader dr = com.ExecuteReader();
                        checkedListBox1.Items.Clear();
                        while (dr.Read())
                            checkedListBox1.Items.Add(dr["Type_Name"].ToString() + "\t" + dr["Type_ID"]);

                        dr.Close();
                        break;

                    case "btnGroup":
                        mTC_Content.SelectedIndex = 4;                       
                        comType.Focus();
                        comType.Text = "";
                        comFactory.Text = "";
                        txtType1.Text = "";
                        txtFactory1.Text = "";
                        txtGroup.Text = "";                      
                        break;

                    case "btnProduct":
                        mTC_Content.SelectedIndex = 3;
                        comTypeProduct.Focus();
                        comTypeProduct.Text = "";
                        comFactoryGroup.Text = "";
                        comGroup.Text = "";
                        txtType2.Text = "";
                        txtFactory2.Text = "";
                        txtGroup2.Text = "";
                        break;
                    case "btnColor":
                        mTC_Content.SelectedIndex = 2;
                        displayColor();
                        comType2.Focus();
                        break;
                    case "btnSize":
                        mTC_Content.SelectedIndex = 1;
                        displaySize();
                        comFactory.Focus();
                        break;
                    case "btnSort":
                        mTC_Content.SelectedIndex = 0;
                        displaySort();
                        txtSort.Focus();
                        break;
                       
                }
                ClearButtonsColor();
                btn.BackColor = Color.FromArgb(194,192,192);
         
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void ProductItems_Load(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string query = "select * from type";       
                MySqlCommand com = new MySqlCommand(query, dbconnection);
                MySqlDataReader dr = com.ExecuteReader();
                while (dr.Read())
                    checkedListBox1.Items.Add(dr["Type_Name"].ToString() + "\t" + dr["Type_ID"]);

                dr.Close();

                MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                comType.DataSource = dt;
                comType.DisplayMember = dt.Columns["Type_Name"].ToString();
                comType.ValueMember = dt.Columns["Type_ID"].ToString();
                comType.Text = "";
               
                comType2.DataSource = dt;
                comType2.DisplayMember = dt.Columns["Type_Name"].ToString();
                comType2.ValueMember = dt.Columns["Type_ID"].ToString();
                comType2.Text = "";

                comTypeProduct.DataSource = dt;
                comTypeProduct.DisplayMember = dt.Columns["Type_Name"].ToString();
                comTypeProduct.ValueMember = dt.Columns["Type_ID"].ToString();
                comTypeProduct.Text = "";

                query = "select distinct * from factory";
                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                //comFactory.DataSource = dt;
                //comFactory.DisplayMember = dt.Columns["Factory_Name"].ToString();
                //comFactory.ValueMember = dt.Columns["Factory_ID"].ToString();
                //comFactory.Text = "";

                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                comFactory2.DataSource = dt;
                comFactory2.DisplayMember = dt.Columns["Factory_Name"].ToString();
                comFactory2.ValueMember = dt.Columns["Factory_ID"].ToString();
                comFactory2.Text = "";

                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                //comFactoryGroup.DataSource = dt;
                //comFactoryGroup.DisplayMember = dt.Columns["Factory_Name"].ToString();
                //comFactoryGroup.ValueMember = dt.Columns["Factory_ID"].ToString();
                //comFactoryGroup.Text = "";
                query = "select distinct * from groupo";
                da = new MySqlDataAdapter(query, dbconnection);
                dt = new DataTable();
                da.Fill(dt);
                comGroup.DataSource = dt;
                comGroup.DisplayMember = dt.Columns["Group_Name"].ToString();
                comGroup.ValueMember = dt.Columns["Group_ID"].ToString();
                comGroup.Text = "";

                load = true;

                dataGridViewStyle();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                row1 = null;
                RestControls();
                switch (mTC_Content.SelectedIndex)
                {
                    case 6 :
                        ClearButtonsColor();
                        btnType.BackColor = Color.FromArgb(194, 192, 192);
                        txtType.Focus();
                        break;
                    case 5:
                        ClearButtonsColor();
                        btnFactory.BackColor = Color.FromArgb(194, 192, 192);
                        txtFactory.Focus();
                        break;
                    case 4:
                        ClearButtonsColor();
                        btnGroup.BackColor = Color.FromArgb(194, 192, 192);
                        comFactory.Focus();
                        break;
                    case 3:
                        ClearButtonsColor();
                        btnProduct.BackColor = Color.FromArgb(194, 192, 192);
                        comFactoryGroup.Focus();
                        break;
                    case 2:
                        ClearButtonsColor();
                        btnColor.BackColor = Color.FromArgb(194, 192, 192);
                        comType2.Focus();
                        break;
                    case 1:
                        ClearButtonsColor();
                        btnSize.BackColor = Color.FromArgb(194, 192, 192);
                        comFactory2.Focus();
                        break;
                    case 0:
                        ClearButtonsColor();
                        btnSort.BackColor = Color.FromArgb(194, 192, 192);
                        txtSort.Focus();
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void comBox_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Close();
                dbconnection.Open();
                if (load)
                {
                    ComboBox comBox = (ComboBox)sender;

                    switch (comBox.Name)
                    {
                        case "comType":
                            txtType1.Text = comType.SelectedValue.ToString();
                            txtFactory1.Text = "";
                            txtFactory1.Focus();
                            break;

                        case "comTypeProduct":
                            txtType2.Text = comTypeProduct.SelectedValue.ToString();
                            break;

                        case "comType2":
                            txtType3.Text = comType2.SelectedValue.ToString();
                            break;

                        case "comFactory":
                            if (flagFactory)
                            {
                                txtFactory1.Text = comFactory.SelectedValue.ToString();
                                txtGroup.Focus();
                            }                  
                            break;

                        case "comFactoryGroup":
                            if (flagFactoryP)
                            {
                                txtFactory2.Text = comFactoryGroup.SelectedValue.ToString();
                            }
                            break;

                        case "comFactory2":
                            txtFactory3.Text = comFactory2.SelectedValue.ToString();
                            break;

                        case "comGroup":
                            txtGroup2.Text = comGroup.SelectedValue.ToString();                          
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
         
            if (e.KeyCode == Keys.Enter)
            {
                TextBox txtBox = (TextBox)sender;
                string query;
                MySqlCommand com;
                string Name;
                try
                {
                    if (txtBox.Text != "")
                    {
                        dbconnection.Open();
                        switch (txtBox.Name)
                        {
                            case "txtType1":
                                query = "select Type_Name from type where Type_ID='" + txtType1.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    displayGroup_Type(Convert.ToInt32(txtType1.Text));
                                    Name = (string)com.ExecuteScalar();
                                    comType.Text = Name;
                                    if (txtType1.Text == "1")
                                        txtGroup.Focus();
                                    txtFactory1.Focus();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtType2":
                                query = "select Type_Name from type where Type_ID='" + txtType2.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comTypeProduct.Text = Name;

                                    txtFactory2.Focus();
                                    dataGridViewProduct.DataSource = displayProduct_Type(Convert.ToInt32(txtType2.Text));
                                    dataGridViewProduct.Columns[0].Width = 50;
                                  
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtType3":
                                query = "select Type_Name from type where Type_ID='" + txtType3.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comType2.Text = Name;

                                    txtColor.Focus();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;

                            case "txtFactory1":
                                query = "select Factory_Name from factory where Factory_ID='" + txtFactory1.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    displayGroup_Factory(Convert.ToInt32(txtFactory1.Text));
                                    Name = (string)com.ExecuteScalar();
                                    comFactory.Text = Name;
                                  
                                    txtGroup.Focus();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtFactory2":
                                query = "select Factory_Name from factory where Factory_ID='" + txtFactory2.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comFactoryGroup.Text = Name;

                                    txtGroup2.Focus();
                                    dataGridViewProduct.DataSource = displayProduct_Factory(Convert.ToInt32(txtFactory2.Text));
                                    dataGridViewProduct.Columns[0].Width = 50;
                                 
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtFactory3":
                                query = "select Factory_Name from factory where Factory_ID='" + txtFactory3.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comFactory2.Text = Name;

                                    txtSize.Focus();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;

                            case "txtGroup2":
                                query = "select Group_Name from groupo where Group_ID='" + txtGroup2.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comGroup.Text = Name;

                                    txtProduct.Focus();
                                    displayProduct_Group(Convert.ToInt32(txtGroup2.Text));
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtGroup_Size":
                                query = "select Group_Name from groupo where Group_ID='" + txtGroup_Size.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comGroup_Size.Text = Name;

                                    txtSize.Focus();
                                    displaySize_group();
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
                catch
                {
                    //  MessageBox.Show(ex.ToString());
                }
                dbconnection.Close();
            }
        }
        private void txt_TextChanged(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Close();
                dbconnection.Open();
                TextBox txtbox = (TextBox)sender;
                MySqlDataAdapter adapter;
                DataTable dtaple;
                string query = "";
                if(txtbox.Text!="")
                switch (txtbox.Name)
                {
                    case "txtType":
                        query = "select Type_ID as 'كود',Type_Name as 'النوع' from type where Type_Name like'" + txtType.Text + "%'";
                        adapter = new MySqlDataAdapter(query, dbconnection);
                        dtaple = new DataTable();
                        adapter.Fill(dtaple);
                        dataGridView1.DataSource = dtaple;
                        break;
                    case "txtType1":
                        query = "select Type_Name from type where Type_ID='" + txtType1.Text + "'";
                        MySqlCommand com = new MySqlCommand(query, dbconnection);
                        if (com.ExecuteScalar() != null)
                        {
                            displayGroup_Type(Convert.ToInt32(txtType1.Text));
                            Name = (string)com.ExecuteScalar();
                            comType.Text = Name;
                            if (txtType1.Text == "1")
                                txtGroup.Focus();
                            txtFactory1.Focus();
                        }
                        else
                        {
                            MessageBox.Show("there is no item with this id");
                            dbconnection.Close();
                            return;
                        }
                        break;
                    case "txtType2":
                        query = "select Type_Name from type where Type_ID='" + txtType2.Text + "'";
                        MySqlCommand com1 = new MySqlCommand(query, dbconnection);
                        if (com1.ExecuteScalar() != null)
                        {
                            Name = (string)com1.ExecuteScalar();
                            comTypeProduct.Text = Name;

                            txtFactory2.Focus();
                            dataGridViewProduct.DataSource = displayProduct_Type(Convert.ToInt32(txtType2.Text));
                            dataGridViewProduct.Columns[0].Width = 50;

                        }
                        else
                        {
                            MessageBox.Show("there is no item with this id");
                            dbconnection.Close();
                            return;
                        }
                        break;
                        case "txtType3":
                            query = "select Type_Name from type where Type_ID='" + txtType3.Text + "'";
                            MySqlCommand com3 = new MySqlCommand(query, dbconnection);
                            if (com3.ExecuteScalar() != null)
                            {
                                displayColor(Convert.ToInt32(txtType3.Text));
                                Name = (string)com3.ExecuteScalar();
                                comType2.Text = Name;
                                txtColor.Focus();
                              
                            }
                            else
                            {
                                MessageBox.Show("there is no item with this id");
                                dbconnection.Close();
                                return;
                            }
                            break;
                        case "txtFactory":
                        query = "select distinct factory.Factory_ID as 'كود',Factory_Name as 'المصنع' from factory inner join type_factory on factory.Factory_ID=type_factory.Factory_ID inner join type  on type.Type_ID=type_factory.Type_ID where Factory_Name like'" + txtFactory.Text + "%'";
                        adapter = new MySqlDataAdapter(query, dbconnection);
                        dtaple = new DataTable();
                        adapter.Fill(dtaple);
                        dataGridViewFactory.DataSource = dtaple;
                        break;
                     case "txtFactory1":
                            query = "select Factory_Name from factory where Factory_ID='" + txtFactory1.Text + "'";
                            com = new MySqlCommand(query, dbconnection);
                            if (com.ExecuteScalar() != null)
                            {
                                displayGroup_Factory(Convert.ToInt32(txtFactory1.Text));
                                Name = (string)com.ExecuteScalar();
                                comFactory.Text = Name;

                                txtGroup.Focus();
                            }
                            else
                            {
                                MessageBox.Show("there is no item with this id");
                                dbconnection.Close();
                                return;
                            }
                            break;
                        case "txtFactory2":
                        query = "select Factory_Name from factory where Factory_ID='" + txtFactory2.Text + "'";
                        com = new MySqlCommand(query, dbconnection);
                        if (com.ExecuteScalar() != null)
                        {
                            Name = (string)com.ExecuteScalar();
                            comFactoryGroup.Text = Name;

                            txtGroup2.Focus();
                            dataGridViewProduct.DataSource = displayProduct_Factory(Convert.ToInt32(txtFactory2.Text));
                            dataGridViewProduct.Columns[0].Width = 50;

                        }
                        else
                        {
                            MessageBox.Show("there is no item with this id");
                            dbconnection.Close();
                            return;
                        }
                        break;
                        case "txtFactory3":
                            query = "select Factory_Name from factory where Factory_ID='" + txtFactory3.Text + "'";
                            com = new MySqlCommand(query, dbconnection);
                            if (com.ExecuteScalar() != null)
                            {
                                Name = (string)com.ExecuteScalar();
                                comFactory2.Text = Name;
                                txtGroup_Size.Focus();

                            }
                            else
                            {
                                MessageBox.Show("there is no item with this id");
                                dbconnection.Close();
                                return;
                            }
                            break;
                        case "txtGroup":
                        if (comFactory2.Text != "")
                        {
                            query = "select Group_ID as 'كود',Group_Name as 'المجموعة' from Groupo where Group_Name like'" + txtGroup.Text + "%' and Factory_ID=" + comFactory2.SelectedValue;
                            adapter = new MySqlDataAdapter(query, dbconnection);
                            dtaple = new DataTable();
                            adapter.Fill(dtaple);
                            dataGridViewGroup.DataSource = dtaple;
                        }
                        break;
                    case "txtGroup2":
                            query = "select Group_Name from groupo where Group_ID='" + txtGroup2.Text + "'";
                            com = new MySqlCommand(query, dbconnection);
                            if (com.ExecuteScalar() != null)
                            {
                                Name = (string)com.ExecuteScalar();
                                comGroup.Text = Name;

                                txtProduct.Focus();
                                displayProduct_Group(Convert.ToInt32(txtGroup2.Text));
                            }
                            else
                            {
                                MessageBox.Show("there is no item with this id");
                                dbconnection.Close();
                                return;
                            }
                            break;
                        case "txtGroup_Size":
                            query = "select Group_Name from groupo where Group_ID='" + txtGroup_Size.Text + "'";
                            com = new MySqlCommand(query, dbconnection);
                            if (com.ExecuteScalar() != null)
                            {
                                Name = (string)com.ExecuteScalar();
                                comGroup_Size.Text = Name;

                                txtSize.Focus();
                            }
                            else
                            {
                                MessageBox.Show("there is no item with this id");
                                dbconnection.Close();
                                return;
                            }
                            break;
                        case "txtProduct":
                            if (comFactoryGroup.Text != "")
                            {
                                query = "select Product_ID as 'كود',Product_Name as 'الصنف' from Product where Product_Name like'" + txtProduct.Text + "%'";
                                MySqlCommand comand = new MySqlCommand(query, dbconnection);
                                MySqlDataReader dr = comand.ExecuteReader();
                                while (dr.Read())
                                {
                                    collection.Add(dr["الصنف"].ToString());
                                }
                                dr.Close();
                                //txtProduct.AutoCompleteMode = AutoCompleteMode.Suggest;
                                //txtProduct.AutoCompleteSource = AutoCompleteSource.CustomSource;
                                //txtProduct.AutoCompleteCustomSource = collection;
                             }
                        break;
                    case "txtSort":
                        query = "select Sort_ID as 'كود',Sort_Value as 'الأسم' from Sort where Sort_Value like'" + txtSort.Text + "%'";
                        adapter = new MySqlDataAdapter(query, dbconnection);
                        dtaple = new DataTable();
                        adapter.Fill(dtaple);
                        dataGridViewSort.DataSource = dtaple;
                        break;
                    case "txtColor":
                            if (comType.Text != "")
                            {
                                query = "select Color_ID as 'كود',Color_Name as 'اللون' from Color where Color_Name like'" + txtColor.Text + "%' and Type_ID=" + comType2.SelectedValue;
                                adapter = new MySqlDataAdapter(query, dbconnection);
                                dtaple = new DataTable();
                                adapter.Fill(dtaple);
                                dataGridViewColor.DataSource = dtaple;
                            }
                            break;
                    case "txtSize":
                        if (comFactory.Text != "")
                        {
                            query = "select Size_ID as 'كود',Size_Value as 'المقاس' from Size where Size_Value like'" + txtSize.Text + "%' and Factory_ID=" + comFactory2.SelectedValue;
                            adapter = new MySqlDataAdapter(query, dbconnection);
                            dtaple = new DataTable();
                            adapter.Fill(dtaple);
                            dataGridViewSize.DataSource = dtaple;
                        }
                        break;
                }
            }
            catch
            {
               // MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }

        #region type panel
        private void btnType_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string query = "select Type_ID from type where Type_Name = '" + txtType.Text + "'";
                MySqlCommand com = new MySqlCommand(query, dbconnection);
                if (com.ExecuteScalar() == null)
                {
                    if (txtType.Text != "")
                    {
                        query = "insert into type (Type_Name,TypeCoding_Method) values (@name,@TypeCoding_Method)";
                        com = new MySqlCommand(query, dbconnection);
                        com.Parameters.AddWithValue("@name", txtType.Text);

                        if(Method1.Checked)
                            com.Parameters.AddWithValue("@TypeCoding_Method", 1);
                        else
                            com.Parameters.AddWithValue("@TypeCoding_Method", 2);

                        com.ExecuteNonQuery();
                        
                        query = "select Type_ID from type order by Type_ID desc limit 1";
                        com = new MySqlCommand(query, dbconnection);
                        
                        UserControl.ItemRecord("type","اضافة",Convert.ToInt32(com.ExecuteScalar().ToString()), DateTime.Now,"",dbconnection);
                      
                        displayType();
                        txtType.Text = "";
                    }
                    else
                    {
                        txtType.Focus();           
                    }
                }
                else
                {
                    MessageBox.Show("This type already exist");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }   
        private void txtType_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter )
                {
                    if (btnSaveUpdateType.Visible == false)
                    {
                        dbconnection.Open();
                        string query = "select Type_ID from type where Type_Name = '" + txtType.Text + "'";
                        MySqlCommand com = new MySqlCommand(query, dbconnection);
                        if (com.ExecuteScalar() == null)
                        {
                            if (txtType.Text != "")
                            {
                                query = "insert into type (Type_Name,TypeCoding_Method) values (@name,@TypeCoding_Method)";
                                com = new MySqlCommand(query, dbconnection);
                                com.Parameters.AddWithValue("@name", txtType.Text);
                                if (Method1.Checked)
                                    com.Parameters.AddWithValue("@TypeCoding_Method", 1);
                                else
                                    com.Parameters.AddWithValue("@TypeCoding_Method", 2);

                                com.ExecuteNonQuery();

                                query = "select Type_ID from type order by Type_ID desc limit 1";
                                com = new MySqlCommand(query, dbconnection);

                                UserControl.ItemRecord("type", "add", Convert.ToInt32(com.ExecuteScalar().ToString()), DateTime.Now, "", dbconnection);

                                displayType();
                                txtType.Text = "";
                            }
                            else
                            {
                                txtType.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("This type already exist");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnUpdateType_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                btnSaveUpdateType.Visible = true;
                row1 = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];
                string Type_Name="";
                if (row1 != null)
                {
                   
                    typeUpdateId = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    string query = "select Type_Name,TypeCoding_Method from type where Type_ID=" + typeUpdateId;
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    MySqlDataReader dr = com.ExecuteReader();
                    while (dr.Read())
                    {
                        string dd = dr[1].ToString();
                        Type_Name = dr[0].ToString();
                        if (dd == "1")
                        {
                            Method1.Checked = true;
                        }
                        else
                        {
                            Method2.Checked = true;
                        }
                    }
                    dr.Close();
                    txtType.Text = Type_Name;
                }
            }
            catch
            {
                MessageBox.Show("اختار النوع المراد تعديله.");
            }
            dbconnection.Close();
        }
        private void btnSaveUpdateType_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                if (typeUpdateId != -1)
                {
                    string query = "update type set Type_Name=@name ,TypeCoding_Method=@TypeCoding_Method where Type_ID=" + typeUpdateId;
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    com.Parameters.AddWithValue("@name", txtType.Text);
                    if (Method1.Checked)
                        com.Parameters.AddWithValue("@TypeCoding_Method", 1);
                    else
                        com.Parameters.AddWithValue("@TypeCoding_Method", 2);
                    com.ExecuteNonQuery();
                    displayType();
                    btnSaveUpdateType.Visible = false;
                    txtType.Text = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnTypeDisplayAll_Click(object sender, EventArgs e)
        {
            try
            {
                displayType();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnDeleteType_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();

                DataGridViewRow storeRow = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];

                if (storeRow != null)
                {
                    int id = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    if (!producItemUsed("Type_ID", id))
                    {
                        DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete the item?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.Yes)
                        {
                            string query = "delete from type where Type_ID=" + id;
                            MySqlCommand com = new MySqlCommand(query, dbconnection);
                            com.ExecuteNonQuery();
                            updateTablesDB("type", "Type_ID", id);

                            UserControl.ItemRecord("type", "حذف", Convert.ToInt32(storeRow.Cells[0].Value.ToString()), DateTime.Now, "", dbconnection);

                            displayType();
                            txtType.Focus();
                        }
                        else if (dialogResult == DialogResult.No)
                        { }
                    }
                    else
                    {
                        MessageBox.Show("هذا العنصر تم استخدامه في تسجيل بنود لا يمكن حذفه.");
                    }
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        #endregion

        #region factory panel
        private void btnFactory_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string query = "select Factory_ID from factory where Factory_Name = '" + txtFactory.Text + "'";
                MySqlCommand com = new MySqlCommand(query, dbconnection);
                if (com.ExecuteScalar() == null)
                {
                    if (checkedListBox1.CheckedItems.Count > 0)
                    {
                        if (txtFactory.Text != "")
                        {
                            query = "insert into factory (Factory_Name) values (@name)";
                            com = new MySqlCommand(query, dbconnection);
                            com.Parameters.AddWithValue("@name", txtFactory.Text);
                            com.ExecuteNonQuery();

                            query = "select Factory_ID from factory order by Factory_ID desc limit 1";
                            com = new MySqlCommand(query, dbconnection);
                            int factory_ID = Convert.ToInt32(com.ExecuteScalar().ToString());

                            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
                            {
                                int Type_ID = Convert.ToInt32(checkedListBox1.CheckedItems[i].ToString().Split('\t')[1]);
                                query = "insert into type_factory (Type_ID,Factory_ID) values (@Type_ID,@Factory_ID)";
                                com = new MySqlCommand(query, dbconnection);
                                com.Parameters.Add("@Type_ID", MySqlDbType.Int16);
                                com.Parameters["@Type_ID"].Value = Type_ID;
                                com.Parameters.Add("@factory_ID", MySqlDbType.Int16);
                                com.Parameters["@factory_ID"].Value = factory_ID;
                                com.ExecuteNonQuery();
                            }
                            
                            query = "select Factory_ID from factory order by Factory_ID desc limit 1";
                            com = new MySqlCommand(query, dbconnection);
                            UserControl.ItemRecord("factory", "اضافة", factory_ID, DateTime.Now,"", dbconnection);

                            displayFactory(Convert.ToInt32(comType.SelectedValue));
                            txtFactory.Text = "";
                            for (int j = 0; j < checkedListBox1.Items.Count; j++)
                            {
                                checkedListBox1.SetItemChecked(j, false);
                            }

                        }
                        else
                        {
                            txtFactory.Focus();
                        }
                    }
                    else
                    {
                        MessageBox.Show("select type");
                    }
                }
                else
                {
                    MessageBox.Show("This factory already exist");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void txtFactory_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (btnSave.Visible == false)
                    {
                        dbconnection.Open();
                        string query = "select Factory_ID from factory where Factory_Name = '" + txtFactory.Text + "'";
                        MySqlCommand com = new MySqlCommand(query, dbconnection);
                        if (com.ExecuteScalar() == null)
                        {
                            if (checkedListBox1.CheckedItems.Count > 0)
                            {
                                if (txtFactory.Text != "")
                                {
                                    query = "insert into factory (Factory_Name) values (@name)";
                                    com = new MySqlCommand(query, dbconnection);
                                    com.Parameters.AddWithValue("@name", txtFactory.Text);
                                    com.ExecuteNonQuery();

                                    query = "select Factory_ID from factory order by Factory_ID desc limit 1";
                                    com = new MySqlCommand(query, dbconnection);
                                    int factory_ID = Convert.ToInt32(com.ExecuteScalar().ToString());

                                    for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
                                    {
                                        int Type_ID = Convert.ToInt32(checkedListBox1.CheckedItems[i].ToString().Split('\t')[1]);
                                        query = "insert into type_factory (Type_ID,Factory_ID) values (@Type_ID,@Factory_ID)";
                                        com = new MySqlCommand(query, dbconnection);
                                        com.Parameters.Add("@Type_ID", MySqlDbType.Int16);
                                        com.Parameters["@Type_ID"].Value = Type_ID;
                                        com.Parameters.Add("@factory_ID", MySqlDbType.Int16);
                                        com.Parameters["@factory_ID"].Value = factory_ID;
                                        com.ExecuteNonQuery();
                                    }


                                    query = "select Factory_ID from factory order by Factory_ID desc limit 1";
                                    com = new MySqlCommand(query, dbconnection);
                                    UserControl.ItemRecord("factory", "اضافة", factory_ID, DateTime.Now, "", dbconnection);

                                   // displayFactory(Convert.ToInt32(comType.SelectedValue));
                                    txtFactory.Text = "";
                                    for (int j = 0; j < checkedListBox1.Items.Count; j++)
                                    {
                                        checkedListBox1.SetItemChecked(j, false);

                                    }

                                }
                                else
                                {
                                    txtFactory.Focus();
                                }
                            }
                            else
                            {
                                MessageBox.Show("select type");
                            }
                        }
                        else
                        {
                            MessageBox.Show("This factory already exist");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
               if (id != -1)
                {                  
                    if (checkedListBox1.CheckedItems.Count > 0)
                    {
                        string query = "select Factory_ID from factory where Factory_Name = '" + txtFactory.Text + "' and Factory_ID not in(" + id + ")";
                        MySqlCommand com = new MySqlCommand(query, dbconnection);
                        if (com.ExecuteScalar() == null)
                        {
                            if (txtFactory.Text != "")
                            {
                                query = "update factory set Factory_Name=@name where Factory_ID=" + id;
                                com = new MySqlCommand(query, dbconnection);
                                com.Parameters.AddWithValue("@name", txtFactory.Text);
                                com.ExecuteNonQuery();

                                query = "delete from type_factory where Factory_ID=" + id;
                                com = new MySqlCommand(query, dbconnection);
                                com.ExecuteNonQuery();
                                for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
                                {
                                    int Type_ID = Convert.ToInt32(checkedListBox1.CheckedItems[i].ToString().Split('\t')[1]);
                                    query = "insert into type_factory (Type_ID,Factory_ID) values (@Type_ID,@Factory_ID)";
                                    com = new MySqlCommand(query, dbconnection);
                                    com.Parameters.Add("@Type_ID", MySqlDbType.Int16);
                                    com.Parameters["@Type_ID"].Value = Type_ID;
                                    com.Parameters.Add("@Factory_ID", MySqlDbType.Int16);
                                    com.Parameters["@Factory_ID"].Value = id;
                                    com.ExecuteNonQuery();
                                }


                                query = "select Factory_ID from factory order by Factory_ID desc limit 1";
                                com = new MySqlCommand(query, dbconnection);
                                UserControl.ItemRecord("factory", "تعديل", id, DateTime.Now,"", dbconnection);

                                displayFactory(Convert.ToInt32(comType.SelectedValue));
                                txtFactory.Text = "";
                                btnSave.Visible = false;
                                for (int j = 0; j < checkedListBox1.Items.Count; j++)
                                {
                                   checkedListBox1.SetItemChecked(j, false);
           
                                }
                            }
                            else
                            {
                                txtFactory.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("This factory already exist");
                        }
                    }
                    else
                    {
                        MessageBox.Show("select type");
                    }
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                btnSave.Visible = true;
                row1 = dataGridViewFactory.Rows[dataGridViewFactory.SelectedCells[0].RowIndex];
                if (row1 != null)
                {
                    id = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    string query = "select Type_ID from type_factory where Factory_ID="+id;
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    MySqlDataReader dr = com.ExecuteReader();
                    List<int> idList = new List<int>();
                    while (dr.Read())
                    {
                        idList.Add(Convert.ToInt32(dr["Type_ID"].ToString()));
                    }
                    dr.Close();
                    dbconnection.Close();
                    txtFactory.Text = row1.Cells[1].Value.ToString();
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        int Type_ID = Convert.ToInt32(checkedListBox1.Items[i].ToString().Split('\t')[1]);
                        for (int j = 0; j < idList.Count; j++)
                        {
                            if (Type_ID == idList[j])
                            {
                                checkedListBox1.SetItemChecked(i, true);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch
            {
                MessageBox.Show("select row");
            }
            dbconnection.Close();
        }
        private void btnDeleteFactory_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                DataGridViewRow row1 = dataGridViewFactory.Rows[dataGridViewFactory.SelectedCells[0].RowIndex];
                if (row1 != null)
                {
                    int id = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    if (!producItemUsed("Factory_ID", id))
                    {
                        DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete the item?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.Yes)
                        {
                            string query = "delete from factory where Factory_ID=" + id;
                            MySqlCommand com = new MySqlCommand(query, dbconnection);
                            com.ExecuteNonQuery();

                            updateTablesDB("factory", "Factory_ID", id);

                            UserControl.ItemRecord("factory", "حذف", Convert.ToInt32(row1.Cells[0].Value.ToString()), DateTime.Now, "", dbconnection);

                            displayFactory();

                            txtFactory.Text = "";
                        }
                        else if (dialogResult == DialogResult.No)
                        { }
                    }
                    else
                    {
                        MessageBox.Show("هذا العنصر تم استخدامه في تسجيل بنود لا يمكن حذفه.");
                    }
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnDisplayAllFactory_Click(object sender, EventArgs e)
        {
            try
            {
                displayFactory();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region group panel
        private void btnGroup_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string query = "select Group_ID from groupo where Group_Name = '" + txtGroup.Text + "'";
                MySqlCommand com = new MySqlCommand(query, dbconnection);
                if (com.ExecuteScalar() == null)
                {
                    if (txtGroup.Text != "")
                    {
                        query = "select TypeCoding_Method from type where Type_ID="+ comType.SelectedValue;
                        com = new MySqlCommand(query, dbconnection);
                        int TypeCoding_Method =(int) com.ExecuteScalar();
                        if (TypeCoding_Method == 1)
                        {
                            query = "insert into groupo (Group_Name,Factory_ID,Type_ID) values (@name,@Factory_ID,@Type_ID)";
                            com = new MySqlCommand(query, dbconnection);
                            com.Parameters.AddWithValue("@name", txtGroup.Text);
                            //if (comType.SelectedValue.ToString() == "2")
                            //{
                            //    com.Parameters.AddWithValue("@Factory_ID", -1);
                            //}
                            //else
                            //{
                                com.Parameters.AddWithValue("@Factory_ID", -Convert.ToInt32(comType.SelectedValue));
                            //}
                            com.Parameters.AddWithValue("@Type_ID", Convert.ToInt32(comType.SelectedValue));
                            com.ExecuteNonQuery();
                            displayGroup(-Convert.ToInt32(comType.SelectedValue));

                        }
                        else
                        {
                            if (comFactory.Text != "")
                            {
                                query = "insert into groupo (Group_Name,Factory_ID,Type_ID) values (@name,@Factory_ID,@Type_ID)";
                                com = new MySqlCommand(query, dbconnection);
                                com.Parameters.AddWithValue("@name", txtGroup.Text);
                                com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactory.SelectedValue));
                                com.Parameters.AddWithValue("@Type_ID", Convert.ToInt32(comType.SelectedValue));
                                com.ExecuteNonQuery();
                                displayGroup(Convert.ToInt32(comFactory.SelectedValue));
                            }
                            else
                            {
                                MessageBox.Show("اختر المصنع");
                            }
                        }
                        dbconnection.Close();
                        dbconnection.Open();
                        query = "select Group_ID from groupo order by Group_ID desc limit 1";
                        com = new MySqlCommand(query, dbconnection);
                        UserControl.ItemRecord("groupo", "اضافة", Convert.ToInt32(com.ExecuteScalar().ToString()), DateTime.Now, "", dbconnection);

                        txtGroup.Text = "";
                    }
                    else
                    {
                        txtGroup.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("This group already exist");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void txtGroup_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (btnSaveGroup.Visible == false)
                    {
                        dbconnection.Open();
                        string query = "select Group_ID from groupo where Group_Name = '" + txtGroup.Text + "'";
                        MySqlCommand com = new MySqlCommand(query, dbconnection);
                        if (com.ExecuteScalar() == null)
                        {
                            if (txtGroup.Text != "")
                            {
                                query = "select TypeCoding_Method from type where Type_ID=" + comType.SelectedValue;
                                com = new MySqlCommand(query, dbconnection);
                                int TypeCoding_Method = (int)com.ExecuteScalar();
                                if (TypeCoding_Method == 1)
                                {
                                    query = "insert into groupo (Group_Name,Factory_ID,Type_ID) values (@name,@Factory_ID,@Type_ID)";
                                    com = new MySqlCommand(query, dbconnection);
                                    com.Parameters.AddWithValue("@name", txtGroup.Text);
                                    //if (comType.SelectedValue.ToString() == "2")
                                    //{
                                    //    com.Parameters.AddWithValue("@Factory_ID", -1);
                                    //}
                                    //else
                                    //{
                                        com.Parameters.AddWithValue("@Factory_ID", -Convert.ToInt32(comType.SelectedValue));
                                   // }
                                    com.Parameters.AddWithValue("@Type_ID", Convert.ToInt32(comType.SelectedValue));
                                    com.ExecuteNonQuery();
                                    displayGroup(-Convert.ToInt32(comType.SelectedValue));

                                }
                                else
                                {
                                    if (comFactory.Text != "")
                                    {
                                        query = "insert into groupo (Group_Name,Factory_ID,Type_ID) values (@name,@Factory_ID,@Type_ID)";
                                        com = new MySqlCommand(query, dbconnection);
                                        com.Parameters.AddWithValue("@name", txtGroup.Text);
                                        com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactory.SelectedValue));
                                        com.Parameters.AddWithValue("@Type_ID", Convert.ToInt32(comType.SelectedValue));
                                        com.ExecuteNonQuery();
                                        displayGroup(Convert.ToInt32(comFactory.SelectedValue));
                                    }
                                    else
                                    {
                                        MessageBox.Show("اختر المصنع");
                                    }
                                }
                                dbconnection.Close();
                                dbconnection.Open();
                                query = "select Group_ID from groupo order by Group_ID desc limit 1";
                                com = new MySqlCommand(query, dbconnection);
                                UserControl.ItemRecord("groupo", "اضافة", Convert.ToInt32(com.ExecuteScalar().ToString()), DateTime.Now, "", dbconnection);

                                txtGroup.Text = "";
                            }
                            else
                            {
                                txtGroup.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("This group already exist");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void comType_SelectedValueChanged_1(object sender, EventArgs e)
        {
            try
            {
                if (load)
                {
                    dbconnection.Open();
                    string query = "select TypeCoding_Method from type where Type_ID=" + comType.SelectedValue;
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    int TypeCoding_Method = (int)com.ExecuteScalar();
                    if (TypeCoding_Method==2)
                    {
                        flagFactory = false;
                        query = "select distinct factory.Factory_ID, Factory_Name from factory inner join type_factory on type_factory.Factory_ID=factory.Factory_ID where type_factory.Type_ID=" + comType.SelectedValue.ToString();
                        MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        comFactory.DataSource = dt;
                        comFactory.DisplayMember = dt.Columns["Factory_Name"].ToString();
                        comFactory.ValueMember = dt.Columns["Factory_ID"].ToString();
                        comFactory.Text = "";
                        txtFactory1.Text = "";

                        flagFactory = true;

                        displayGroup_Type(Convert.ToInt32(comType.SelectedValue.ToString()));
                        
                        label6.Visible = true;
                        comFactory.Visible = true;
                        txtFactory1.Visible = true;
                    }
                    else
                    {
                        displayGroup_Type(Convert.ToInt32(comType.SelectedValue.ToString()));

                        label6.Visible = false;
                        comFactory.Visible = false;
                        txtFactory1.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void comFactory_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    displayGroup((int)comFactory.SelectedValue);
                    txtGroup.Focus();
                    txtGroup.Select();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void comFactory_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (flagFactory)
                {
                    displayGroup_Factory(Convert.ToInt32(comFactory.SelectedValue));
                    
                    txtFactory1.Text = "";
                    txtGroup.Focus();
                    txtGroup.SelectAll();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnSaveGroup_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string query = "select Group_ID from groupo where Group_Name = '" + txtGroup.Text + "' and Group_ID not in("+groupUpdateId+")";
                MySqlCommand com = new MySqlCommand(query, dbconnection);
                if (com.ExecuteScalar() == null)
                {
                    if (txtGroup.Text != "")
                    {
                        query = "select TypeCoding_Method from type where Type_ID=" + comType.SelectedValue;
                        com = new MySqlCommand(query, dbconnection);
                        int TypeCoding_Method = (int)com.ExecuteScalar();
                        if (TypeCoding_Method==1)
                        {
                            query = "update  groupo set Group_Name=@Group_Name,Factory_ID=@Factory_ID,Type_ID=@Type_ID where Group_ID="+ groupUpdateId;
                            com = new MySqlCommand(query, dbconnection);
                            com.Parameters.AddWithValue("@Group_Name", txtGroup.Text);
                            //if (comType.SelectedValue.ToString() == "2")
                            //{
                            //    com.Parameters.AddWithValue("@Factory_ID", -1);
                            //}
                            //else
                            //{
                                com.Parameters.AddWithValue("@Factory_ID", -Convert.ToInt32(comType.SelectedValue));
                            //}
                            com.Parameters.AddWithValue("@Type_ID", comType.SelectedValue);
                            com.ExecuteNonQuery();
                            displayGroup(-(int)comType.SelectedValue);

                        }
                        else
                        {
                            if (comFactory.Text != "")
                            {
                                query = "update  groupo set Group_Name=@Group_Name,Factory_ID=@Factory_ID,Type_ID=@Type_ID where Group_ID=" + groupUpdateId;
                                com = new MySqlCommand(query, dbconnection);
                                com.Parameters.AddWithValue("@Group_Name", txtGroup.Text);
                                com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactory.SelectedValue));
                                com.Parameters.AddWithValue("@Type_ID", Convert.ToInt32(comType.SelectedValue));
                                com.ExecuteNonQuery();
                                displayGroup(Convert.ToInt32(comFactory.SelectedValue));
                            }
                            else
                            {
                                MessageBox.Show("اختر المصنع");
                            }
                        }
                        dbconnection.Close();
                        dbconnection.Open();
                        UserControl.ItemRecord("groupo", "تعديل", groupUpdateId, DateTime.Now, "", dbconnection);
                        label6.Visible = true;
                        comFactory.Visible = true;
                        txtFactory1.Visible = true;
                        comType.Text = "";
                        txtType1.Text = "";
                        comFactory.Text = "";
                        txtFactory1.Text = "";
                        txtGroup.Text = "";
                        btnSaveGroup.Visible = false;
                    }
                    else
                    {
                        txtGroup.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("This group already exist");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnUpdateGroup_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                String type="", factory="", group="";
                btnSaveGroup.Visible = true;
                row1 = dataGridViewGroup.Rows[dataGridViewGroup.SelectedCells[0].RowIndex];
                if (row1 != null)
                {
                    groupUpdateId = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    string query = "select distinct * from groupo  where Group_ID=" + groupUpdateId;
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    MySqlDataReader drGroup = com.ExecuteReader();
                    while (drGroup.Read())
                    {
                        type = drGroup["Type_ID"].ToString();
                        group = drGroup["Group_Name"].ToString();
                        factory = drGroup["Factory_ID"].ToString();
                    }
                    drGroup.Close();
                    txtType1.Text = type;
                    txtGroup.Text = group;
                    if (Convert.ToInt32(factory) > 0)
                        txtFactory1.Text = factory;
                    else
                    {
                        label6.Visible = false;
                        comFactory.Visible = false;
                        txtFactory1.Visible = false;
                    }
                }
                else
                {
                    MessageBox.Show("اختار المجموعة المراد تعديلها.");
                }
            }
            catch
            {
               // MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnDeleteGroup_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                DataGridViewRow row1 = dataGridViewGroup.Rows[dataGridViewGroup.SelectedCells[0].RowIndex];

                if (row1 != null)
                {
                    int id = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    if (!producItemUsed("Group_ID", id))
                    {
                        DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete the item?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.Yes)
                        {
                            string query = "delete from groupo where Group_ID=" + id;
                            MySqlCommand com = new MySqlCommand(query, dbconnection);
                            com.ExecuteNonQuery();

                            updateTablesDB("groupo", "Group_ID", id);

                            UserControl.ItemRecord("groupo", "حذف", Convert.ToInt32(row1.Cells[0].Value.ToString()), DateTime.Now, "", dbconnection);
                            if (comFactory.Text != "")
                            {
                                displayGroup(Convert.ToInt32(comFactory.SelectedValue));
                            }
                            else
                            {
                                displayGroup();
                            }

                            txtGroup.Text = "";
                        }
                        else if (dialogResult == DialogResult.No)
                        { }
                    }
                    else
                    {
                        MessageBox.Show("هذا العنصر تم استخدامه في تسجيل بنود لا يمكن حذفه.");
                    }
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnGroupDisplayAll_Click(object sender, EventArgs e)
        {
            try
            {
                displayGroup();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region product panel
        private void btnProduct_Click(object sender, EventArgs e)
        {
            try
            {
                if ((txtGroup2.Text != "" || chListBoxGroup.CheckedItems.Count > 0) && comFactoryGroup.Text != "" && comTypeProduct.Text != "")
                {
                    dbconnection.Open();
                    string query = "select Product_ID from product where Product_Name = '" + txtProduct.Text + "'";
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    if (com.ExecuteScalar() == null)
                    {
                        if (txtProduct.Text != "")
                        {
                            query = "insert into product (Product_Name,Type_ID) values (@name,@Type_ID)";
                            com = new MySqlCommand(query, dbconnection);
                            com.Parameters.AddWithValue("@name", txtProduct.Text);
                            com.Parameters.AddWithValue("@Type_ID", Convert.ToInt32(comTypeProduct.SelectedValue));
                            com.ExecuteNonQuery();
                            query = "select Product_ID from product order by Product_ID desc limit 1";
                            com = new MySqlCommand(query, dbconnection);
                            int id = Convert.ToInt32(com.ExecuteScalar());

                            query = "select TypeCoding_Method from type where Type_ID=" + comType.SelectedValue;
                            com = new MySqlCommand(query, dbconnection);
                            int TypeCoding_Method = (int)com.ExecuteScalar();
                            if (TypeCoding_Method == 1)
                            {
                                int Group_ID = 0;
                                for (int i = 0; i < chListBoxGroup.CheckedItems.Count; i++)
                                {
                                    Group_ID = Convert.ToInt32(chListBoxGroup.CheckedItems[i].ToString().Split('\t')[1]);
                                    query = "insert into product_factory_group (Product_ID,Factory_ID,Group_ID) values (@Product_ID,@Factory_ID,@Group_ID)";
                                    com = new MySqlCommand(query, dbconnection);
                                    com.Parameters.Add("@Product_ID", MySqlDbType.Int16);
                                    com.Parameters["@Product_ID"].Value = id;
                                    com.Parameters.Add("@Group_ID", MySqlDbType.Int16);
                                    com.Parameters["@Group_ID"].Value = Group_ID;
                                    com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactoryGroup.SelectedValue));

                                    com.ExecuteNonQuery();
                                }
                                displayProduct_Group(Group_ID);
                            }
                            else
                            {
                                query = "insert into product_factory_group (Product_ID,Factory_ID,Group_ID) values (@Product_ID,@Factory_ID,@Group_ID)";
                                com = new MySqlCommand(query, dbconnection);
                                com.Parameters.AddWithValue("@Product_ID", id);
                                com.Parameters.AddWithValue("@Group_ID", Convert.ToInt32(comGroup.SelectedValue));
                                com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactoryGroup.SelectedValue));

                                com.ExecuteNonQuery();
                                displayProduct_Group(Convert.ToInt32(comGroup.SelectedValue));
                            }

                            UserControl.ItemRecord("product", "اضافة", id , DateTime.Now,"", dbconnection);
                         
                            txtProduct.Text = "";
                            txtProduct.Focus();
                        }
                        else
                        {
                            txtProduct.Focus();

                        }
                    }
                    else
                    {
                        int Product_ID = Convert.ToInt32(com.ExecuteScalar());

                        for (int i = 0; i < chListBoxGroup.CheckedItems.Count; i++)
                        {
                            int Group_ID = Convert.ToInt32(chListBoxGroup.CheckedItems[i].ToString().Split('\t')[1]);
                            query = "insert into product_factory_group (Product_ID,Factory_ID,Group_ID) values (@Product_ID,@Factory_ID,@Group_ID)";
                            com = new MySqlCommand(query, dbconnection);
                            com.Parameters.Add("@Product_ID", MySqlDbType.Int16);
                            com.Parameters["@Product_ID"].Value = Product_ID;
                            com.Parameters.Add("@Group_ID", MySqlDbType.Int16);
                            com.Parameters["@Group_ID"].Value = Group_ID;
                            com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactoryGroup.SelectedValue));

                            com.ExecuteNonQuery();
                        }

                        txtProduct.Text = "";
                        txtProduct.Focus();
                        displayProduct_Factory(Convert.ToInt32(txtFactory2.Text));
                    }
                }
                else
                {
                    MessageBox.Show("select group");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void txtProduct_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (btnSave_productUpdate.Visible == false)
                    {
                        if ((txtGroup2.Text != "" || chListBoxGroup.CheckedItems.Count > 0) && comFactoryGroup.Text != "" && comTypeProduct.Text != "")
                        {
                            dbconnection.Open();
                            string query = "select Product_ID from product where Product_Name = '" + txtProduct.Text + "'";
                            MySqlCommand com = new MySqlCommand(query, dbconnection);
                            if (com.ExecuteScalar() == null)
                            {
                                if (txtProduct.Text != "")
                                {
                                    query = "insert into product (Product_Name,Type_ID) values (@name,@Type_ID)";
                                    com = new MySqlCommand(query, dbconnection);
                                    com.Parameters.AddWithValue("@name", txtProduct.Text);
                                    com.Parameters.AddWithValue("@Type_ID", Convert.ToInt32(comTypeProduct.SelectedValue));
                                    com.ExecuteNonQuery();
                                    query = "select Product_ID from product order by Product_ID desc limit 1";
                                    com = new MySqlCommand(query, dbconnection);
                                    int id = Convert.ToInt32(com.ExecuteScalar());
                                    query = "select TypeCoding_Method from type where Type_ID=" + comType.SelectedValue;
                                    com = new MySqlCommand(query, dbconnection);
                                    int TypeCoding_Method = (int)com.ExecuteScalar();
                                    if (TypeCoding_Method == 1)
                                    {
                                        int Group_ID = 0;
                                        for (int i = 0; i < chListBoxGroup.CheckedItems.Count; i++)
                                        {
                                            Group_ID = Convert.ToInt32(chListBoxGroup.CheckedItems[i].ToString().Split('\t')[1]);
                                            query = "insert into product_factory_group (Product_ID,Factory_ID,Group_ID) values (@Product_ID,@Factory_ID,@Group_ID)";
                                            com = new MySqlCommand(query, dbconnection);
                                            com.Parameters.Add("@Product_ID", MySqlDbType.Int16);
                                            com.Parameters["@Product_ID"].Value = id;
                                            com.Parameters.Add("@Group_ID", MySqlDbType.Int16);
                                            com.Parameters["@Group_ID"].Value = Group_ID;
                                            com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactoryGroup.SelectedValue));

                                            com.ExecuteNonQuery();
                                        }
                                        displayProduct_Group(Group_ID);
                                    }
                                    else
                                    {
                                        query = "insert into product_factory_group (Product_ID,Factory_ID,Group_ID) values (@Product_ID,@Factory_ID,@Group_ID)";
                                        com = new MySqlCommand(query, dbconnection);
                                        com.Parameters.AddWithValue("@Product_ID", id);
                                        com.Parameters.AddWithValue("@Group_ID", Convert.ToInt32(comGroup.SelectedValue));
                                        com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactoryGroup.SelectedValue));

                                        com.ExecuteNonQuery();
                                        displayProduct_Group(Convert.ToInt32(comGroup.SelectedValue));
                                    }
                                    UserControl.ItemRecord("product", "اضافة", id, DateTime.Now, "", dbconnection);


                                    txtProduct.Text = "";
                                    txtProduct.Focus();
                                }
                                else
                                {
                                    txtProduct.Focus();

                                }
                            }
                            else
                            {
                                int Product_ID = Convert.ToInt32(com.ExecuteScalar());

                                for (int i = 0; i < chListBoxGroup.CheckedItems.Count; i++)
                                {
                                    int Group_ID = Convert.ToInt32(chListBoxGroup.CheckedItems[i].ToString().Split('\t')[1]);
                                    query = "insert into product_factory_group (Product_ID,Factory_ID,Group_ID) values (@Product_ID,@Factory_ID,@Group_ID)";
                                    com = new MySqlCommand(query, dbconnection);
                                    com.Parameters.Add("@Product_ID", MySqlDbType.Int16);
                                    com.Parameters["@Product_ID"].Value = Product_ID;
                                    com.Parameters.Add("@Group_ID", MySqlDbType.Int16);
                                    com.Parameters["@Group_ID"].Value = Group_ID;
                                    com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactoryGroup.SelectedValue));

                                    com.ExecuteNonQuery();
                                }


                                txtProduct.Text = "";
                                txtProduct.Focus();
                                displayProduct_Factory(Convert.ToInt32(txtFactory2.Text));
                            }
                        }
                        else
                        {
                            MessageBox.Show("select group");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }   
        private void comGroup2_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    displayProduct_Group(Convert.ToInt32(txtGroup2.Text));
                    txtProduct.Focus();
                    txtProduct.SelectAll();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void comFactoryGroup_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (flagFactoryP)
                {
                    if (txtType2.Text != "1" && txtType2.Text != "2" && txtType2.Text != "3" && txtType2.Text != "4")
                    {
                        string query = "select * from groupo where Factory_ID=" + comFactoryGroup.SelectedValue;
                        MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        comGroup.DataSource = dt;
                        comGroup.DisplayMember = dt.Columns["Group_Name"].ToString();
                        comGroup.ValueMember = dt.Columns["Group_ID"].ToString();
                        comGroup.Text = "";
                        txtGroup2.Text = "";
                        comGroup.Focus();
                    }
                  
                    dataGridViewProduct.DataSource = null;
                    dataGridViewProduct.DataSource = displayProduct_Factory(Convert.ToInt32(comFactoryGroup.SelectedValue));
                    dataGridViewProduct.Columns[0].Width = 50;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void comGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (flagGroup)
                {
                    displayProduct_Group(Convert.ToInt32(txtGroup2.Text));
                    txtProduct.Focus();
                    txtProduct.SelectAll();
                }         
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnDeleteProduct_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                DataGridViewRow row1 = dataGridViewProduct.Rows[dataGridViewProduct.SelectedCells[0].RowIndex];

                if (row1 != null)
                {
                    int id = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    if (!producItemUsed("Product_ID", id))
                    {
                        DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete the item?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.Yes)
                        {
                            string query = "delete from product where Product_ID=" + id;
                            MySqlCommand com = new MySqlCommand(query, dbconnection);
                            com.ExecuteNonQuery();
                            updateTablesDB("product", "Product_ID", id);

                            UserControl.ItemRecord("product", "حذف", Convert.ToInt32(row1.Cells[0].Value.ToString()), DateTime.Now, "", dbconnection);

                            dataGridViewProduct.DataSource = null;
                            displayProductAll();

                            txtProduct.Focus();
                        }
                        else if (dialogResult == DialogResult.No)
                        { }
                    }
                    else
                    {
                        MessageBox.Show("هذا العنصر تم استخدامه في تسجيل بنود لا يمكن حذفه.");
                    }
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnDisplayAllProduct_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridViewProduct.DataSource = null;
                displayProductAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void comTypeProduct_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                if (mTC_Content.SelectedIndex == 3)
                {
                    flagFactoryP = false;
                    flagGroup = false;
                    txtType2.Text = comTypeProduct.SelectedValue.ToString();
                    string query = "select distinct factory.Factory_ID, factory.Factory_Name from factory inner join type_factory on type_factory.Factory_ID=factory.Factory_ID where type_factory.Type_ID=" + comTypeProduct.SelectedValue.ToString();
                    MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    comFactoryGroup.DataSource = dt;
                    comFactoryGroup.DisplayMember = dt.Columns["Factory_Name"].ToString();
                    comFactoryGroup.ValueMember = dt.Columns["Factory_ID"].ToString();
                    comFactoryGroup.Text = "";
                    txtFactory2.Text = "";
                    dbconnection.Close();
                    dbconnection.Open();
                    query = "select TypeCoding_Method from type where Type_ID=" + comTypeProduct.SelectedValue;
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    int TypeCoding_Method = (int)com.ExecuteScalar();
                    if (TypeCoding_Method==1)
                    {
                        //if (comTypeProduct.SelectedValue.ToString() == "2" || comTypeProduct.SelectedValue.ToString() == "1")
                        //{
                        //    query = "select distinct groupo.Group_ID, Group_Name from groupo  where Factory_ID=-1";
                        //}
                        //else
                        //{
                            query = "select distinct groupo.Group_ID, Group_Name from groupo  where Type_ID=" + comTypeProduct.SelectedValue + " and Factory_ID=" + -(int)comTypeProduct.SelectedValue;
                     //   }

                        com = new MySqlCommand(query, dbconnection);
                        MySqlDataReader dr = com.ExecuteReader();

                        chListBoxGroup.Items.Clear();
                        while (dr.Read())
                            chListBoxGroup.Items.Add(dr["Group_Name"].ToString() + "\t" + dr["Group_ID"]);

                        dr.Close();
                        
                        chListBoxGroup.Visible = true;
                        comGroup.Visible = false;
                        txtGroup2.Visible = false;
                        label11.Visible = false;

                    }
                    else
                    {
                        chListBoxGroup.Visible = false;
                        comGroup.Visible = true;
                        txtGroup2.Visible = true;
                        label11.Visible = true;
                    }
                    dataGridViewProduct.DataSource = displayProduct_Type(Convert.ToInt32(txtType2.Text));
                    dataGridViewProduct.Columns[0].Width = 50;
                    flagFactoryP = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnProductUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                string  Type_ID="", Factory_ID="", Product_Name="",Group_ID="";
                btnSave_productUpdate.Visible = true;
                dbconnection.Close();
                dbconnection.Open();
                DataGridViewRow row1 = dataGridViewProduct.Rows[dataGridViewProduct.SelectedCells[0].RowIndex];
                productUpdateId = Convert.ToInt32(row1.Cells[0].Value.ToString());
                string queryx = "select distinct * from  product_factory_group inner join product on product_factory_group.Product_ID=product.Product_ID where product_factory_group.Product_ID=" + productUpdateId;
                MySqlCommand com = new MySqlCommand(queryx,dbconnection);
                MySqlDataReader dr = com.ExecuteReader();
                while (dr.Read())
                {
                    Type_ID = dr["Type_ID"].ToString();
                    Product_Name = dr["Product_Name"].ToString();
                    Factory_ID = dr["Factory_ID"].ToString();
                    Group_ID = dr["Group_ID"].ToString();
                }
                dr.Close();
                txtProduct.Text = Product_Name;
                txtType2.Text = Type_ID;
                txtFactory2.Text = Factory_ID;
                txtGroup2.Text = Group_ID;
                dbconnection.Close();
                dbconnection.Open();
                string query = "select TypeCoding_Method from type where Type_ID=" + comType.SelectedValue;
                com = new MySqlCommand(query, dbconnection);
                int TypeCoding_Method = (int)com.ExecuteScalar();
                if (TypeCoding_Method==1)
                {
                    string query1 = "";
                    if (comType.SelectedValue.ToString() == "2")
                    {
                        query1 = "select distinct groupo.Group_ID, Group_Name from groupo inner join type on type.Type_ID=groupo.Type_ID  where type.Type_ID=" + 1 + " and Factory_ID=" + -1;
                    }
                    else
                    {
                        query1 = "select distinct groupo.Group_ID, Group_Name from groupo inner join type on type.Type_ID=groupo.Type_ID  where type.Type_ID=" + comType.SelectedValue + " and Factory_ID=" + -(int)comType.SelectedValue;

                    }

                    chListBoxGroup.Visible = true;
                    labGroup.Visible = true;
                    comGroup.Visible = false;
                    txtGroup2.Visible = false;
                    label11.Visible = false;
                    MySqlCommand com1 = new MySqlCommand(query1, dbconnection);
                    MySqlDataReader dr1 = com1.ExecuteReader();

                    chListBoxGroup.Items.Clear();
                    while (dr1.Read())
                        chListBoxGroup.Items.Add(dr1["Group_Name"].ToString() + "\t" + dr1["Group_ID"]);
                    dr1.Close();

                    fun();
                }
                else
                {
                    chListBoxGroup.Items.Clear();
                    chListBoxGroup.Visible = false;
                    labGroup.Visible = false;
                    comGroup.Visible = true;
                    txtGroup2.Visible = true;
                    label11.Visible = true;
                }
                
              
                dbconnection.Close();
                dbconnection.Open();
                comTypeProduct.Enabled = false;
                txtType2.ReadOnly = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnSave_productUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                if ((txtGroup2.Text != "" || chListBoxGroup.CheckedItems.Count > 0) && comTypeProduct.Text != "" && checkGroupsForAllFactories())
                {
                    dbconnection.Open();
                    string query = "select Product_ID from product where Product_Name = '" + txtProduct.Text + "' and Product_ID not in ("+ productUpdateId + ")";
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    if (com.ExecuteScalar() == null)
                    {
                        if (txtProduct.Text != "")
                        {
                            query = "update  product set Product_Name=@name,Type_ID=@Type_ID where Product_ID="+ productUpdateId;
                            com = new MySqlCommand(query, dbconnection);
                            com.Parameters.AddWithValue("@name", txtProduct.Text);
                            com.Parameters.AddWithValue("@Type_ID", Convert.ToInt32(txtType2.Text));
                           com.ExecuteNonQuery();

                            query = "delete from product_factory_group where Product_ID=" + productUpdateId;
                            com = new MySqlCommand(query, dbconnection);
                            com.ExecuteNonQuery();

                            query = "select TypeCoding_Method from type where Type_ID=" + comType.SelectedValue;
                            com = new MySqlCommand(query, dbconnection);
                            int TypeCoding_Method = (int)com.ExecuteScalar();
                            if (TypeCoding_Method == 1)
                            {
                                int Group_ID = 0;
                                for (int i = 0; i < chListBoxGroup.CheckedItems.Count; i++)
                                {
                                    Group_ID = Convert.ToInt32(chListBoxGroup.CheckedItems[i].ToString().Split('\t')[1]);
                                    query = "insert into product_factory_group (Product_ID,Factory_ID,Group_ID) values (@Product_ID,@Factory_ID,@Group_ID)";
                                    com = new MySqlCommand(query, dbconnection);
                                    com.Parameters.Add("@Product_ID", MySqlDbType.Int16);
                                    com.Parameters["@Product_ID"].Value = productUpdateId;
                                    com.Parameters.Add("@Group_ID", MySqlDbType.Int16);
                                    com.Parameters["@Group_ID"].Value = Group_ID;
                                    com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactoryGroup.SelectedValue));

                                    com.ExecuteNonQuery();
                                }
                                displayProduct_Group(Group_ID);
                            }
                            else
                            {
                                query = "insert into product_factory_group (Product_ID,Factory_ID,Group_ID) values (@Product_ID,@Factory_ID,@Group_ID)";
                                com = new MySqlCommand(query, dbconnection);
                                com.Parameters.AddWithValue("@Product_ID", productUpdateId);
                                com.Parameters.AddWithValue("@Group_ID", Convert.ToInt32(comGroup.SelectedValue));
                                com.Parameters.AddWithValue("@Factory_ID", Convert.ToInt32(comFactoryGroup.SelectedValue));

                                com.ExecuteNonQuery();
                                displayProduct_Group(Convert.ToInt32(comGroup.SelectedValue));
                            }

                            //int FactoryID = 0;
                            //for (int j = 0; j < listFactory_Group.Count; j++)
                            //{
                            //    FactoryID = listFactory_Group[j].Factory_ID;
                            //    factory_Group mfactory_Group = new factory_Group();
                            //    foreach (factory_Group item in listFactory_Group)
                            //    {
                            //        if (item.Factory_ID == FactoryID)
                            //        {
                            //            mfactory_Group = item;
                            //            break;
                            //        }
                            //    }
                            //    int Group_ID = 0;
                            //    for (int i = 0; i < mfactory_Group.Factory_GroupsID.Count; i++)
                            //    {
                            //        Group_ID = mfactory_Group.Factory_GroupsID[i];
                            //        query = "insert into product_factory_group (Product_ID,Group_ID,Factory_ID) values (@Product_ID,@Group_ID,@Factory_ID)";
                            //        com = new MySqlCommand(query, dbconnection);
                            //        com.Parameters.Add("@Product_ID", MySqlDbType.Int16);
                            //        com.Parameters["@Product_ID"].Value = productUpdateId;
                            //        com.Parameters.Add("@Group_ID", MySqlDbType.Int16);
                            //        com.Parameters["@Group_ID"].Value = Group_ID;
                            //        com.Parameters.AddWithValue("@Factory_ID", FactoryID);

                            //        com.ExecuteNonQuery();
                            //    }
                            //}

                            UserControl.ItemRecord("product", "تعديل", id , DateTime.Now,"", dbconnection);

                            txtProduct.Text = "";
                            txtType2.Text = "";
                            txtFactory2.Text = "";
                            txtGroup2.Text = "";
                            comTypeProduct.Text = "";
                            comFactoryGroup.Text = "";
                            comGroup.Text = "";
                            txtGroup.Text = "";

                            comGroup.Visible = true;
                            txtGroup2.Visible = true;
                            label11.Visible = true;
                            comFactoryGroup.Visible = true;
                            txtFactory2.Visible = true;
                            label5.Visible = true;

                            btnSave_productUpdate.Visible = false;
                            chListBoxGroup.Visible = false;
                            labGroup.Visible = false;
                            dataGridViewProduct.DataSource = null;

                            comTypeProduct.Enabled = true;
                            txtType2.ReadOnly = false;

                            displayProductAll();
                           
                        }
                        else
                        {
                            txtProduct.Focus();

                        }
                    }
                    else
                    {
                        MessageBox.Show("This product already exist");
                    }
                }
                else
                {
                    MessageBox.Show("select group");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void chListBoxGroup_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            try
            {
                if (btnSave_productUpdate.Visible == true&& flagItemCheckGroup)
                {
                    int FactoryID = Convert.ToInt32(txtFactory2.Text);
                    List<int> idList = new List<int>();
                    factory_Group mfactory_Group = new factory_Group();
                    foreach (factory_Group item in listFactory_Group)
                    {
                        if (item.Factory_ID == FactoryID)
                        {
                            mfactory_Group = item;
                            break;
                        }
                    }
                    if (e.CurrentValue == CheckState.Unchecked)
                        mfactory_Group.Factory_GroupsID.Add(Convert.ToInt32(chListBoxGroup.Items[e.Index].ToString().Split('\t')[1]));
                    else
                        mfactory_Group.Factory_GroupsID.Remove(Convert.ToInt32(chListBoxGroup.Items[e.Index].ToString().Split('\t')[1]));

                    for (int i = 0; i < listFactory_Group.Count; i++)
                    {
                        if (listFactory_Group[i].Factory_ID == FactoryID)
                        {
                            listFactory_Group[i].Factory_GroupsID = mfactory_Group.Factory_GroupsID;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < chListBoxGroup.Items.Count; i++)
                {
                    chListBoxGroup.SetItemChecked(i,true);
                }
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region color panel
        private void btnColor_Click(object sender, EventArgs e)
        {
            try
            {
                if (comType2.Text != "")
                {
                    dbconnection.Open();
                    string query = "select Color_ID from color where Color_Name = '" + txtColor.Text + "' and Type_ID=" + Convert.ToInt32(comType2.SelectedValue);
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    if (com.ExecuteScalar() == null)
                    {
                        if (txtColor.Text != "")
                        {
                            query = "insert into color (Color_Name,Type_ID) values (@name,@id)";
                            com = new MySqlCommand(query, dbconnection);
                            com.Parameters.AddWithValue("@name", txtColor.Text);
                            com.Parameters.AddWithValue("@id", Convert.ToInt32(comType2.SelectedValue));
                            com.ExecuteNonQuery();
                            query = "select Color_ID from color order by Color_ID desc limit 1";
                            com = new MySqlCommand(query, dbconnection);
                            UserControl.ItemRecord("color", "اضافة",Convert.ToInt32(com.ExecuteScalar().ToString()), DateTime.Now,"", dbconnection);

                            displayColor(Convert.ToInt32(comType2.SelectedValue));
                            txtColor.Text = "";
                        }
                        else
                        {
                            txtColor.Focus();
                            
                        }
                    }
                    else
                    {
                        MessageBox.Show("This Color already exist");
                    }
                }
                else
                {
                    MessageBox.Show("select Type");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void txtColor_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (btnSaveUpdateColor.Visible == false)
                    {
                        if (comType2.Text != "")
                        {
                            dbconnection.Open();
                            string query = "select Color_ID from color where Color_Name = '" + txtColor.Text + "' and Type_ID=" + Convert.ToInt32(comType2.SelectedValue);
                            MySqlCommand com = new MySqlCommand(query, dbconnection);
                            if (com.ExecuteScalar() == null)
                            {
                                if (txtColor.Text != "")
                                {
                                    query = "insert into color (Color_Name,Type_ID) values (@name,@id)";
                                    com = new MySqlCommand(query, dbconnection);
                                    com.Parameters.AddWithValue("@name", txtColor.Text);
                                    com.Parameters.AddWithValue("@id", Convert.ToInt32(comType2.SelectedValue));
                                    com.ExecuteNonQuery();
                                    query = "select Color_ID from color order by Color_ID desc limit 1";
                                    com = new MySqlCommand(query, dbconnection);
                                    UserControl.ItemRecord("color", "اضافة",Convert.ToInt32(com.ExecuteScalar().ToString()), DateTime.Now,"", dbconnection);

                                    displayColor(Convert.ToInt32(comType2.SelectedValue));
                                    txtColor.Text = "";
                                    txtColor.Focus();
                                }
                                else
                                {
                                    txtColor.Focus();
                                }
                            }
                            else
                            {
                                MessageBox.Show("This Color already exist");
                            }
                        }
                        else
                        {
                            MessageBox.Show("select Type");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnUpdateColor_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string type="";
                btnSaveUpdateColor.Visible = true;
                row1 = dataGridViewColor.Rows[dataGridViewColor.SelectedCells[0].RowIndex];
              
                if (row1 != null)
                {
                    colorUpdateId = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    txtColor.Text = row1.Cells[1].Value.ToString();
                    dbconnection.Open();
                    
                    string query = "select Type_ID from color where Color_ID=" + colorUpdateId;
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    type = com.ExecuteScalar().ToString();
                    txtType3.Text =type;
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnSaveUpdateColor_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string query = "select Color_ID from color where Color_Name = '" + txtColor.Text + "' and Color_ID not in(" + colorUpdateId + ")";
                MySqlCommand com = new MySqlCommand(query, dbconnection);
                if (com.ExecuteScalar() == null)
                {
                    if (txtColor.Text != "")
                    {
                        query = "update color set Color_Name=@name ,Type_ID=@Type_ID  where Color_ID=" + colorUpdateId;
                        com = new MySqlCommand(query, dbconnection);
                        com.Parameters.AddWithValue("@name", txtColor.Text);
                        com.Parameters.AddWithValue("@Type_ID", txtType3.Text);
                        com.ExecuteNonQuery();
                    }
                }
                btnSaveUpdateColor.Visible = false;
                comType2.Text = "";
                txtType3.Text = "";
                txtColor.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void comType2_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    displayColor((int)comType2.SelectedValue);
                    txtColor.Focus();
                    txtColor.SelectAll();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void comType2_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (load)
                {
                    displayColor((int)comType2.SelectedValue);
                    txtColor.Focus();
                    txtColor.SelectAll();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }  
        private void btnDeleteColor_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                DataGridViewRow row1 = dataGridViewColor.Rows[dataGridViewColor.SelectedCells[0].RowIndex];

                if (row1 != null)
                {
                    int id = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    if (!producItemUsed("Color_ID", id))
                    {
                        DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete the item?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.Yes)
                        {
                            string query = "delete from color where Color_ID=" + id;
                            MySqlCommand com = new MySqlCommand(query, dbconnection);
                            com.ExecuteNonQuery();

                            updateTablesDB("color", "Color_ID", id);

                            UserControl.ItemRecord("color", "حذف", Convert.ToInt32(row1.Cells[0].Value), DateTime.Now, "", dbconnection);

                            if (comType2.Text != "")
                            {
                                displayColor(Convert.ToInt32(comType2.SelectedValue));
                            }
                            else
                            {
                                displayColor();
                            }

                            txtColor.Focus();
                        }
                        else if (dialogResult == DialogResult.No)
                        { }
                    }
                    else
                    {
                        MessageBox.Show("هذا العنصر تم استخدامه في تسجيل بنود لا يمكن حذفه.");
                    }
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch
            {
                MessageBox.Show("select row");
            }
            dbconnection.Close();
        }
        private void btnColorDisplayAll_Click(object sender, EventArgs e)
        {
            try
            {
                displayColor();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region size panel
        private void btnSize_Click(object sender, EventArgs e)
        {
            try
            {
                if (comFactory2.Text != "")
                {
                    dbconnection.Open();
                    string query = "select Size_ID from size where Size_Value = '" + txtSize.Text + "' and Factory_ID=" + Convert.ToInt32(comFactory2.SelectedValue) + " and Group_ID="+ txtGroup_Size.Text;
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    if (com.ExecuteScalar() == null)
                    {
                        if (txtSize.Text != "")
                        {
                            query = "insert into size (Size_Value,Factory_ID,Group_ID) values (@name,@id,@Group_ID)";
                            com = new MySqlCommand(query, dbconnection);
                            com.Parameters.AddWithValue("@name", txtSize.Text);
                            com.Parameters.AddWithValue("@id", Convert.ToInt32(txtFactory3.Text));
                            com.Parameters.AddWithValue("@Group_ID", Convert.ToInt32(txtGroup_Size.Text));
                            com.ExecuteNonQuery();
                            query = "select Product_ID from product order by Product_ID desc limit 1";
                            com = new MySqlCommand(query, dbconnection);
                            UserControl.ItemRecord("product", "اضافة",Convert.ToInt32(com.ExecuteScalar().ToString()), DateTime.Now,"", dbconnection);

                            displaySize_group();
                            txtSize.Text = "";
                            txtSize.Focus();
                        }
                        else 
                        {
                            txtSize.Focus();                          
                        }
                    }
                    else
                    {
                        MessageBox.Show("This Size already exist");
                    }
                }
                else
                {
                    MessageBox.Show("select Factory");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void txtSize_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter && btnSaveSizeUpdate.Visible == false)
                {
                    if (btnSaveSizeUpdate.Visible == false)
                    {
                        if (comFactory2.Text != "")
                        {
                            dbconnection.Open();
                            string query = "select Size_ID from size where Size_Value = '" + txtSize.Text + "' and Factory_ID=" + Convert.ToInt32(comFactory2.SelectedValue)+ " and Group_ID=" + txtGroup_Size.Text; ;
                            MySqlCommand com = new MySqlCommand(query, dbconnection);
                            if (com.ExecuteScalar() == null)
                            {
                                if (txtSize.Text != "")
                                {
                                    query = "insert into size (Size_Value,Factory_ID,Group_ID) values (@name,@id,@Group_ID)";
                                    com = new MySqlCommand(query, dbconnection);
                                    com.Parameters.AddWithValue("@name", txtSize.Text);
                                    com.Parameters.AddWithValue("@id", Convert.ToInt32(txtFactory3.Text));
                                    com.Parameters.AddWithValue("@Group_ID", Convert.ToInt32(txtGroup_Size.Text));
                                    com.ExecuteNonQuery();
                                    query = "select Product_ID from product order by Product_ID desc limit 1";
                                    com = new MySqlCommand(query, dbconnection);
                                    UserControl.ItemRecord("product", "اضافة", Convert.ToInt32(com.ExecuteScalar().ToString()), DateTime.Now, "", dbconnection);

                                    displaySize_group();
                                    txtSize.Text = "";
                                    txtSize.Focus();
                                }
                                else
                                {
                                    txtSize.Focus();

                                }
                            }
                            else
                            {
                                MessageBox.Show("This Size already exist");
                            }
                        }
                        else
                        {
                            MessageBox.Show("select Factory");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnSizeUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Close();
                dbconnection.Open();
                btnSaveSizeUpdate.Visible = true;
                string factory_ID="", group_ID="",size="";
                row1 = dataGridViewSize.Rows[dataGridViewSize.SelectedCells[0].RowIndex];
                size = row1.Cells[1].Value.ToString();
                if (row1 != null)
                {
                    SizeUpdateId = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    string query = "select Factory_ID,Group_ID from size where Size_ID=" + SizeUpdateId;
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    MySqlDataReader dr = com.ExecuteReader();
                    while (dr.Read())
                    {
                        factory_ID = dr[0].ToString();
                        group_ID = dr[1].ToString();
                    }
                    dr.Close();
                    txtFactory3.Text = factory_ID;
                    txtGroup_Size.Text = group_ID;
                    txtSize.Text = size;
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnSaveSizeUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string query = "select Size_ID from size where Size_Value = '" + txtSize.Text + "' and Size_ID not in(" + SizeUpdateId + ") and Factory_ID="+ txtFactory3.Text + " and Group_ID="+ txtGroup_Size.Text;
                MySqlCommand com = new MySqlCommand(query, dbconnection);
                if (com.ExecuteScalar() == null)
                {
                    if (txtSize.Text != "")
                    {
                        query = "update size set Size_Value=@name ,Factory_ID=@Factory_ID,Group_ID=@Group_ID  where Size_ID=" + SizeUpdateId;
                        com = new MySqlCommand(query, dbconnection);
                        com.Parameters.AddWithValue("@name", txtSize.Text);
                        com.Parameters.AddWithValue("@Factory_ID", txtFactory3.Text);
                        com.Parameters.AddWithValue("@Group_ID", txtGroup_Size.Text);
                        com.ExecuteNonQuery();
                    }
                }
                btnSaveSizeUpdate.Visible = false;
                comFactory2.Text = "";
                txtFactory3.Text = "";
                comGroup_Size.Text = "";
                txtGroup_Size.Text = "";
                txtSize.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void comFactory2_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    displaySize_factory((int)comFactory2.SelectedValue);
                    filterGroup_Factory((int)comFactory2.SelectedValue);
                    comGroup_Size.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void comFactory2_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Close();
                dbconnection.Open();
                if (load)
                {
                  
                    filterGroup_Factory((int)comFactory2.SelectedValue);
                    displaySize_factory((int)comFactory2.SelectedValue);
                    comGroup_Size.Focus();
                  
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void comGroup_Size_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtFactory3.Text != ""&& Convert.ToInt32(comGroup_Size.SelectedValue.ToString())>0)
                {
                   
                    txtGroup_Size.Text = comGroup_Size.SelectedValue.ToString();
                    displaySize_group();
                    txtSize.Focus();
                }
            }
            catch
            {
               // MessageBox.Show(ex.Message);
            }
        }
        private void btnDisplayAll_Click(object sender, EventArgs e)
        {
            try
            {
                displaySize();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnDeleteSize_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                DataGridViewRow row1 = dataGridViewSize.Rows[dataGridViewSize.SelectedCells[0].RowIndex];

                if (row1 != null)
                {
                    int id = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    if (!producItemUsed("Size_ID", id))
                    {
                        DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete the item?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.Yes)
                        {
                            string query = "delete from size where Size_ID=" + id;
                            MySqlCommand com = new MySqlCommand(query, dbconnection);
                            com.ExecuteNonQuery();

                            updateTablesDB("size", "Size_ID", id);

                            UserControl.ItemRecord("size", "حذف", Convert.ToInt32(row1.Cells[0].Value.ToString()), DateTime.Now, "", dbconnection);

                            if (comFactory2.Text != "" && comGroup_Size.Text != "")
                            {
                                displaySize_group();
                            }
                            else
                            {
                                displaySize();
                            }

                            txtSize.Focus();

                        }
                        else if (dialogResult == DialogResult.No)
                        {

                        }
                    }
                    else
                    {
                        MessageBox.Show("هذا العنصر تم استخدامه في تسجيل بنود لا يمكن حذفه.");
                    }
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        #endregion

        #region sort panel
        private void btnSort_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string query = "select Sort_ID from sort where Sort_Value = '" + txtSort.Text + "'";
                MySqlCommand com = new MySqlCommand(query, dbconnection);
                if (com.ExecuteScalar() == null)
                {
                    if (txtSort.Text != "")
                    {
                        query = "insert into sort (Sort_Value) values (@name)";
                        com = new MySqlCommand(query, dbconnection);
                        com.Parameters.AddWithValue("@name", txtSort.Text);
                        com.ExecuteNonQuery();
                        displaySort();
                        txtSort.Text = "";
                        txtSort.Focus();
                    }
                    else
                    {
                        txtSort.Focus();
                     
                    }
                }
                else
                {
                    MessageBox.Show("This Sort already exist");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void txtSort_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (btnSaveSortUpdate.Visible == false)
                    {
                        dbconnection.Open();
                        string query = "select Sort_ID from sort where Sort_Value = '" + txtSort.Text + "'";
                        MySqlCommand com = new MySqlCommand(query, dbconnection);
                        if (com.ExecuteScalar() == null)
                        {
                            if (txtSort.Text != "")
                            {
                                query = "insert into sort (Sort_Value) values (@name)";
                                com = new MySqlCommand(query, dbconnection);
                                com.Parameters.AddWithValue("@name", txtSort.Text);
                                com.ExecuteNonQuery();

                                displaySort();
                                txtSort.Text = "";
                                txtSort.Focus();
                            }
                            else
                            {
                                txtSort.Focus();

                            }
                        }
                        else
                        {
                            MessageBox.Show("This Sort already exist");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnUpdateSort_Click(object sender, EventArgs e)
        {

            try
            {
                dbconnection.Open();
                btnSaveSortUpdate.Visible = true;
                row1 = dataGridViewSort.Rows[dataGridViewSort.SelectedCells[0].RowIndex];
                if (row1 != null)
                {
                    SortUpdateId = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    string query = "select Sort_Value from sort where Sort_ID=" + SortUpdateId;
                    MySqlCommand com = new MySqlCommand(query, dbconnection);
                    txtSort.Text = com.ExecuteScalar().ToString();

                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnSaveSortUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                string query = "select Sort_ID from sort where Sort_Value = '" + txtSort.Text + "' and Sort_ID not in(" + SortUpdateId + ")";
                MySqlCommand com = new MySqlCommand(query, dbconnection);
                if (com.ExecuteScalar() == null)
                {
                    if (txtSort.Text != "")
                    {
                        query = "update sort set Sort_Value=@name   where Sort_ID=" + SortUpdateId;
                        com = new MySqlCommand(query, dbconnection);
                        com.Parameters.AddWithValue("@name", txtSort.Text);
                        com.ExecuteNonQuery();
                    }
                }
                btnSaveSortUpdate.Visible = false;
                txtSort.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnDeleteSort_Click(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                //  DataRowView row1 = (DataRowView)(((GridView)dataGridViewSort.MainView).GetRow(((GridView)dataGridViewSort.MainView).GetSelectedRows()[0]));
                DataGridViewRow row1 = dataGridViewSort.Rows[dataGridViewSort.SelectedCells[0].RowIndex];
                row1.Selected = true;
                if (row1 != null)
                {
                    int id = Convert.ToInt32(row1.Cells[0].Value.ToString());
                    if (!producItemUsed("Sort_ID", id))
                    {
                        DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete the item?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.Yes)
                        {

                            string query = "delete from sort where Sort_ID=" + id;
                            MySqlCommand com = new MySqlCommand(query, dbconnection);
                            com.ExecuteNonQuery();

                            updateTablesDB("sort", "Sort_ID", id);

                            UserControl.ItemRecord("sort", "حذف", (int)row1.Cells[0].Value, DateTime.Now, "", dbconnection);

                            displaySort();
                            txtSort.Focus();
                        }
                        else if (dialogResult == DialogResult.No)
                        { }
                    }
                    else
                    {
                        MessageBox.Show("هذا العنصر تم استخدامه في تسجيل بنود لا يمكن حذفه.");
                    }
                }
                else
                {
                    MessageBox.Show("select row");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void btnSortDisplayAll_Click(object sender, EventArgs e)
        {
            try
            {
                displaySort();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        //function
        #region Type Tap
        public void displayType()
        {
            string query = "select distinct Type_ID as 'كود',Type_Name as 'النوع' from type";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            query = "select distinct * from type";
            da = new MySqlDataAdapter(query, dbconnection);
            dt = new DataTable();
            da.Fill(dt);
            comType2.DataSource = dt;
            comType2.DisplayMember = dt.Columns["Type_Name"].ToString();
            comType2.ValueMember = dt.Columns["Type_ID"].ToString();
        }
        #endregion

        #region Factory Tap
        public void displayFactory()
        {
            string query = "select distinct factory.Factory_ID as 'كود',Factory_Name as 'المصنع' from factory inner join type_factory on factory.Factory_ID=type_factory.Factory_ID inner join type  on type.Type_ID=type_factory.Type_ID order by factory.Factory_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewFactory.DataSource = null;
            dataGridViewFactory.DataSource = dt;
            dataGridViewFactory.Columns[0].Width = 50;
            dataGridViewFactory.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        public void displayFactory(int id)
        {
            string query = "select distinct Factory_ID as 'كود',Factory_Name as 'المصنع' from factory where Type_ID=" + id + " order by Factory_ID ";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewFactory.DataSource = null;
            dataGridViewFactory.DataSource = dt;
            dataGridViewFactory.Columns[0].Width = 50;
            dataGridViewFactory.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            query = "select distinct * from factory";
            da = new MySqlDataAdapter(query, dbconnection);
            dt = new DataTable();
            da.Fill(dt);
            comFactory.DataSource = dt;
            comFactory.DisplayMember = dt.Columns["Factory_Name"].ToString();
            comFactory.ValueMember = dt.Columns["Factory_ID"].ToString();

            da = new MySqlDataAdapter(query, dbconnection);
            dt = new DataTable();
            da.Fill(dt);
            comFactory2.DataSource = dt;
            comFactory2.DisplayMember = dt.Columns["Factory_Name"].ToString();
            comFactory2.ValueMember = dt.Columns["Factory_ID"].ToString();
            comFactory2.Text = "";

            da = new MySqlDataAdapter(query, dbconnection);
            dt = new DataTable();
            da.Fill(dt);
            comFactoryGroup.DataSource = dt;
            comFactoryGroup.DisplayMember = dt.Columns["Factory_Name"].ToString();
            comFactoryGroup.ValueMember = dt.Columns["Factory_ID"].ToString();
            comFactoryGroup.Text = "";
        }
        #endregion

        #region Group Tap
        public void displayGroup_Type(int type_id)
        {
            string query = "select distinct Group_ID as 'كود',Group_Name as 'المجموعة' from groupo left join factory on factory.Factory_ID=groupo.Factory_ID  where groupo.Type_ID=" + type_id + " order by Group_ID";
            //if (type_id==2|| type_id==1)
            //     query = "select distinct Group_ID as 'كود',Group_Name as 'المجموعة' from groupo left join factory on factory.Factory_ID=groupo.Factory_ID  where groupo.Factory_ID=" + -1 + " order by Group_ID";

            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewGroup.DataSource = dt;
            dataGridViewGroup.Columns[0].Width = 50;
            dataGridViewGroup.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        public void displayGroup_Factory(int factory_id)
        {
            string query = "select distinct Group_ID as 'كود',Group_Name as 'المجموعة' from groupo left join factory on factory.Factory_ID=groupo.Factory_ID  where groupo.Type_ID=" + txtType1.Text + " and groupo.Factory_ID=" + factory_id + "  order by Group_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewGroup.DataSource = dt;
            dataGridViewGroup.Columns[0].Width = 50;
            dataGridViewGroup.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        public void displayGroup(int id)
        {
            string query = "";
            //if (txtType1.Text=="2"|| txtType1.Text == "1")
            //  query = "select distinct Group_ID as 'كود',Group_Name as 'المجموعة' from groupo   where  groupo.Factory_ID=-1 order by Group_ID";
            //else
                query = "select distinct Group_ID as 'كود',Group_Name as 'المجموعة' from groupo   where groupo.Type_ID=" + txtType1.Text + " and groupo.Factory_ID=" + id + " order by Group_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewGroup.DataSource = dt;
            dataGridViewGroup.Columns[0].Width = 50;
            dataGridViewGroup.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            query = "select distinct * from groupo";
            da = new MySqlDataAdapter(query, dbconnection);
            dt = new DataTable();
            da.Fill(dt);
            comGroup.DataSource = dt;
            comGroup.DisplayMember = dt.Columns["Group_Name"].ToString();
            comGroup.ValueMember = dt.Columns["Group_ID"].ToString();
            comGroup.Text = "";
        }
        public void displayGroup()
        {
            string query = "select distinct Group_ID as 'كود',Group_Name as 'المجموعة' from groupo left join factory on factory.Factory_ID=groupo.Factory_ID inner join type on type.Type_ID=groupo.Type_ID order by Group_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewGroup.DataSource = dt;
            dataGridViewGroup.Columns[0].Width = 50;
            dataGridViewGroup.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        #endregion

        #region functions of product tap    
        public void displayProductAll()
        {
            string query = "select distinct  product.Product_ID as 'كود', Product_Name as 'الصنف' from product inner join product_factory_group on product.Product_ID=product_factory_group.Product_ID   inner join  groupo on product_factory_group.Group_ID=groupo.Group_ID inner join  factory on product_factory_group.Factory_ID=factory.Factory_ID inner join  type on product.Type_ID=type.Type_ID   order by product.Product_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewProduct.DataSource = dt;
            dataGridViewProduct.Columns[0].Width = 50;
            dataGridViewProduct.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        public object displayProduct_Type(int type_id)
        {
            string query = "select distinct product.Product_ID as 'كود' ,Product_Name as 'الصنف' from product inner join product_factory_group on product.Product_ID=product_factory_group.Product_ID  inner join  groupo on product_factory_group.Group_ID=groupo.Group_ID  inner join  factory on product_factory_group.Factory_ID=factory.Factory_ID inner join  type on product.Type_ID=type.Type_ID where product.Type_ID=" + type_id + " order by product.Product_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public object displayProduct_Factory(int factory_id)
        {
            string query = "select distinct  product.Product_ID as 'كود' ,Product_Name as 'الصنف' from product join product_factory_group on product.Product_ID=product_factory_group.Product_ID  inner join  groupo on product_factory_group.Group_ID=groupo.Group_ID   where product.Type_ID=" + txtType2.Text + " and product_factory_group.Factory_ID=" + factory_id+ " order by product.Product_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public void displayProduct_Group(int groupID)
        {
            if (txtType2.Text != "" && txtFactory2.Text != "" && (txtGroup2.Text != ""||chListBoxGroup.CheckedItems.Count>0))
            {
                string query = "select distinct  product.Product_ID as 'كود' ,Product_Name as 'الصنف' from product inner join product_factory_group on product.Product_ID=product_factory_group.Product_ID   where product.Type_ID=" + txtType2.Text + " and product_factory_group.Factory_ID=" + txtFactory2.Text + " and product_factory_group.Group_ID=" + groupID + "  order by product.Product_ID";
                MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridViewProduct.DataSource = dt;
                dataGridViewProduct.Columns[0].Width = 50;
                dataGridViewProduct.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
        }
        public void displayProduct()
        {
            string query = "select distinct Product_ID as 'كود' ,Product_Name as 'الصنف' from product join product_factory_group on product.Product_ID=product_factory_group.Product_ID   where product.Type_ID=" + txtType2.Text + " and product_factory_group.Factory_ID=" + txtFactory2.Text + " and  product_factory_group.Group_ID=" + txtGroup2.Text + " order by product.Product_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewProduct.DataSource = dt;
            dataGridViewProduct.Columns[0].Width = 50;
            dataGridViewProduct.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        #endregion

        #region The optional Taps
        public void displayColor(int id)
        {
            string query = "select distinct Color_ID as 'كود',Color_Name as 'اللون' from color where Type_ID=" + id + " order by Color_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewColor.DataSource = dt;
            dataGridViewColor.Columns[0].Width = 50;
            dataGridViewColor.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        public void displayColor()
        {
            string query = "select distinct Color_ID as 'كود',Color_Name as 'اللون' from color inner join type on color.Type_ID=type.Type_ID  order by Color_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewColor.DataSource = dt;
            dataGridViewColor.Columns[0].Width = 50;
            dataGridViewColor.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        public void displaySize_factory(int id)
        {
            string query = "select distinct Size_ID as 'كود',Size_Value as 'المقاس' from size inner join groupo on groupo.Group_ID=size.Group_ID where size.Factory_ID=" + id + " order by Size_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewSize.DataSource = dt;
            dataGridViewSize.Columns[0].Width = 50;
            dataGridViewSize.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        public void displaySize_group()
        {
            string query = "select distinct Size_ID as 'كود',Size_Value as 'المقاس' from size  where size.Factory_ID="+txtFactory3.Text+" and size.Group_ID="+txtGroup_Size.Text+" order by Size_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewSize.DataSource = dt;
            dataGridViewSize.Columns[0].Width = 50;
            dataGridViewSize.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        public void displaySize()
        {
            string query = "select distinct Size_ID as 'كود',Size_Value as 'المقاس' from size inner join factory on factory.Factory_ID=size.Factory_ID inner join groupo on groupo.Group_ID=size.Group_ID order by Size_ID";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewSize.DataSource = dt;
            dataGridViewSize.Columns[0].Width = 50;
            dataGridViewSize.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        public void filterGroup_Factory(int id)
        {
            string query = "",q="" ;
            List<int> ids = new List<int>();
            string fQuery = "select TypeCoding_Method,type_Factory.Type_ID from type_Factory inner join type on type.Type_ID=type_Factory.Type_ID where Factory_ID=" + id;
            MySqlCommand com =new MySqlCommand(fQuery, dbconnection);
            MySqlDataReader dr = com.ExecuteReader();
            while (dr.Read())
            {
                if (dr[0].ToString() == "1")
                {
                    //if (dr[1].ToString() == "1" || dr[1].ToString() == "2")
                    //{
                    //    ids.Add(-1);
                    //}
                    //else
                    //{
                        ids.Add(-Convert.ToInt32(dr[1].ToString()));
                    //}
                }
                else
                {
                    ids.Add(id);
                }
            }
            dr.Close();
            for (int i = 0; i < ids.Count-1; i++)
            {
                q += ids[i] + ",";
            }
            q += ids[ids.Count - 1] + "";
                query = "select * from groupo where Factory_ID in (" + q + ")";
            
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            comGroup_Size.DataSource = dt;
            comGroup_Size.DisplayMember = dt.Columns["Group_Name"].ToString();
            comGroup_Size.ValueMember = dt.Columns["Group_ID"].ToString();
            comGroup_Size.Text = "";
            txtGroup_Size.Text = "";     
        }
        public void displaySort()
        {
            string query = "select Sort_ID as 'كود',Sort_Value as 'الفرز' from sort";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewSort.DataSource = dt;
        }
        #endregion

        #region Hellper Functions
        public void RestControls()
        {
            comFactory.Text = "";
            comFactory2.Text = "";
            comFactoryGroup.Text = "";
            comType.Text = "";
            comType2.Text = "";
            comTypeProduct.Text = "";
            comGroup.Text = "";
            comGroup_Size.Text = "";

            txtType.Text = "";
            txtType2.Text = "";
            txtType3.Text = "";
            txtSort.Text = "";
            txtProduct.Text = "";
            txtGroup.Text = "";
            txtGroup2.Text = "";
            txtGroup_Size.Text = "";
            txtFactory.Text = "";
            txtFactory2.Text = "";
            txtFactory3.Text = "";
            txtColor.Text = "";
            txtSize.Text = "";
            dataGridViewSize.DataSource = null;
            dataGridViewSort.DataSource = null;
            dataGridViewColor.DataSource = null;
            dataGridView1.DataSource = null;
            dataGridViewFactory.DataSource = null;
            dataGridViewGroup.DataSource = null;
            dataGridViewProduct.DataSource = null;
        }
        public void updateTablesDB(string tableName, string ColumnName, int id)
        {
            string query = "ALTER TABLE " + tableName + " AUTO_INCREMENT = 1;";
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            com.ExecuteNonQuery();
        }
        public void ClearButtonsColor()
        {
            foreach (Control control in this.tableLayoutPanel2.Controls)
            {
                if (control is Button)
                    control.BackColor = Color.FromArgb(229, 229, 229);
            }
        }
        public void dataGridViewStyle()
        {
            foreach (DataGridView item in this.DataGridViewList)
            {
                DataGridViewCellStyle style = new DataGridViewCellStyle(dataGridView1.ColumnHeadersDefaultCellStyle);
                style.ForeColor = Color.AliceBlue;
                item.ColumnHeadersDefaultCellStyle = style;
                item.RowHeadersDefaultCellStyle = style;
            }
        }
        #endregion
        
        public void fun()
        {
            dbconnection.Close();
            dbconnection.Open();
            // string queryx = "select distinct * from factory  inner join  product_factory_group on product_factory_group.Factory_ID=factory.Factory_ID group by product_factory_group.Factory_ID having  product_factory_group.Product_ID=" + productUpdateId;
            string queryx = "select distinct * from product_factory_group where Factory_ID="+txtFactory2.Text+" and  Product_ID=" + productUpdateId;

            MySqlCommand com = new MySqlCommand(queryx, dbconnection);
            MySqlDataReader dr = com.ExecuteReader();
            List<int> idList = new List<int>();
            int count = 0;
            while (dr.Read())
            {
                idList.Add(Convert.ToInt32(dr["Group_ID"].ToString()));
                count++;
            }
            dr.Close();
            for (int i = 0; i < chListBoxGroup.Items.Count; i++)
            {
                int GroupID = Convert.ToInt32(chListBoxGroup.Items[i].ToString().Split('\t')[1]);
                for (int j = 0; j < idList.Count; j++)
                {
                    if (GroupID == idList[j])
                    {
                        chListBoxGroup.SetItemChecked(i, true);
                    }
                }
            }
            dbconnection.Close();

        }
        public void setProductFactoriesGroups()
        {
            listFactory_Group.Clear();
            string query = "select DISTINCT Factory_ID from product_factory_group where Product_ID=" + productUpdateId;
            MySqlCommand com = new MySqlCommand(query,dbconnection);
            MySqlDataReader dr = com.ExecuteReader();
            while (dr.Read())
            {
                factory_Group factory_Group = new factory_Group(Convert.ToInt32(dr[0].ToString()),new List<int>());
                listFactory_Group.Add(factory_Group);
            }
            dr.Close();
            foreach (factory_Group item in listFactory_Group)
            {
                factory_Group mfactory_Group = listFactory_Group[listFactory_Group.IndexOf(item)];
                query = "select DISTINCT Group_ID from product_factory_group where Product_ID=" + productUpdateId+ " and Factory_ID="+item.Factory_ID;
                com = new MySqlCommand(query, dbconnection);
                dr = com.ExecuteReader();
                while (dr.Read())
                {
                    mfactory_Group.Factory_GroupsID.Add(Convert.ToInt32(dr[0].ToString()));

                }
                dr.Close();
            }
        }
        public bool checkGroupsForAllFactories()
        {
            foreach (factory_Group item in listFactory_Group)
            {
                if (item.Factory_GroupsID.Count == 0)
                    return false;
            }
            return true;
        }
        private void btnNewChoice_Click(object sender, EventArgs e)
        {
            try
            {
                txtProduct.Text = "";
                txtType2.Text = "";
                txtFactory2.Text = "";
                txtGroup2.Text = "";
                comTypeProduct.Text = "";
                comFactoryGroup.Text = "";
                comGroup.Text = "";
                txtGroup.Text = "";

                comGroup.Visible = true;
                txtGroup2.Visible = true;
                label11.Visible = true;
                comFactoryGroup.Visible = true;
                txtFactory2.Visible = true;
                label5.Visible = true;

                btnSave_productUpdate.Visible = false;
                chListBoxGroup.Visible = false;
                labGroup.Visible = false;
                dataGridViewProduct.DataSource = null;

                comTypeProduct.Enabled = true;
                txtType2.ReadOnly = false;

                displayProductAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void DeletePermission()
        {
            if (UserControl.userType != 1)
            {
                btnDelete.Enabled = false;
                btnTypeDelete.Enabled = false;
                btnFactoryDelete.Enabled = false;
                btnDeleteSize.Enabled = false;
                btnDeleteGroup.Enabled = false;
                btnDeleteProduct.Enabled = false;
                bunifuTileButton1.Enabled = false;
            }
        }
        public bool producItemUsed(string columnName, int ProductItem_ID)
        {
            string query = "select Data_ID from data where " + columnName + "=" + ProductItem_ID + " limit 1";
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            if (com.ExecuteScalar() != null)
            {
                return true;
            }
            return false;
        }

    }

    public class factory_Group
    {
        public int Factory_ID { get; set; }
        public  List<int> Factory_GroupsID { get; set; }

        public factory_Group()
        {
            Factory_GroupsID = new List<int>();
        }
        public factory_Group(int Factory_ID, List<int> Factory_GroupsID)
        {
            this.Factory_ID = Factory_ID;
            this.Factory_GroupsID = Factory_GroupsID;
        }
      
    }

 
}
