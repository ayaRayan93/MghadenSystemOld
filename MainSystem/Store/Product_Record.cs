using DevExpress.XtraTab;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MainSystem
{
    public partial class Product_Record : Form
    {
        byte[] selectedImage;
        MySqlConnection conn;
        bool loaded = false;
        bool factoryFlage = false;
        bool groupFlage = false;
        bool  flagProduct = false;
        Products products = null;
        XtraTabControl xtraTabControlStoresContent = null;

        public Product_Record(Products products, XtraTabControl xtraTabControlStoresContent)
        {
            try
            {
                InitializeComponent();
                conn = new MySqlConnection(connection.connectionString);
                selectedImage = null;
                this.products = products;
                this.xtraTabControlStoresContent = xtraTabControlStoresContent;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Product_Record_Load(object sender, EventArgs e)
        {
            try
            {
                conn.Open();

                string query = "select * from type";
                MySqlDataAdapter da = new MySqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                comType.DataSource = dt;
                comType.DisplayMember = dt.Columns["Type_Name"].ToString();
                comType.ValueMember = dt.Columns["Type_ID"].ToString();
                comType.Text = "";
                txtType.Text = "";

                query = "select * from sort";
                da = new MySqlDataAdapter(query, conn);
                dt = new DataTable();
                da.Fill(dt);
                comSort.DataSource = dt;
                comSort.DisplayMember = dt.Columns["Sort_Value"].ToString();
                comSort.ValueMember = dt.Columns["Sort_ID"].ToString();
                comSort.Text = "";

                VScrollBar myScrollBar = new VScrollBar();

                myScrollBar.Height = panel1.Height;

                myScrollBar.Left = panel1.Width - myScrollBar.Width;

                myScrollBar.Top = 0;

                myScrollBar.Enabled = false;

                panel3.Controls.Add(myScrollBar);

                loaded = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            conn.Close();
        }

        private void comBox_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (loaded)
                {
                    conn.Close();
                    conn.Open();
                    ComboBox comBox = (ComboBox)sender;

                    switch (comBox.Name)
                    {
                        case "comType":
                            if (loaded)
                            {
                                txtType.Text = comType.SelectedValue.ToString();
                                string query = "select * from factory inner join type_factory on factory.Factory_ID=type_factory.Factory_ID inner join type on type_factory.Type_ID=type.Type_ID where type_factory.Type_ID=" + txtType.Text;
                                MySqlDataAdapter da = new MySqlDataAdapter(query, conn);
                                DataTable dt = new DataTable();
                                da.Fill(dt);
                                comFactory.DataSource = dt;
                                comFactory.DisplayMember = dt.Columns["Factory_Name"].ToString();
                                comFactory.ValueMember = dt.Columns["Factory_ID"].ToString();
                                comFactory.Text = "";
                                txtFactory.Text = "";
                                query = "select TypeCoding_Method from type where Type_ID=" + txtType.Text;
                                MySqlCommand com = new MySqlCommand(query, conn);
                                int TypeCoding_Method = (int)com.ExecuteScalar();
                                if (TypeCoding_Method==1)
                                {
                                    string query2 = "";
                                    //if (txtType.Text == "2"|| txtType.Text == "1")
                                    //{
                                    //    query2 = "select * from groupo where Factory_ID=-1";
                                    //}
                                    //else
                                    //{
                                        query2 = "select * from groupo where Factory_ID="+-Convert.ToInt32(txtType.Text) + " and Type_ID="+ txtType.Text;
                                    //}

                                    MySqlDataAdapter da2 = new MySqlDataAdapter(query2, conn);
                                    DataTable dt2 = new DataTable();
                                    da2.Fill(dt2);
                                    comGroup.DataSource = dt2;
                                    comGroup.DisplayMember = dt2.Columns["Group_Name"].ToString();
                                    comGroup.ValueMember = dt2.Columns["Group_ID"].ToString();
                                    comGroup.Text = "";
                                    txtGroup.Text = "";
                                }
                                factoryFlage = true;

                                query = "select * from color where Type_ID=" + txtType.Text;
                                da = new MySqlDataAdapter(query, conn);
                                dt = new DataTable();
                                da.Fill(dt);
                                comColour.DataSource = dt;
                                comColour.DisplayMember = dt.Columns["Color_Name"].ToString();
                                comColour.ValueMember = dt.Columns["Color_ID"].ToString();
                                comColour.Text = "";
                                txtColor.Text = "";

                                comProduct.Text = "";
                                txtProduct.Text = "";

                                comFactory.Focus();
                            }
                            break;
                        case "comFactory":
                            if (factoryFlage)
                            {
                                txtFactory.Text = comFactory.SelectedValue.ToString();
                                string query = "select TypeCoding_Method from type where Type_ID=" + txtType.Text;
                                MySqlCommand com = new MySqlCommand(query, conn);
                                int TypeCoding_Method = (int)com.ExecuteScalar();
                                if (TypeCoding_Method==2)
                                {
                                    string query2f = "select * from groupo where Type_ID="+txtType.Text+" and Factory_ID=" + txtFactory.Text;
                                    MySqlDataAdapter da2f = new MySqlDataAdapter(query2f, conn);
                                    DataTable dt2f = new DataTable();
                                    da2f.Fill(dt2f);
                                    comGroup.DataSource = dt2f;
                                    comGroup.DisplayMember = dt2f.Columns["Group_Name"].ToString();
                                    comGroup.ValueMember = dt2f.Columns["Group_ID"].ToString();
                                    comGroup.Text = "";
                                    txtGroup.Text = "";
                                }

                                groupFlage = true;

                                string query2 = "select * from size where Factory_ID=" + txtFactory.Text;
                                MySqlDataAdapter  da2 = new MySqlDataAdapter(query2, conn);
                                DataTable  dt2 = new DataTable();
                                da2.Fill(dt2);
                                comSize.DataSource = dt2;
                                comSize.DisplayMember = dt2.Columns["Size_Value"].ToString();
                                comSize.ValueMember = dt2.Columns["Size_ID"].ToString();
                                comSize.Text = "";
                                txtSize.Text = "";

                                comProduct.Text = "";
                                txtProduct.Text = "";

                                comGroup.Focus();
                            }
                            break;
                        case "comGroup":
                            if (groupFlage)
                            {
                                txtGroup.Text = comGroup.SelectedValue.ToString();

                                string query3 = "select distinct  product.Product_ID  ,Product_Name  from product inner join product_factory_group on product.Product_ID=product_factory_group.Product_ID  where product.Type_ID=" + txtType.Text + " and product_factory_group.Factory_ID=" + txtFactory.Text + " and product_factory_group.Group_ID=" + txtGroup.Text + "  order by product.Product_ID";
                                MySqlDataAdapter da3 = new MySqlDataAdapter(query3, conn);
                                DataTable dt3 = new DataTable();
                                da3.Fill(dt3);
                                comProduct.DataSource = dt3;
                                comProduct.DisplayMember = dt3.Columns["Product_Name"].ToString();
                                comProduct.ValueMember = dt3.Columns["Product_ID"].ToString();
                                comProduct.Text = "";
                                txtProduct.Text = "";
                                
                                string query2 = "select * from size where Factory_ID=" + txtFactory.Text+" and Group_ID="+txtGroup.Text;
                                MySqlDataAdapter da2 = new MySqlDataAdapter(query2, conn);
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
                            if (flagProduct)
                            {
                                flagProduct = false;
                                txtProduct.Text = comProduct.SelectedValue.ToString();
                                comColour.Focus();
                            }
                            break;

                        case "comColour":
                            txtColor.Text = comColour.SelectedValue.ToString();
                            comSize.Focus();
                            break;

                        case "comSize":
                            txtSize.Text = comSize.SelectedValue.ToString();
                            txtClassification.Focus();
                            break;

                        case "comSort":
                            txtSort.Text = comSort.SelectedValue.ToString();
                            txtCarton.Focus();
                            break;
                    }
                }
            }
            catch
            {
              //  MessageBox.Show(ex.Message);
            }
            conn.Close();
            conn.Open();
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
                    conn.Close();
                    conn.Open();
                    switch (txtBox.Name)
                    {
                        case "txtType":
                            query = "select Type_Name from type where Type_ID='" + txtType.Text + "'";
                            com = new MySqlCommand(query, conn);
                            if (com.ExecuteScalar() != null)
                            {
                                Name = (string)com.ExecuteScalar();
                                factoryFlage = true;
                                comBox_SelectedValueChanged(comType, e);
                                comType.Text = Name;
                                txtFactory.Focus();
                            }
                            else
                            {
                                MessageBox.Show("there is no item with this id");
                                txtType.Text = "";
                                comType.Text = "";
                                conn.Close();
                                return;
                            }
                            break;
                        case "txtFactory":
                            if (txtType.Text != "")
                            {
                                query = "select Factory_Name from factory inner join type_factory on type_factory.Factory_ID=factory.Factory_ID where factory.Factory_ID='" + txtFactory.Text + "' and type_factory.Type_ID=" + txtType.Text;
                                com = new MySqlCommand(query, conn);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    groupFlage = true;
                                    comBox_SelectedValueChanged(comFactory, e);
                                    comFactory.Text = Name;
                                    txtGroup.Focus();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    txtFactory.Text = "";
                                    comFactory.Text = "";
                                    conn.Close();
                                    return;
                                }
                            }
                            else
                            {
                                txtFactory.Text = "";
                            }
                                break;

                        case "txtGroup":
                            if (txtType.Text != "" && txtFactory.Text != "")
                            {
                                //if(txtType.Text=="2")
                                //    query = "select Group_Name from groupo inner join product_factory_group on product_factory_group.Group_ID=groupo.Group_ID where groupo.Group_ID='" + txtGroup.Text + "' and product_factory_group.Factory_ID=" + txtFactory.Text + " and Type_ID=1";
                                //else
                                    query = "select Group_Name from groupo inner join product_factory_group on product_factory_group.Group_ID=groupo.Group_ID where groupo.Group_ID='" + txtGroup.Text + "' and product_factory_group.Factory_ID=" + txtFactory.Text + " and Type_ID=" + txtType.Text;
                                com = new MySqlCommand(query, conn);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    flagProduct = true;
                                    comBox_SelectedValueChanged(comGroup, e);
                                    comGroup.Text = Name;
                                    txtProduct.Focus();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    txtGroup.Text = "";
                                    comGroup.Text = "";
                                    conn.Close();
                                    return;
                                }
                            }
                            else
                            {
                                txtGroup.Text = "";
                            }
                            break;
                        case "txtProduct":
                            if (txtType.Text != "" && txtFactory.Text != "" && txtGroup.Text != "")
                            {
                                query = "select Product_Name from product inner join product_factory_group on product_factory_group.Product_ID=product.Product_ID where product.Product_ID='" + txtProduct.Text + "' and Group_ID=" + txtGroup.Text + " and Factory_ID=" + txtFactory.Text + " and Type_ID=" + txtType.Text;
                                com = new MySqlCommand(query, conn);
                                if (com.ExecuteScalar() != null)
                                {
                                    Name = (string)com.ExecuteScalar();
                                    comBox_SelectedValueChanged(comProduct, e);
                                    comProduct.Text = Name;
                                    txtColor.Focus();
                                }
                                else
                                {
                                    MessageBox.Show("there is no item with this id");
                                    txtProduct.Text = "";
                                    comProduct.Text = "";
                                    conn.Close();
                                    return;
                                }
                            }
                            else
                            {
                                txtProduct.Text = "";
                            }
                            break;
                        case "txtColor":
                            query = "select Color_Name from color where Color_ID='" + txtColor.Text + "' and Type_ID='" + txtType.Text + "'";
                            com = new MySqlCommand(query, conn);
                            if (com.ExecuteScalar() != null)
                            {
                                Name = (string)com.ExecuteScalar();
                                comColour.Text = Name;
                                txtSize.Focus();
                            }
                            else
                            {
                                MessageBox.Show("there is no item with this id");
                                txtColor.Text = "";
                                comColour.Text = "";
                                conn.Close();
                                return;
                            }
                            break;
                        case "txtSize":
                            query = "select Size_Value from size where Size_ID='" + txtSize.Text + "' and Factory_ID='" + txtFactory.Text + "'";
                            com = new MySqlCommand(query, conn);
                            if (com.ExecuteScalar() != null)
                            {
                                Name = (string)com.ExecuteScalar();
                                comSize.Text = Name;
                                txtClassification.Focus();
                            }
                            else
                            {
                                MessageBox.Show("there is no item with this id");
                                txtSize.Text = "";
                                comSize.Text = "";
                                txtClassification.Focus();
                                conn.Close();
                                return;
                            }
                            break;
                        case "txtClassification":
                            txtDescription.Focus();

                            break;
                        case "txtDescription":
                            txtDescription.Focus();
                            txtSort.Focus();
                            break;
                        case "txtSort":
                            query = "select Sort_Value from sort where Sort_ID='" + txtSort.Text + "'";
                            com = new MySqlCommand(query, conn);
                            if (com.ExecuteScalar() != null)
                            {
                                Name = (string)com.ExecuteScalar();
                                comSort.Text = Name;
                                txtCarton.Focus();
                            }
                            else
                            {
                                MessageBox.Show("there is no item with this id");
                                txtSort.Text = "";
                                comSort.Text = "";
                                txtCarton.Focus();
                                conn.Close();
                                return;
                            }
                            break;
                        case "txtCarton":
                            DialogResult dialogResult = MessageBox.Show("هل تريد حفظ البيانات?", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                            if (dialogResult == DialogResult.OK)
                            {
                                saveData();
                            }
                            else
                            {
                                txtType.Focus();
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                conn.Close();
            }
        }

        private void btnAddProduct_Click(object sender, EventArgs e)
        {
            try
            {
                saveData();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            conn.Close();
        }

        private void btnUploadPhoto_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string selectedFile = openFileDialog1.FileName;

                    ImageProduct.Image = Image.FromFile(openFileDialog1.FileName);
                    selectedImage = File.ReadAllBytes(selectedFile);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private void ImageProduct_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string selectedFile = openFileDialog1.FileName;

                    ImageProduct.Image = Image.FromFile(openFileDialog1.FileName);
                    selectedImage = File.ReadAllBytes(selectedFile);
                }
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
                if (loaded)
                {
                    XtraTabPage xtraTabPage = getTabPage("أضافة بند");
                    if (!IsClear())
                        xtraTabPage.ImageOptions.Image = Properties.Resources.unsave;
                    else
                        xtraTabPage.ImageOptions.Image = null;
                }
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
                comType.Text = "";
                comFactory.Text = "";
                comGroup.Text = "";
                comProduct.Text = "";
                comColour.Text = "";
                comSort.Text = "";
                comSize.Text = "";

                txtType.Text = "";
                txtFactory.Text = "";
                txtGroup.Text = "";
                txtProduct.Text = "";
                txtColor.Text = "";
                txtSize.Text = "";
                txtSort.Text = "";
                txtClassification.Text = "";
                txtCarton.Text = "";
                txtDescription.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void saveData()
        {
            if (txtFactory.Text != "" && txtGroup.Text != "" && txtProduct.Text != "" && txtType.Text != "")
            {
                int color_id = 0;
                int size_id = 0;
                int sort_id = 0;
                double carton = 0;
                string classification, description = "";
                if (comColour.Text != "")
                {
                    color_id = Convert.ToInt32(txtColor.Text);
                }
                if (comSize.Text != "")
                {
                    size_id = Convert.ToInt32(txtSize.Text);
                }
                if (comSort.Text != "")
                {
                    sort_id = Convert.ToInt32(txtSort.Text);
                }
                if (txtCarton.Text != "")
                {
                    if (double.TryParse(txtCarton.Text, out carton))
                    {
                    }
                    else
                    {
                        MessageBox.Show("carton must be numeric");
                        return;
                    }
                }
                classification = txtClassification.Text;
                description = txtDescription.Text;

                string q = "SELECT Data_ID from data where Color_ID=" + color_id + " and Size_ID=" + size_id + " and Sort_ID=" + sort_id + " and Description='" + description + "' and Carton=" + carton + " and Type_ID=" + txtType.Text + " and Factory_ID=" + txtFactory.Text + " and Group_ID=" + txtGroup.Text + " and Product_ID=" + txtProduct.Text + " and Classification='" + classification + "'";
                MySqlCommand comand = new MySqlCommand(q, conn);
                conn.Close();
                conn.Open();
                var resultValue = comand.ExecuteReader();
                if (!resultValue.HasRows)
                {
                    conn.Close();

                    MySqlCommand command = conn.CreateCommand();

                    string code = txtType.Text;

                    int typecount = txtType.Text.Length;

                    int factorycount = txtFactory.Text.Length;

                    int groupcount = txtGroup.Text.Length;

                    int productcount = txtProduct.Text.Length;

                    while (typecount < 4)
                    {
                        code = "0" + code;
                        typecount++;
                    }

                    string code2 = txtFactory.Text;

                    while (factorycount < 4)
                    {
                        code2 = "0" + code2;
                        factorycount++;
                    }

                    code = code + code2;

                    string code3 = txtGroup.Text;

                    while (groupcount < 4)
                    {
                        code3 = "0" + code3;
                        groupcount++;
                    }

                    code = code + code3;

                    string code4 = txtProduct.Text;

                    while (productcount < 4)
                    {
                        code4 = "0" + code4;
                        productcount++;
                    }

                    code = code + code4;

                    string query2 = "SELECT IFNULL(max(Code),1) FROM data where Type_ID=" + txtType.Text + " and Factory_ID=" + txtFactory.Text + " and Group_ID=" + txtGroup.Text + " and Product_ID=" + txtProduct.Text;
                    conn.Open();
                    MySqlCommand adpt = new MySqlCommand(query2, conn);
                    string maxCode = adpt.ExecuteScalar().ToString();
                    char [] arrCode = maxCode.ToCharArray();
                    int result = 1;
                    if (maxCode != "1")
                    {
                        string part5 = arrCode[16].ToString() + arrCode[17].ToString() + arrCode[18].ToString() + arrCode[19].ToString() + "";
                        result = Convert.ToInt32(part5);
                        result = result + 1;
                    }
                    
                    conn.Close();

                    int resultcount = result.ToString().Length;

                    string code5 = result.ToString();

                    while (resultcount < 4)
                    {
                        code5 = "0" + code5;
                        resultcount++;
                    }

                    code = code + code5;

                    command.CommandText = "INSERT INTO data (Color_ID,Size_ID,Sort_ID,Description,Carton,Code,Type_ID,Factory_ID,Group_ID,Product_ID,Classification,Data_Date) VALUES (?Color_ID,?Size_ID,?Sort_ID,?Description,?Carton,?Code,?Type_ID,?Factory_ID,?Group_ID,?Product_ID,?Classification,?Data_Date)";

                    command.Parameters.AddWithValue("?Color_ID", color_id);
                    command.Parameters.AddWithValue("?Size_ID", size_id);
                    command.Parameters.AddWithValue("?Sort_ID", sort_id);
                    command.Parameters.AddWithValue("?Description", description);
                    command.Parameters.AddWithValue("?Carton", carton);
                    command.Parameters.AddWithValue("?Code", code);
                    command.Parameters.AddWithValue("?Type_ID", int.Parse(txtType.Text));
                    command.Parameters.AddWithValue("?Factory_ID", int.Parse(txtFactory.Text));
                    command.Parameters.AddWithValue("?Group_ID", int.Parse(txtGroup.Text));
                    command.Parameters.AddWithValue("?Product_ID", int.Parse(txtProduct.Text));
                    command.Parameters.AddWithValue("?Classification", classification);
                    command.Parameters.AddWithValue("?Data_Date", DateTime.Now.Date);
                    conn.Open();
                    command.ExecuteNonQuery();

                    query2 = "select Data_ID from data order by Data_ID desc limit 1";
                    comand = new MySqlCommand(query2, conn);
                    int data_ID = Convert.ToInt32(comand.ExecuteScalar());
                    //save image as bytes
                    command = conn.CreateCommand();
                    if (selectedImage != null)
                    {
                        command.CommandText = "INSERT INTO data_photo (Data_ID,Photo) VALUES (?Data_ID,?Photo)";
                        command.Parameters.AddWithValue("?Data_ID", data_ID);
                        command.Parameters.AddWithValue("?Photo", selectedImage);
                        command.ExecuteNonQuery();
                    }
                    if (products != null)
                    {
                        products.displayProducts();
                    }
                    UserControl.ItemRecord("data", "اضافة", data_ID, DateTime.Now, "", conn);
                    clear();
                    comType.Focus();
                }
                else
                {
                    MessageBox.Show("This item already exist");
                }
            }
            else
            {
                MessageBox.Show("you should fill all fields");
            }
        }

        //clear function
        public void clear()
        {
            foreach (Control co in this.panel1.Controls)
            {
                if (co is GroupBox)
                {
                    foreach (Control item in co.Controls)
                    {
                        if (item is ComboBox)
                        {
                            item.Text = "";
                        }
                        else if (item is TextBox)
                        {
                            item.Text = "";
                        }
                    }
                }
                ImageProduct.Image = null;
            }
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
            bool flag = false;
            foreach (Control co in this.panel1.Controls)
            {
                if (co is GroupBox)
                {
                    foreach (Control item in co.Controls)
                    {
                        if (item is ComboBox)
                        {
                            if (item.Text == "")
                                flag = true;
                            else
                               return false;
                        }
                        else if (item is TextBox)
                        {
                            if (item.Text == "")
                                flag = true;
                            else
                                return false;
                        }
                    }
                }
                else if (co is PictureBox)
                {
                    if (ImageProduct.Image == null)
                        flag = true;
                    else
                        return false;
                }
            }

            return flag;
        }
    }
    
}
