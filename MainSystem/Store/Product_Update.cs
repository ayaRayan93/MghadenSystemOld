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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MainSystem
{
    public partial class Product_Update : Form
    {
        byte[] selectedImage;
        bool loaded = false,flag=false, factoryFlage=false, groupFlage=false;
        bool flagProduct = false;
        MySqlConnection dbconnection, dbconnection1;
        string code = "",lastPartCode="";
        int Data_ID;
        int color_id = -1;
        int size_id = -1;
        int sort_id = -1;
        DataRowView prodRow;
        Products products = null;
        XtraTabControl xtraTabControlStoresContent = null;

        public Product_Update(DataRowView prodRow,Products products,XtraTabControl xtraTabControlStoresContent)
        {
            try
            {
                InitializeComponent();
                this.products = products;
                this.xtraTabControlStoresContent = xtraTabControlStoresContent;
                this.prodRow = prodRow;
                dbconnection = new MySqlConnection(connection.connectionString);
                dbconnection1 = new MySqlConnection(connection.connectionString);
       
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void Product_Update_Load(object sender, EventArgs e)
        {
            try
            {
                dbconnection.Open();
                loadData();
                code = prodRow["الكود"].ToString();
                displayData(prodRow);             
                loaded = true;
                flag = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void comBox_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (loaded)
                {
                    dbconnection.Close();
                    dbconnection.Open();
                    ComboBox comBox = (ComboBox)sender;

                    switch (comBox.Name)
                    {
                        case "comType":
                            if (loaded)
                            {
                                txtType.Text = comType.SelectedValue.ToString();
                                fillTypeCombox();
                                comProduct.Text = "";
                                txtProduct.Text = "";
                            }
                            break;
                        case "comFactory":
                            if (factoryFlage)
                            {
                                txtFactory.Text = comFactory.SelectedValue.ToString();
                                fillFactoryCombox();
                                comProduct.Text = "";
                                txtProduct.Text = "";
                            }
                            break;
                        case "comGroup":
                            if (groupFlage)
                            {
                                txtGroup.Text = comGroup.SelectedValue.ToString();
                                fillGroupCombox();
                            }
                            break;

                        case "comProduct":
                            if (flagProduct)
                            {
                               // flagProduct = false;
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
            dbconnection.Close();
            dbconnection.Open();
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
                        dbconnection.Close();
                        dbconnection.Open();
                        switch (txtBox.Name)
                        {
                            case "txtType":
                                query = "select Type_Name from type where Type_ID='" + txtType.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
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
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtFactory":
                                if (txtType.Text != "")
                                {
                                    query = "select Factory_Name from factory inner join type_factory on type_factory.Factory_ID=factory.Factory_ID where factory.Factory_ID='" + txtFactory.Text + "' and type_factory.Type_ID=" + txtType.Text;
                                    com = new MySqlCommand(query, dbconnection);
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
                                        dbconnection.Close();
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
                                    //if (txtType.Text == "2")
                                    //    query = "select Group_Name from groupo inner join product_factory_group on product_factory_group.Group_ID=groupo.Group_ID where groupo.Group_ID='" + txtGroup.Text + "' and product_factory_group.Factory_ID=" + txtFactory.Text + " and Type_ID=1";
                                    //else
                                        query = "select Group_Name from groupo inner join product_factory_group on product_factory_group.Group_ID=groupo.Group_ID where groupo.Group_ID='" + txtGroup.Text + "' and product_factory_group.Factory_ID=" + txtFactory.Text + " and Type_ID=" + txtType.Text;
                                    com = new MySqlCommand(query, dbconnection);
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
                                        dbconnection.Close();
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
                                    com = new MySqlCommand(query, dbconnection);
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
                                        dbconnection.Close();
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
                                com = new MySqlCommand(query, dbconnection);
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
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                            case "txtSize":
                                query = "select Size_Value from size where Size_ID='" + txtSize.Text + "' and Factory_ID='" + txtFactory.Text + "'";
                                com = new MySqlCommand(query, dbconnection);
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
                                    dbconnection.Close();
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
                                com = new MySqlCommand(query, dbconnection);
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
                                    dbconnection.Close();
                                    return;
                                }
                                break;
                       
                        }

                    }

                }
                catch
                {
                    // MessageBox.Show(ex.ToString());
                }
                dbconnection.Close();
            }

        }
        private void btnUpdateProduct_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtFactory.Text != "" && txtGroup.Text != "" && txtProduct.Text != "" && txtType.Text != "")
                {
                    double carton = 0;
                    string classification, description = "";
                    if (comColour.Text != "")
                    {
                        try
                        {
                            color_id = Convert.ToInt32(comColour.SelectedValue.ToString());
                        }
                        catch
                        {

                        }
                    }
                    if (comSize.Text != "")
                    {
                        try
                        {
                            size_id = Convert.ToInt32(comSize.SelectedValue.ToString());
                        }
                        catch
                        {

                        }
                    }
                    if (comSort.Text != "")
                    {
                        try
                        {
                            sort_id = Convert.ToInt32(comSort.SelectedValue.ToString());
                        }
                        catch
                        {

                        }

                    }
                    if (txtCarton.Text != "")
                    {
                        if (double.TryParse(txtCarton.Text, out carton))
                        {
                        }
                        else
                        {
                            MessageBox.Show("carton must be numeric");
                            dbconnection.Close();
                            return;
                        }
                    }
                    classification = txtClassification.Text;
                    description = txtDescription.Text;


                    dbconnection.Close();

                    MySqlCommand command = dbconnection.CreateCommand();

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
                    string code5 = "";

                    if (!IsMainChang())
                    {
                        string query2 = "SELECT max(Code) FROM data where Type_ID=" + txtType.Text + " and Factory_ID=" + txtFactory.Text + " and Group_ID=" + txtGroup.Text + " and Product_ID=" + txtProduct.Text;
                        dbconnection.Open();
                        MySqlCommand adpt = new MySqlCommand(query2, dbconnection);
                        string maxCode = adpt.ExecuteScalar().ToString();
                        if (maxCode != "")
                        {
                            char[] arrCode = maxCode.ToCharArray();
                            string part5 = arrCode[16].ToString() + arrCode[17].ToString() + arrCode[18].ToString() + arrCode[19].ToString() + "";

                            int result = Convert.ToInt32(part5);
                            result = result + 1;
                            dbconnection.Close();

                            int resultcount = result.ToString().Length;

                            code5 = result.ToString();

                            while (resultcount < 4)
                            {
                                code5 = "0" + code5;
                                resultcount++;
                            }
                        }
                        else
                        {
                            code5 = "0001";
                        }
                    }
                    else
                    {
                        code5 = lastPartCode;
                    }
                    code = code + code5;

                    string query = "update  data set Color_ID=?Color_ID,Size_ID=?Size_ID,Sort_ID=?Sort_ID,Description=?Description,Carton=?Carton,Code=?Code,Type_ID=?Type_ID,Factory_ID=?Factory_ID,Group_ID=?Group_ID,Product_ID=?Product_ID,Classification=?Classification where Data_ID=" + Data_ID;
                    command = new MySqlCommand(query, dbconnection);
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
                    dbconnection.Close();
                    dbconnection.Open();
                    command.ExecuteNonQuery();


                    //save image as bytes
                    query = "select DataPhoto_ID from data_photo where  Data_ID=" + Data_ID;
                    command = new MySqlCommand(query, dbconnection);
                    try
                    {
                        string DataPhoto_ID = command.ExecuteScalar().ToString();
                        string q = "update data_photo set Photo =?Photo where DataPhoto_ID=" + DataPhoto_ID;
                        command = new MySqlCommand(q, dbconnection);
                        command.Parameters.AddWithValue("?Photo", selectedImage);
                        command.ExecuteNonQuery();
                    }
                    catch
                    {
                        string q = "insert into  data_photo (Photo,Data_ID) values (@Photo,@Data_ID)";
                        command = new MySqlCommand(q, dbconnection);
                        command.Parameters.AddWithValue("?Photo", selectedImage);
                        command.Parameters.AddWithValue("?Data_ID", Data_ID);
                        dbconnection.Close();
                        dbconnection.Open();
                        command.ExecuteNonQuery();
                    }
                    UserControl.ItemRecord("data", "تعديل", Data_ID, DateTime.Now, "", dbconnection);
                    dbconnection.Close();
                    MessageBox.Show("updated");
                    products.displayProducts();

                    XtraTabPage xtraTabPage = getTabPage("تعديل بند");
                    xtraTabPage.ImageOptions.Image = null;
                    comType.Focus();

                }
                else
                {
                    MessageBox.Show("you should fill all fields");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbconnection.Close();
        }
        private void txtClassification_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtDescription.Focus();
            }
        }
        private void txtDescription_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                comSort.Focus();
            }
        }    
        private void btnUploadPhoto_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFile = openFileDialog1.FileName;
              
                ImageProduct.Image = Image.FromFile(openFileDialog1.FileName);                
                selectedImage= File.ReadAllBytes(selectedFile);
            }
        }
        private void txtBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (flag)
                {
                    XtraTabPage xtraTabPage = getTabPage("تعديل بند");
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

        //function
        public void loadData()
        {
            string query = "select * from type";
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            comType.DataSource = dt;
            comType.DisplayMember = dt.Columns["Type_Name"].ToString();
            comType.ValueMember = dt.Columns["Type_ID"].ToString();
            comType.Text = "";

            query = "select * from factory";
            da = new MySqlDataAdapter(query, dbconnection);
            dt = new DataTable();
            da.Fill(dt);
            comFactory.DataSource = dt;
            comFactory.DisplayMember = dt.Columns["Factory_Name"].ToString();
            comFactory.ValueMember = dt.Columns["Factory_ID"].ToString();
            comFactory.Text = "";

            query = "select * from groupo";
            da = new MySqlDataAdapter(query, dbconnection);
            dt = new DataTable();
            da.Fill(dt);
            comGroup.DataSource = dt;
            comGroup.DisplayMember = dt.Columns["Group_Name"].ToString();
            comGroup.ValueMember = dt.Columns["Group_ID"].ToString();
            comGroup.Text = "";

            query = "select * from product";
            da = new MySqlDataAdapter(query, dbconnection);
            dt = new DataTable();
            da.Fill(dt);
            comProduct.DataSource = dt;
            comProduct.DisplayMember = dt.Columns["Product_Name"].ToString();
            comProduct.ValueMember = dt.Columns["Product_ID"].ToString();
            comProduct.Text = "";

            query = "select * from sort";
            da = new MySqlDataAdapter(query, dbconnection);
            dt = new DataTable();
            da.Fill(dt);
            comSort.DataSource = dt;
            comSort.DisplayMember = dt.Columns["Sort_Value"].ToString();
            comSort.ValueMember = dt.Columns["Sort_ID"].ToString();
            comSort.Text = "";

        }
        public void displayData(DataRowView row1)
        {
            ImageProduct.Image = Properties.Resources.animated_gif_loading;
            Thread t1 = new Thread(displayImage);
            t1.Start();
            displayCode();
            comType.Text = row1["النوع"].ToString();                      
            comFactory.Text = row1["المصنع"].ToString();       
            comGroup.Text = row1["المجموعة"].ToString();
            comProduct.Text = row1["الصنف"].ToString();
         
            TypeColor();
            FactorySize();

            txtCarton.Text = row1["الكرتنة"].ToString();

            comColour.Text =row1["اللون"].ToString();
            if(comColour.Text!="")
                txtColor.Text = row1["Color_ID"].ToString();

            comSize.Text = row1["المقاس"].ToString();
            if (comSize.Text != "")
                txtSize.Text = row1["Size_ID"].ToString();

            comSort.Text = row1["الفرز"].ToString();
            if (comSort.Text != "")
                txtSort.Text = row1["Sort_ID"].ToString();



            txtClassification.Text = row1["التصنيف"].ToString();
            txtDescription.Text = row1["الوصف"].ToString();

            Data_ID =Convert.ToInt32(row1[0]);
        }
        public void displayImage()
        {
            try
            {
                dbconnection1.Open();
                string query = "select Photo from data_photo where Data_ID=" + Convert.ToInt32(prodRow[0]);
                MySqlCommand com = new MySqlCommand(query, dbconnection1);
                byte[] photo = (byte[])com.ExecuteScalar();
               
                if (photo != null)
                {
                    MemoryStream ms = new MemoryStream(photo);
                    ImageProduct.Image = Image.FromStream(ms);
                    selectedImage = photo;
                }
                else
                {
                    ImageProduct.Image = null;// Properties.Resources.notFound;
                }
            }
            catch
            {
               //MessageBox.Show(ex.Message);
               ImageProduct.Image = Properties.Resources.notFound;
            }
            dbconnection1.Close();
        }
        public void displayCode()
        {
            char[] arrCode = code.ToCharArray();
            txtType.Text =Convert.ToInt32(arrCode[0].ToString() + arrCode[1].ToString() + arrCode[2].ToString() + arrCode[3].ToString() + "").ToString();
            fillTypeCombox();
            txtFactory.Text = Convert.ToInt32(arrCode[4].ToString() + arrCode[5].ToString() + arrCode[6].ToString() + arrCode[7].ToString() + "").ToString();
            fillFactoryCombox();
            txtGroup.Text = Convert.ToInt32(arrCode[8].ToString() + arrCode[9].ToString() + arrCode[10].ToString() + arrCode[11].ToString() + "").ToString();
            fillGroupCombox();
            txtProduct.Text = Convert.ToInt32(arrCode[12].ToString() + arrCode[13].ToString() + arrCode[14].ToString() + arrCode[15].ToString() + "").ToString();
            lastPartCode = "" + arrCode[16] + arrCode[17] + arrCode[18] + arrCode[19];
        }
        private void TypeColor()
        {
            string query = "select * from color where Type_ID=" + Convert.ToInt32(comType.SelectedValue);
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            comColour.DataSource = dt;
            comColour.DisplayMember = dt.Columns["Color_Name"].ToString();
            comColour.ValueMember = dt.Columns["Color_ID"].ToString();
            comColour.Text = "";
           
           // comFactory.Focus();
        }
        private void FactorySize()
        {
            string query2 = "select * from size where Factory_ID=" + Convert.ToInt32(comFactory.SelectedValue)+" and Group_ID="+txtGroup.Text;
            MySqlDataAdapter da2 = new MySqlDataAdapter(query2, dbconnection);
            DataTable dt2 = new DataTable();
            da2.Fill(dt2);
            comSize.DataSource = dt2;
            comSize.DisplayMember = dt2.Columns["Size_Value"].ToString();
            comSize.ValueMember = dt2.Columns["Size_ID"].ToString();
            comSize.Text = "";
            //comGroup.Focus();
        }      
        public bool IsClear()
        {
            if (comType.Text == prodRow["النوع"].ToString() &&
            comFactory.Text == prodRow["المصنع"].ToString() &&
            comGroup.Text == prodRow["المجموعة"].ToString() &&
            comProduct.Text == prodRow["الصنف"].ToString() &&
            txtCarton.Text == prodRow["الكرتنة"].ToString() &&
            comColour.Text == prodRow["اللون"].ToString() &&
            comSize.Text == prodRow["المقاس"].ToString() &&
            comSort.Text == prodRow["الفرز"].ToString() &&
            txtDescription.Text == prodRow["التصنيف"].ToString() &&
            txtClassification.Text == prodRow["الوصف"].ToString())
                return true;

            return false;
        }
        public bool IsMainChang()
        {
            if (comType.Text == prodRow[5].ToString() &&
            comFactory.Text == prodRow[6].ToString() &&
            comGroup.Text == prodRow[7].ToString() &&
            comProduct.Text == prodRow[8].ToString() )
                return true;

            return false;
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

        public void fillTypeCombox()
        {
            string query = "select * from factory inner join type_factory on factory.Factory_ID=type_factory.Factory_ID inner join type on type_factory.Type_ID=type.Type_ID where type_factory.Type_ID=" + txtType.Text;
            MySqlDataAdapter da = new MySqlDataAdapter(query, dbconnection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            comFactory.DataSource = dt;
            comFactory.DisplayMember = dt.Columns["Factory_Name"].ToString();
            comFactory.ValueMember = dt.Columns["Factory_ID"].ToString();
            comFactory.Text = "";
            txtFactory.Text = "";
            query = "select TypeCoding_Method from type where Type_ID=" + txtType.Text;
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            int TypeCoding_Method = (int)com.ExecuteScalar();
            if (TypeCoding_Method == 1)
            {
                string query2 = "";
                //if (txtType.Text == "2" || txtType.Text == "1")
                //{
                //    query2 = "select * from groupo where Factory_ID=-1";
                //}
                //else
                //{
                    query2 = "select * from groupo where Factory_ID=" + -Convert.ToInt32(txtType.Text) + " and Type_ID=" + txtType.Text;
                //}

                MySqlDataAdapter da2 = new MySqlDataAdapter(query2, dbconnection);
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
            da = new MySqlDataAdapter(query, dbconnection);
            dt = new DataTable();
            da.Fill(dt);
            comColour.DataSource = dt;
            comColour.DisplayMember = dt.Columns["Color_Name"].ToString();
            comColour.ValueMember = dt.Columns["Color_ID"].ToString();
            comColour.Text = "";
            txtColor.Text = "";
            comFactory.Focus();
        }
        public void fillFactoryCombox()
        {
            string query = "select TypeCoding_Method from type where Type_ID=" + txtType.Text;
            MySqlCommand com = new MySqlCommand(query, dbconnection);
            int TypeCoding_Method = (int)com.ExecuteScalar();
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
        }
        public void fillGroupCombox()
        {
            string query3 = "select distinct  product.Product_ID  ,Product_Name  from product inner join product_factory_group on product.Product_ID=product_factory_group.Product_ID  where product.Type_ID=" + txtType.Text + " and product_factory_group.Factory_ID=" + txtFactory.Text + " and product_factory_group.Group_ID=" + txtGroup.Text + "  order by product.Product_ID";
            MySqlDataAdapter da3 = new MySqlDataAdapter(query3, dbconnection);
            DataTable dt3 = new DataTable();
            da3.Fill(dt3);
            comProduct.DataSource = dt3;
            comProduct.DisplayMember = dt3.Columns["Product_Name"].ToString();
            comProduct.ValueMember = dt3.Columns["Product_ID"].ToString();
            comProduct.Text = "";
            txtProduct.Text = "";

            string query2 = "select * from size where Factory_ID=" + txtFactory.Text + " and Group_ID=" + txtGroup.Text;
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
        public void fillProductCombox()
        {

        }
    }
    
}
