using System;
using System.Collections;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Windows.Forms;
using System.Security.Principal;

namespace ChangeUserInfoInAD
{
    public partial class MainWindow : Form
    {
        public MainWindow()
        {
            InitializeComponent();
            FieldDisable();
            comboBox1.Text = "Пользователь";
            Text = "Поиск пользователя v." + (System.Reflection.Assembly.GetExecutingAssembly()).GetName().Version + " - (" + (WindowsIdentity.GetCurrent().Name) + ")";
            TopLevel = true;
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            ArrayList User_atrib = new ArrayList();
            String typeSearch = null; // user or contact
            if (comboBox1.Text == "Пользователь") { typeSearch = "user"; } else { typeSearch = "contact"; }
            if (comboBox1.Text == String.Empty || textBox1.Text == String.Empty) { MessageBox.Show("Не указан критерий поиска"); }
            else
            {
                try
                {
                    DirectoryContext DomainContext = new DirectoryContext(DirectoryContextType.Domain);
                    Domain DomainName = Domain.GetDomain(DomainContext);
                    DirectoryEntry ADEntry = new DirectoryEntry("LDAP://" + DomainName, null, null, AuthenticationTypes.Secure);
                    DirectorySearcher ADSearchUser = new DirectorySearcher(ADEntry);
                    ADSearchUser.Filter = "(&(objectClass=" + typeSearch + ")(objectCategory=Person)(|(displayName=" + textBox1.Text + "*)(mail=" + textBox1.Text + "*)))";
                    SearchResult result = ADSearchUser.FindOne();
                    if (result != null)
                    {
                        User_atrib.Clear();
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["sn"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["givenName"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["displayName"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["mail"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["l"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["streetAddress"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["company"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["department"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["title"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["telephoneNumber"].Value.ToString()); }
                        catch { User_atrib.Add(""); }
                            try { User_atrib.Add(result.GetDirectoryEntry().Properties["drink"].Value.ToString()); }
                            catch { User_atrib.Add(""); }
                        try { User_atrib.Add(result.GetDirectoryEntry().Properties["distinguishedname"].Value.ToString()); }
                        catch { User_atrib.Add(null); }
                        if (typeSearch == "user")
                        {
                            try { User_atrib.Add(result.GetDirectoryEntry().Properties["userAccountControl"].Value.ToString()); }
                            catch { User_atrib.Add(null); }
                        }
                        User_atrib.TrimToSize();
                        textBox2.Text = User_atrib[0].ToString();
                        textBox3.Text = User_atrib[1].ToString();
                        textBox4.Text = User_atrib[2].ToString();
                        textBox5.Text = User_atrib[3].ToString();
                        textBox6.Text = User_atrib[4].ToString();
                        textBox7.Text = User_atrib[5].ToString();
                        textBox8.Text = User_atrib[6].ToString();
                        textBox9.Text = User_atrib[7].ToString();
                        textBox10.Text = User_atrib[8].ToString();
                        textBox11.Text = User_atrib[9].ToString();
                        textBox12.Text = User_atrib[10].ToString();
                        textBox13.Text = User_atrib[11].ToString();
                        textBox14.Text = User_atrib[12].ToString();
                        FieldEnable();
                        if (typeSearch == "user")
                        {
                            if (User_atrib[13].ToString() == "66048" || User_atrib[13].ToString() == "512") { checkBox1.Checked = false; }
                            else { checkBox1.Checked = true; }
                            DirectoryEntry user = new DirectoryEntry("LDAP://" + textBox14.Text, null, null, AuthenticationTypes.Secure);
                            user.RefreshCache(new string[] { "msDS-User-Account-Control-Computed" });
                            int val = (int)user.Properties["msDS-User-Account-Control-Computed"].Value;
                            // ADS_UF_LOCKOUT
                            const int UF_LOCKOUT = 0x0010;
                            if (Convert.ToBoolean(val & UF_LOCKOUT)) { checkBox2.Checked = true;checkBox2.Enabled = true; }
                            else { checkBox2.Checked = false; checkBox2.Enabled = false; }
                        }
                        else { checkBox1.Enabled = false; checkBox2.Enabled = false; textBox13.Enabled = false; }
                    }
                    else { MessageBox.Show("По Вашему запросу: " + textBox1.Text + "\nничего не найдено\nИзмените критерий поиска"); }
                    ADEntry.Dispose();
                    ADEntry.Close();
                    ADSearchUser.Dispose();
                }
                catch (DirectoryServicesCOMException ex) { MessageBox.Show(ex.Message.Trim()); FieldDisable(); }
                catch (UnauthorizedAccessException) { MessageBox.Show("Отказано в доступе к объекту\r\nУ Вас нет прав на проведение данной операции."); FieldDisable(); }
                catch (Exception ex) { MessageBox.Show(ex.Message.Trim() + "\r\n" + ex.StackTrace); FieldDisable(); }
            }
        }
        private void btnChange_Click(object sender, EventArgs e)
        {
            if (textBox14.Text != String.Empty)
            {
                ArrayList Usr_att = new ArrayList();
                Usr_att.Clear();
                try { Usr_att.Add(textBox2.Text); }
                catch { Usr_att.Add(null); }
                try { Usr_att.Add(textBox3.Text); }
                catch { Usr_att.Add(null); }
                try { Usr_att.Add(textBox4.Text); }
                catch { Usr_att.Add(null); }
                try { Usr_att.Add(textBox6.Text); }
                catch { Usr_att.Add(null); }
                try { Usr_att.Add(textBox7.Text); }
                catch { Usr_att.Add(null); }
                try { Usr_att.Add(textBox8.Text); }
                catch { Usr_att.Add(null); }
                try { Usr_att.Add(textBox9.Text); }
                catch { Usr_att.Add(null); }
                try { Usr_att.Add(textBox10.Text); }
                catch { Usr_att.Add(null); }
                try { Usr_att.Add(textBox11.Text); }
                catch { Usr_att.Add(null); }
                try { Usr_att.Add(textBox12.Text); }
                catch { Usr_att.Add(null); }
                try { Usr_att.Add(textBox13.Text); }
                catch { Usr_att.Add(null); }
                Usr_att.TrimToSize();
                try
                {
                    DirectoryEntry UserEntry = new DirectoryEntry("LDAP://" + textBox14.Text, null, null, AuthenticationTypes.Secure);
                    if (Usr_att[0].ToString() != String.Empty) { UserEntry.Properties["sn"].Value = Usr_att[0].ToString(); }
                    else { UserEntry.Properties["sn"].Value = null; }
                    if (Usr_att[1].ToString() != String.Empty) { UserEntry.Properties["givenName"].Value = Usr_att[1].ToString(); }
                    else { UserEntry.Properties["givenName"].Value = null; }
                    if (Usr_att[2].ToString() != String.Empty) { UserEntry.Properties["displayName"].Value = Usr_att[2].ToString(); }
                    else { UserEntry.Properties["displayName"].Value = null; }
                    if (Usr_att[3].ToString() != String.Empty) { UserEntry.Properties["l"].Value = Usr_att[3].ToString(); }
                    else { UserEntry.Properties["l"].Value = null; }
                    if (Usr_att[4].ToString() != String.Empty) { UserEntry.Properties["streetAddress"].Value = Usr_att[4].ToString(); }
                    else { UserEntry.Properties["streetAddress"].Value = null; }
                    if (Usr_att[5].ToString() != String.Empty) { UserEntry.Properties["company"].Value = Usr_att[5].ToString(); }
                    else { UserEntry.Properties["company"].Value = null; }
                    if (Usr_att[6].ToString() != String.Empty) { UserEntry.Properties["department"].Value = Usr_att[6].ToString(); }
                    else { UserEntry.Properties["department"].Value = null; }
                    if (Usr_att[7].ToString() != String.Empty) { UserEntry.Properties["title"].Value = Usr_att[7].ToString(); }
                    else { UserEntry.Properties["title"].Value = null; }
                    if (Usr_att[8].ToString() != String.Empty) { UserEntry.Properties["physicalDeliveryOfficeName"].Value = Usr_att[8].ToString(); }
                    else { UserEntry.Properties["physicalDeliveryOfficeName"].Value = null; }
                    if (Usr_att[9].ToString() != String.Empty) { UserEntry.Properties["telephoneNumber"].Value = Usr_att[9].ToString(); }
                    else { UserEntry.Properties["telephoneNumber"].Value = null; }
                    if (Usr_att[10].ToString() != String.Empty) { UserEntry.Properties["drink"].Value = Usr_att[10].ToString(); }
                    else { UserEntry.Properties["drink"].Value = null; }
                    UserEntry.CommitChanges();
                    UserEntry.RefreshCache();
                    UserEntry.Dispose();
                    UserEntry.Close();
                    MessageBox.Show("Изменения успешно внесены");
                    FieldDisable();
                }
                catch (DirectoryServicesCOMException ex) { MessageBox.Show(ex.Message.Trim()); FieldDisable(); }
                catch (UnauthorizedAccessException) { MessageBox.Show("Отказано в доступе к объекту\r\nУ Вас нет прав на проведение данной операции."); FieldDisable(); }
                catch (Exception ex) { MessageBox.Show(ex.Message.Trim() + "\r\n" + ex.StackTrace); FieldDisable(); }
            }
        }
        private void chkBoxChange1(object sender, EventArgs e)
        {
            if (textBox14.Text != String.Empty)
            {
                if (checkBox1.Checked == true)
                {
                    // Disable a User Account
                    try
                    {
                        DirectoryEntry user = new DirectoryEntry("LDAP://" + textBox14.Text, null, null, AuthenticationTypes.Secure);
                        int val = (int)user.Properties["userAccountControl"].Value;
                        user.Properties["userAccountControl"].Value = val | 0x2;
                        //ADS_UF_ACCOUNTDISABLE;
                        user.CommitChanges();
                        user.Dispose();
                        user.Close();
                        MessageBox.Show("Учетная запись ВЫКЛЮЧЕНА");
                    }
                    catch (DirectoryServicesCOMException ex) { MessageBox.Show(ex.Message.Trim()); FieldDisable(); }
                    catch (UnauthorizedAccessException) { MessageBox.Show("Отказано в доступе к объекту\r\nУ Вас нет прав на проведение данной операции."); FieldDisable(); }
                    catch (Exception ex) { MessageBox.Show(ex.Message.Trim() + "\r\n" + ex.StackTrace); FieldDisable(); }
                }
                else
                {
                    // Enable a User Account
                    try
                    {
                        DirectoryEntry user = new DirectoryEntry("LDAP://" + textBox14.Text, null, null, AuthenticationTypes.Secure);
                        int val = (int)user.Properties["userAccountControl"].Value;
                        user.Properties["userAccountControl"].Value = val & ~0x2;
                        //ADS_UF_NORMAL_ACCOUNT;
                        user.CommitChanges();
                        user.Dispose();
                        user.Close();
                        MessageBox.Show("Учетная запись ВКЛЮЧЕНА");
                    }
                    catch (DirectoryServicesCOMException ex) { MessageBox.Show(ex.Message.Trim()); FieldDisable(); }
                    catch (UnauthorizedAccessException) { MessageBox.Show("Отказано в доступе к объекту\r\nУ Вас нет прав на проведение данной операции."); FieldDisable(); }
                    catch (Exception ex) { MessageBox.Show(ex.Message.Trim() + "\r\n" + ex.StackTrace); FieldDisable(); }
                }
            }
        }
        private void chkBoxChange2(object sender, EventArgs e)
        {
            if (textBox14.Text != String.Empty)
            {
                if (checkBox2.Checked == false)
                {
                    try
                    {
                        DirectoryEntry user = new DirectoryEntry("LDAP://" + textBox14.Text, null, null, AuthenticationTypes.Secure);
                        user.Properties["lockoutTime"].Value = 0;
                        user.CommitChanges();
                        user.Dispose();
                        user.Close();
                        checkBox2.Enabled = false;
                        MessageBox.Show("Учетная запись РАЗБЛОКИРОВАНА");
                    }
                    catch (DirectoryServicesCOMException ex) { MessageBox.Show(ex.Message.Trim()); FieldDisable(); }
                    catch (UnauthorizedAccessException) { MessageBox.Show("Отказано в доступе к объекту\r\nУ Вас нет прав на проведение данной операции."); FieldDisable(); }
                    catch (Exception ex) { MessageBox.Show(ex.Message.Trim() + "\r\n" + ex.StackTrace); FieldDisable(); }
                }
            }

        }
        private void btnClear_Click(object sender, EventArgs e)
        {
            FieldDisable();
            textBox1.Focus();
        }
        private void FieldDisable()
        {
            textBox1.Text = null;
            textBox1.Focus();
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;
            textBox14.Text = null;
            checkBox1.Checked = false;
            checkBox2.Checked = false;

            label1.Enabled = true;
            comboBox1.Enabled = true;
            textBox1.Enabled = true;
            btnSearch.Enabled = true;

            label2.Enabled = false;
            label3.Enabled = false;
            label4.Enabled = false;
            label5.Enabled = false;
            label6.Enabled = false;
            label7.Enabled = false;
            label8.Enabled = false;
            label9.Enabled = false;
            label10.Enabled = false;
            label11.Enabled = false;
            label12.Enabled = false;
            label13.Enabled = false;
            label14.Enabled = false;

            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            textBox7.Enabled = false;
            textBox8.Enabled = false;
            textBox9.Enabled = false;
            textBox10.Enabled = false;
            textBox11.Enabled = false;
            textBox12.Enabled = false;
            textBox13.Enabled = false;
            textBox14.Enabled = false;
            checkBox1.Checked = false;
            checkBox1.Enabled = false;
            checkBox2.Enabled = false;
            btnClear.Enabled = false;
            btnChange.Enabled = false;
        }
        private void FieldEnable()
        {
            label1.Enabled = false;
            comboBox1.Enabled = false;
            textBox1.Enabled = false;
            btnSearch.Enabled = false;
            label2.Enabled = true;
            label3.Enabled = true;
            label4.Enabled = true;
            label5.Enabled = true;
            label6.Enabled = true;
            label7.Enabled = true;
            label8.Enabled = true;
            label9.Enabled = true;
            label10.Enabled = true;
            label11.Enabled = true;
            label12.Enabled = true;
            label13.Enabled = true;
            label14.Enabled = false;
            textBox2.Enabled = true;
            textBox2.Focus();
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox5.Enabled = false;
            textBox6.Enabled = true;
            textBox7.Enabled = true;
            textBox8.Enabled = true;
            textBox9.Enabled = true;
            textBox10.Enabled = true;
            textBox11.Enabled = true;
            textBox12.Enabled = true;
            textBox13.Enabled = true;
            textBox14.Enabled = false;
            checkBox1.Enabled = true;
            btnClear.Enabled = true;
            btnChange.Enabled = true;
        }
        private void cmbBox_changed(object sender, EventArgs e)
        {
            textBox1.Text = "";
            if (comboBox1.Text != String.Empty)
            {
                textBox1.AutoCompleteSource = AutoCompleteSource.CustomSource; ;
                textBox1.AutoCompleteMode = AutoCompleteMode.Suggest;
                AutoCompleteStringCollection preLoadName = new AutoCompleteStringCollection();
                String preSearch = null; // user or contact
                if (comboBox1.Text == "Пользователь") { preSearch = "user"; } else { preSearch = "contact"; }
                try
                {
                    DirectoryContext preDomainContext = new DirectoryContext(DirectoryContextType.Domain);
                    Domain preDomainName = Domain.GetDomain(preDomainContext);
                    DirectoryEntry preADEntry = new DirectoryEntry("LDAP://" + preDomainName, null, null, AuthenticationTypes.Secure);
                    DirectorySearcher preADLoadUser = new DirectorySearcher(preADEntry);
                    preADLoadUser.Filter = "(&(objectClass=" + preSearch + ")(objectCategory=Person))";
                    SearchResultCollection ADSearchResultsUser = preADLoadUser.FindAll();
                    if (ADSearchResultsUser != null)
                    {
                        foreach (SearchResult result in ADSearchResultsUser)
                        {
                            try { preLoadName.Add(result.GetDirectoryEntry().Properties["displayName"].Value.ToString()); }
                            catch { preLoadName.Add("нет совпадений"); }
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message.Trim() + "\r\n" + ex.StackTrace); }
                textBox1.AutoCompleteCustomSource = preLoadName;
            }
        }
    }
}
