using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace MailConverter
{
    public class ContactSelectionForm : Form
    {
        private DataGridView _dataGrid;
        private Button _btnSelectAll;
        private Button _btnDeselectAll;
        private Button _btnOK;
        private Button _btnCancel;
        private Label _lblInfo;
        private TextBox _txtSearch;
        private List<ContactItem> _allContacts;

        public List<string> SelectedUrls { get; private set; } = new List<string>();
        public List<string> SelectedNames { get; private set; } = new List<string>();

        public ContactSelectionForm(List<CardDavContact> contacts)
        {
            _allContacts = new List<ContactItem>();
            foreach (var contact in contacts)
            {
                // 清理数据：去掉"(无邮箱)"等占位符
                var name = contact.Name ?? "";
                var email = contact.Email ?? "";
                var phone = contact.Phone ?? "";

                if (email.Contains("无邮箱")) email = "";
                if (string.IsNullOrEmpty(name)) name = "(无姓名)";

                _allContacts.Add(new ContactItem
                {
                    Url = contact.Url,
                    Name = name,
                    Phone = phone,
                    Email = email,
                    IsSelected = true
                });
            }
            InitializeComponent();
            RefreshGrid();

            _lblInfo.Text = $"共 {contacts.Count} 个联系人，勾选要同步的联系人(双击行可切换勾选)";
        }

        private void RefreshGrid()
        {
            var searchText = _txtSearch.Text.Trim();
            if (searchText == "搜索联系人...")
                searchText = "";
            searchText = searchText.ToLower();

            _dataGrid.Rows.Clear();

            foreach (var contact in _allContacts)
            {
                // 搜索姓名、电话或邮箱
                bool match = string.IsNullOrEmpty(searchText) ||
                    contact.Name.ToLower().Contains(searchText) ||
                    contact.Phone.ToLower().Contains(searchText) ||
                    contact.Email.ToLower().Contains(searchText);

                if (match)
                {
                    int rowIndex = _dataGrid.Rows.Add(contact.Name, contact.Phone, contact.Email, contact.IsSelected);
                    _dataGrid.Rows[rowIndex].Tag = contact;
                }
            }
        }

        private void InitializeComponent()
        {
            this.Text = "选择联系人";
            this.Size = new Size(750, 650);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            _lblInfo = new Label
            {
                Location = new Point(10, 10),
                Size = new Size(710, 25),
                Text = "选择要同步的联系人:",
                Font = new Font("Microsoft Sans Serif", 10F)
            };

            _txtSearch = new TextBox
            {
                Location = new Point(10, 40),
                Size = new Size(200, 25),
                Text = "搜索联系人...",
                ForeColor = Color.Gray
            };
            _txtSearch.Enter += (s, e) =>
            {
                if (_txtSearch.Text == "搜索联系人...")
                {
                    _txtSearch.Text = "";
                    _txtSearch.ForeColor = Color.Black;
                }
            };
            _txtSearch.Leave += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(_txtSearch.Text))
                {
                    _txtSearch.Text = "搜索联系人...";
                    _txtSearch.ForeColor = Color.Gray;
                }
            };
            _txtSearch.TextChanged += (s, e) => RefreshGrid();

            _dataGrid = new DataGridView
            {
                Location = new Point(10, 70),
                Size = new Size(720, 480),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = false,
                MultiSelect = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            // 添加列
            var colName = new DataGridViewTextBoxColumn { Name = "Name", HeaderText = "姓名", Width = 200, ReadOnly = true };
            var colPhone = new DataGridViewTextBoxColumn { Name = "Phone", HeaderText = "电话", Width = 180, ReadOnly = true };
            var colEmail = new DataGridViewTextBoxColumn { Name = "Email", HeaderText = "邮箱", Width = 280, ReadOnly = true };
            var colSelect = new DataGridViewCheckBoxColumn { Name = "Select", HeaderText = "选择", Width = 60 };

            _dataGrid.Columns.AddRange(new DataGridViewColumn[] { colName, colPhone, colEmail, colSelect });

            _dataGrid.CellClick += (s, e) =>
            {
                if (e.ColumnIndex == 3 && e.RowIndex >= 0)
                {
                    var row = _dataGrid.Rows[e.RowIndex];
                    if (row.Tag is ContactItem contact)
                    {
                        var isChecked = (bool)_dataGrid.Rows[e.RowIndex].Cells[0].Value;
                        contact.IsSelected = !isChecked;
                        _dataGrid.Rows[e.RowIndex].Cells[0].Value = !isChecked;
                    }
                }
            };

            _dataGrid.CellDoubleClick += (s, e) =>
            {
                if (e.RowIndex >= 0)
                {
                    var row = _dataGrid.Rows[e.RowIndex];
                    if (row.Tag is ContactItem contact)
                    {
                        contact.IsSelected = !contact.IsSelected;
                        row.Cells[0].Value = contact.IsSelected;
                    }
                }
            };

            _btnSelectAll = new Button
            {
                Text = "全选",
                Location = new Point(220, 40),
                Size = new Size(80, 25)
            };
            _btnSelectAll.Click += (s, e) =>
            {
                foreach (var contact in _allContacts)
                    contact.IsSelected = true;
                RefreshGrid();
            };

            _btnDeselectAll = new Button
            {
                Text = "取消全选",
                Location = new Point(310, 40),
                Size = new Size(80, 25)
            };
            _btnDeselectAll.Click += (s, e) =>
            {
                foreach (var contact in _allContacts)
                    contact.IsSelected = false;
                RefreshGrid();
            };

            _btnOK = new Button
            {
                Text = "确定同步",
                Location = new Point(530, 560),
                Size = new Size(90, 30),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                DialogResult = DialogResult.OK
            };
            _btnOK.Click += BtnOK_Click;

            _btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(630, 560),
                Size = new Size(90, 30),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.AddRange(new Control[] { _lblInfo, _txtSearch, _btnSelectAll, _btnDeselectAll, _dataGrid, _btnOK, _btnCancel });
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            SelectedUrls.Clear();
            SelectedNames.Clear();
            foreach (var contact in _allContacts)
            {
                if (contact.IsSelected)
                {
                    SelectedUrls.Add(contact.Url);
                    SelectedNames.Add(contact.Name);
                }
            }
        }

        private class ContactItem
        {
            public string Url { get; set; }
            public string Name { get; set; }
            public string Phone { get; set; }
            public string Email { get; set; }
            public bool IsSelected { get; set; }
        }
    }
}
