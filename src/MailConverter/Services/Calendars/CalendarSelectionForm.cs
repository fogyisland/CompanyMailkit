using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MailConverter.Services.Calendars
{
    public class CalendarSelectionForm : Form
    {
        private List<CalendarData> _allCalendars;
        private List<CalendarData> _selectedCalendars;
        private CheckBox chkSelectAll;
        private DataGridView dgvCalendars;
        private Button btnOK;
        private Button btnCancel;

        public List<CalendarData> SelectedCalendarData
        {
            get { return _selectedCalendars; }
        }

        public CalendarSelectionForm(List<CalendarData> calendars)
        {
            _allCalendars = calendars ?? new List<CalendarData>();
            _selectedCalendars = new List<CalendarData>();
            InitializeComponent();
            LoadCalendars();
        }

        private void InitializeComponent()
        {
            this.Text = "选择要导入的日历";
            this.Size = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(600, 400);

            chkSelectAll = new CheckBox
            {
                Text = "全选",
                Location = new Point(20, 15),
                AutoSize = true,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold)
            };
            chkSelectAll.CheckedChanged += ChkSelectAll_CheckedChanged;

            dgvCalendars = new DataGridView
            {
                Location = new Point(20, 50),
                Size = new Size(640, 350),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false
            };

            var colSelect = new DataGridViewCheckBoxColumn
            {
                Name = "Select",
                HeaderText = "选择",
                Width = 60
            };
            var colSubject = new DataGridViewTextBoxColumn { Name = "Subject", HeaderText = "主题", ReadOnly = true };
            var colStart = new DataGridViewTextBoxColumn { Name = "Start", HeaderText = "开始时间", ReadOnly = true };
            var colEnd = new DataGridViewTextBoxColumn { Name = "End", HeaderText = "结束时间", ReadOnly = true };
            var colLocation = new DataGridViewTextBoxColumn { Name = "Location", HeaderText = "地点", ReadOnly = true };
            var colAllDay = new DataGridViewTextBoxColumn { Name = "AllDay", HeaderText = "全天", ReadOnly = true };

            dgvCalendars.Columns.AddRange(new DataGridViewColumn[] {
                colSelect, colSubject, colStart, colEnd, colLocation, colAllDay
            });
            dgvCalendars.CellValueChanged += DgvCalendars_CellValueChanged;
            dgvCalendars.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgvCalendars.IsCurrentCellDirty)
                    dgvCalendars.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            btnOK = new Button
            {
                Text = "确定",
                Location = new Point(490, 415),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(580, 415),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.Add(chkSelectAll);
            this.Controls.Add(dgvCalendars);
            this.Controls.Add(btnOK);
            this.Controls.Add(btnCancel);
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadCalendars()
        {
            dgvCalendars.Rows.Clear();
            foreach (var cal in _allCalendars)
            {
                dgvCalendars.Rows.Add(
                    false,
                    cal.Subject ?? "(无主题)",
                    cal.StartTime?.ToString("yyyy-MM-dd HH:mm") ?? "",
                    cal.EndTime?.ToString("yyyy-MM-dd HH:mm") ?? "",
                    cal.Location ?? "",
                    cal.IsAllDayEvent ? "是" : "否"
                );
            }
        }

        private void DgvCalendars_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex != 0) return;
            // 同步 Select 列状态到业务数据 (UI-only, 提交时再筛选)
        }

        private void ChkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            bool sel = chkSelectAll.Checked;
            foreach (DataGridViewRow row in dgvCalendars.Rows)
            {
                row.Cells["Select"].Value = sel;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            _selectedCalendars.Clear();
            for (int i = 0; i < dgvCalendars.Rows.Count; i++)
            {
                bool isSelected = Convert.ToBoolean(dgvCalendars.Rows[i].Cells["Select"].Value ?? false);
                if (isSelected && i < _allCalendars.Count)
                {
                    _selectedCalendars.Add(_allCalendars[i]);
                }
            }
        }
    }
}
