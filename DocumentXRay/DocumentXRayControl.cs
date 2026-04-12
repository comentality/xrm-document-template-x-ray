using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentXRay.Logic;
using XrmToolBox.Extensibility;

namespace DocumentXRay
{
    public partial class DocumentXRayControl : PluginControlBase
    {
        private List<FieldInfo> _currentFields;

        // Controls
        private Panel _dropZonePanel;
        private Label _dropLabel;
        private Button _btnBrowse;
        private Panel _toolbarPanel;
        private Label _lblFilePath;
        private RadioButton _rbFlat;
        private RadioButton _rbTree;
        private Label _lblFieldCount;
        private Panel _resultsPanel;
        private ListView _lvFields;
        private TreeView _tvFields;

        public DocumentXRayControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            // -- Drop zone --
            _dropZonePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 120,
                AllowDrop = true,
                BackColor = Color.FromArgb(245, 245, 250)
            };
            _dropZonePanel.Paint += DropZonePanel_Paint;
            _dropZonePanel.DragEnter += DropZonePanel_DragEnter;
            _dropZonePanel.DragLeave += DropZonePanel_DragLeave;
            _dropZonePanel.DragDrop += DropZonePanel_DragDrop;

            _dropLabel = new Label
            {
                Text = "Drag && drop a .docx template here",
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 11f),
                ForeColor = Color.FromArgb(100, 100, 100)
            };

            _btnBrowse = new Button
            {
                Text = "Browse...",
                Width = 100,
                Height = 30,
                Anchor = AnchorStyles.Bottom
            };
            _btnBrowse.Click += BtnBrowse_Click;

            // Position browse button at bottom-center of drop zone
            var browsePanel = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            _btnBrowse.Dock = DockStyle.None;
            browsePanel.Controls.Add(_btnBrowse);
            browsePanel.Resize += (s, e) =>
            {
                _btnBrowse.Left = (browsePanel.Width - _btnBrowse.Width) / 2;
                _btnBrowse.Top = 5;
            };

            _dropZonePanel.Controls.Add(_dropLabel);
            _dropZonePanel.Controls.Add(browsePanel);

            // -- Toolbar --
            _toolbarPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 35,
                Padding = new Padding(5, 5, 5, 0)
            };

            _lblFilePath = new Label
            {
                Text = "No file selected",
                AutoSize = true,
                Location = new Point(5, 9),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 9f)
            };

            _rbFlat = new RadioButton
            {
                Text = "Flat List",
                Checked = true,
                AutoSize = true,
                Location = new Point(400, 7)
            };
            _rbFlat.CheckedChanged += DisplayMode_Changed;

            _rbTree = new RadioButton
            {
                Text = "Tree View",
                AutoSize = true,
                Location = new Point(490, 7)
            };

            _lblFieldCount = new Label
            {
                Text = "",
                AutoSize = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                ForeColor = Color.FromArgb(0, 120, 0),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };

            _toolbarPanel.Controls.AddRange(new Control[] { _lblFilePath, _rbFlat, _rbTree, _lblFieldCount });
            _toolbarPanel.Resize += ToolbarPanel_Resize;

            // -- Results area --
            _resultsPanel = new Panel { Dock = DockStyle.Fill };

            _lvFields = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Font = new Font("Segoe UI", 9f)
            };
            _lvFields.Columns.Add("Field Path", 350);
            _lvFields.Columns.Add("Tag", 200);
            _lvFields.Columns.Add("Alias", 200);
            _lvFields.Columns.Add("Location", 180);

            _tvFields = new TreeView
            {
                Dock = DockStyle.Fill,
                Visible = false,
                Font = new Font("Segoe UI", 9f),
                ShowLines = true,
                ShowPlusMinus = true,
                ShowRootLines = true
            };

            _resultsPanel.Controls.Add(_lvFields);
            _resultsPanel.Controls.Add(_tvFields);

            // -- Assemble (order matters for Dock: Fill must be added first) --
            Controls.Add(_resultsPanel);
            Controls.Add(_toolbarPanel);
            Controls.Add(_dropZonePanel);

            Name = "DocumentXRayControl";
            Size = new Size(900, 600);

            ResumeLayout(false);
        }

        private void ToolbarPanel_Resize(object sender, EventArgs e)
        {
            // Keep radio buttons and field count positioned relative to panel width
            var w = _toolbarPanel.Width;
            _rbFlat.Left = w - 240;
            _rbTree.Left = w - 150;
            _lblFieldCount.Left = w - _lblFieldCount.Width - 10;
            _lblFieldCount.Top = 9;
        }

        private void DropZonePanel_Paint(object sender, PaintEventArgs e)
        {
            var rect = new Rectangle(10, 10, _dropZonePanel.Width - 21, _dropZonePanel.Height - 21);
            using (var pen = new Pen(Color.FromArgb(180, 180, 200), 2f))
            {
                pen.DashStyle = DashStyle.Dash;
                e.Graphics.DrawRectangle(pen, rect);
            }
        }

        private void DropZonePanel_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Any(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
                {
                    e.Effect = DragDropEffects.Copy;
                    _dropZonePanel.BackColor = Color.FromArgb(220, 230, 250);
                    return;
                }
            }
            e.Effect = DragDropEffects.None;
        }

        private void DropZonePanel_DragLeave(object sender, EventArgs e)
        {
            _dropZonePanel.BackColor = Color.FromArgb(245, 245, 250);
        }

        private void DropZonePanel_DragDrop(object sender, DragEventArgs e)
        {
            _dropZonePanel.BackColor = Color.FromArgb(245, 245, 250);
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var docx = files.FirstOrDefault(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase));
            if (docx != null)
                LoadDocument(docx);
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Word Documents (*.docx)|*.docx";
                dlg.Title = "Select a Dynamics 365 Word Template";
                if (dlg.ShowDialog() == DialogResult.OK)
                    LoadDocument(dlg.FileName);
            }
        }

        private void LoadDocument(string path)
        {
            try
            {
                _currentFields = DocxFieldExtractor.ExtractFields(path);
                _lblFilePath.Text = Path.GetFileName(path);
                _lblFilePath.ForeColor = Color.Black;

                if (_currentFields.Count == 0)
                {
                    _lblFieldCount.Text = "No Dynamics fields found";
                    _lblFieldCount.ForeColor = Color.FromArgb(180, 120, 0);
                }
                else
                {
                    _lblFieldCount.Text = $"{_currentFields.Count} field(s) found";
                    _lblFieldCount.ForeColor = Color.FromArgb(0, 120, 0);
                }

                DisplayResults();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error reading template", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DisplayMode_Changed(object sender, EventArgs e)
        {
            if (_currentFields != null)
                DisplayResults();
        }

        private void DisplayResults()
        {
            if (_rbFlat.Checked)
            {
                _tvFields.Visible = false;
                _lvFields.Visible = true;
                PopulateListView();
            }
            else
            {
                _lvFields.Visible = false;
                _tvFields.Visible = true;
                PopulateTreeView();
            }
        }

        private void PopulateListView()
        {
            _lvFields.Items.Clear();
            foreach (var f in _currentFields)
            {
                var item = new ListViewItem(f.FieldPath ?? "");
                item.SubItems.Add(f.Tag ?? "");
                item.SubItems.Add(f.Alias ?? "");
                item.SubItems.Add(f.Location ?? "");
                _lvFields.Items.Add(item);
            }
        }

        private void PopulateTreeView()
        {
            _tvFields.BeginUpdate();
            _tvFields.Nodes.Clear();

            var uniquePaths = _currentFields
                .Select(f => f.FieldPath)
                .Where(p => p != null)
                .Distinct()
                .OrderBy(p => p)
                .ToList();

            foreach (var path in uniquePaths)
            {
                var segments = path.Split('/');
                var nodes = _tvFields.Nodes;

                foreach (var segment in segments)
                {
                    var existing = nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == segment);
                    if (existing != null)
                    {
                        nodes = existing.Nodes;
                    }
                    else
                    {
                        var newNode = nodes.Add(segment);
                        nodes = newNode.Nodes;
                    }
                }
            }

            _tvFields.TreeViewNodeSorter = new TreeNodeAlphabeticSorter();
            _tvFields.Sort();
            _tvFields.ExpandAll();
            _tvFields.EndUpdate();
        }

        private class TreeNodeAlphabeticSorter : System.Collections.IComparer
        {
            public int Compare(object x, object y)
            {
                return string.Compare(((TreeNode)x).Text, ((TreeNode)y).Text, StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
