using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentTemplateXRay.Logic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using McTools.Xrm.Connection;
using XrmToolBox.Extensibility;
using Label = System.Windows.Forms.Label;

namespace DocumentTemplateXRay
{
    public partial class DocumentTemplateXRayControl : PluginControlBase
    {
        private List<FieldInfo> _currentFields;
        private readonly List<TemplateItem> _templates = new List<TemplateItem>();

        // Controls - left panel
        private SplitContainer _splitContainer;
        private Panel _leftPanel;
        private Panel _leftToolbar;
        private Button _btnFetch;
        private Button _btnBrowseLocal;
        private ListView _lvTemplates;

        // Controls - right panel
        private Panel _rightPanel;
        private Panel _dropZonePanel;
        private Label _dropLabel;
        private Panel _toolbarPanel;
        private Label _lblFilePath;
        private RadioButton _rbFlat;
        private RadioButton _rbTree;
        private Label _lblFieldCount;
        private Panel _resultsPanel;
        private ListView _lvFields;
        private TreeView _tvFields;

        public DocumentTemplateXRayControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            // -- Split container --
            _splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 280,
                FixedPanel = FixedPanel.Panel1
            };

            // ===== LEFT PANEL: Template list =====
            _leftPanel = new Panel { Dock = DockStyle.Fill };

            _leftToolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 70,
                Padding = new Padding(5)
            };

            _btnFetch = new Button
            {
                Text = "Fetch from Dynamics",
                Location = new Point(5, 5),
                Width = 160,
                Height = 28
            };
            _btnFetch.Click += BtnFetch_Click;

            _btnBrowseLocal = new Button
            {
                Text = "Add Local File...",
                Location = new Point(5, 37),
                Width = 160,
                Height = 28
            };
            _btnBrowseLocal.Click += BtnBrowseLocal_Click;

            _leftToolbar.Controls.Add(_btnFetch);
            _leftToolbar.Controls.Add(_btnBrowseLocal);

            _lvTemplates = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                Font = new Font("Segoe UI", 9f),
                AllowDrop = true
            };
            _lvTemplates.Columns.Add("Template Name", 180);
            _lvTemplates.Columns.Add("Table", 80);
            _lvTemplates.SelectedIndexChanged += LvTemplates_SelectedIndexChanged;
            _lvTemplates.DragEnter += DropZone_DragEnter;
            _lvTemplates.DragLeave += DropZone_DragLeave;
            _lvTemplates.DragDrop += DropZone_DragDrop;

            _leftPanel.Controls.Add(_lvTemplates);
            _leftPanel.Controls.Add(_leftToolbar);

            // ===== RIGHT PANEL: Results =====
            _rightPanel = new Panel { Dock = DockStyle.Fill };

            // -- Drop zone (shown when no template selected) --
            _dropZonePanel = new Panel
            {
                Dock = DockStyle.Fill,
                AllowDrop = true,
                BackColor = Color.FromArgb(245, 245, 250)
            };
            _dropZonePanel.Paint += DropZonePanel_Paint;
            _dropZonePanel.DragEnter += DropZone_DragEnter;
            _dropZonePanel.DragLeave += DropZone_DragLeave;
            _dropZonePanel.DragDrop += DropZone_DragDrop;

            _dropLabel = new Label
            {
                Text = "Select a template from the list\nor drag && drop a .docx file",
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 11f),
                ForeColor = Color.FromArgb(100, 100, 100)
            };
            _dropZonePanel.Controls.Add(_dropLabel);

            // -- Toolbar --
            _toolbarPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 35,
                Padding = new Padding(5, 5, 5, 0),
                Visible = false
            };

            _lblFilePath = new Label
            {
                Text = "",
                AutoSize = true,
                Location = new Point(5, 9),
                ForeColor = Color.Black,
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
            _resultsPanel = new Panel { Dock = DockStyle.Fill, Visible = false };

            _lvFields = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Font = new Font("Segoe UI", 9f)
            };
            _lvFields.Columns.Add("Field Path", 300);
            _lvFields.Columns.Add("Table", 160);
            _lvFields.Columns.Add("Column", 180);
            _lvFields.Columns.Add("Tag", 180);
            _lvFields.Columns.Add("Alias", 180);
            _lvFields.Columns.Add("Repeating Section", 160);
            _lvFields.Columns.Add("Location", 150);

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

            // Assemble right panel (order matters for Dock: Fill must be added first)
            _rightPanel.Controls.Add(_resultsPanel);
            _rightPanel.Controls.Add(_dropZonePanel);
            _rightPanel.Controls.Add(_toolbarPanel);

            // Assemble split container
            _splitContainer.Panel1.Controls.Add(_leftPanel);
            _splitContainer.Panel2.Controls.Add(_rightPanel);

            Controls.Add(_splitContainer);

            Name = "DocumentTemplateXRayControl";
            Size = new Size(1000, 600);

            ResumeLayout(false);
        }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);
            FetchTemplatesFromDynamics();
        }

        private void BtnFetch_Click(object sender, EventArgs e)
        {
            FetchTemplatesFromDynamics();
        }

        private void FetchTemplatesFromDynamics()
        {
            if (Service == null)
            {
                MessageBox.Show("Not connected to Dynamics. Please connect first.", "No Connection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Fetching document templates...",
                Work = (worker, args) =>
                {
                    var query = new QueryExpression("documenttemplate")
                    {
                        ColumnSet = new ColumnSet("name", "documenttype", "associatedentitytypecode", "content"),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("documenttype", ConditionOperator.Equal, 2) // 2 = Word
                            }
                        }
                    };
                    args.Result = Service.RetrieveMultiple(query);
                },
                PostWorkCallBack = result =>
                {
                    if (result.Error != null)
                    {
                        MessageBox.Show(result.Error.Message, "Error fetching templates", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var entities = (EntityCollection)result.Result;

                    // Remove existing Dynamics items, keep local ones
                    _templates.RemoveAll(t => !t.IsLocal);

                    foreach (var entity in entities.Entities)
                    {
                        var name = entity.GetAttributeValue<string>("name") ?? "(unnamed)";
                        var entityType = entity.GetAttributeValue<string>("associatedentitytypecode") ?? "";
                        var content = entity.GetAttributeValue<string>("content");

                        _templates.Add(new TemplateItem
                        {
                            Name = name,
                            EntityType = entityType,
                            Base64Content = content,
                            IsLocal = false
                        });
                    }

                    RefreshTemplateList();
                }
            });
        }

        private void BtnBrowseLocal_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Word Documents (*.docx)|*.docx";
                dlg.Title = "Select a Dynamics 365 Word Template";
                if (dlg.ShowDialog() == DialogResult.OK)
                    AddLocalFile(dlg.FileName);
            }
        }

        private void AddLocalFile(string path)
        {
            // Don't add duplicates
            if (_templates.Any(t => t.IsLocal && string.Equals(t.LocalPath, path, StringComparison.OrdinalIgnoreCase)))
            {
                // Select the existing one
                for (int i = 0; i < _lvTemplates.Items.Count; i++)
                {
                    var item = (TemplateItem)_lvTemplates.Items[i].Tag;
                    if (item.IsLocal && string.Equals(item.LocalPath, path, StringComparison.OrdinalIgnoreCase))
                    {
                        _lvTemplates.Items[i].Selected = true;
                        _lvTemplates.Items[i].EnsureVisible();
                        return;
                    }
                }
                return;
            }

            // Extract root entity from fields
            string rootEntity = null;
            try
            {
                var fields = DocxFieldExtractor.ExtractFields(path);
                var firstPath = fields.FirstOrDefault(f => f.FieldPath != null)?.FieldPath;
                if (firstPath != null)
                    rootEntity = firstPath.Split('/')[0];
            }
            catch { }

            _templates.Add(new TemplateItem
            {
                Name = Path.GetFileName(path),
                EntityType = rootEntity,
                LocalPath = path,
                IsLocal = true
            });

            RefreshTemplateList();

            // Select the newly added item (last one)
            var lastIndex = _lvTemplates.Items.Count - 1;
            if (lastIndex >= 0)
            {
                _lvTemplates.Items[lastIndex].Selected = true;
                _lvTemplates.Items[lastIndex].EnsureVisible();
            }
        }

        private void RefreshTemplateList()
        {
            _lvTemplates.Items.Clear();
            foreach (var t in _templates)
            {
                var displayName = t.IsLocal ? $"{t.Name} (local)" : t.Name;
                var item = new ListViewItem(displayName);
                item.SubItems.Add(t.EntityType ?? "");
                item.Tag = t;

                if (t.IsLocal)
                    item.ForeColor = Color.FromArgb(0, 120, 60);

                _lvTemplates.Items.Add(item);
            }
        }

        private void LvTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_lvTemplates.SelectedItems.Count == 0) return;

            var template = (TemplateItem)_lvTemplates.SelectedItems[0].Tag;
            LoadTemplate(template);
        }

        private void LoadTemplate(TemplateItem template)
        {
            try
            {
                string tempPath = null;
                bool cleanupTemp = false;

                if (template.IsLocal)
                {
                    tempPath = template.LocalPath;
                }
                else
                {
                    if (string.IsNullOrEmpty(template.Base64Content))
                    {
                        MessageBox.Show("Template has no content.", "Empty Template", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Decode base64 to temp file
                    tempPath = Path.Combine(Path.GetTempPath(), $"xray_{Guid.NewGuid():N}.docx");
                    File.WriteAllBytes(tempPath, Convert.FromBase64String(template.Base64Content));
                    cleanupTemp = true;
                }

                try
                {
                    _currentFields = DocxFieldExtractor.ExtractFields(tempPath);
                }
                finally
                {
                    if (cleanupTemp && File.Exists(tempPath))
                    {
                        try { File.Delete(tempPath); } catch { }
                    }
                }

                // Show results
                _dropZonePanel.Visible = false;
                _toolbarPanel.Visible = true;
                _resultsPanel.Visible = true;

                _lblFilePath.Text = template.IsLocal
                    ? template.Name
                    : $"{template.Name} ({template.EntityType})";

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
                ResolveDisplayNames();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error reading template", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResolveDisplayNames()
        {
            if (_currentFields == null || _currentFields.Count == 0) return;
            if (Service == null) return;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Resolving field display names...",
                Work = (worker, args) =>
                {
                    var resolver = new MetadataResolver(Service);
                    resolver.ResolveDisplayNames(_currentFields);
                },
                PostWorkCallBack = result =>
                {
                    if (result.Error != null)
                    {
                        _lblFieldCount.Text += " (metadata lookup failed)";
                    }

                    DisplayResults();
                }
            });
        }

        // -- Drag & drop --
        private void DropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Any(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
                {
                    e.Effect = DragDropEffects.Copy;
                    if (sender == _dropZonePanel)
                        _dropZonePanel.BackColor = Color.FromArgb(220, 230, 250);
                    return;
                }
            }
            e.Effect = DragDropEffects.None;
        }

        private void DropZone_DragLeave(object sender, EventArgs e)
        {
            if (sender == _dropZonePanel)
                _dropZonePanel.BackColor = Color.FromArgb(245, 245, 250);
        }

        private void DropZone_DragDrop(object sender, DragEventArgs e)
        {
            _dropZonePanel.BackColor = Color.FromArgb(245, 245, 250);
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var docx = files.FirstOrDefault(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase));
            if (docx != null)
                AddLocalFile(docx);
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

        // -- Display --
        private void DisplayMode_Changed(object sender, EventArgs e)
        {
            if (_currentFields != null)
                DisplayResults();
        }

        private void ToolbarPanel_Resize(object sender, EventArgs e)
        {
            var w = _toolbarPanel.Width;
            _rbFlat.Left = w - 240;
            _rbTree.Left = w - 150;
            _lblFieldCount.Left = w - _lblFieldCount.Width - 10;
            _lblFieldCount.Top = 9;
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
                item.SubItems.Add(f.TableDisplayName ?? "");
                item.SubItems.Add(f.ColumnDisplayName ?? "");
                item.SubItems.Add(f.Tag ?? "");
                item.SubItems.Add(f.Alias ?? "");

                string repeatInfo = "";
                if (f.IsRepeatingSection)
                    repeatInfo = "(section)";
                else if (f.RepeatingSectionName != null)
                    repeatInfo = f.RepeatingSectionName;
                item.SubItems.Add(repeatInfo);

                item.SubItems.Add(f.Location ?? "");

                if (f.IsRepeatingSection)
                {
                    item.Font = new Font(_lvFields.Font, FontStyle.Bold);
                    item.ForeColor = Color.FromArgb(0, 100, 160);
                }
                else if (f.RepeatingSectionName != null)
                {
                    item.ForeColor = Color.FromArgb(0, 100, 160);
                }

                _lvFields.Items.Add(item);
            }
        }

        private void PopulateTreeView()
        {
            _tvFields.BeginUpdate();
            _tvFields.Nodes.Clear();

            var repeatingSectionPaths = new HashSet<string>(
                _currentFields
                    .Where(f => f.IsRepeatingSection && f.FieldPath != null)
                    .Select(f => f.FieldPath),
                StringComparer.OrdinalIgnoreCase);

            var displayNameLookup = _currentFields
                .Where(f => f.FieldPath != null && f.ColumnDisplayName != null)
                .GroupBy(f => f.FieldPath, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First().ColumnDisplayName, StringComparer.OrdinalIgnoreCase);

            var tableNameLookup = _currentFields
                .Where(f => f.FieldPath != null && f.TableDisplayName != null)
                .GroupBy(f =>
                {
                    var lastSlash = f.FieldPath.LastIndexOf('/');
                    return lastSlash > 0 ? f.FieldPath.Substring(0, lastSlash) : f.FieldPath;
                }, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First().TableDisplayName, StringComparer.OrdinalIgnoreCase);

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
                var builtPath = "";

                foreach (var segment in segments)
                {
                    builtPath = builtPath.Length == 0 ? segment : builtPath + "/" + segment;
                    var existing = nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text.StartsWith(segment));
                    if (existing != null)
                    {
                        nodes = existing.Nodes;
                    }
                    else
                    {
                        var isRepeating = repeatingSectionPaths.Contains(builtPath);
                        var isLeaf = (Array.IndexOf(segments, segment) == segments.Length - 1);
                        string displayText;
                        if (isRepeating)
                            displayText = segment + " (repeating)";
                        else if (isLeaf && displayNameLookup.TryGetValue(path, out var dn))
                            displayText = segment + "  [" + dn + "]";
                        else if (!isLeaf && tableNameLookup.TryGetValue(builtPath, out var tn))
                            displayText = segment + "  [" + tn + "]";
                        else
                            displayText = segment;
                        var newNode = nodes.Add(displayText);
                        if (isRepeating)
                        {
                            newNode.ForeColor = Color.FromArgb(0, 100, 160);
                            newNode.NodeFont = new Font(_tvFields.Font, FontStyle.Bold);
                        }
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

        private class TemplateItem
        {
            public string Name { get; set; }
            public string EntityType { get; set; }
            public string Base64Content { get; set; }
            public string LocalPath { get; set; }
            public bool IsLocal { get; set; }
        }
    }
}
