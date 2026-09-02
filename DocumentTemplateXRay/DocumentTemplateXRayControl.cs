using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentTemplateXRay.Logic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
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

        // ===== what is on the wire =====
        //
        // Nothing disables the window while the tool waits - the panel XrmToolBox draws is a small
        // box in the middle of the tab - so on a slow link every button stays live for the seconds
        // an answer takes, and the tool has to know for itself what is outstanding and whose
        // answer an arriving one is.

        /// <summary>Which fetch of the template list is current. Bumped by each new one.</summary>
        private int _fetchGeneration;

        /// <summary>Which template is being read. Bumped whenever another one is opened.</summary>
        private int _readGeneration;

        private bool _fetching;
        private bool _reading;
        private bool _resolving;

        /// <summary>What went wrong with the template list, or null. Outlives its dialog.</summary>
        private string _fetchTrouble;

        /// <summary>What the pane has to admit about the names in it, or null.</summary>
        private string _fieldNote;

        /// <summary>How many fields the template on screen turned out to have, in words.</summary>
        private string _fieldCount = "";

        /// <summary>
        /// Everything looked up about a table, kept for the life of the tab. A
        /// RetrieveEntityRequest carries every attribute and every relationship of a table, and
        /// three templates on the same four tables used to pay for all of it three times.
        /// </summary>
        private readonly Dictionary<string, EntityMetadata> _metadata =
            new Dictionary<string, EntityMetadata>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// How a resolve is called off. The worker holds the gate it was started with and asks it
        /// between tables; opening another template shuts that gate and opens a new one. A field
        /// rather than a token because it is read from a worker thread and written from this one.
        /// </summary>
        private class Gate
        {
            public volatile bool Closed;
        }

        private Gate _resolveGate = new Gate();

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
        private Label _lblNote;
        private Panel _resultsPanel;
        private ListView _lvFields;
        private TreeView _tvFields;

        public DocumentTemplateXRayControl()
        {
            InitializeComponent();
            Load += (s, e) =>
            {
                if (_splitContainer.Width > 0)
                    _splitContainer.SplitterDistance = _splitContainer.Width / 2;
            };
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            // -- Split container --
            _splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                FixedPanel = FixedPanel.None
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
                Text = DropZoneText,
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
                ForeColor = Color.FromArgb(0, 120, 0),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };

            // What the count cannot say for itself: that it is still being worked out, or that
            // the names beside it could not be read. It gets the space between the file name and
            // the view radios and ellipsises inside it, so however long it is it never pushes
            // anything else off the toolbar.
            _lblNote = new Label
            {
                Text = "",
                AutoSize = false,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.FromArgb(120, 120, 120),
                Font = new Font("Segoe UI", 9f)
            };

            _toolbarPanel.Controls.AddRange(new Control[] { _lblFilePath, _lblNote, _rbFlat, _rbTree, _lblFieldCount });
            _toolbarPanel.Resize += (s, e) => LayoutToolbar();

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

        /// <summary>
        /// A different environment, handed over by XrmToolBox.
        ///
        /// Everything the window is holding for the org being left goes first, before the base
        /// class is called. A connection made because somebody pressed a button arrives carrying
        /// that button's method name, and running it is the last thing the base class does - so a
        /// reset afterwards would throw away the very fetch the connection was made for, and
        /// leave an empty list and a live button behind.
        /// </summary>
        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            var busy = _fetching || _reading || _resolving;

            // Whatever is still on the wire was asked of the org being left, and its answer is
            // not this org's to show.
            _fetchGeneration++;
            _readGeneration++;
            _fetching = false;
            _reading = false;
            _resolving = false;
            _fetchTrouble = null;

            // Nor are its templates. A file somebody dragged in is theirs, not the org's, and
            // stays where it is - so only what came from Dynamics goes.
            var open = SelectedTemplate();
            _templates.RemoveAll(t => !t.IsLocal);
            RefreshTemplateList();

            // The pane is showing one template's fields. If that template came from the org
            // being left, it is not a thing this window can still claim to be looking at.
            if (open != null && !open.IsLocal)
            {
                ClearResults();
                ShowDropZone();
            }

            // The fetch pulls every template's whole file, and the display names behind it are
            // more round trips again. What is left of that is worth stopping rather than letting
            // it finish for an org nobody is looking at any more. Before the base class, or it
            // is the new org's fetch that gets cancelled.
            if (busy) CancelWorker();

            base.UpdateConnection(newService, detail, actionName, parameter);

            // Nobody opens this tool to look at an empty list. Not when the base class has just
            // run the fetch this connection was asked for, though - that would be two of them,
            // and each one carries every template's file.
            if (newService != null && string.IsNullOrEmpty(actionName))
            {
                FetchTemplatesFromDynamics();
            }
            else
            {
                UpdateState();
            }
        }

        private void BtnFetch_Click(object sender, EventArgs e)
        {
            // Through XrmToolBox rather than straight at the org: pressed with no connection,
            // this asks for one and is run again once there is one, instead of telling somebody
            // to go and connect and then forgetting they ever asked.
            ExecuteMethod(FetchTemplatesFromDynamics);
        }

        private void FetchTemplatesFromDynamics()
        {
            if (Service == null)
            {
                MessageBox.Show("Not connected to Dynamics. Please connect first.", "No Connection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var generation = ++_fetchGeneration;
            _fetchTrouble = null;
            _fetching = true;
            UpdateState();

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Fetching document templates...",
                // Every Word template in the environment comes back with its whole file attached,
                // which on a slow link is the longest wait the tool asks anybody to sit through.
                IsCancelable = true,
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
                    var found = Service.RetrieveMultiple(query);
                    if (worker.CancellationPending)
                    {
                        args.Cancel = true;
                        return;
                    }

                    args.Result = found;
                },
                PostWorkCallBack = result =>
                {
                    // The tab may have been closed while this was in the air, and on a slow link
                    // that is a window seconds wide rather than milliseconds. Nobody is owed an
                    // answer, a dialog least of all.
                    if (IsDisposed || generation != _fetchGeneration) return;

                    _fetching = false;

                    // Cancelled before Result, which rethrows after a cancel or an error.
                    if (result.Cancelled)
                    {
                        _fetchTrouble = "Cancelled.\nPress Fetch from Dynamics to ask again.";
                        UpdateState();
                        return;
                    }

                    if (result.Error != null)
                    {
                        // An environment that could not be reached must not go on reading as an
                        // environment with no templates in it.
                        _fetchTrouble = "The templates could not be fetched:\n" + result.Error.Message;
                        UpdateState();
                        MessageBox.Show(result.Error.Message, "Error fetching templates", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var entities = (EntityCollection)result.Result;

                    // The list is about to be rebuilt from scratch, which loses the selection -
                    // and the pane would go on showing a template that is no longer selected,
                    // no longer highlighted, and possibly no longer in the environment.
                    var open = SelectedTemplate();

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
                    Reselect(open);
                    UpdateState();
                }
            });
        }

        private TemplateItem SelectedTemplate()
        {
            return _lvTemplates.SelectedItems.Count == 0
                ? null
                : (TemplateItem)_lvTemplates.SelectedItems[0].Tag;
        }

        /// <summary>
        /// Puts the selection back on the template that was open, now that the list holds
        /// different objects for the same templates. Selecting it reads it again, which is the
        /// point: a refresh is asking whether it has changed.
        ///
        /// A template that is no longer in the environment cannot be reselected, and then the
        /// pane has to let go of it too rather than describing something nothing points at.
        /// </summary>
        private void Reselect(TemplateItem open)
        {
            if (open == null) return;

            foreach (ListViewItem row in _lvTemplates.Items)
            {
                var candidate = (TemplateItem)row.Tag;
                var same = open.IsLocal
                    ? candidate.IsLocal && string.Equals(candidate.LocalPath, open.LocalPath, StringComparison.OrdinalIgnoreCase)
                    : !candidate.IsLocal && candidate.Name == open.Name && candidate.EntityType == open.EntityType;

                if (!same) continue;

                row.Selected = true;
                row.EnsureVisible();
                return;
            }

            ClearResults();
            ShowDropZone();
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

            // The table the template is about is the first segment of its field paths, and the
            // only way to know it is to read the file. That used to happen here, on the thread
            // that draws the window, and again a moment later when the row was selected. Now it
            // happens once, on a worker, and the row is filled in when the answer comes back.
            _templates.Add(new TemplateItem
            {
                Name = Path.GetFileName(path),
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
            // A different template than the one being read, or resolved, or both. Whatever is on
            // its way belongs to the one before this and its answer is nobody's when it lands.
            var generation = ++_readGeneration;
            _resolveGate.Closed = true;
            _resolveGate = new Gate();
            _resolving = false;
            _reading = true;
            _fieldNote = null;

            // The pane describes one template. It stops describing the last one now, rather than
            // when the new one arrives, because the gap between those two moments is the whole of
            // what a slow disk or a slow link adds.
            ClearResults();
            ShowResults();
            _lblFilePath.Text = template.IsLocal
                ? template.Name
                : $"{template.Name} ({template.EntityType})";
            UpdateState();

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Reading the template...",
                // Decoding it, writing it out, unzipping it and scanning it: none of that belongs
                // on the thread that draws the window. A template is a Word file, and a Word file
                // is as likely to be on a share or a synced folder as on the local disk.
                Work = (worker, args) => args.Result = Read(template),
                PostWorkCallBack = result =>
                {
                    if (IsDisposed || generation != _readGeneration) return;

                    _reading = false;

                    if (result.Error != null)
                    {
                        _fieldCount = "";
                        _fieldNote = "the template could not be read";
                        UpdateState();
                        MessageBox.Show(result.Error.Message, "Error reading template", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    _currentFields = (List<FieldInfo>)result.Result;

                    // A file somebody dropped in only says which table it is about once it has
                    // been read, and now it has been.
                    if (template.IsLocal && string.IsNullOrEmpty(template.EntityType))
                    {
                        var firstPath = _currentFields.FirstOrDefault(f => f.FieldPath != null)?.FieldPath;
                        if (firstPath != null) template.EntityType = firstPath.Split('/')[0];
                        NoteTableOf(template);
                    }

                    _fieldCount = _currentFields.Count == 0
                        ? "No Dynamics fields found"
                        : $"{_currentFields.Count} field(s) found";

                    DisplayResults();
                    UpdateState();
                    ResolveDisplayNames(generation);
                }
            });
        }

        /// <summary>
        /// The fields in a template, off the UI thread. Everything it touches is its own: the
        /// bytes it was handed, a temp file of its own, and the list it gives back.
        /// </summary>
        private static List<FieldInfo> Read(TemplateItem template)
        {
            if (template.IsLocal) return DocxFieldExtractor.ExtractFields(template.LocalPath);

            if (string.IsNullOrEmpty(template.Base64Content))
                throw new InvalidOperationException("This template has no content.");

            var tempPath = Path.Combine(Path.GetTempPath(), $"xray_{Guid.NewGuid():N}.docx");
            File.WriteAllBytes(tempPath, Convert.FromBase64String(template.Base64Content));
            try
            {
                return DocxFieldExtractor.ExtractFields(tempPath);
            }
            finally
            {
                try { File.Delete(tempPath); } catch { }
            }
        }

        /// <summary>Fills in the Table cell of a dropped file, without rebuilding the list.</summary>
        private void NoteTableOf(TemplateItem template)
        {
            foreach (ListViewItem row in _lvTemplates.Items)
            {
                if (!ReferenceEquals(row.Tag, template)) continue;
                if (row.SubItems.Count > 1) row.SubItems[1].Text = template.EntityType ?? "";
                return;
            }
        }

        private void ResolveDisplayNames(int generation)
        {
            if (_currentFields == null || _currentFields.Count == 0) return;
            if (Service == null) return;

            var gate = _resolveGate;
            var paths = _currentFields.Select(f => f.FieldPath).Where(p => p != null).Distinct().ToList();
            // Copied here, on the thread that owns it, so the worker reads nothing this one writes.
            var known = new Dictionary<string, EntityMetadata>(_metadata, StringComparer.OrdinalIgnoreCase);

            _resolving = true;
            UpdateState();

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Resolving field display names...",
                // A table each, one after another. The one on the wire cannot be recalled; the
                // rest can, and on a slow link those are most of the wait.
                IsCancelable = true,
                Work = (worker, args) =>
                {
                    var resolver = new MetadataResolver(Service, known);
                    var names = resolver.Resolve(paths, () => gate.Closed || worker.CancellationPending);
                    if (gate.Closed || worker.CancellationPending)
                    {
                        args.Cancel = true;
                        return;
                    }

                    args.Result = new object[] { names, resolver.Cache, resolver.Unavailable };
                },
                PostWorkCallBack = result =>
                {
                    if (IsDisposed || generation != _readGeneration) return;

                    _resolving = false;

                    if (result.Cancelled)
                    {
                        _fieldNote = "cancelled, so the names below are logical names";
                        UpdateState();
                        return;
                    }

                    if (result.Error != null)
                    {
                        _fieldNote = "the display names could not be read";
                        UpdateState();
                        return;
                    }

                    var parts = (object[])result.Result;
                    var names = (Dictionary<string, ResolvedName>)parts[0];
                    var learned = (Dictionary<string, EntityMetadata>)parts[1];
                    var unavailable = (List<string>)parts[2];

                    foreach (var pair in learned) _metadata[pair.Key] = pair.Value;

                    foreach (var field in _currentFields)
                    {
                        ResolvedName name;
                        if (field.FieldPath == null || !names.TryGetValue(field.FieldPath, out name)) continue;
                        field.TableDisplayName = name.Table;
                        field.ColumnDisplayName = name.Column;
                    }

                    // A blank Table cell means one of two things, and they are not the same: the
                    // environment says there is no such column, or the environment could not be
                    // asked. Telling those apart is the whole point of the tool.
                    _fieldNote = unavailable.Count == 0
                        ? null
                        : "the display names could not be read for " + string.Join(", ", unavailable);

                    DisplayResults();
                    UpdateState();
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

        /// <summary>What the pane holds about one template, and nothing to do with which one.</summary>
        private void ClearResults()
        {
            _currentFields = null;
            _fieldCount = "";
            _lvFields.Items.Clear();
            _tvFields.Nodes.Clear();
        }

        private void ShowResults()
        {
            _dropZonePanel.Visible = false;
            _toolbarPanel.Visible = true;
            _resultsPanel.Visible = true;
        }

        private void ShowDropZone()
        {
            _dropZonePanel.Visible = true;
            _toolbarPanel.Visible = false;
            _resultsPanel.Visible = false;
            _lblFilePath.Text = "";
        }

        private const string DropZoneText = "Select a template from the list\nor drag && drop a .docx file";

        /// <summary>
        /// The buttons and the two lines of text, made to agree with what is actually true. One
        /// place, called from every path in and out of a fetch, because the failure paths are
        /// where a button that disables itself gets left dead for the rest of the session.
        /// </summary>
        private void UpdateState()
        {
            _btnFetch.Enabled = !_fetching;
            _btnBrowseLocal.Enabled = !_reading;

            // What is on screen while the list is empty, which is exactly when a failed fetch
            // would otherwise pass for an environment with no templates in it.
            _dropLabel.Text = _fetchTrouble ?? DropZoneText;

            _lblFieldCount.Text = _reading ? "Reading the template..." : _fieldCount;
            _lblNote.Text =
                _reading ? "" :
                _resolving ? "resolving display names..." :
                _fieldNote ?? "";

            var settled = !_reading && !_resolving && _fieldNote == null;
            _lblFieldCount.ForeColor =
                !settled ? Color.FromArgb(120, 120, 120) :
                _currentFields != null && _currentFields.Count > 0
                    ? Color.FromArgb(0, 120, 0)
                    : Color.FromArgb(180, 120, 0);

            LayoutToolbar();
        }

        // The field count sits hard against the right edge and the view radios pack in to its
        // left, so a long message ("No Dynamics fields found") pushes them along instead of
        // printing over them. The label auto-sizes, so this has to run whenever its text
        // changes, not only when the panel is resized.
        private void LayoutToolbar()
        {
            _lblFieldCount.Left = _toolbarPanel.Width - _lblFieldCount.Width - 10;
            _lblFieldCount.Top = 9;
            _rbTree.Left = _lblFieldCount.Left - _rbTree.Width - 15;
            _rbFlat.Left = _rbTree.Left - _rbFlat.Width - 10;

            var from = _lblFilePath.Right + 12;
            var to = _rbFlat.Left - 12;
            _lblNote.SetBounds(from, 7, Math.Max(to - from, 0), 20);
        }

        private void DisplayResults()
        {
            if (_currentFields == null)
            {
                _lvFields.Items.Clear();
                _tvFields.Nodes.Clear();
                return;
            }

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
