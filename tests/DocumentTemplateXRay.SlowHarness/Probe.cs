using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
// The tool has a FieldInfo of its own and reflection has one too, and this file uses both.
using FieldInfo = DocumentTemplateXRay.Logic.FieldInfo;

namespace DocumentTemplateXRay.SlowHarness
{
    /// <summary>
    /// The control from the outside, which is where a user is. Everything a scenario does is done
    /// for real - PerformClick, a row selected, a file handed over the way a drop hands one over -
    /// and everything it reads is what is on screen at that moment.
    ///
    /// It gets there through reflection: the fields are private because nothing outside the
    /// control has any business with them, and widening them for a test bench would be a worse
    /// thing to do to the tool than this is.
    /// </summary>
    public class Probe
    {
        private const BindingFlags Priv = BindingFlags.NonPublic | BindingFlags.Instance;

        private readonly DocumentTemplateXRayControl _control;

        public Probe(DocumentTemplateXRayControl control)
        {
            _control = control;
        }

        public T Field<T>(string name)
        {
            var field = typeof(DocumentTemplateXRayControl).GetField(name, Priv);
            if (field == null) throw new MissingFieldException("DocumentTemplateXRayControl", name);
            return (T)field.GetValue(_control);
        }

        // ===== what is on screen =====

        public Button Fetch { get { return Field<Button>("_btnFetch"); } }
        public Button Browse { get { return Field<Button>("_btnBrowseLocal"); } }

        public ListView Templates { get { return Field<ListView>("_lvTemplates"); } }
        public ListView Fields { get { return Field<ListView>("_lvFields"); } }
        public TreeView Tree { get { return Field<TreeView>("_tvFields"); } }

        public string FilePath { get { return Field<Label>("_lblFilePath").Text; } }
        /// <summary>Everything the pane says about the fields: the count, and what it admits.</summary>
        public string FieldCount
        {
            get
            {
                var note = Field<Label>("_lblNote").Text;
                var count = Field<Label>("_lblFieldCount").Text;
                return note.Length == 0 ? count : count + " - " + note;
            }
        }
        public string DropText { get { return Field<Label>("_dropLabel").Text; } }

        public bool ResultsShowing { get { return Field<Panel>("_resultsPanel").Visible; } }

        public List<FieldInfo> CurrentFields { get { return Field<List<FieldInfo>>("_currentFields"); } }

        public List<string> TemplateNames()
        {
            return Templates.Items.Cast<ListViewItem>().Select(i => i.Text).ToList();
        }

        public string SelectedTemplate()
        {
            return Templates.SelectedItems.Count == 0 ? null : Templates.SelectedItems[0].Text;
        }

        /// <summary>The field paths the flat list is showing, in the order they are drawn.</summary>
        public List<string> FieldPaths()
        {
            return Fields.Items.Cast<ListViewItem>().Select(i => i.Text).ToList();
        }

        /// <summary>The rows whose Table and Column columns actually say something.</summary>
        public List<string> ResolvedPaths()
        {
            return Fields.Items.Cast<ListViewItem>()
                .Where(i => i.SubItems.Count > 2
                            && i.SubItems[1].Text.Length > 0
                            && i.SubItems[2].Text.Length > 0)
                .Select(i => i.Text)
                .ToList();
        }

        // ===== what a user does =====

        public void PressFetch() { Fetch.PerformClick(); }

        public void SelectTemplate(string name)
        {
            var row = Templates.Items.Cast<ListViewItem>()
                .FirstOrDefault(i => i.Text.IndexOf(name, StringComparison.Ordinal) >= 0);
            if (row == null)
                throw new InvalidOperationException("No template named " + name + " in the list ("
                                                    + string.Join(", ", TemplateNames()) + ")");
            row.Selected = true;
            row.Focused = true;
        }

        /// <summary>
        /// A .docx handed to the tool, which is what a drag and drop amounts to: the drop handler
        /// picks the file out of the data object and calls this. Synthesising the drop itself
        /// would be testing OLE rather than the tool.
        /// </summary>
        public void AddLocal(string path)
        {
            var method = typeof(DocumentTemplateXRayControl).GetMethod("AddLocalFile", Priv);
            if (method == null) throw new MissingMethodException("DocumentTemplateXRayControl", "AddLocalFile");
            method.Invoke(_control, new object[] { path });
        }

        /// <summary>
        /// What the panel's Cancel button does. Only reachable through reflection because the
        /// panel is XrmToolBox's, drawn over the tool by code the tool does not own.
        /// </summary>
        public void PressCancel()
        {
            var method = _control.GetType().GetMethod("CancelWorker",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (method == null) throw new MissingMethodException("PluginControlBase", "CancelWorker");
            method.Invoke(_control, null);
        }
    }
}
