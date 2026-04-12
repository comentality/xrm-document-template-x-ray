using System.ComponentModel.Composition;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace DocumentXRay
{
    [Export(typeof(IXrmToolBoxPlugin))]
    [ExportMetadata("Name", "Document X-Ray")]
    [ExportMetadata("Description", "Extract Dynamics 365 field references from Word document templates")]
    [ExportMetadata("SmallImageBase64", null)]
    [ExportMetadata("BigImageBase64", null)]
    [ExportMetadata("BackgroundColor", "White")]
    [ExportMetadata("PrimaryFontColor", "#000000")]
    [ExportMetadata("SecondaryFontColor", "#999999")]
    public class DocumentXRayPlugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new DocumentXRayControl();
        }
    }
}
