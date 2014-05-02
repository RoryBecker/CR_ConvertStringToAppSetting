using System.ComponentModel.Composition;
using DevExpress.CodeRush.Common;

namespace CR_ConvertStringToAppSetting
{
    [Export(typeof(IVsixPluginExtension))]
    public class CR_ConvertStringToAppSettingExtension : IVsixPluginExtension { }
}