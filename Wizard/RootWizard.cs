using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio.TemplateWizard;

namespace VisioComCppAddinWizard
{
    // Root wizard is used by root project vstemplate

    public class RootWizard : IWizard
    {
        // Use to communicate $saferootprojectname$ to ChildWizard
        public static Dictionary<string, string> GlobalDictionary =
            new Dictionary<string,string>();

        // Add global replacement parameters
        public void RunStarted(object automationObject, 
            Dictionary<string, string> replacementsDictionary, 
            WizardRunKind runKind, object[] customParams)
        {
            var dte = automationObject as DTE;

            GlobalDictionary["$PLATFORMTOOLSET$"] = "v" + dte.Version.Replace(".", "");

            // Place "$saferootprojectname$ in the global dictionary.
            GlobalDictionary["$VCXPROJECTNAME$"] = replacementsDictionary["$safeprojectname$"];

            GlobalDictionary["$PROGID$"] = replacementsDictionary["$safeprojectname$"] + ".VisioConnect";
            GlobalDictionary["$LIBID$"] = replacementsDictionary["$guid1$"];
            GlobalDictionary["$CLSID$"] = replacementsDictionary["$guid2$"];
            GlobalDictionary["$ITFID$"] = replacementsDictionary["$guid3$"];

            GlobalDictionary["$WIXPROJECTGUID$"] = replacementsDictionary["$guid4$"];
            GlobalDictionary["$VCXPROJECTGUID$"] = replacementsDictionary["$guid5$"];
        }

        public void RunFinished()
        {
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectFinishedGenerating(Project project)
        {
        }

        public bool ShouldAddProjectItem(string filePath)
        {
            return true;
        }

        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {
        }
    }
}
