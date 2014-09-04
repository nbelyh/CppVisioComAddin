using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TemplateWizard;

namespace VisioComCppAddinWizard
{
    // Child wizard is used by child project vstemplates

    public class ChildWizard : IWizard
    {
        // Retrieve global replacement parameters
        public void RunStarted(object automationObject, 
            Dictionary<string, string> replacementsDictionary, 
            WizardRunKind runKind, object[] customParams)
        {
            // Add custom parameters.
            replacementsDictionary.Add("$FILTERGUID1$", Guid.NewGuid().ToString());
            replacementsDictionary.Add("$FILTERGUID2$", Guid.NewGuid().ToString());
            
            replacementsDictionary.Add("$PLATFORMTOOLSET$", RootWizard.GlobalDictionary["$PLATFORMTOOLSET$"]);
            replacementsDictionary.Add("$VCXPROJECTNAME$", RootWizard.GlobalDictionary["$VCXPROJECTNAME$"]);

            replacementsDictionary.Add("$PROGID$", RootWizard.GlobalDictionary["$PROGID$"]);
            replacementsDictionary.Add("$CLSID$", RootWizard.GlobalDictionary["$CLSID$"]);
            replacementsDictionary.Add("$LIBID$", RootWizard.GlobalDictionary["$LIBID$"]);
            replacementsDictionary.Add("$ITFID$", RootWizard.GlobalDictionary["$ITFID$"]);

            replacementsDictionary.Add("$VCXPROJECTGUID$", RootWizard.GlobalDictionary["$VCXPROJECTGUID$"]);
            replacementsDictionary.Add("$WIXPROJECTGUID$", RootWizard.GlobalDictionary["$WIXPROJECTGUID$"]);
        }

        public void RunFinished()
        {
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectFinishedGenerating(Project project)
        {
            /*
            if (Path.GetExtension(project.FileName) == ".csproj")
                RootWizard.GlobalDictionary["$VCXPROJECTGUID$"] = GetProjectGuid(project);
             */
        }

        /*
        public static string GetProjectGuid(Project project)
        {
            IServiceProvider serviceProvider = new ServiceProvider(project.DTE as
                Microsoft.VisualStudio.OLE.Interop.IServiceProvider);

            var sol = (IVsSolution)serviceProvider.GetService(typeof(SVsSolution));

            IVsHierarchy proj;

            ErrorHandler.ThrowOnFailure(sol.GetProjectOfUniqueName(project.UniqueName, out proj));

            Guid projGuid;

            proj.GetGuidProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_ProjectIDGuid, out projGuid);

            return projGuid.ToString();
        }
         * */

        public bool ShouldAddProjectItem(string filePath)
        {
            return true;
        }

        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {
        }
    }
}
