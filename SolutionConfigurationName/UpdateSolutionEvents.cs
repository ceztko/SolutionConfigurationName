using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using EnvDTE80;
using Microsoft.Build.Evaluation;

namespace SolutionConfigurationName
{
    public class UpdateSolutionEvents : IVsUpdateSolutionEvents
    {
        public int OnActiveProjectCfgChange(IVsHierarchy pIVsHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int UpdateSolution_Begin(ref int pfCancelUpdate)
        {
            // Contrary to what Microsoft documentation seems to allude, this is really
            // the first syncronous event that gets fired when Solution update begins,
            // and not UpdateSolution_StartUpdate 
            SolutionConfiguration2 configuration =
                (SolutionConfiguration2)MainSite.DTE2.Solution.SolutionBuild.ActiveConfiguration;

            ProjectCollection global = ProjectCollection.GlobalProjectCollection;
            global.SetGlobalProperty("SolutionConfigurationName", configuration.Name);
            global.SetGlobalProperty("SolutionPlatformName", configuration.PlatformName);
            
            return VSConstants.S_OK;
        }

        public int UpdateSolution_Cancel()
        {
            return VSConstants.S_OK;
        }

        public int UpdateSolution_Done(int fSucceeded, int fModified, int fCancelCommand)
        {
            return VSConstants.S_OK;
        }

        public int UpdateSolution_StartUpdate(ref int pfCancelUpdate)
        {
            // Not setting variables here, build may begin before. Also it may be asyncronous
            return VSConstants.S_OK;
        }
    }
}
