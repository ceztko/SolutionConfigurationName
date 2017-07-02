// Copyright (c) 2013-2014 Francesco Pretto
// This file is subject to the MIT license

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
    public class UpdateSolutionEvents : IVsUpdateSolutionEvents3, IVsUpdateSolutionEvents2
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
            // This event may be asyncronous
            return VSConstants.S_OK;
        }


        public int UpdateProjectCfg_Begin(IVsHierarchy pHierProj, IVsCfg pCfgProj, IVsCfg pCfgSln, uint dwAction, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int UpdateProjectCfg_Done(IVsHierarchy pHierProj, IVsCfg pCfgProj, IVsCfg pCfgSln, uint dwAction, int fSuccess, int fCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterActiveSolutionCfgChange(IVsCfg pOldActiveSlnCfg, IVsCfg pNewActiveSlnCfg)
        {
            string previousConfiguration;
            if (pOldActiveSlnCfg == null)
                previousConfiguration = null;
            else
                pOldActiveSlnCfg.get_DisplayName(out previousConfiguration);

            // Set variables according the new active configuration
            MainSite.SetConfigurationProperties(previousConfiguration);

            return VSConstants.S_OK;
        }

        public int OnBeforeActiveSolutionCfgChange(IVsCfg pOldActiveSlnCfg, IVsCfg pNewActiveSlnCfg)
        {
            string solutionConfiguration;
            pNewActiveSlnCfg.get_DisplayName(out solutionConfiguration);
            MainSite.BeforeChangeActiveSolutionConfiguration(solutionConfiguration);
            return VSConstants.S_OK;
        }
    }
}
