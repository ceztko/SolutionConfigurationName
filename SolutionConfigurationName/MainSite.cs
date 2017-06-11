// Copyright (c) 2013-2014 Francesco Pretto
// This file is subject to the MIT license

using System;
using System.Linq;
using System.Diagnostics;
using System.Globalization;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using EnvDTE80;
using Microsoft.Build.Evaluation;
using System.Reflection;
using BuildProject = Microsoft.Build.Evaluation.Project;
using DTEProject = EnvDTE.Project;
using System.IO;

namespace SolutionConfigurationName
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidSolutionConfigurationNamePkgString)]
    //[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)] // Load if solution exists
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]       // Load if no solution
    public sealed partial class MainSite : Package                   // Sealed is needed in VS2013
    {
        private const string SCN_DUMMY_PROPERTY = "__SCNDummy";
        private const string SOLUTION_CONFIGURATION_MACRO = "SolutionConfiguration";
        private const string SOLUTION_PLATFORM_MACRO = "SolutionPlatform";

        private static DTE2 _DTE2;
        private static UpdateSolutionEvents _UpdateSolutionEvents;
        private static SolutionEvents _SolutionEvents;

        public MainSite() { }

        protected override void Initialize()
        {
            base.Initialize();
            IVsExtensibility extensibility = GetService<IVsExtensibility>();
            _DTE2 = (DTE2)extensibility.GetGlobalsObject(null).DTE;

            IVsSolution solution = GetService<SVsSolution>() as IVsSolution;
            IVsCfgProvider2 test = solution as IVsCfgProvider2;
            _SolutionEvents = new SolutionEvents();
            int hr;
            uint pdwCookie;
            hr = solution.AdviseSolutionEvents(_SolutionEvents, out pdwCookie);
            Marshal.ThrowExceptionForHR(hr);

            IVsSolutionBuildManager3 vsSolutionBuildManager = (IVsSolutionBuildManager3)GetService<SVsSolutionBuildManager>();
            _UpdateSolutionEvents = new UpdateSolutionEvents();
            hr = vsSolutionBuildManager.AdviseUpdateSolutionEvents3(_UpdateSolutionEvents, out pdwCookie);
            Marshal.ThrowExceptionForHR(hr);
        }
        public static void SetConfigurationProperties(string previousConfiguration)
        {
            SolutionConfiguration2 configuration =
                (SolutionConfiguration2)_DTE2.Solution.SolutionBuild.ActiveConfiguration;
            if (configuration == null)
                return;

            List<DTEProject> projectsToInvalidate = DetermineProjectsToInvalidate(previousConfiguration);

            List<DTEProject> projectsToInvalidateSaved = new List<DTEProject>();
            foreach(DTEProject project in projectsToInvalidate)
            {
                if (project.Saved)
                    projectsToInvalidateSaved.Add(project);
            }

            string configurationName = configuration.Name;
            string platformName = configuration.PlatformName;

            ProjectCollection global = ProjectCollection.GlobalProjectCollection;
            ConfigureCollection(global, configurationName, platformName);
#if VS12
            SetVCProjectsConfigurationProperties(configurationName, platformName, projectsToInvalidate);
#endif

            InvalidateProjects(projectsToInvalidate);
        }

        private static void InvalidateProjects(IEnumerable<DTEProject> projectsToInvalidate)
        {
            // Set and remove a dummy property and remove it immediately
            // to mark the project as dirty properly
            // CHECK-ME While the BuildProject is indeed reeavaluated, setting the
            // global property does not really mark the project as dirty everywhere:
            // in some circustances the VCProject is still considered up-to date.
            // For example, when $(SolutionConfiguration) is used in $(OutDir)
            // and the project is set to build in Relase both in Debug/Release
            // solution configurations targets, the build system doesn't realize
            // it has to recompile the project when switching from Release to
            // Debug. Understand if this is a bug or find a better way to ensure
            // the VCProject is marked dirty. Check Resources\Test project
            foreach (DTEProject dteproject in projectsToInvalidate)
            {
                dteproject.Globals[SCN_DUMMY_PROPERTY] = SCN_DUMMY_PROPERTY;
                // The following is needed to truly invalidate the project
                dteproject.Globals.VariablePersists[SCN_DUMMY_PROPERTY] = true;
                dteproject.Globals.VariablePersists[SCN_DUMMY_PROPERTY] = false;
                dteproject.Save();
            }
        }

        private static void ConfigureCollection(ProjectCollection collection,
            string configurationName, string platformName)
        {
            collection.SkipEvaluation = true;

            collection.SetGlobalProperty(SOLUTION_CONFIGURATION_MACRO, configurationName);
            collection.SetGlobalProperty(SOLUTION_PLATFORM_MACRO, platformName);

            collection.SkipEvaluation = false;
        }
        
        // Visual Studio doesn't invalidate build objects just by changing global
        // properties. When project build configuration is the same when switching
        // solution configuration we need to invalidate them manually
        private static List<DTEProject> DetermineProjectsToInvalidate(string previousConfiguration)
        {
            List<DTEProject> ret = new List<DTEProject>();
            if (previousConfiguration == null)
                return ret;

            string[] tokens = previousConfiguration.Split('|');
            string prevSolCfgName = tokens[0];
            string prevSolCfgPlatform = tokens[1];

            SolutionConfiguration prevSolCfg = null;
            foreach (SolutionConfiguration2 solCfg in _DTE2.Solution.SolutionBuild.SolutionConfigurations)
            {
                // Item() indexer doesn't work with full config name
                if (solCfg.Name == prevSolCfgName && solCfg.PlatformName == prevSolCfgPlatform)
                {
                    prevSolCfg = solCfg;
                    break;
                }
            }

            SolutionConfiguration currSolCfg = _DTE2.Solution.SolutionBuild.ActiveConfiguration;
            foreach (DTEProject project in _DTE2.Solution.Projects.AllProjects())
            {
                if (project.Kind == EnvDTE.Constants.vsProjectKindSolutionItems || project.ConfigurationManager == null)
                    continue;

                try
                {
                    string prevPrjCfgName = currSolCfg.SolutionContexts.Item(project.UniqueName).ConfigurationName;
                    string currPrjCfgName = prevSolCfg.SolutionContexts.Item(project.UniqueName).ConfigurationName;

                    if (prevPrjCfgName == currPrjCfgName)
                        ret.Add(project);
                }
                catch
                {
                    // Safeguard, just in case SolutionContexts Item() indexer fails
                    continue;
                }
            }

            return ret;
        }

        private void GetService<T>(out T service)
        {
            service = (T)GetService(typeof(T));
        }

        private T GetService<T>()
        {
            return (T)GetService(typeof(T));
        }

        public static DTE2 DTE2
        {
            get { return _DTE2; }
        }

        public static DTEVersion Version
        {
            get
            {
                return GetVersion(_DTE2.Version);
            }
        }

        public static DTEVersion GetVersion(String str)
        {
            switch (str)
            {
                case "10.0":
                    return DTEVersion.VS10;
                case "11.0":
                    return DTEVersion.VS11;
                case "12.0":
                    return DTEVersion.VS12;
                case "14.0":
                    return DTEVersion.VS14;
                case "15.0":
                    return DTEVersion.VS15;
                default:
                    throw new Exception();
            }
        }
    }

    public enum DTEVersion
    {
        Latest,
        VS10,
        VS11,
        VS12,
        VS14,
        VS15,
    }
}
