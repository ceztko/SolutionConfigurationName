using System;
using System.Diagnostics;
using System.Globalization;
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

namespace SolutionConfigurationName
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidSolutionConfigurationNamePkgString)]
    //[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)] // Load if solution exists
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]       // Load if no solution
    public sealed partial class MainSite : Package                   // Sealed is needed in VS2013
    {
        private const string SCN_DUMMY_PROPERTY = "SCNDummy";
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

        public static void SetConfigurationProperties()
        {
            SolutionConfiguration2 configuration =
                (SolutionConfiguration2)_DTE2.Solution.SolutionBuild.ActiveConfiguration;
            if (configuration == null)
                return;

            string configurationName = configuration.Name;
            string platformName = configuration.PlatformName;

            ProjectCollection global = ProjectCollection.GlobalProjectCollection;
            ConfigureCollection(global, null, configurationName, platformName);

#if VS12
            SetVCProjectsConfigurationProperties(configurationName, platformName);
#endif
        }

        public static void ConfigureCollection(ProjectCollection collection,
            BuildProject singleproj, string configurationName, string platformName)
        {
            collection.DisableMarkDirty = true;
            collection.SkipEvaluation = true;

            collection.SetGlobalProperty(SOLUTION_CONFIGURATION_MACRO, configurationName);
            collection.SetGlobalProperty(SOLUTION_PLATFORM_MACRO, platformName);

            if (singleproj == null)
            {
                foreach (BuildProject project in collection.LoadedProjects)
                {
                    // Set and remove a dummy property and remove it immediately
                    // to mark the project as dirty properly

                    // CHECK-ME While the Project is indeed marked as dirty, setting the
                    // global property does not really mark the project as dirty at all
                    // levels: in some circustances the VCProject is still considered up-to
                    // date. For example, when $(SolutionConfiguration) is used in $(OutDir)
                    // and the project is set to build in Relase both in Debug/Release
                    // solution configurations targets, the build system doesn't realize
                    // it has to recompile the project when switching from Release to
                    // Debug. Understand if this is a bug or find a better way to ensure
                    // the VCProject is marked dirty. Check Resources\Test project

                    ProjectProperty prop = project.SetProperty(SCN_DUMMY_PROPERTY, SCN_DUMMY_PROPERTY);
                    project.RemoveProperty(prop);
                }
            }
            else
            {
                // Same as above but for a single specified project
                ProjectProperty prop = singleproj.SetProperty(SCN_DUMMY_PROPERTY, SCN_DUMMY_PROPERTY);
                singleproj.RemoveProperty(prop);
            }

            collection.SkipEvaluation = false;
            collection.DisableMarkDirty = false;
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
    }
}
