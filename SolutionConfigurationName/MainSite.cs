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
#if VS12
using Microsoft.VisualStudio.ProjectSystem;
using BuildProject = Microsoft.Build.Evaluation.Project;
#endif

using DTEProject = EnvDTE.Project;

namespace SolutionConfigurationName
{
#if VS12
    extern alias VC;
    using VCProjectEngineShim=VC::Microsoft.VisualStudio.Project.VisualC.VCProjectEngine.VCProjectEngineShim;
    using Microsoft.VisualStudio.ProjectSystem.Designers;
    using Microsoft.VisualStudio.VCProjectEngine;
#endif

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidSolutionConfigurationNamePkgString)]
    //[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)] // Load if solution exists
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)] // Load if no solution
    public sealed class MainSite : Package
    {
        private const string SOLUTION_CONFIGURATION_MACRO = "SolutionConfiguration";
        private const string SOLUTION_PLATFORM_MACRO = "SolutionPlatform";

        private static DTE2 _DTE2;
        private static UpdateSolutionEvents _UpdateSolutionEvents;
        private static SolutionEvents _SolutionEvents;
#if VS12
        private static volatile bool _VCProjectCollectionLoaded;

        static MainSite()
        {
            _VCProjectCollectionLoaded = false;
        }
#endif

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
            global.SkipEvaluation = true;
            global.SetGlobalProperty(SOLUTION_CONFIGURATION_MACRO, configurationName);
            global.SetGlobalProperty(SOLUTION_PLATFORM_MACRO, platformName);
            global.SkipEvaluation = false;

#if VS12
            SetVCProjectsConfigurationProperties(configurationName, platformName);
#endif
        }

#if VS12
        public static void EnsureVCProjectsPropertiesConfigured(IVsHierarchy hiearchy)
        {
            if (_VCProjectCollectionLoaded)
                return;

            DTEProject project = hiearchy.GetProject();
            if (project == null || !(project.Object is VCProject))
                return;

            SolutionConfiguration2 configuration =
                (SolutionConfiguration2)_DTE2.Solution.SolutionBuild.ActiveConfiguration;

            SetVCProjectsConfigurationProperties(project, configuration.Name, configuration.PlatformName);
        }

        private static async void SetVCProjectsConfigurationProperties(DTEProject project,
            string configurationName, string platformName)
        {
            IVsBrowseObjectContext context = project.Object as IVsBrowseObjectContext;
            UnconfiguredProject unconfiguredProject = context.UnconfiguredProject;
            IProjectLockService service = unconfiguredProject.ProjectService.Services.ProjectLockService;

            using (ProjectWriteLockReleaser releaser = await service.WriteLockAsync())
            {
                ProjectCollection collection = releaser.ProjectCollection;
                collection.SkipEvaluation = true;
                collection.SetGlobalProperty(SOLUTION_CONFIGURATION_MACRO, configurationName);
                collection.SetGlobalProperty(SOLUTION_PLATFORM_MACRO, platformName);
                collection.SkipEvaluation = false;

                _VCProjectCollectionLoaded = true;

                await releaser.ReleaseAsync();
            }
        }

        public static void SetVCProjectsConfigurationProperties(string configurationName, string platformName)
        {
            foreach (DTEProject project in _DTE2.Solution.Projects)
            {
                if (!(project.Object is VCProject))
                    continue;

                SetVCProjectsConfigurationProperties(project, configurationName, platformName);

                break;
            }
        }

        /* Alternative method to obtain the VCProject(s) collection
        private static void foo()
        {
            Type type = typeof(VCProjectEngineShim);
            object engine = type.GetProperty("Instance", BindingFlags.NonPublic | BindingFlags.Static).GetValue(null, null);
            if (engine == null)
                return;

            ProjectCollection collection = (ProjectCollection)type.GetProperty("ProjectCollection", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(engine, null);

            IProjectLockService service = (IProjectLockService)type.GetProperty("ProjectLockService", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(engine, null);
        }
        */
#endif
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
