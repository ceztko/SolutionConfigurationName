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
#if VS12
using Microsoft.VisualStudio.ProjectSystem;
using Microsoft.VisualStudio.ProjectSystem.Designers;
using Microsoft.VisualStudio.VCProjectEngine;
#endif

using DTEProject = EnvDTE.Project;

namespace SolutionConfigurationName
{
#if VS12 && DEBUG
    extern alias VC;
    using VCProjectShim = VC::Microsoft.VisualStudio.Project.VisualC.VCProjectEngine.VCProjectShim;
#endif

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidSolutionConfigurationNamePkgString)]
    //[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)] // Load if solution exists
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)] // Load if no solution
    public sealed class MainSite : Package
    {
        private const string SCN_DUMMY_PROPERTY = "SCNDummy";
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
            ConfigureCollection(global, configurationName, platformName, true);

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

            // This is the first VC Project loaded, so we don't need to take
            // measures to ensure all projects are correctly marked as dirty
            SetVCProjectsConfigurationProperties(project, configuration.Name, configuration.PlatformName, false);
        }

        private static async void SetVCProjectsConfigurationProperties(DTEProject project,
            string configurationName, string platformName, bool allprojects)
        {
            // Inspired from Nuget: https://github.com/Haacked/NuGet/blob/master/src/VisualStudio12/ProjectHelper.cs
            IVsBrowseObjectContext context = project.Object as IVsBrowseObjectContext;
            UnconfiguredProject unconfiguredProject = context.UnconfiguredProject;
            IProjectLockService service = unconfiguredProject.ProjectService.Services.ProjectLockService;

            using (ProjectWriteLockReleaser releaser = await service.WriteLockAsync())
            {
                ProjectCollection collection = releaser.ProjectCollection;
                ConfigureCollection(collection, configurationName, platformName, allprojects);

                if (!allprojects)
                {
                    await releaser.CheckoutAsync(unconfiguredProject.FullPath);
                    ConfiguredProject configuredProject = await unconfiguredProject.GetSuggestedConfiguredProjectAsync();
                    BuildProject buildproj = await releaser.GetProjectAsync(configuredProject);

                    // Check ConfigureCollection() method for explanation
                    ProjectProperty prop = buildproj.SetProperty(SCN_DUMMY_PROPERTY, SCN_DUMMY_PROPERTY);
                    buildproj.RemoveProperty(prop);
                }

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

                SetVCProjectsConfigurationProperties(project, configurationName, platformName, true);
#if DEBUG
                // The VCProject should be dirty when switching soulution configuration
                VCProjectShim shim = project.Object as VCProjectShim;
                bool test = shim.IsDirty;
#endif
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

        public static void ConfigureCollection(ProjectCollection collection,
            string configurationName, string platformName, bool allprojects)
        {
            collection.SkipEvaluation = true;

            collection.SetGlobalProperty(SOLUTION_CONFIGURATION_MACRO, configurationName);
            collection.SetGlobalProperty(SOLUTION_PLATFORM_MACRO, platformName);

            if (allprojects)
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
                    // the VCProject is marked dirty

                    ProjectProperty prop = project.SetProperty(SCN_DUMMY_PROPERTY, SCN_DUMMY_PROPERTY);
                    project.RemoveProperty(prop);
                }
            }

            collection.SkipEvaluation = false;
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
