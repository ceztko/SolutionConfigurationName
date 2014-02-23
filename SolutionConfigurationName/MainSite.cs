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
using Microsoft.VisualStudio.ProjectSystem;

namespace SolutionConfigurationName
{
    extern alias VC;

    using VCProjectEngineShim=VC::Microsoft.VisualStudio.Project.VisualC.VCProjectEngine.VCProjectEngineShim;

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidSolutionConfigurationNamePkgString)]
    //[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)] // Load if solution exists
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)] // Load if no solution
    public sealed class MainSite : Package
    {
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

        public static void SetConfigurationVariables()
        {
            SolutionConfiguration2 configuration =
                (SolutionConfiguration2)_DTE2.Solution.SolutionBuild.ActiveConfiguration;

            ProjectCollection global = ProjectCollection.GlobalProjectCollection;
            global.DisableMarkDirty = true;
            global.SetGlobalProperty("SolutionConfigurationName", configuration.Name);
            global.SetGlobalProperty("SolutionPlatformName", configuration.PlatformName);
            global.DisableMarkDirty = true;

#if VS12
            Type type = typeof(VCProjectEngineShim);
            object engine = type.GetProperty("Instance", BindingFlags.NonPublic | BindingFlags.Static).GetValue(null, null);
            if (engine == null)
                return;

            IProjectLockService service = (IProjectLockService)type.GetProperty("ProjectLockService", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(engine, null);
            // CHECK-ME How to acquire the lock?
            //ProjectWriteLockReleaser test = await service.WriteLockAsync(ProjectLockFlags.StickyWrite);
            //ProjectWriteLockAwaiter test2 = test.GetAwaiter();

            ProjectCollection vccollection = (ProjectCollection)type.GetProperty("ProjectCollection", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(engine, null);
            // CHECK-ME I need to acquire the lock first
            //vccollection.DisableMarkDirty = true;
            //vccollection.SetGlobalProperty("SolutionConfigurationName", configuration.Name);
            //vccollection.SetGlobalProperty("SolutionPlatformName", configuration.PlatformName);
            //vccollection.DisableMarkDirty = false;

            //test3.ReleaseAsync();
#endif
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
