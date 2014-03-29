// Copyright (c) 2013-2014 Francesco Pretto. Subject to the MIT license

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using ATask = System.Threading.Tasks.Task;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using EnvDTE80;
using Microsoft.Build.Evaluation;
using BuildProject = Microsoft.Build.Evaluation.Project;
using Microsoft.VisualStudio.ProjectSystem;
using Microsoft.VisualStudio.ProjectSystem.Designers;
using Microsoft.VisualStudio.VCProjectEngine;
using DTEProject = EnvDTE.Project;

namespace SolutionConfigurationName
{
#if DEBUG
    extern alias VC;
    using VCProjectShim = VC::Microsoft.VisualStudio.Project.VisualC.VCProjectEngine.VCProjectShim;
#endif

    partial class MainSite
    {
        private static volatile bool _VCProjectCollectionLoaded;
        private static AsyncLock _lock;

        static MainSite()
        {
            _VCProjectCollectionLoaded = false;
            _lock = new AsyncLock();
        }

        public static async void EnsureVCProjectsPropertiesConfigured(IVsHierarchy hiearchy)
        {
            using (await _lock.LockAsync())
            {
                if (_VCProjectCollectionLoaded)
                    return;

                DTEProject project = hiearchy.GetProject();
                if (project == null || !(project.Object is VCProject))
                    return;

                SolutionConfiguration2 configuration =
                    (SolutionConfiguration2)_DTE2.Solution.SolutionBuild.ActiveConfiguration;

                // When creating a completely new project the solution doesn't exist yet
                // so the ActiveConfiguration is null
                if (configuration == null)
                    return;

                // This is the first VC Project loaded, so we don't need to take
                // measures to ensure all projects are correctly marked as dirty
                await SetVCProjectsConfigurationProperties(project, configuration.Name, configuration.PlatformName, null);
            }
        }

        private static async ATask SetVCProjectsConfigurationProperties(DTEProject project,
            string configurationName, string platformName, HashSet<string> projectsToInvalidate)
        {
            // Inspired from Nuget: https://github.com/Haacked/NuGet/blob/master/src/VisualStudio12/ProjectHelper.cs
            IVsBrowseObjectContext context = project.Object as IVsBrowseObjectContext;
            UnconfiguredProject unconfiguredProject = context.UnconfiguredProject;
            IProjectLockService service = unconfiguredProject.ProjectService.Services.ProjectLockService;

            using (ProjectWriteLockReleaser releaser = await service.WriteLockAsync())
            {
                ProjectCollection collection = releaser.ProjectCollection;

                ConfigureCollection(collection, configurationName, platformName, projectsToInvalidate);

                _VCProjectCollectionLoaded = true;

                // The following was present in the NuGet code: it seesms unecessary,
                // as the lock it's release anyway after the using block (check
                // service.IsAnyLockHeld). Also it seemed to cause a deadlock sometimes
                // when switching solution configuration
                //await releaser.ReleaseAsync();
            }
        }

        private static async void SetVCProjectsConfigurationProperties(string configurationName, string platformName,
            HashSet<string> projectsToInvalidate)
        {
            foreach (DTEProject project in _DTE2.Solution.Projects.AllProjects())
            {
                if (!(project.Object is VCProject))
                    continue;

                await SetVCProjectsConfigurationProperties(project, configurationName, platformName, projectsToInvalidate);

                break;
            }

#if DEBUG
            foreach (DTEProject project in _DTE2.Solution.Projects.AllProjects())
            {
                VCProjectShim shim = project.Object as VCProjectShim;
                if (shim == null)
                    continue;

                bool test = shim.IsDirty;
            }
#endif
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
    }
}
