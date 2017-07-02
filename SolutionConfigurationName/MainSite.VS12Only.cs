// Copyright (c) 2013-2014 Francesco Pretto
// This file is subject to the MIT license

using System;
using Microsoft.Build.Evaluation;
using System.Runtime.CompilerServices;
using ATask = System.Threading.Tasks.Task;
using DTEProject = EnvDTE.Project;

namespace SolutionConfigurationName
{
    extern alias VS12;
    using VS12::Microsoft.VisualStudio.ProjectSystem.Designers;
    using VS12::Microsoft.VisualStudio.ProjectSystem;

    partial class MainSite
    {
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static async ATask SetVCProjectsConfigurationProperties12(DTEProject project,
            string configurationName, string platformName)
        {
            // Inspired from Nuget: https://github.com/Haacked/NuGet/blob/master/src/VisualStudio12/ProjectHelper.cs
            IVsBrowseObjectContext context = project.Object as IVsBrowseObjectContext;
            UnconfiguredProject unconfiguredProject = context.UnconfiguredProject;
            IProjectLockService service = unconfiguredProject.ProjectService.Services.ProjectLockService;

            using (ProjectWriteLockReleaser releaser = await service.WriteLockAsync())
            {
                ProjectCollection collection = releaser.ProjectCollection;

                ConfigureCollection(collection, configurationName, platformName);

                _VCProjectCollectionLoaded = true;

                // The following was present in the NuGet code: it seesms unecessary,
                // as the lock it's release anyway after the using block (check
                // service.IsAnyLockHeld). Also it seemed to cause a deadlock sometimes
                // when switching solution configuration
                //await releaser.ReleaseAsync();
            }
        }
    }
}
