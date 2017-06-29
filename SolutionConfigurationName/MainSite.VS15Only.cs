using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Build.Evaluation;
using System.Runtime.CompilerServices;
using ATask = System.Threading.Tasks.Task;
using DTEProject = EnvDTE.Project;
using MsBuildProject = Microsoft.Build.Evaluation.Project;

namespace SolutionConfigurationName
{
    using Microsoft.VisualStudio.ProjectSystem.Properties;
    using Microsoft.VisualStudio.ProjectSystem;

    partial class MainSite
    {
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static async ATask SetVCProjectsConfigurationProperties15(DTEProject project,
            string configurationName, string platformName)
        {
            // Inspired from Nuget: https://github.com/Haacked/NuGet/blob/master/src/VisualStudio12/ProjectHelper.cs
            IVsBrowseObjectContext context = project.Object as IVsBrowseObjectContext;
            UnconfiguredProject unconfiguredProject = context.UnconfiguredProject;
            IProjectLockService service = unconfiguredProject.ProjectService.Services.ProjectLockService;

            using (ProjectWriteLockReleaser releaser = await service.WriteLockAsync())
            {
                ConfiguredProject configuredProject = await unconfiguredProject.GetSuggestedConfiguredProjectAsync();
                MsBuildProject buildProject = await releaser.GetProjectAsync(configuredProject);

                ConfigureProject(buildProject, configurationName, platformName);

                // The following was present in the NuGet code: it seesms unecessary,
                // as the lock it's release anyway after the using block (check
                // service.IsAnyLockHeld). Also it seemed to cause a deadlock sometimes
                // when switching solution configuration
                //await releaser.ReleaseAsync();
            }
        }
    }
}
