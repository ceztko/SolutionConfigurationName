using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BuildProject = Microsoft.Build.Evaluation.Project;
using DTEProject = EnvDTE.Project;

namespace SolutionConfigurationName
{
    partial class MainSite
    {
        public static void ExecuteWithinLock(DTEProject project, Action<BuildProject> action, object data)
        {
            using (var projectlock = project.GetVCProject().GetConfiguredProject().GetProjectService().GlobalCheckout(true))
            {
                action(projectlock.Project);
            }
        }
    }
}
