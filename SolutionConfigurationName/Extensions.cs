// Copyright (c) 2013-2014 Francesco Pretto. Subject to the MIT license

using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolutionConfigurationName
{
    public static class Extensions
    {
        public static Project GetProject(this IVsHierarchy hierarchy)
        {
            object project;
             if (!ErrorHandler.Succeeded(hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_ExtObject, out project)))
                 return null;

            return project as Project;
        }

        public static IEnumerable<Project> AllProjects(this Projects projects)
        {
            foreach (Project project in projects)
            {
                foreach (Project subproject in AllProjects(project))
                    yield return subproject;
            }
        }

        private static IEnumerable<Project> AllProjects(this Project project)
        {
            if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
            {
                // The Project is a Solution folder
                foreach (ProjectItem projectItem in project.ProjectItems)
                {
                    if (projectItem.SubProject != null)
                    {
                        // The ProjectItem is actually a Project
                        foreach (Project subproject in AllProjects(projectItem.SubProject))
                            yield return subproject;
                    }
                }
            }
#if MORE_PEDANTIC
            else if (project.ConfigurationManager != null)
#else
            else
#endif
                yield return project;
        }
    }
}
