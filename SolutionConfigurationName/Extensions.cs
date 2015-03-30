// Copyright (c) 2013-2014 Francesco Pretto
// This file is subject to the MIT license

using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.VCProjectEngine;

namespace SolutionConfigurationName
{
    public static partial class Extensions
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

        public static VCProject GetVCProject(this Project project)
        {
            return project.Object as VCProject;
        }

        public static Lazy<T, IDictionary<string, object>> FindExport<T>(IEnumerable<Lazy<T, IDictionary<string, object>>> collection, string metadataName, string metadataValue)
        {
            foreach (Lazy<T, IDictionary<string, object>> lazy in collection)
            {
                object obj = lazy.Metadata[metadataName];
                string str = obj as string;
                if (str != null)
                {
                    if (str.Equals(metadataValue, StringComparison.OrdinalIgnoreCase))
                        return lazy;
                }
                else
                {
                    string[] array = obj as string[];
                    if (array.Contains(metadataValue, StringComparer.OrdinalIgnoreCase))
                        return lazy;
                }
            }

            return null;
        }
    }
}
