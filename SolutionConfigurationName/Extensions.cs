using EnvDTE;
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
    }
}
