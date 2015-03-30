// Copyright (c) 2013-2014 Francesco Pretto
// This file is subject to the MIT license

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.VCProjectEngine;
using Microsoft.Build.Evaluation;
using System.Reflection;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Project.Framework.INTERNAL.VS2010ONLY;
using Microsoft.VisualStudio.Project.Contracts.INTERNAL.VS2010ONLY;

namespace SolutionConfigurationName
{
    extern alias VC;
    using VCProjectShim = VC::Microsoft.VisualStudio.Project.VisualC.VCProjectEngine.VCProjectShim;
    using VCConfigurationShim = VC::Microsoft.VisualStudio.Project.VisualC.VCProjectEngine.VCConfigurationShim;

    public static partial class Extensions
    {
        static PropertyInfo _PropGetterConf;
        static PropertyInfo _PropGetterProj;

        static Extensions()
        {
            _PropGetterConf = typeof(VCConfigurationShim).GetProperty("ConfiguredProject", BindingFlags.Instance | BindingFlags.NonPublic);
            _PropGetterProj = typeof(VCProjectShim).GetProperty("ConfiguredProject", BindingFlags.Instance | BindingFlags.NonPublic);
        }

        public static IVsProject GetVsProject(this IVsHierarchy pHierarchy)
        {
            return pHierarchy as IVsProject;
        }

        public static ConfiguredProject GetConfiguredProject(this VCProject project)
        {
            return _PropGetterProj.GetValue(project, null) as ConfiguredProject;
        }

        public static ConfiguredProject GetConfiguredProject(this VCConfiguration configuration)
        {
            return _PropGetterConf.GetValue(configuration, null) as ConfiguredProject;
        }

        public static MSBuildProjectService GetProjectService(this ConfiguredProject confproj)
        {
            return confproj.GetServiceFeature<MSBuildProjectService>();
        }

        public static MSBuildProjectCollectionXmlService GetProjectCollectionXmlService(this ConfiguredProject confproj)
        {
            return confproj.GetServiceFeature<MSBuildProjectCollectionXmlService>();
        }

        public static MSBuildProjectXmlService GetProjectXmlService(this ConfiguredProject confproj)
        {
            return confproj.GetServiceFeature<MSBuildProjectXmlService>();
        }

        public static IProjectPropertiesProvider GetProjectPropertiesProvider(this ConfiguredProject project)
        {
            return GetFeature<IProjectPropertiesProvider>(project, "Name", "ProjectFile");
        }

        public static IProjectPropertiesProvider GetUserPropertiesProvider(this ConfiguredProject project)
        {
            return GetFeature<IProjectPropertiesProvider>(project, "Name", "UserFile");
        }

        public static T GetFeature<T>(IProjectContractQuery contract, string metadataName, string metadataValue)
        {
            Lazy<T, IDictionary<string, object>> export = FindExport(contract.GetFeature<T, IDictionary<string, object>>(), metadataName, metadataValue);
            return export.Value;
        }
    }
}
