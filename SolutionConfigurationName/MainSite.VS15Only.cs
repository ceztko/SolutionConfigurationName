// Copyright (c) 2013-2014 Francesco Pretto
// This file is subject to the MIT license

using System;
using System.Runtime.CompilerServices;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Reflection;
using SolutionConfigurationNameMef;

namespace SolutionConfigurationName
{
    partial class MainSite : IDisposable
    {
        CompositionContainer _mefContainer;

        void LoadMef()
        {
            var assembly = Assembly.Load("SolutionConfigurationNameMef");
            var catalog = new AggregateCatalog(
                    new AssemblyCatalog(assembly));

            //Create a composition container
            _mefContainer = new CompositionContainer(catalog);
            _mefContainer.ComposeParts();
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        static void UpdateSolutionConfigurationMEF(string configurationName, string platformName)
        {
            SolutionConfigurationPropertiesProvider.SetSolutionConfiguration(configurationName, platformName);
        }

        public void Dispose()
        {
            _mefContainer.Dispose();
        }
    }
}
