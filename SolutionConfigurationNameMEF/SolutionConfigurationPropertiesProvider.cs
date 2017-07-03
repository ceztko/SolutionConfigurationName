using System;
using Microsoft.VisualStudio.ProjectSystem.Properties;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio.ProjectSystem;
using Microsoft.VisualStudio.ProjectSystem.Build;
using System.Threading.Tasks.Dataflow;
using System.Collections.Immutable;
using System.Threading;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SolutionConfigurationNameMef
{
    [Export(typeof(IProjectGlobalPropertiesProvider))]
    [AppliesTo("VisualC + VCProjectEngineFactory")] // Copied from internal VCGlobalPropertiesProvider
    public class SolutionConfigurationPropertiesProvider :
        ProjectValueDataSourceBase<IImmutableDictionary<string, string>>,
        IProjectGlobalPropertiesProvider
    {
        const string NAME_IDENTITY = "SolutionConfiguration";

        private static object _lock;
        private static SolutionConfigurationPropertiesProvider _instance;   // Lock guarded
        private static IImmutableDictionary<string, string> _properties;    // Lock guarded
        private static IComparable _version;                                // Lock guarded
        private static BroadcastBlock<IProjectVersionedValue<IImmutableDictionary<string, string>>> _brodcastBlock; // Lock guarded

        private static IReceivableSourceBlock<IProjectVersionedValue<IImmutableDictionary<string, string>>> _publicBlock;
        private static NamedIdentity _dataSourceKey;

        static SolutionConfigurationPropertiesProvider()
        {
            _version = 0L;
            _lock = new object();
            _properties = ImmutableDictionary<string, string>.Empty;
            _dataSourceKey = new NamedIdentity(NAME_IDENTITY);
        }

        // The scope is the solution ProjectCollection for VC projects
        // https://github.com/Microsoft/VSProjectSystem/blob/master/doc/overview/scopes.md
        [ImportingConstructor]
        protected SolutionConfigurationPropertiesProvider(IProjectService service)
            : base(service.Services)
        {
            lock (_lock)
            {
                _instance = this;
            }
        }

        public static void SetSolutionConfiguration(string configurationName, string platformName)
        {
            lock (_lock)
            {
                _properties = GetDictionary(configurationName, platformName);

                if (_instance != null)
                    _instance.PostSolutionConfiguration();
            }
        }

        private void PostSolutionConfiguration()
        {
            _version = (long)_version + 1;
            postSolutionConfiguration();
        }

        private void postSolutionConfiguration()
        {
            _brodcastBlock.Post(
                new ProjectVersionedValue<IImmutableDictionary<string, string>>(_properties,
                    ImmutableDictionary<NamedIdentity, IComparable>.Empty.Add(_dataSourceKey, _version)));
        }

        public override NamedIdentity DataSourceKey => _dataSourceKey;

        public override IComparable DataSourceVersion
        {
            get { lock (_lock) { return _version; } }
        }

        public override IReceivableSourceBlock<IProjectVersionedValue<IImmutableDictionary<string, string>>> SourceBlock
        {
            get
            {
                EnsureInitialized();
                return _publicBlock;
            }
        }

        public Task<IImmutableDictionary<string, string>> GetGlobalPropertiesAsync(CancellationToken cancellationToken)
        {
            lock (_lock)
            {
                return Task.FromResult(_properties);
            }
        }

        static IImmutableDictionary<string, string> GetDictionary(string configurationName, string platformName)
        {
            var builder = Microsoft.VisualStudio.ProjectSystem.Empty.PropertiesMap.ToBuilder();
            builder.Add("SolutionConfiguration", configurationName);
            builder.Add("SolutionPlatform", platformName);
            return builder.ToImmutable();
        }

        protected override void Initialize()
        {
            base.Initialize();

            lock (_lock)
            {
                _brodcastBlock = new BroadcastBlock<IProjectVersionedValue<IImmutableDictionary<string, string>>>(null,
                    new DataflowBlockOptions() { NameFormat = NAME_IDENTITY + ": {1}" });
                _publicBlock = _brodcastBlock.SafePublicize();

                postSolutionConfiguration();
            }
        }
    }
}
