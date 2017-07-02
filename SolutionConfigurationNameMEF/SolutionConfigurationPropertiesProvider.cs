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
    [AppliesTo("VisualC + VCProjectEngineFactory")]
    public class SolutionConfigurationPropertiesProvider :
        ProjectValueDataSourceBase<IImmutableDictionary<string, string>>,
        IProjectGlobalPropertiesProvider
    {
        private static object _lock;
        private static SolutionConfigurationPropertiesProvider _instance;   // Lock guarded
        private static IImmutableDictionary<string, string> _properties;    // Lock guarded
        private static long _version;                                       // Lock guarded

        private static NamedIdentity _dataSourceKey;

        /// <summary>The block to post to when publishing new values.</summary>
        private static BroadcastBlock<IProjectVersionedValue<IImmutableDictionary<string, string>>> _brodcastBlock;

        /// <summary>The backing field for the <see cref="SourceBlock"/> property.</summary>
        private IReceivableSourceBlock<IProjectVersionedValue<IImmutableDictionary<string, string>>> _publicBlock;

        static SolutionConfigurationPropertiesProvider()
        {
            _version = -1;
            _lock = new object();
            _properties = ImmutableDictionary<string, string>.Empty;
            _dataSourceKey = new NamedIdentity("SolutionConfiguration");
        }

        [ImportingConstructor]
        protected SolutionConfigurationPropertiesProvider(IProjectServices services)
            : base(services)
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
                _properties = ImmutableDictionary<string, string>.Empty.AddRange(new KeyValuePair<string, string>[] {
                        new KeyValuePair<string, string>("SolutionConfiguration", configurationName),
                        new KeyValuePair<string, string>("SolutionPlatform", platformName)});

                if (_instance != null)
                    _instance.PostSolutionConfiguration();
            }
        }

        private void PostSolutionConfiguration()
        {
            _version++;
            bool test = _brodcastBlock.Post(
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

        protected override void Initialize()
        {
            base.Initialize();
            lock (_lock)
            {
                if (_brodcastBlock == null)
                {
                    _brodcastBlock = new BroadcastBlock<IProjectVersionedValue<IImmutableDictionary<string, string>>>(null,
                    new DataflowBlockOptions() { NameFormat = "SolutionConfiguration: {1}" });

                    _publicBlock = _brodcastBlock.SafePublicize();
                    PostSolutionConfiguration();
                }
                else
                {
                    _publicBlock = _brodcastBlock.SafePublicize();
                }
            }
        }
    }
}
