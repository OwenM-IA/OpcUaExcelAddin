using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Opc.Ua;
using Opc.Ua.Client;
using Opc.Ua.Configuration;

namespace OpcUaExcelAddin
{
    public sealed class OpcUaService
    {
        private static readonly Lazy<OpcUaService> LazyInstance = new Lazy<OpcUaService>(() => new OpcUaService());
        private readonly SemaphoreSlim _connectGate = new SemaphoreSlim(1, 1);
        private Session _session;
        private CancellationTokenSource _shutdown;
        private int _connected;
        private int _started;

        private OpcUaService()
        {
        }

        public static OpcUaService Instance => LazyInstance.Value;

        public event Action<Session> Connected;
        public event Action Disconnected;

        public bool IsConnected => Interlocked.CompareExchange(ref _connected, 0, 0) == 1;

        public Session CurrentSession => Volatile.Read(ref _session);

        public void Start()
        {
            if (Interlocked.Exchange(ref _started, 1) != 0)
            {
                return;
            }

            _shutdown = new CancellationTokenSource();
            Task.Run(() => ReconnectLoopAsync(_shutdown.Token));
        }

        public async Task<bool> ConnectAsync()
        {
            await _connectGate.WaitAsync().ConfigureAwait(false);
            try
            {
                if (IsConnected && _session != null)
                {
                    return true;
                }

                CloseSession(_session);
                _session = null;
                Interlocked.Exchange(ref _connected, 0);

                string[] endpointUrls = ConfigLoader.EndpointUrls;
                if (endpointUrls.Length == 0)
                {
                    return false;
                }

                ApplicationConfiguration configuration = await CreateConfigurationAsync().ConfigureAwait(false);
                ITelemetryContext telemetry = configuration.CreateMessageContext(false).Telemetry;
                var application = new ApplicationInstance(telemetry)
                {
                    ApplicationName = "OpcUaExcelAddin",
                    ApplicationType = ApplicationType.Client,
                    ApplicationConfiguration = configuration
                };

                if (!await application.CheckApplicationInstanceCertificatesAsync(true).ConfigureAwait(false))
                {
                    throw new InvalidOperationException("The OPC UA client application certificate could not be created.");
                }

                var sessionFactory = new DefaultSessionFactory(telemetry);
                foreach (string endpointUrl in endpointUrls)
                {
                    try
                    {
                        EndpointDescription selectedEndpoint;
                        try
                        {
                            selectedEndpoint = await CoreClientUtils.SelectEndpointAsync(
                                configuration,
                                endpointUrl,
                                true,
                                telemetry,
                                CancellationToken.None).ConfigureAwait(false);
                        }
                        catch
                        {
                            selectedEndpoint = await CoreClientUtils.SelectEndpointAsync(
                                configuration,
                                endpointUrl,
                                false,
                                telemetry,
                                CancellationToken.None).ConfigureAwait(false);
                        }

                        var endpoint = new ConfiguredEndpoint(
                            null,
                            selectedEndpoint,
                            EndpointConfiguration.Create(configuration));
                        Session session = (Session)await sessionFactory.CreateAsync(
                            configuration,
                            endpoint,
                            false,
                            "OpcUaExcelAddin",
                            60000,
                            new UserIdentity(new AnonymousIdentityToken()),
                            null,
                            CancellationToken.None).ConfigureAwait(false);

                        session.KeepAlive += OnKeepAlive;
                        _session = session;
                        Interlocked.Exchange(ref _connected, 1);
                        NotifyConnected(session);
                        Debug.WriteLine("Connected to OPC UA endpoint " + endpointUrl);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("OPC UA connection to " + endpointUrl + " failed: " + ex.Message);
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("OPC UA connection failed: " + ex);
                Interlocked.Exchange(ref _connected, 0);
                return false;
            }
            finally
            {
                _connectGate.Release();
            }
        }

        public object ReadValue(string nodeId)
        {
            if (!ConfigLoader.TryResolveNodeId(nodeId, out string resolvedNodeId))
            {
                return "#BAD_NODEID";
            }

            _ = Task.Run(() => SubscriptionManager.Instance.EnsureSubscribedAsync(resolvedNodeId));
            return TagCache.Instance.GetValue(resolvedNodeId);
        }

        public void Disconnect()
        {
            _shutdown?.Cancel();
            _connectGate.Wait();
            try
            {
                Session current = _session;
                _session = null;
                Interlocked.Exchange(ref _connected, 0);
                CloseSession(current);
                NotifyDisconnected();
            }
            finally
            {
                _connectGate.Release();
            }
        }

        private static async Task<ApplicationConfiguration> CreateConfigurationAsync()
        {
            string pkiDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OpcUaExcelAddin",
                "pki");

            var configuration = new ApplicationConfiguration
            {
                ApplicationName = "OpcUaExcelAddin",
                ApplicationUri = "urn:" + Environment.MachineName + ":OpcUaExcelAddin",
                ApplicationType = ApplicationType.Client,
                SecurityConfiguration = new SecurityConfiguration
                {
                    ApplicationCertificate = new CertificateIdentifier
                    {
                        StoreType = CertificateStoreType.Directory,
                        StorePath = Path.Combine(pkiDirectory, "own"),
                        SubjectName = "CN=OpcUaExcelAddin"
                    },
                    TrustedPeerCertificates = new CertificateTrustList
                    {
                        StoreType = CertificateStoreType.Directory,
                        StorePath = Path.Combine(pkiDirectory, "trusted")
                    },
                    TrustedIssuerCertificates = new CertificateTrustList
                    {
                        StoreType = CertificateStoreType.Directory,
                        StorePath = Path.Combine(pkiDirectory, "issuer")
                    },
                    RejectedCertificateStore = new CertificateTrustList
                    {
                        StoreType = CertificateStoreType.Directory,
                        StorePath = Path.Combine(pkiDirectory, "rejected")
                    },
                    AutoAcceptUntrustedCertificates = true,
                    AddAppCertToTrustedStore = true
                },
                TransportConfigurations = new TransportConfigurationCollection(),
                TransportQuotas = new TransportQuotas { OperationTimeout = 15000 },
                ClientConfiguration = new ClientConfiguration { DefaultSessionTimeout = 60000 },
                TraceConfiguration = new TraceConfiguration()
            };

            await configuration.Validate(ApplicationType.Client).ConfigureAwait(false);
            configuration.CertificateValidator.CertificateValidation += (sender, args) => args.Accept = true;
            return configuration;
        }

        private async Task ReconnectLoopAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                if (!IsConnected)
                {
                    await ConnectAsync().ConfigureAwait(false);
                }

                try
                {
                    await Task.Delay(IsConnected ? 2000 : 5000, cancellationToken).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
            }
        }

        private void OnKeepAlive(ISession session, KeepAliveEventArgs args)
        {
            if (!ReferenceEquals(session, _session))
            {
                return;
            }

            if (ServiceResult.IsBad(args.Status) && Interlocked.Exchange(ref _connected, 0) == 1)
            {
                NotifyDisconnected();
            }
        }

        private void NotifyConnected(Session session)
        {
            Action<Session> handlers = Connected;
            if (handlers == null)
            {
                return;
            }

            foreach (Action<Session> handler in handlers.GetInvocationList())
            {
                try
                {
                    handler(session);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("OPC UA connected handler failed: " + ex.Message);
                }
            }
        }

        private void NotifyDisconnected()
        {
            Action handlers = Disconnected;
            if (handlers == null)
            {
                return;
            }

            foreach (Action handler in handlers.GetInvocationList())
            {
                try
                {
                    handler();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("OPC UA disconnected handler failed: " + ex.Message);
                }
            }
        }

        private static void CloseSession(Session session)
        {
            if (session == null)
            {
                return;
            }

            try
            {
                session.KeepAlive -= Instance.OnKeepAlive;
                session.Close();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("OPC UA session close failed: " + ex.Message);
            }
            finally
            {
                session.Dispose();
            }
        }
    }
}