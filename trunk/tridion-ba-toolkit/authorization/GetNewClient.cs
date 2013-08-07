using System;
using Tridion.ContentManager.CoreService.Client;
using System.ServiceModel;

// adapted from http://code.google.com/p/tridion-practice/wiki/GetCoreServiceClientWithoutConfigFile

namespace BA_Toolkit
{
    public class CoreServiceHandler : IDisposable
    {
        private ChannelFactory<ICoreService> _factory;


        public ICoreService GetNewClient(string hostname, string username, string password)
    {
        var binding = new BasicHttpBinding()
        {
            MaxBufferSize = 4194304, // 4MB
            MaxBufferPoolSize = 4194304,
            MaxReceivedMessageSize = 4194304,
            ReaderQuotas = new System.Xml.XmlDictionaryReaderQuotas()
            {
                MaxStringContentLength = 4194304, // 4MB
                MaxArrayLength = 4194304,
            },
            Security = new BasicHttpSecurity()
            {
                Mode = BasicHttpSecurityMode.TransportCredentialOnly,
                Transport = new HttpTransportSecurity()
                {
                    ClientCredentialType = HttpClientCredentialType.Windows,
                }
            }
        };
        hostname = string.Format("{0}{1}{2}", hostname.StartsWith("http") ? "" : "http://", hostname, hostname.EndsWith("/") ? "" : "/");
        // var endpoint = new EndpointAddress(hostname + "/webservices/CoreService.svc/basicHttp_2010");
        
        // updated for SDL Tridion 2011 SP1-HR1    
        var endpoint = new EndpointAddress(hostname + "/webservices/CoreService2011.svc/basicHttp");
        _factory = new ChannelFactory<ICoreService>(binding, endpoint);
        _factory.Credentials.Windows.ClientCredential = new System.Net.NetworkCredential(username, password);
        return _factory.CreateChannel();
    }


        public void Dispose()
        {
            _factory.Close();
        }
    }
}


