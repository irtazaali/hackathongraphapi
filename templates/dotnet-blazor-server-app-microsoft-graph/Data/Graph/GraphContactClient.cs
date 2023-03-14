using Microsoft.Graph;
using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using System.Linq;
using System.Net;
using System.Net.Http;

namespace BlazorSample.Data.Graph
{
    public class GraphContactClient
    {
        private readonly ILogger<GraphContactClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;

        public GraphContactClient(ILogger<GraphContactClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<IEnumerable<Contact>> GetContacts()
        {
            try
            {
                var contacts = await _graphServiceClient.Me.Contacts
                            .Request()
                            .GetAsync();

                return contacts;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error calling Graph /me/contacts: {ex.Message}");
                throw;
            }
        }
    }
}
