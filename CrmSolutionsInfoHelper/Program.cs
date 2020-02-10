using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Deployment;
using Microsoft.Xrm.Sdk.Deployment.Proxy;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Sdk.Query;

namespace CrmSolutionsInfoHelper
{
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                var crmUrl = ConfigurationManager.AppSettings["CrmUrl"];
                var url = new Uri(crmUrl + "/XRMServices/2011/Discovery.svc");
                var config = ServiceConfigurationFactory.CreateConfiguration<IDiscoveryService>(url);
                var credentials = new ClientCredentials();
                credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;
                var discoveryService = new DiscoveryServiceProxy(config, credentials);
                discoveryService.Authenticate();
                var request = new RetrieveOrganizationsRequest();
                var response = (RetrieveOrganizationsResponse)discoveryService.Execute(request);

                foreach (var detail in response.Details)
                {
                    var organizationServiceUrl = detail.Endpoints[EndpointType.OrganizationService];
                    var organizationName = detail.UniqueName;
                    GetOrganizationInfo(organizationName, organizationServiceUrl);
                }
                Console.WriteLine("Press any key to close the window.");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occured." + Environment.NewLine + ex.ToString());
            }
        }

        private static void GetOrganizationInfo(string organizationName, string organizationServiceUrl)
        {
            var crmUrl = ConfigurationManager.AppSettings["CrmUrl"];
            var uri = new Uri(crmUrl + "/XRMDeployment/2011/Deployment.svc");
            var deploymentService = ProxyClientHelper.CreateClient(uri);
            deploymentService.ClientCredentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;

            var orgRetRequest = new RetrieveRequest();
            orgRetRequest.EntityType = DeploymentEntityType.Organization;

            orgRetRequest.InstanceTag = new EntityInstanceId();
            orgRetRequest.InstanceTag.Name = organizationName;

            try
            {
                var response = (RetrieveResponse)deploymentService.Execute(orgRetRequest);
                Console.WriteLine("---------------" + organizationName + "---------------");
                Console.WriteLine("State: " + ((Organization)response.Entity).State.ToString());
                Console.WriteLine("Version: " + ((Organization)response.Entity).Version);
                GetSolutionsInfo(organizationServiceUrl);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to get organization solutions info." + Environment.NewLine + ex.ToString());
            }
        }

        private static void GetSolutionsInfo(string organizationServiceUrl)
        {
            var uri = new Uri(organizationServiceUrl);
            var clientCredentials = new ClientCredentials();
            clientCredentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;
            var service = new OrganizationServiceProxy(uri, null, clientCredentials, null);

            var query = new QueryExpression("solution");
            query.ColumnSet.AddColumns("uniquename", "version", "ismanaged", "installedon", "modifiedon");

            bool publisherExists = true;
            var publisherName = ConfigurationManager.AppSettings["PublisherName"];
            if (!string.IsNullOrWhiteSpace(publisherName))
            {
                try
                {
                    var queryPublisher = new QueryExpression("publisher");
                    queryPublisher.ColumnSet = new ColumnSet(true);
                    queryPublisher.Criteria.AddCondition("uniquename", ConditionOperator.Equal, publisherName);
                    var publishers = service.RetrieveMultiple(queryPublisher);
                    if (publishers.Entities.Count > 0)
                    {
                        query.Criteria.AddCondition("publisherid", ConditionOperator.Equal, publishers[0].Id);
                    }
                    else
                    {
                        publisherExists = false;
                        Console.WriteLine("No such publisher: " + publisherName);
                        Console.WriteLine();
                    }
                }
                catch
                {
                    Console.WriteLine();
                    Console.WriteLine("Unable to get publisher");
                }
            }

            if (publisherExists)
            {
                var errorWhileGettingSolutions = false;
                var solutions = new List<Entity>();
                try
                {
                    var result = service.RetrieveMultiple(query);
                    solutions = result.Entities.ToList();
                }
                catch
                {
                    errorWhileGettingSolutions = true;
                    Console.WriteLine("Unable to get solutions");
                    Console.WriteLine();
                }

                if (solutions.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine("\t--SOLUTIONS INFO--");
                    foreach (var solution in solutions)
                    {
                        var uniqueName = solution.GetAttributeValue<string>("uniquename");
                        var version = solution.GetAttributeValue<string>("version");
                        var isManaged = solution.GetAttributeValue<bool>("ismanaged").ToString();
                        var installedon = solution.GetAttributeValue<DateTime>("installedon");
                        var modifiedon = solution.GetAttributeValue<DateTime>("modifiedon");
                        if (uniqueName != "System" && uniqueName != "Active" && uniqueName != "Basic" && uniqueName != "ActivityFeeds")
                        {
                            Console.WriteLine("\tName: " + uniqueName);
                            Console.WriteLine("\tVersion: " + version);
                            Console.WriteLine("\tIs Managed: " + isManaged);
                            Console.WriteLine("\tInstall Date: " + installedon);
                            Console.WriteLine("\tModify Date: " + modifiedon);
                            Console.WriteLine();
                        }
                    }
                }
                else if(!errorWhileGettingSolutions)
                {
                    Console.WriteLine("No active solutions for this publisher");
                    Console.WriteLine();
                }
            }
        }
    }
}