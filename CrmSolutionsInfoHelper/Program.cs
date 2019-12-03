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
using Microsoft.Xrm.Sdk.Query;

namespace CrmSolutionsInfoHelper
{
    public class Program
    {
        private static List<string> nonexistentOrganizations = new List<string>();

        public static void Main(string[] args)
        {
            var organizationsString = ConfigurationManager.AppSettings["Organizations"];
            var organizations = new List<string>();
            if (string.IsNullOrWhiteSpace(organizationsString))
            {
                Console.Write("Enter Organization Name: ");
                var input = Console.ReadLine();
                organizations.Add(input);
                Console.WriteLine();
            }
            else
            {
                organizations = organizationsString.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            }

            foreach (var organization in organizations)
            {
                GetOrganizationInfo(organization);
            }

            if (nonexistentOrganizations.Count > 0)
            {
                Console.WriteLine();
                Console.WriteLine("------------------------------");
                Console.WriteLine("--NONEXISTENT ORGANIZATIONS--");
                Console.WriteLine(string.Join(",", nonexistentOrganizations));
            }
            Console.ReadKey();
        }

        private static void GetOrganizationInfo(string organizationName)
        {
            var crmUrl = ConfigurationManager.AppSettings["CrmUrl"];
            if (!string.IsNullOrWhiteSpace(crmUrl))
            {
                var uri = new Uri(crmUrl + "/XRMDeployment/2011/Deployment.svc");
                var deploymentService = ProxyClientHelper.CreateClient(uri);
                deploymentService.ClientCredentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;

                var orgRetRequest = new RetrieveRequest();
                orgRetRequest.EntityType = DeploymentEntityType.Organization;

                orgRetRequest.InstanceTag = new EntityInstanceId();
                orgRetRequest.InstanceTag.Name = organizationName;

                var organizationExists = true;
                try
                {
                    var response = (RetrieveResponse)deploymentService.Execute(orgRetRequest);
                    Console.WriteLine("---------------" + organizationName.ToUpper() + "---------------");
                    Console.WriteLine("State: " + ((Organization)response.Entity).State.ToString());
                    Console.WriteLine("Version: " + ((Organization)response.Entity).Version);
                }
                catch
                {
                    nonexistentOrganizations.Add(organizationName);
                    organizationExists = false;
                }

                if (organizationExists)
                {
                    var organizationServiceUrl = crmUrl + "/" + organizationName + "/XRMServices/2011/Organization.svc";
                    GetSolutionsInfo(organizationServiceUrl);
                }
            }
            else
            {
                Console.WriteLine("CRM url is empty");
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